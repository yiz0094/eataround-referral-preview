/**
 * 이더라운드 레퍼럴 프로모션 — Apps Script 백엔드 (Google Sheets = DB)
 * ------------------------------------------------------------------
 * 단일 페이지 바이럴 루프:
 *   - 누구나 닉네임+전화 제출 → 고유 ref_code(개인 링크) 발급
 *   - ?ref=CODE 로 들어와 제출하면 CODE 소유자의 점수로 집계(쿠키/전화 중복 제외)
 *   - 리더보드 = referred_by 별 "유효 제출수"
 *
 * 설치/배포: backend/SETUP.md 참고
 * ------------------------------------------------------------------
 */

// ===== 설정 =====
var SHEET_ID  = '1ZNoyGTrYPvl_Y9BBlP6U5RtsNcguu5oLWvTXvgAMTMk';   // 대상 시트(비공개 유지 OK, 웹앱이 '실행:나'로 접근)
var ENTRIES   = 'entries';
var CLICKS    = 'clicks';
var BASE_LINK = 'https://eataround-referral-preview.vercel.app/'; // 발급 링크 베이스 (추후 도메인 바뀌면 여기 수정 후 재배포)
var TOP_N     = 50;   // 리더보드 노출 인원
var CACHE_SEC = 7;    // 리더보드 캐시 (폴링 부하/쿼터 완화)

// ===== 알림톡 변수용 고정값 =====
// 제출 1건 = 시트 1행에 아래 값들이 함께 기록됨 → 알림톡 자동발송 툴이 행을 그대로 변수에 매핑.
// (닉네임=nickname 열, 전화번호=phone 열, 추천링크=아래 링크 열을 사용)
var PROMO_NAME     = '2026 마을 여행 기획전';             // #{프로모션명}
var PROMO_PERIOD   = '2026.08.05(수) ~ 2026.08.18(화)';   // #{기간}
var PROMO_ANNOUNCE = '2026.08.21(금)';                    // #{발표일}
var FOODLIST_LINK  = 'https://jeju-matjip-map.vercel.app/'; // #{맛집리스트링크}

// entries 시트 헤더(순서 고정). 9번째까지는 기존 컬럼 — 인덱스 참조 로직이 의존하므로 변경 금지.
var HEADERS = ['entry_id','ref_code','nickname','phone','cookie_id','referred_by','user_agent','created_at','valid_for_rank',
               '추천링크','프로모션명','기간','발표일','맛집리스트링크'];

// ===== 시트 헬퍼 =====
function ss_(){ return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name){ var s = ss_().getSheetByName(name); if(!s){ s = ss_().insertSheet(name); } return s; }

function ensureHeaders_(){
  var e = sheet_(ENTRIES);
  if(e.getLastRow() === 0){
    e.appendRow(HEADERS);
  }
  var c = sheet_(CLICKS);
  if(c.getLastRow() === 0){
    c.appendRow(['ref_code','cookie_id','user_agent','created_at']);
  }
}

/** 기존 시트(구 9컬럼)에 알림톡 변수 헤더를 보강. doPost에서 이미 읽은 헤더행을 넘겨 추가 호출 없이 처리. */
function ensureAlimtalkHeaders_(sh, headerRow){
  var n = (headerRow || []).length;
  if(n > 0 && n < HEADERS.length){
    sh.getRange(1, n + 1, 1, HEADERS.length - n).setValues([HEADERS.slice(n)]);
  }
}

/** 시트 메뉴에서 한 번 실행하면 탭/헤더가 즉시 생성됩니다(선택). */
function setup(){ ensureHeaders_(); }

// ===== 유틸 =====
function normPhone_(p){
  var d = (p || '').toString().replace(/[^0-9]/g, '');
  if(d.length === 11 && d.indexOf('010') === 0) return d.slice(0,3) + '-' + d.slice(3,7) + '-' + d.slice(7);
  return null; // 형식 오류
}
function phoneTail_(p){ var d = ('' + (p || '')).replace(/[^0-9]/g, ''); return d ? d.slice(-4) : ''; }

function genCode_(existingSet){
  var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 헷갈리는 0/O/1/l/I 제외
  for(var t=0; t<50; t++){
    var s = '';
    for(var i=0;i<6;i++){ s += chars.charAt(Math.floor(Math.random()*chars.length)); }
    if(!existingSet[s]) return s;
  }
  return 'r' + String(Date.now()).slice(-7);
}

function isTrue_(v){ return v === true || v === 'TRUE' || v === 'true'; }

// ===== 응답 (JSON / JSONP 둘 다 지원) =====
function respond_(e, obj){
  var cb = e && e.parameter && e.parameter.callback;
  if(cb){
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== GET: 리더보드 / 내 통계 / 클릭 로깅 =====
function doGet(e){
  ensureHeaders_();
  var p = (e && e.parameter) || {};
  var action = p.action || 'leaderboard';

  if(action === 'leaderboard'){
    var b = getBoardCached_();
    return respond_(e, { ok:true, leaderboard: b.top, total: b.total, activity: b.activity });
  }
  if(action === 'me'){
    var full = getLeaderboardFull_();
    var ref = (p.ref || '').toString();
    var row = full.byCode[ref];
    return respond_(e, { ok:true, ref:ref, count: row?row.count:0, rank: row?row.rank:null, total: full.total });
  }
  if(action === 'click'){
    logClick_((p.ref||'').toString(), (p.cookie||'').toString(), (p.ua||'').toString());
    return respond_(e, { ok:true });
  }
  return respond_(e, { ok:false, error:'unknown action' });
}

// ===== POST: 제출 =====
function doPost(e){
  ensureHeaders_();
  var body = {};
  try { body = JSON.parse(e.postData.contents); }
  catch(err){ return respond_(e, { ok:false, error:'bad json' }); }

  if((body.action || 'submit') !== 'submit') return respond_(e, { ok:false, error:'unknown action' });

  var nickname   = (body.nickname || '').toString().trim().slice(0,20);
  var phone      = normPhone_(body.phone);
  var cookie     = (body.cookie || '').toString().slice(0,40);
  var referredBy = (body.ref || '').toString().slice(0,12);
  var ua         = (body.ua || '').toString().slice(0,180);

  if(!nickname) return respond_(e, { ok:false, error:'닉네임을 입력해주세요' });
  if(!phone)    return respond_(e, { ok:false, error:'전화번호 형식이 올바르지 않습니다 (010-0000-0000)' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(err){ return respond_(e, { ok:false, error:'잠시 후 다시 시도해주세요' }); }

  try {
    var sh = sheet_(ENTRIES);
    var data = sh.getDataRange().getValues(); // [0]=헤더
    var codeSet = {}, cookieSeen = {}, phoneRow = -1;
    for(var i=1;i<data.length;i++){
      var rcode = data[i][1], rphone = data[i][3], rcookie = data[i][4];
      if(rcode)   codeSet[rcode] = true;
      if(rcookie) cookieSeen[rcookie] = true;
      if(rphone === phone) phoneRow = i;
    }

    // 이미 가입한 전화 → 기존 고유 링크 그대로 반환(멱등)
    if(phoneRow >= 0){
      var myCode = data[phoneRow][1];
      var lbR = buildLeaderboard_(data); var meR = lbR.byCode[myCode] || { count:0, rank:null };
      lock.releaseLock();
      return respond_(e, { ok:true, returning:true, ref_code:myCode, link: BASE_LINK + '?ref=' + myCode, count: meR.count, rank: meR.rank });
    }

    // 신규 → 고유 코드 발급 + 행 추가
    var newCode = genCode_(codeSet);
    var newLink = BASE_LINK + '?ref=' + newCode;
    // 유효 집계 조건: 추천인이 있고, (쿠키 없음 OR 처음 보는 쿠키)  ← 같은 기기 다중 번호 패딩 차단
    var validForRank = (!!referredBy) && (!cookie || !cookieSeen[cookie]);
    ensureAlimtalkHeaders_(sh, data[0]); // 구 9컬럼 시트면 알림톡 변수 헤더 보강
    // 뒤 5개 = 알림톡 변수(추천링크/프로모션명/기간/발표일/맛집리스트링크)
    var newRow = [ Utilities.getUuid(), newCode, nickname, phone, cookie, referredBy || '', ua, new Date(), validForRank,
                   newLink, PROMO_NAME, PROMO_PERIOD, PROMO_ANNOUNCE, FOODLIST_LINK ];
    sh.appendRow(newRow);
    CacheService.getScriptCache().remove('lb'); // 리더보드 캐시 무효화
    data.push(newRow); // 방금 추가분 포함해 내 순위 계산
    var lbN = buildLeaderboard_(data); var meN = lbN.byCode[newCode] || { count:0, rank:null };
    lock.releaseLock();

    return respond_(e, { ok:true, returning:false, ref_code:newCode, link: newLink, count: meN.count, rank: meN.rank, credited: validForRank });
  } catch(err){
    try { lock.releaseLock(); } catch(e2){}
    return respond_(e, { ok:false, error: String(err) });
  }
}

// ===== 집계 (모든 참여자 포함 — 추천 0명도 순위에 노출) =====
// data = entries 전체(헤더 포함). 각 참여자 행을 count 0으로 등록 후, 유효 추천을 referred_by에 가산.
function buildLeaderboard_(data){
  var rec = {}, order = [], total = 0;
  for(var i=1;i<data.length;i++){
    var code=data[i][1], nick=data[i][2], phone=data[i][3], rb=data[i][5], created=data[i][7], v=data[i][8];
    if(!code) continue;
    total++;
    if(!rec[code]){ rec[code] = { code:code, nickname:nick, phone:phone, count:0, created:created }; order.push(code); }
    else { rec[code].nickname = nick; rec[code].phone = phone; }
    if(rb && isTrue_(v)){
      if(!rec[rb]){ rec[rb] = { code:rb, nickname:'', phone:'', count:0, created:'' }; order.push(rb); }
      rec[rb].count++;
    }
  }
  var arr = order.map(function(k){ return rec[k]; });
  arr.sort(function(a,b){ if(b.count !== a.count) return b.count - a.count; return (a.created < b.created ? -1 : (a.created > b.created ? 1 : 0)); });
  var byCode = {};
  for(var j=0;j<arr.length;j++){ arr[j].rank = j+1; byCode[arr[j].code] = arr[j]; }
  return { arr:arr, byCode:byCode, total:total };
}

function getLeaderboardFull_(){ return buildLeaderboard_(sheet_(ENTRIES).getDataRange().getValues()); }

function getBoardCached_(){
  var cache = CacheService.getScriptCache();
  var hit = cache.get('lb');
  if(hit) return JSON.parse(hit);
  var data = sheet_(ENTRIES).getDataRange().getValues();
  var full = buildLeaderboard_(data);
  var top = full.arr.slice(0, TOP_N).map(function(o){ return { rank:o.rank, nickname:o.nickname, count:o.count, ref:o.code, tail: phoneTail_(o.phone) }; });
  var out = { top:top, total:full.total, activity: buildActivity_(data, full) };
  cache.put('lb', JSON.stringify(out), CACHE_SEC);
  return out;
}

// ===== 최근 1시간 실시간 활동(실데이터) → 표시용 메시지(최신순) + 실데이터 요약 =====
function buildActivity_(data, full){
  var now = Date.now(), WINDOW = 60*60*1000, byCode = full.byCode, ev = [];
  for(var i=1;i<data.length;i++){
    var nick=data[i][2], rb=data[i][5], created=data[i][7], v=data[i][8];
    if(!nick) continue;
    var t = (created instanceof Date) ? created.getTime() : (created ? new Date(created).getTime() : 0);
    if(!t || (now - t) > WINDOW) continue;        // 최근 60분만
    ev.push({ nick:nick, rb:rb, valid:isTrue_(v), t:t });
  }
  ev.sort(function(a,b){ return b.t - a.t; });     // 최신순
  ev = ev.slice(0, 20);
  var msgs = [], usedRef = {};
  for(var k=0;k<ev.length;k++){
    var x = ev[k], mins = Math.floor((now - x.t)/60000), when = (mins<1 ? '방금' : (mins+'분 전'));
    var r = (x.rb && x.valid) ? byCode[x.rb] : null;
    if(r && r.count>0 && !usedRef[x.rb]){          // 추천인 마일스톤(추천인별 1회)
      usedRef[x.rb] = true;
      msgs.push('🔥 <b>'+esc_(r.nickname)+'</b>님이 친구 '+r.count+'명을 모았어요 · '+when);
    } else {                                       // 신규 참여
      msgs.push('🎉 <b>'+esc_(x.nick)+'</b>님이 참여했어요 · '+when);
    }
  }
  // 실데이터 요약(최근 활동이 적어도 티커가 비지 않도록)
  if(full.total>0) msgs.push('🙌 지금까지 <b>'+full.total+'</b>명이 참여했어요');
  if(full.arr.length && full.arr[0].count>0) msgs.push('🏆 <b>'+esc_(full.arr[0].nickname)+'</b>님이 친구 '+full.arr[0].count+'명으로 1위!');
  return msgs;
}
function esc_(s){ return ('' + (s||'')).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function getTotalCount_(){ return Math.max(0, sheet_(ENTRIES).getLastRow() - 1); }

function logClick_(ref, cookie, ua){
  if(!ref) return;
  sheet_(CLICKS).appendRow([ ref, cookie, ua, new Date() ]);
}
