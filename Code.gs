/*************************************************
 * 문파 빨시 장부 v2 — 백엔드
 * 구글시트에 붙여넣고 [배포 > 새 배포 > 웹 앱] 하세요.
 *  - 실행 사용자: 나
 *  - 액세스 권한: 모든 사용자
 *************************************************/

var S_OUT  = '불출기록';
var S_IN   = '입고기록';
var S_DAY  = '일별설정';
var S_SALE = '판매신청';
var S_FIX  = '정정요청';
var S_CHAR = '캐릭명단';
var S_CFG  = '설정';
var S_FUND = '자금기록';

var HEAD = {};
HEAD[S_OUT]  = ['기록ID','일자','시각','캐릭명','구분','실사용자','유형','수량','경로','비고','상태'];
HEAD[S_IN]   = ['기록ID','일자','시각','판매자','수량','개당단가','총액','등록자','비고'];
HEAD[S_DAY]  = ['일자','총불출량','마감여부','마감시각','메모'];
HEAD[S_SALE] = ['신청ID','일시','캐릭명','수량','희망단가','비고','상태','처리시각'];
HEAD[S_FIX]  = ['요청ID','일시','캐릭명','대상기록ID','내용','상태','처리시각'];
HEAD[S_CHAR] = ['캐릭명','구분','최초등록일'];
HEAD[S_CFG]  = ['키','값'];
HEAD[S_FUND] = ['기록ID','일자','시각','유형','출처','금액','메모','등록자'];

var CFG_DEFAULT = [
  ['자금초기잔고','0'],
  ['미반납경고일수','2']
];

/* 비밀번호는 시트가 아니라 스크립트 속성에 저장됩니다.
   시트를 모두에게 공개해도 비밀번호는 보이지 않습니다.
   바꾸려면 아래 함수의 값을 고치고 한 번 실행하세요. */
function 비밀번호바꾸기(){
  var NEW = '1234';   // <-- 여기를 원하는 비밀번호로 바꾸고 실행
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', String(NEW));
  SpreadsheetApp.getUi().alert('운영진 비밀번호를 바꿨습니다.');
}
function adminPw(){
  var v = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!v) { v = '1234'; PropertiesService.getScriptProperties().setProperty('ADMIN_PW', v); }
  return String(v);
}

/* ---------- 최초 1회 실행 ---------- */
function 시트만들기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEAD).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,HEAD[name].length).setValues([HEAD[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,HEAD[name].length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    }
  });
  var cfgSh = ss.getSheetByName(S_CFG);
  if (cfgSh.getLastRow() < 2) cfgSh.getRange(2,1,CFG_DEFAULT.length,2).setValues(CFG_DEFAULT);
  // 시트에 비밀번호가 남아 있으면 스크립트 속성으로 옮기고 시트에서는 지웁니다
  if (cfgSh.getLastRow() > 1) {
    var cv = cfgSh.getRange(2,1,cfgSh.getLastRow()-1,2).getValues();
    for (var i = cv.length - 1; i >= 0; i--) {
      if (String(cv[i][0]).trim() === '운영진비밀번호') {
        PropertiesService.getScriptProperties().setProperty('ADMIN_PW', String(cv[i][1]).trim() || '1234');
        cfgSh.deleteRow(i + 2);
      }
    }
  }
  adminPw();
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('시트 준비 완료.\n\n· 자금 초기잔고는 [설정] 탭에서 바꾸세요.\n· 운영진 비밀번호는 [비밀번호바꾸기] 함수를 실행해 바꾸세요.\n  (시트에는 저장되지 않아 공개해도 안전합니다)');
}

/* ---------- 공용 ---------- */
function sh(name){
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!s) throw new Error('시트가 없습니다: ' + name + ' — 먼저 [시트만들기]를 실행하세요.');
  return s;
}
var _cache = {};
function rows(name){
  if (_cache[name]) return _cache[name];
  var s = sh(name);
  if (s.getLastRow() < 2) return [];
  var vals = s.getRange(2,1,s.getLastRow()-1,HEAD[name].length).getValues();
  var out2 = vals.map(function(r){
    var o = {};
    HEAD[name].forEach(function(h,i){ o[h] = r[i]; });
    return o;
  }).filter(function(o){ return String(o[HEAD[name][0]]).length > 0; });
  _cache[name] = out2;
  return out2;
}
function push(name, obj){
  var s = sh(name);
  s.appendRow(HEAD[name].map(function(h){ return obj[h] !== undefined ? obj[h] : ''; }));
  touch();
}
function bust(){ _cache = {}; _ledger = null; }
function touch(){
  bust();
  try { PropertiesService.getScriptProperties().setProperty('REV', String(Date.now())); } catch(e){}
}
function rev(){
  var v = PropertiesService.getScriptProperties().getProperty('REV');
  return v ? Number(v) : 0;
}
function cfg(){
  var o = {};
  rows(S_CFG).forEach(function(r){ o[String(r['키']).trim()] = String(r['값']).trim(); });
  CFG_DEFAULT.forEach(function(d){ if(o[d[0]]===undefined) o[d[0]] = d[1]; });
  return o;
}
function today(){ return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }
function now(){ return Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm'); }
function stamp(){ return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }
function uid(p){ return p + '-' + Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMddHHmmss') + '-' + Math.floor(Math.random()*900+100); }
function d(v){ return v instanceof Date ? Utilities.formatDate(v,'Asia/Seoul','yyyy-MM-dd') : String(v||'').slice(0,10); }
function n(v){ var x = Number(v); return isNaN(x) ? 0 : x; }
function daysBetween(a,b){ return Math.round((new Date(b+'T00:00:00+09:00') - new Date(a+'T00:00:00+09:00'))/86400000); }

/* ---------- 진입점 ---------- */
function doGet(e){ return handle(e); }
function doPost(e){ return handle(e); }

function handle(e){
  var p = (e && e.parameter) || {};
  var out;
  try {
    var payload = p.payload ? JSON.parse(p.payload) : {};
    out = { ok:true, data: route(p.action || 'bootstrap', payload) };
  } catch(err){
    out = { ok:false, error: String((err && err.message) || err) };
  }
  var json = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function route(action, x){
  switch(action){
    case 'rev':       return { rev: rev() };
    case 'bootstrap': return bootstrap();
    case 'submit':    return submit(x);
    case 'me':        return me(x.캐릭명);
    case 'board':     return board();
    case 'sale':      return sale(x);
    case 'fix':       return fix(x);
    case 'admin':     return admin(x);
    default: throw new Error('알 수 없는 요청: ' + action);
  }
}

/* ---------- 계산 ---------- */
function dayRow(date){
  var r = rows(S_DAY).filter(function(v){ return d(v['일자']) === date; })[0];
  return r ? { 일자:date, 총불출량:n(r['총불출량']), 마감:String(r['마감여부'])==='true'||r['마감여부']===true, 마감시각:String(r['마감시각']||'') }
           : { 일자:date, 총불출량:0, 마감:false, 마감시각:'' };
}

var _ledger = null;
function ledger(){
  if (_ledger) return _ledger;
  var out = rows(S_OUT).filter(function(r){ return String(r['상태']) !== '취소'; });
  out.sort(function(a,b){ return (d(a['일자'])+String(a['시각'])).localeCompare(d(b['일자'])+String(b['시각'])); });
  var per = {};
  out.forEach(function(r){
    var name = String(r['캐릭명']).trim();
    if(!name) return;
    if(!per[name]) per[name] = { 캐릭명:name, 받음:0, 반납:0, 보유:0, 시작일:'', 마지막:'', 구분:String(r['구분']||'') };
    var p = per[name], q = n(r['수량']);
    if (String(r['유형']) === '반납') { p.반납 += q; p.보유 -= q; }
    else { p.받음 += q; if (p.보유 <= 0) p.시작일 = d(r['일자']); p.보유 += q; }
    if (p.보유 <= 0) { p.보유 = Math.max(0, p.보유); p.시작일 = ''; }
    p.마지막 = d(r['일자']);
    p.구분 = String(r['구분']||p.구분);
  });
  var t = today();
  Object.keys(per).forEach(function(k){
    var p = per[k];
    p.소진율 = p.받음 > 0 ? Math.round((p.받음 - p.보유) / p.받음 * 100) : 0;
    p.미반납일수 = (p.보유 > 0 && p.시작일) ? daysBetween(p.시작일, t) : 0;
  });
  _ledger = { records: out, per: per };
  return _ledger;
}

function stock(){
  var L = ledger();
  var inn = rows(S_IN);
  var 입고 = 0, 입고액 = 0;
  inn.forEach(function(r){ 입고 += n(r['수량']); 입고액 += n(r['총액']); });
  var 불출 = 0, 반납 = 0;
  L.records.forEach(function(r){
    if (String(r['유형']) === '반납') 반납 += n(r['수량']);
    else if (String(r['경로']) === '문파지급') 불출 += n(r['수량']);
  });
  var c = cfg();
  var 수입 = 0, 지출 = 0, 성별 = {};
  rows(S_FUND).forEach(function(r){
    var amt = n(r['금액']);
    if (String(r['유형']) === '지출') { 지출 += amt; }
    else {
      수입 += amt;
      var src = String(r['출처'] || '기타').trim() || '기타';
      성별[src] = (성별[src] || 0) + amt;
    }
  });
  return {
    재고: 입고 - 불출 + 반납,
    입고합: 입고, 불출합: 불출, 반납합: 반납,
    자금잔고: n(c['자금초기잔고']) + 수입 - 지출 - 입고액,
    입고총액: 입고액, 자금수입: 수입, 자금지출: 지출, 성별수입: 성별
  };
}

/* ---------- 문원 기능 ---------- */
function bootstrap(){
  var c = cfg();
  return {
    rev: rev(),
    today: today(),
    chars: rows(S_CHAR).map(function(r){ return String(r['캐릭명']).trim(); }).filter(String),
    day: dayRow(today()),
    경고일수: n(c['미반납경고일수']) || 2
  };
}

function submit(x){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var name = String(x.캐릭명 || '').trim();
    if (!name) throw new Error('캐릭명을 입력하세요.');
    var qty = n(x.수량);
    if (qty <= 0) throw new Error('수량은 1 이상이어야 합니다.');
    var type = x.유형 === '반납' ? '반납' : '받음';
    var date = today();
    var day = dayRow(date);
    if (day.마감 && type === '받음') throw new Error('오늘은 마감되었습니다. 받음 기록은 운영진에게 문의하세요.');

    if (type === '받음' && String(x.구분) === '용병' && !String(x.실사용자||'').trim())
      throw new Error('용병은 실사용자 이름을 적어주세요.');

    push(S_OUT, {
      기록ID: uid('O'), 일자: date, 시각: now(), 캐릭명: name,
      구분: String(x.구분 || '일반'), 실사용자: String(x.실사용자 || '').trim(),
      유형: type, 수량: qty,
      경로: type === '받음' ? String(x.경로 || '문파지급') : '',
      비고: String(x.비고 || '').trim(), 상태: '정상'
    });

    var known = rows(S_CHAR).map(function(r){ return String(r['캐릭명']).trim(); });
    if (known.indexOf(name) === -1)
      push(S_CHAR, { 캐릭명: name, 구분: String(x.구분 || '일반'), 최초등록일: date });

    return me(name);
  } finally { lock.releaseLock(); }
}

function me(name){
  name = String(name || '').trim();
  var L = ledger();
  var p = L.per[name] || { 캐릭명:name, 받음:0, 반납:0, 보유:0, 소진율:0, 미반납일수:0, 시작일:'' };
  var t = today();
  var 오늘받음 = 0, 오늘반납 = 0;
  var recent = L.records.filter(function(r){ return String(r['캐릭명']).trim() === name; }).slice(-12).reverse()
    .map(function(r){
      return { 기록ID:String(r['기록ID']), 일자:d(r['일자']), 시각:String(r['시각']),
               유형:String(r['유형']), 수량:n(r['수량']), 경로:String(r['경로']||''),
               구분:String(r['구분']||''), 실사용자:String(r['실사용자']||''), 비고:String(r['비고']||'') };
    });
  L.records.forEach(function(r){
    if (String(r['캐릭명']).trim() !== name || d(r['일자']) !== t) return;
    if (String(r['유형']) === '반납') 오늘반납 += n(r['수량']); else 오늘받음 += n(r['수량']);
  });
  var sales = rows(S_SALE).filter(function(r){ return String(r['캐릭명']).trim() === name; }).slice(-5).reverse()
    .map(function(r){ return { 일시:String(r['일시']), 수량:n(r['수량']), 희망단가:n(r['희망단가']), 상태:String(r['상태']) }; });

  var 판매수량 = 0, 판매금액 = 0;
  var 판매이력 = rows(S_IN).filter(function(r){ return String(r['판매자']).trim() === name; })
    .map(function(r){
      판매수량 += n(r['수량']); 판매금액 += n(r['총액']);
      return { 일자:d(r['일자']), 수량:n(r['수량']), 개당단가:n(r['개당단가']), 총액:n(r['총액']), 비고:String(r['비고']||'') };
    }).slice(-10).reverse();

  return {
    캐릭명: name, 보유: p.보유, 받음: p.받음, 반납: p.반납,
    소진율: p.소진율, 미반납일수: p.미반납일수, 미정산시작일: p.시작일 || '',
    오늘받음: 오늘받음, 오늘반납: 오늘반납, 최근: recent, 판매신청: sales,
    판매수량: 판매수량, 판매금액: 판매금액, 판매이력: 판매이력
  };
}

function board(){
  var t = today();
  var day = dayRow(t);
  var L = ledger();
  var 오늘기입 = 0;
  L.records.forEach(function(r){
    if (d(r['일자']) !== t) return;
    if (String(r['유형']) === '받음' && String(r['경로']) === '문파지급') 오늘기입 += n(r['수량']);
  });
  var c = cfg();
  var warn = n(c['미반납경고일수']) || 2;
  var list = Object.keys(L.per).map(function(k){ return L.per[k]; })
    .filter(function(p){ return p.보유 > 0; })
    .sort(function(a,b){ return b.보유 - a.보유; })
    .map(function(p){
      return { 캐릭명:p.캐릭명, 보유:p.보유, 받음:p.받음, 반납:p.반납,
               소진율:p.소진율, 미반납일수:p.미반납일수, 경고: p.미반납일수 >= warn };
    });
  var s = stock();
  return {
    rev: rev(),
    today: t, day: day, 오늘기입: 오늘기입, 차이: day.총불출량 - 오늘기입,
    재고: s.재고, 자금잔고: s.자금잔고, 미반납: list, 경고일수: warn,
    자금수입: s.자금수입, 자금지출: s.자금지출, 빨시매입: s.입고총액, 성별수입: s.성별수입
  };
}

function sale(x){
  var name = String(x.캐릭명 || '').trim();
  if (!name) throw new Error('캐릭명을 입력하세요.');
  if (n(x.수량) <= 0) throw new Error('수량을 입력하세요.');
  push(S_SALE, {
    신청ID: uid('S'), 일시: stamp(), 캐릭명: name, 수량: n(x.수량),
    희망단가: n(x.희망단가), 비고: String(x.비고 || '').trim(), 상태: '대기', 처리시각: ''
  });
  return { ok: true };
}

function fix(x){
  var name = String(x.캐릭명 || '').trim();
  if (!name) throw new Error('캐릭명을 입력하세요.');
  if (!String(x.내용 || '').trim()) throw new Error('무엇이 잘못됐는지 적어주세요.');
  push(S_FIX, {
    요청ID: uid('F'), 일시: stamp(), 캐릭명: name,
    대상기록ID: String(x.대상기록ID || ''), 내용: String(x.내용).trim(), 상태: '대기', 처리시각: ''
  });
  return { ok: true };
}

/* ---------- 운영진 ---------- */
function admin(x){
  if (String(x.pw || '') !== adminPw()) throw new Error('비밀번호가 다릅니다.');
  switch(x.sub){
    case 'auth':     return { ok: true };
    case 'dash':     return adminDash();
    case 'setTotal': return setTotal(x);
    case 'stockIn':  return stockIn(x);
    case 'saleAct':  return saleAct(x);
    case 'fixAct':   return fixAct(x);
    case 'close':    return closeDay(x);
    case 'cancel':   return cancelRecord(x);
    case 'setPw':    return setPw(x);
    case 'fundAdd':  return fundAdd(x);
    case 'fundDel':  return fundDel(x);
    case 'setCfg':   return setCfg(x);
    default: throw new Error('알 수 없는 운영 요청');
  }
}

function adminDash(){
  var b = board();
  var s = stock();
  var L = ledger();
  var t = today();
  var 오늘기록 = L.records.filter(function(r){ return d(r['일자']) === t; }).reverse()
    .map(function(r){
      return { 기록ID:String(r['기록ID']), 시각:String(r['시각']), 캐릭명:String(r['캐릭명']),
               구분:String(r['구분']||''), 실사용자:String(r['실사용자']||''), 유형:String(r['유형']),
               수량:n(r['수량']), 경로:String(r['경로']||''), 비고:String(r['비고']||'') };
    });
  var sales = rows(S_SALE).filter(function(r){ return String(r['상태']) === '대기'; })
    .map(function(r){ return { 신청ID:String(r['신청ID']), 일시:String(r['일시']), 캐릭명:String(r['캐릭명']),
                               수량:n(r['수량']), 희망단가:n(r['희망단가']), 비고:String(r['비고']||'') }; });
  var fixes = rows(S_FIX).filter(function(r){ return String(r['상태']) === '대기'; })
    .map(function(r){ return { 요청ID:String(r['요청ID']), 일시:String(r['일시']), 캐릭명:String(r['캐릭명']),
                               대상기록ID:String(r['대상기록ID']||''), 내용:String(r['내용']) }; });
  var all = Object.keys(L.per).map(function(k){ return L.per[k]; })
    .sort(function(a,b){ return b.보유 - a.보유 || b.받음 - a.받음; });
  var 최근입고 = rows(S_IN).slice(-8).reverse().map(function(r){
    return { 일자:d(r['일자']), 판매자:String(r['판매자']), 수량:n(r['수량']),
             개당단가:n(r['개당단가']), 총액:n(r['총액']) };
  });
  return {
    rev: rev(),
    today: t, day: b.day, 오늘기입: b.오늘기입, 차이: b.차이,
    재고: s.재고, 자금잔고: s.자금잔고, 입고합: s.입고합, 불출합: s.불출합, 반납합: s.반납합,
    자금수입: s.자금수입, 자금지출: s.자금지출, 빨시매입: s.입고총액, 성별수입: s.성별수입,
    초기잔고: n(cfg()['자금초기잔고']), 경고일수: n(cfg()['미반납경고일수']) || 2,
    자금기록: rows(S_FUND).slice(-15).reverse().map(function(r){
      return { 기록ID:String(r['기록ID']), 일자:d(r['일자']), 유형:String(r['유형']),
               출처:String(r['출처']||''), 금액:n(r['금액']), 메모:String(r['메모']||'') };
    }),
    미반납: b.미반납, 전체: all, 오늘기록: 오늘기록, 판매대기: sales, 정정대기: fixes, 최근입고: 최근입고
  };
}

function setTotal(x){
  var date = String(x.일자 || today());
  var s = sh(S_DAY);
  var data = rows(S_DAY);
  var idx = -1;
  data.forEach(function(r,i){ if (d(r['일자']) === date) idx = i + 2; });
  if (idx === -1) push(S_DAY, { 일자:date, 총불출량:n(x.총불출량), 마감여부:false, 마감시각:'', 메모:String(x.메모||'') });
  else {
    s.getRange(idx, 2).setValue(n(x.총불출량)); touch();
    if (x.메모 !== undefined) s.getRange(idx, 5).setValue(String(x.메모||''));
  }
  return adminDash();
}

function stockIn(x){
  if (!String(x.판매자||'').trim()) throw new Error('판매자 캐릭명을 적어주세요.');
  if (n(x.수량) <= 0) throw new Error('수량을 입력하세요.');
  var qty = n(x.수량), unit = n(x.개당단가);
  push(S_IN, {
    기록ID: uid('I'), 일자: today(), 시각: now(), 판매자: String(x.판매자).trim(),
    수량: qty, 개당단가: unit, 총액: qty * unit,
    등록자: String(x.등록자 || '운영진'), 비고: String(x.비고 || '').trim()
  });
  if (x.신청ID) markSale(String(x.신청ID), '승인');
  return adminDash();
}

function markSale(id, status){
  var s = sh(S_SALE);
  var data = rows(S_SALE);
  data.forEach(function(r,i){
    if (String(r['신청ID']) === id) {
      s.getRange(i+2, 7).setValue(status);
      s.getRange(i+2, 8).setValue(stamp());
      touch();
    }
  });
}
function saleAct(x){
  var id = String(x.신청ID);
  if (x.결정 === '승인') {
    var req = rows(S_SALE).filter(function(r){ return String(r['신청ID']) === id; })[0];
    if (!req) throw new Error('신청을 찾을 수 없습니다.');
    if (String(req['상태']) !== '대기') throw new Error('이미 처리된 신청입니다.');
    var qty = n(req['수량']), unit = n(x.개당단가 !== undefined && x.개당단가 !== '' ? x.개당단가 : req['희망단가']);
    push(S_IN, {
      기록ID: uid('I'), 일자: today(), 시각: now(), 판매자: String(req['캐릭명']).trim(),
      수량: qty, 개당단가: unit, 총액: qty * unit,
      등록자: '판매승인', 비고: String(req['비고'] || '')
    });
    markSale(id, '승인');
  } else {
    markSale(id, '거절');
  }
  return adminDash();
}

function fixAct(x){
  var s = sh(S_FIX);
  var data = rows(S_FIX);
  data.forEach(function(r,i){
    if (String(r['요청ID']) === String(x.요청ID)) {
      s.getRange(i+2, 6).setValue(x.결정 === '처리' ? '처리완료' : '반려');
      s.getRange(i+2, 7).setValue(stamp());
      touch();
    }
  });
  return adminDash();
}

function cancelRecord(x){
  var s = sh(S_OUT);
  var data = rows(S_OUT);
  data.forEach(function(r,i){
    if (String(r['기록ID']) === String(x.기록ID)) s.getRange(i+2, 11).setValue('취소'); touch();
  });
  return adminDash();
}

function fundAdd(x){
  var amt = n(x.금액);
  if (amt <= 0) throw new Error('금액을 입력하세요.');
  var type = x.유형 === '지출' ? '지출' : '수입';
  push(S_FUND, {
    기록ID: uid('M'), 일자: String(x.일자 || today()), 시각: now(),
    유형: type, 출처: String(x.출처 || '기타').trim(), 금액: amt,
    메모: String(x.메모 || '').trim(), 등록자: '운영진'
  });
  return adminDash();
}

function fundDel(x){
  var s = sh(S_FUND);
  var data = rows(S_FUND);
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i]['기록ID']) === String(x.기록ID)) { s.deleteRow(i + 2); touch(); break; }
  }
  return adminDash();
}

function setCfg(x){
  var s = sh(S_CFG);
  var data = rows(S_CFG);
  var pairs = [];
  if (x.현재잔고 !== undefined) {
    // 지금 금고에 있는 실제 금액을 넣으면, 기록된 수입·지출·매입을 되돌려 기준값을 맞춥니다
    var st = stock();
    var flow = (st.자금수입 || 0) - (st.자금지출 || 0) - (st.입고총액 || 0);
    pairs.push(['자금초기잔고', String(n(x.현재잔고) - flow)]);
  }
  else if (x.자금초기잔고 !== undefined) pairs.push(['자금초기잔고', String(n(x.자금초기잔고))]);
  if (x.미반납경고일수 !== undefined) pairs.push(['미반납경고일수', String(Math.max(1, n(x.미반납경고일수)))]);
  pairs.forEach(function(pr){
    var found = -1;
    data.forEach(function(r,i){ if (String(r['키']).trim() === pr[0]) found = i + 2; });
    if (found === -1) { s.appendRow([pr[0], pr[1]]); }
    else { s.getRange(found, 2).setValue(pr[1]); }
  });
  touch();
  return adminDash();
}

function setPw(x){
  var np = String(x.새비밀번호 || '').trim();
  if (np.length < 4) throw new Error('비밀번호는 4자 이상으로 정해주세요.');
  if (np.length > 30) throw new Error('비밀번호가 너무 깁니다.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', np);
  return { 변경됨: true };
}

function closeDay(x){
  var date = String(x.일자 || today());
  var s = sh(S_DAY);
  var data = rows(S_DAY);
  var idx = -1;
  data.forEach(function(r,i){ if (d(r['일자']) === date) idx = i + 2; });
  var on = x.해제 ? false : true;
  if (idx === -1) push(S_DAY, { 일자:date, 총불출량:0, 마감여부:on, 마감시각: on ? stamp() : '', 메모:'' });
  else {
    s.getRange(idx, 3).setValue(on);
    s.getRange(idx, 4).setValue(on ? stamp() : '');
    touch();
  }
  return adminDash();
}

/* 시트를 직접 수정해도 앱이 알아채도록 */
function onEdit(e){ touch(); }

/* ---------- 시트 메뉴 ---------- */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('빨시 장부')
    .addItem('시트 만들기 / 점검', '시트만들기')
    .addItem('운영진 비밀번호 바꾸기', '비밀번호바꾸기')
    .addToUi();
}
