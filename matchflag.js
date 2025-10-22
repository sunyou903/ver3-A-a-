/*
 * matchflag.js — 브라우저용 교차검증(ABCDE)
 * 요구사항:
 *  - matchflag.html 에서 SheetJS(xlsx.full.min.js) 로드 후 본 스크립트를 defer로 로드
 *  - window.run() 을 제공 (버튼 클릭 시 실행)
 *  - 업로드된 단일 엑셀 통합문서 내 여러 시트를 대상으로 A~E 검사 수행
 *  - 결과는 하나의 통합 결과 엑셀(요약 + 각 검사: 전체/불일치/일치 시트)로 다운로드
 *
 * 주의:
 *  - 파서는 수식 원문(cell.f) 접근이 필요하므로 XLSX.read(..., {cellFormula:true}) 옵션이 켜져 있어야 함
 */
(function(){
  // ------------------------------
  // 기본 유틸 및 로깅
  // ------------------------------
  const LOG_ID = 'mfLog';
  function log(msg){
    const el = document.getElementById(LOG_ID);
    if(!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
  }

  function safeGet(obj, path, dflt=null){
    try{ return path.split('.').reduce((o,k)=>o&&o[k], obj) ?? dflt; } catch(_){ return dflt; }
  }

  // 알파벳 열번호 <-> 숫자 변환
  function colToNum(col){
    let n = 0; for (let i=0;i<col.length;i++){ n = n*26 + (col.charCodeAt(i)-64); } return n;
  }
  function numToCol(num){
    let col = ''; while(num>0){ const rem=(num-1)%26; col=String.fromCharCode(65+rem)+col; num = Math.floor((num-1)/26); } return col;
  }

  // ------------------------------
  // 정규화/키/라벨 유틸
  // ------------------------------
  const stripWS = s => s==null? '' : String(s).replace(/[\u3000\s]+/g,'').trim(); // 전각 포함 공백 제거
  const normWS = s => s==null? '' : String(s).replace(/,/g,' ').replace(/[\u3000\s]+/g,' ').trim(); // 쉼표→공백, 다중공백 축약
  const normKey = (a,b) => `${normWS(a)}|${normWS(b)}`.trim();

  // 라벨 동의어 맵(공통)
  const LABELS_COMMON = {
    품명: ['품명','품 명','자재명','품 목','공종명'],
    규격: ['규격','규 격','사양','규격/사양'],
    단위: ['단위'],
    수량: ['수량','수 량','물 량'],
    재료비단가: ['재료비단가','재료비 단가','재료비 적용단가','재료비적용단가'],
    노무비단가: ['노무비','노 무 비','노무비 단가','노무비적용단가'],
    경비단가: ['경비','경비 단가','경비 적용단가','경비적용단가'],
    합계단가: ['합계단가','합 계 단 가','총단가','총 단가']
  };
  function nearestHeaderLikeUp(ws, startRow0, colMap, maxScan=200){
    // ws: SheetJS ws, startRow0: 0-based, colMap: {품명, 단위, 수량}는 0-based col index
    const blank = new Set([null, '', 0, '-', '—']);
    for (let R = startRow0; R >= Math.max(0, startRow0 - maxScan); R--){
      const pname = safeGet(ws[XLSX.utils.encode_cell({r:R, c:colMap['품명']})], 'v');
      if (!String(pname||'').trim()) continue;
      const unit  = safeGet(ws[XLSX.utils.encode_cell({r:R, c:colMap['단위']})], 'v');
      const qty   = safeGet(ws[XLSX.utils.encode_cell({r:R, c:colMap['수량']})], 'v');
      const s = String(pname||'').replace(/\s|\u3000/g,'');
      if (blank.has(unit) && blank.has(qty) && !s.includes('합계')) return R; // 0-based header row
    }
    return null;
  }
   function buildHeaderKey(ws, row0, colMap){
    // PY 로직과 동일: "원문(raw)에서 먼저 분할 보정" → 그 다음 정규화
    const rawP = safeGet(ws[XLSX.utils.encode_cell({ r: row0, c: colMap['품명'] })], 'v');
    const rawS = safeGet(ws[XLSX.utils.encode_cell({ r: row0, c: colMap['규격'] })], 'v');
  
    const pname = normWS(rawP);
    let   spec  = normWS(rawS);
  
    if (!spec) {
      // 전/반각 공백 2칸 이상 기준으로 '원문'에서 분리 → 각 토막을 정규화 후 사용
      const toks = String(rawP ?? '')
        .split(/[\u3000 ]{2,}/)
        .map(s => normWS(s))
        .filter(Boolean);
      if (toks.length >= 2) {
        return `${toks[0]}|${toks[1]}`.replace(/^\|+|\|+$/g, '');
      }
    }
    return `${pname}|${spec}`.replace(/^\|+|\|+$/g, '');
  }
  /** 병합된 셀을 실제 값으로 채워넣는 유틸 */
  function fillMergedCells(ws) {
    const merges = ws['!merges'] || [];
    for (const m of merges) {
      const topLeft = XLSX.utils.encode_cell({r: m.s.r, c: m.s.c});
      const val = ws[topLeft]?.v;
      if (val == null) continue;
      for (let R = m.s.r; R <= m.e.r; R++) {
        for (let C = m.s.c; C <= m.e.c; C++) {
          const addr = XLSX.utils.encode_cell({r:R, c:C});
          if (!ws[addr]) ws[addr] = {};
          if (ws[addr].v == null) ws[addr].v = val;
        }
      }
    }
  }

  function findHeaderRowAndCols(ws, wants, scanRows=12, scanCols=40){
    fillMergedCells(ws);  
    // ws: SheetJS worksheet, wants: string[] of logical labels (keys in LABELS_COMMON)
    // returns {headerRow, colMap:{want->colIndex(number starting 0)}} or null
    const range = XLSX.utils.decode_range(ws['!ref']||'A1:Z100');
    const maxR = Math.min(range.e.r, scanRows-1);
    const maxC = Math.min(range.e.c, scanCols-1);

    // 미리 AOA로 뽑아두면 string 비교가 쉬움
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
    const synonyms = wants.map(w => [w, (LABELS_COMMON[w]||[w]).map(stripWS)]);

    // 각 행별로 "매칭된 라벨 수"를 세고 최다인 행을 헤더로
    const rowHits = new Map();
    for (let r=0; r<=Math.min(maxR, aoa.length-1); r++){
      const row = aoa[r] || [];
      let hitCount = 0; const localHits = {};
      for (let c=0; c<=Math.min(maxC, row.length-1); c++){
        const v = row[c]; if (typeof v !== 'string') continue; const key = stripWS(v);
        for (const [want, syns] of synonyms){
          if (syns.includes(key)) { hitCount++; localHits[want] = (localHits[want]??[]).concat(c); }
        }
      }
      if (hitCount>0) rowHits.set(r, {hitCount, localHits});
    }
    if (rowHits.size===0) return null;
    const headerRow = [...rowHits.entries()].sort((a,b)=>b[1].hitCount-a[1].hitCount)[0][0];
    const bestHits = rowHits.get(headerRow).localHits;

    const colMap = {};
    for (const want of wants){
      const cols = bestHits[want]||[];
      if (!cols.length) return null; // 필수 라벨 누락
      colMap[want] = cols[0];
    }
    return {headerRow, colMap};
  }

  // 수식 문자열 f에서 모든 A1 참조를 뽑아내 sheet/row 배열로 반환
// ===== Utilities: formula reference parser =====
  // 수식 문자열 f에서 모든 A1 참조를 뽑아내 sheet/row 배열로 반환
  function extractRefsFromFormula(f){
    const out = [];
    if (typeof f !== 'string' || !f) return out;
  
    // 'Sheet Name'!A1  |  [Book]Sheet!A1  |  Sheet!A1
    // (영역 참조는 첫 셀만 row 취득; $ 절대참조 허용)
    const re = /'(.*?)'\s*!\s*\$?([A-Z]{1,3})\$?(\d+)(?::[A-Z]{1,3}\d+)?|(?:\[[^\]]+\])?([^\s!'"\(]+)\s*!\s*\$?([A-Z]{1,3})\$?(\d+)(?::[A-Z]{1,3}\d+)?/g;
  
    let m;
    while ((m = re.exec(f)) !== null) {
      // 그룹 1: 따옴표 시트 / 그룹 4: 비따옴표 시트
      let rawSheet = (m[1] ?? m[4] ?? '').trim();
      // 열/행 (2,3) 또는 (5,6)
      const rowStr = (m[3] ?? m[6] ?? '').trim();
      const row = parseInt(rowStr, 10);
  
      // 시트명 정규화(함수 래핑, =, [Book] 등 제거)
      const sheet = normalizeSheetNameForLookup(rawSheet);
  
      if (sheet && Number.isFinite(row)) {
        out.push({ sheet, row });
      }
    }
    return out;
  }

  // (선택 권장) 여러 체크에서 공용으로 쓰는 범용 수집기
  function collectRowRefs(ws, R, cols, selfName){
    const refs = [];
    if (!ws || !Array.isArray(cols) || cols.length===0) return refs;
  
    const selfNorm = normalizeSheetNameForLookup(selfName || '');
  
    for (const c of cols){
      const addr = XLSX.utils.encode_cell({r:R, c});
      const cell = ws[addr];
      const f = cell && typeof cell.f==='string' ? cell.f : null;
      if (!f) continue;
  
      const found = extractRefsFromFormula(f);
      for (const r of found){
        const sheetNorm = normalizeSheetNameForLookup(r.sheet);
        if (!sheetNorm) continue;
        if (selfNorm && sheetNorm.toLowerCase() === selfNorm.toLowerCase()) continue;
        refs.push({ sheet: sheetNorm, row: r.row });
      }
    }
  
    // 중복 제거
    const uniq = []; const seen = new Set();
    for (const it of refs){
      const key = `${it.sheet}#${it.row}`;
      if (!seen.has(key)){ seen.add(key); uniq.push(it); }
    }
    return uniq;
  }


// 행 R(0-based)에서 지정한 열들(cols 배열)의 모든 수식을 훑어 외부 참조 목록을 수집
// selfName: 현재 검사 대상 시트 이름(자기참조 제외 목적)
  function collectRowRefs(ws, R, cols, selfName){
    const refs = [];
    if (!ws || !Array.isArray(cols) || cols.length===0) return refs;
  
    for (const c of cols){
      const addr = XLSX.utils.encode_cell({r:R, c});
      const cell = ws[addr];
      const f = cell && typeof cell.f==='string' ? cell.f : null;
      if (!f) continue;
  
      // 수식에서 A1 참조들을 뽑아온다
      const found = extractRefsFromFormula(f);
      for (const r of found){
        // 자기 시트 참조는 제외
        if (!r.sheet || (selfName && r.sheet === selfName)) continue;
        // 최소 정보만 보존: {sheet, row}
        refs.push({ sheet: r.sheet, row: r.row });
      }
    }
  
    // 중복 제거(같은 셀을 여러 번 참조하는 경우)
    const uniq = [];
    const seen = new Set();
    for (const it of refs){
      const key = `${it.sheet}#${it.row}`;
      if (!seen.has(key)){ seen.add(key); uniq.push(it); }
    }
    return uniq;
  }

  function buildRowKeyMap(ws, headerRow, colMap, keySpec){
    // keySpec: {left:'품명', right:'규격'} 등 논리키
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
    const start = headerRow+1;
    const out = new Map(); // rowIndex(1-based Excel row) -> key string
    for (let r=start; r<aoa.length; r++){
      const row = aoa[r]||[];
      const a = row[colMap[keySpec.left]]; const b = keySpec.right? row[colMap[keySpec.right]] : '';
      const k = normKey(a,b);
      if (k && k !== '|' && normWS(a)){
        const excelRow = r+1; // 0-based -> Excel 1-based
        out.set(excelRow, k);
      }
    }
    return out;
  }

  // ------------------------------
  // 수식/셀 접근 유틸
  // ------------------------------
  // '시트명'!$A$123 패턴(작은따옴표 포함/미포함, 절대/상대 참조)
  const SHEET_REF_RE = /(?:'([^']+)'|([^'!:\]\[]+))!\$?([A-Z]{1,3})\$?(\d+)/g;

  function eachCell(ws, cb){
    const ref = ws['!ref']; if (!ref) return;
    const r = XLSX.utils.decode_range(ref);
    for (let R=r.s.r; R<=r.e.r; R++){
      for (let C=r.s.c; C<=r.e.c; C++){
        const addr = {r:R, c:C};
        const A1 = XLSX.utils.encode_cell(addr);
        const cell = ws[A1];
        cb(cell, A1, R, C);
      }
    }
  }

  function collectExternalRefs(formula, selfSheet){
    if (!formula || typeof formula!=='string') return [];
    const out = [];
    SHEET_REF_RE.lastIndex = 0;
    let m;
    while((m = SHEET_REF_RE.exec(formula))){
      const sheet = (m[1]||m[2]||'').trim();
      if (!sheet || sheet===selfSheet) continue; // 자기시트 참조는 제외(외부참조만)
      const col = m[3], row = parseInt(m[4],10);
      out.push({sheet, col, row});
    }
    return out;
  }

  function mostFrequentRef(refs){
    // refs: [{sheet,col,row}, ...] -> 대표 (시트,row) 최빈값
    if (!refs.length) return null;
    const freq = new Map();
    for (const r of refs){
      const key = `${r.sheet}|${r.row}`;
      freq.set(key, (freq.get(key)||0)+1);
    }
    let bestKey=null, bestN=-1;
    for (const [k,v] of freq.entries()) if (v>bestN){ bestN=v; bestKey=k; }
    if (!bestKey) return null;
    const [sheet, rowStr] = bestKey.split('|');
    return {sheet, row: parseInt(rowStr,10), count: bestN};
  }

  // 시트 가져오기 by 이름(대소문자 무시형)
  function getSheetCaseInsensitive(wb, name){
    if (wb.SheetNames.includes(name)) return wb.Sheets[name];
    const low = name.toLowerCase();
    for (const sn of wb.SheetNames){ if (sn.toLowerCase()===low) return wb.Sheets[sn]; }
    return null;
  }
  // 일반화: 어떤 시트든 주어진 행의 '품명|규격' 키를 생성
  function getKeyFromAnySheet(wb, sheetName, row1){
    const ws = getSheetCaseInsensitive(wb, sheetName);
    if (!ws) return '';
    const def = findHeaderRowAndCols(ws, ['품명','규격']);
    if (!def) return '';
    const map = buildRowKeyMap(ws, def.headerRow, def.colMap, {left:'품명', right:'규격'});
    return map.get(row1) || '';
  }


  // "[원본.xlsx]단가대비표" -> "단가대비표"
  // "'단가대비표'" -> "단가대비표"
  // "=TRUNC(단가대비표" -> "단가대비표"
  // "IFERROR(TRUNC(일위대가목록" -> "일위대가목록"
  function normalizeSheetNameForLookup(name){
    let s = String(name ?? '').trim();
  
    // 1) 앞/뒤 작은따옴표 제거
    s = s.replace(/^'+|'+$/g, '');
  
    // 2) 수식의 '=' 제거 (있다면)
    s = s.replace(/^=/, '');
  
    // 3) 함수 래핑 제거: FUNC( …  , 중첩 함수도 반복 제거
    //    예: IFERROR(TRUNC(단가대비표 -> 단가대비표
    while (/^[A-Za-z_][A-Za-z0-9_]*\(/.test(s)) {
      s = s.replace(/^[A-Za-z_][A-Za-z0-9_]*\(/, '');
    }
  
    // 4) [통합문서] 접두 제거 (이번 파일에선 없어도 안전)
    s = s.replace(/^\[[^\]]+\]/, '');
  
    // 5) 잔여 공백/괄호 정리
    s = s.replace(/^\s+/, '').replace(/\s+$/, '');
    s = s.replace(/^\(+/, ''); // 혹시 남은 여는 괄호
    return s.trim();
  }

  // ------------------------------
  // 검사 A: 일위대가 ↔ (단가대비표|일위대가목록) — "행 단위" 집계로 수정
  // - 각 데이터 행에서 특정 열 집합만 스캔하여 외부참조를 모으고 대표(최빈) 1개로 판정
  // - 참조 없음 && 수량 값이 상수(수식 아님) && 0이 아니면 불일치
  // ------------------------------
  
  function checkA(wb){
    const S_A = getSheetCaseInsensitive(wb,'일위대가');
    const S_B = getSheetCaseInsensitive(wb,'단가대비표');
    const S_C = getSheetCaseInsensitive(wb,'일위대가목록');
    const out = [];
    if (!S_A || !(S_B || S_C)) return {name:'A', rows:[], summary:{note:'필수 시트 미존재'}};
  
    // 헤더 및 키 맵
    const Adef = findHeaderRowAndCols(S_A, ['품명','규격','단위','수량']);
    const Bdef = S_B ? findHeaderRowAndCols(S_B, ['품명','규격','단위']) : null;
    const Cdef = S_C ? findHeaderRowAndCols(S_C, ['품명','규격'])           : null;
    if (!Adef) return {name:'A', rows:[], summary:{note:'일위대가 헤더 미검출'}};
  
    const Amap = buildRowKeyMap(S_A, Adef.headerRow, Adef.colMap, {left:'품명', right:'규격'});
    const Bmap = (S_B && Bdef) ? buildRowKeyMap(S_B, Bdef.headerRow, Bdef.colMap, {left:'품명', right:'규격'}) : new Map();
    const Cmap = (S_C && Cdef) ? buildRowKeyMap(S_C, Cdef.headerRow, Cdef.colMap, {left:'품명', right:'규격'}) : new Map();
  
    // 행 전체 열 스캔
    const range = XLSX.utils.decode_range(S_A['!ref']);
    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S_A) || '일위대가';
  
    for (let R = Adef.headerRow+1; R <= range.e.r; R++){
      const excelRow = R+1;
      const myKey = Amap.get(excelRow);
      if (!myKey) continue;
  
      const pnameA1 = XLSX.utils.encode_cell({r:R, c:Adef.colMap['품명']});
      const gnameA1 = XLSX.utils.encode_cell({r:R, c:Adef.colMap['규격']});
      const pname = normWS(safeGet(S_A[pnameA1],'v'));
      const gname  = normWS(safeGet(S_A[gnameA1],'v'));
      const hasPercent = (pname && pname.includes('%')) || (gname && gname.includes('%'));
  
      let foundAnyRef = false;
  
      // 열 전부 스캔
      for (let C = 0; C <= range.e.c; C++){
        const A1 = XLSX.utils.encode_cell({r:R, c:C});
        const f = safeGet(S_A[A1],'f');
        if (!f) continue;
        const refs = collectExternalRefs(f, selfName);
        if (!refs.length) continue;
  
        for (const {sheet, col, row} of refs){
          // 대상 시트 한정: 단가대비표 / 일위대가목록
          const sheetLower = String(sheet||'').toLowerCase();
          const isDV = sheetLower === '단가대비표';
          const isLS = sheetLower === '일위대가목록';
          if (!isDV && !isLS) continue;
  
          foundAnyRef = true;
          const refKey = isDV ? (Bmap.get(row)||'') : (Cmap.get(row)||'');
          let status = '일치';
          if (!refKey || normWS(refKey.split('|')[0])!==normWS(myKey.split('|')[0]) || normWS(refKey.split('|')[1])!==normWS(myKey.split('|')[1])){
            status = hasPercent ? '제외' : '불일치';
          }
  
          // Py와 유사한 필드 구성
          out.push({
            시트: '일위대가',
            행: excelRow,
            '일위대가_품명|규격': myKey,
            참조시트: sheet,
            참조셀: `${sheet}!${col}${row}`,
            참조키: refKey || '',
            결과: status
          });
        }
      }
  
      // 참조 전혀 없는데 수량 값이 입력(0/빈칸 제외) → 불일치
      if (!foundAnyRef) {
        const qCol = Adef.colMap['수량'];
        if (typeof qCol === 'number') {
          const qA1 = XLSX.utils.encode_cell({r:R, c:qCol});
          const qCell = S_A[qA1];
          const val = safeGet(qCell,'v');
          const isEmpty = (val===null || val===undefined || val==='' || Number(val)===0);
          if (!isEmpty){
            out.push({
              시트: '일위대가',
              행: excelRow,
              '일위대가_품명|규격': myKey,
              참조시트: '',
              참조셀: '',
              참조키: '',
              결과: '불일치'
            });
          }
        }
      }
    }
  
    // 요약(참조 개수 기준 집계)
    const mismatches = out.filter(r=>r.결과==='불일치');
    const matches    = out.filter(r=>r.결과==='일치');
    const summary = {검사:'A', 전체:out.length, 일치:matches.length, 불일치:mismatches.length};
    return {name:'A', rows:out, summary, matches, mismatches};
  }

  // ------------------------------
// 검사 B: 일위대가목록 ↔ 일위대가 — 파이썬(run_check_B) 정합
// 검사 B : 일위대가목록 ↔ 일위대가 (PY run_check_B 동일 로직)
function checkB(wb){
  const S_C = getSheetCaseInsensitive(wb,'일위대가목록');
  const S_A = getSheetCaseInsensitive(wb,'일위대가');
  const out = [];
  if (!S_C || !S_A) return {name:'B', rows:[], summary:{note:'필수 시트 미존재'}};

  function normalizeSheetName(s){
    return String(s ?? '')
      .replace(/^'+|'+$/g, '')   // '일위대가' → 일위대가
      .trim()
      .toLowerCase();
  }

  // PY와 동일하게: 목록은 품명/규격만 필수, 일위대가는 품명/규격/단위/수량
  const Cdef = findHeaderRowAndCols(S_C, ['품명','규격']);
  const Adef = findHeaderRowAndCols(S_A, ['품명','규격','단위','수량']);
  if (!Cdef || !Adef) return {name:'B', rows:[], summary:{note:'헤더 미검출'}};

  // 키 맵(행→"품명|규격"), 기존 구조 유지
  const Cmap = buildRowKeyMap(S_C, Adef ? Cdef.headerRow : Cdef.headerRow, Cdef.colMap, {left:'품명', right:'규격'});
  const Amap = buildRowKeyMap(S_A, Adef.headerRow, Adef.colMap, {left:'품명', right:'규격'});

  // 시트 실명 확보(자기 시트명 / 일위대가 시트명) → 필터에 사용
  const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S_C) || '일위대가목록';
  const Aname   = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S_A) || '일위대가';

  const range = XLSX.utils.decode_range(S_C['!ref']);

  // PY처럼 전열 스캔
  const colsToScan = [];
  for (let c = 0; c <= range.e.c; c++) colsToScan.push(c);

  for (let R = Cdef.headerRow+1; R <= range.e.r; R++){
    const excelRow = R+1;
    const myKey = Cmap.get(excelRow);
    if (!myKey) continue;

    const refsAll = collectRowRefs(S_C, R, colsToScan, selfName);
    const refsToA = refsAll.filter(r => normalizeSheetName(r.sheet) === normalizeSheetName(Aname));
    if (!refsToA.length) continue;

    const rep = mostFrequentRef(refsToA);
    if (!rep) continue;

    // PY: 참조행 기준 위로 올라가며 '헤더처럼 보이는' 행 찾기 → 그 '헤더행' 자체로 키 구성
    const headerRowLike = nearestHeaderLikeUp(S_A, rep.row-1, Adef.colMap, 80);
    const refKey = (headerRowLike != null)
      ? buildHeaderKey(S_A, headerRowLike, Adef.colMap)
      : (Amap.get(rep.row) || '');

    const status = (!refKey || normWS(refKey)!==normWS(myKey)) ? '불일치' : '일치';
    out.push({
      시트:'일위대가목록',
      행: excelRow,
      키: myKey,
      참조시트: Aname,
      참조행: rep.row,
      참조키: refKey || '',
      결과: status
    });
  }

  // 기존 출력 형식 그대로 유지
  return summarize('B', out);
}

  // ------------------------------
  // 검사 C: 공종별내역서 — 행 단위로 대표 참조, 합계단가 상수 판정은 해당 열만 체크
  // ------------------------------
 // ------------------------------
// 검사 C: 공종별내역서 ↔ 단가대비표 (PY run_check_C 동일 로직)
// ------------------------------
// ------------------------------
// 검사 C: 공종별내역서 ↔ 단가대비표 (PY run_check_C 동등 + 참조행 직접키/위로보정)
// ------------------------------

  function checkC(wb){
    const S = getSheetCaseInsensitive(wb,'공종별내역서');
    const T = getSheetCaseInsensitive(wb,'단가대비표');
    const out = [];
    if (!S || !T) return {name:'C', rows:[], summary:{note:'필수 시트 미존재'}};

    const wantsS = ['품명','규격','합계단가'];
    const wantsT = ['품명','규격','단위'];
    const Sdef = findHeaderRowAndCols(S, wantsS);
    const Tdef = findHeaderRowAndCols(T, wantsT);
    if (!Sdef || !Tdef) return {name:'C', rows:[], summary:{note:'헤더 미검출'}};

    const Smap = buildRowKeyMap(S, Sdef.headerRow, Sdef.colMap, {left:'품명', right:'규격'});
    const Tmap = buildRowKeyMap(T, Tdef.headerRow, Tdef.colMap, {left:'품명', right:'규격'});
    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S) || '공종별내역서';

    const range = XLSX.utils.decode_range(S['!ref']);
    // 합계단가 열과 그 주변 우측 금액열 위주로 스캔(불필요한 셀 스캔 제거)
    const colsToScan = new Set();
    for (let c = 0; c <= range.e.c; c++) colsToScan.add(c);  // 전체 열 스캔
    const sumCol = Sdef.colMap['합계단가']; // 아래 상수판정 시에만 사용

    for (let R = Sdef.headerRow+1; R <= range.e.r; R++){
      const excelRow = R+1;
      const myKey = Smap.get(excelRow); if (!myKey) continue;
      const refs = collectRowRefs(S, R, [...colsToScan], selfName);
      const rep = mostFrequentRef(refs);
      let status='일치', refKey='', refSheet='', refRow='';
      if (rep){
        refSheet = rep.sheet; refRow = rep.row;
        const tKey = getKeyFromAnySheet(wb, rep.sheet, rep.row);
        refKey = tKey||'';
        if (!tKey || normWS(tKey)!==normWS(myKey)) status='불일치';
      } else {
        // 외부참조 없음 → 합계단가 셀만 검사하여 상수 여부 판단
        if (typeof sumCol==='number'){
          const sumA1 = XLSX.utils.encode_cell({r:R, c:sumCol});
          const scell = S[sumA1];
          const sF = safeGet(scell,'f');
          const sV = Number(safeGet(scell,'v'));
          const EPS = 1e-9; // 미세 부동소수 오차 허용치
          if (!sF && Number.isFinite(sV) && Math.abs(sV) > EPS){ status='불일치'; } else { status='제외'; }
        }
      }
      out.push({시트:'공종별내역서', 행:excelRow, 키:myKey, 참조시트:refSheet, 참조행:refRow, 참조키:refKey||'', 결과:status});
    }

    return summarize('C', out);
  }

  // ------------------------------
  // 검사 D: 공종별집계표 — 재/노/경 단가 외부참조 대표 ↔ 품명만 느슨 비교
  // ------------------------------
 
  // ------------------------------
 // ------------------------------
// 검사 D (공종별집계표): 열 단위 유지 + 합계행 헤더탐색 + 진단 카운터
// ------------------------------
 // D 검사 (B와 동일한 스타일 + 헤더 판정 강화)
// - 공종별집계표의 재/노/경 각 열에서 외부 참조를 모아 대표(시트,행) 선택
// - 대표 참조 시트에서 "윗방향"으로 헤더처럼 보이는 행을 탐색(강화 조건)
// - 헤더 행의 품명|규격 키를 만들어 공종별집계표의 품명(좌측)과 비교

function checkD(wb){
  const S = getSheetCaseInsensitive(wb,'공종별집계표');
  if (!S) return {name:'D', rows:[], summary:{note:'공종별집계표 미존재'}};

  const Sdef = findHeaderRowAndCols(S, ['품명','재료비단가','노무비단가','경비단가']);
  if (!Sdef) return {name:'D', rows:[], summary:{note:'공종별집계표 헤더 미검출'}};

  const MAX_SCAN = 120; // 필요 시 150 등으로 상향 가능
  const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S) || '공종별집계표';

  const nameCol = Sdef.colMap['품명'];
  const getMyName = (r)=>{
    const a1 = XLSX.utils.encode_cell({r, c:nameCol});
    return normWS(safeGet(S[a1],'v'));
  };

  // ---- 대상(참조) 시트: B와 동일하게 품명/규격/단위/수량을 요구 ----
  const wantsT = ['품명','규격','단위','수량'];

  // 시트별 캐시
  const targetCache = new Map(); // sheetName -> {ws, def, rowKeyMap, numericCols}
  const getTargetDef = (sheetName)=>{
    if (!sheetName) return null;
    const key = String(sheetName);
    if (targetCache.has(key)) return targetCache.get(key);

    const ws = getSheetCaseInsensitive(wb, sheetName);
    if (!ws){ targetCache.set(key, null); return null; }

    const def = findHeaderRowAndCols(ws, wantsT);
    const rowKeyMap = def ? buildRowKeyMap(ws, def.headerRow, def.colMap, {left:'품명', right:'규격'}) : null;
    const numericCols = def ? detectNumericColsAtHeader(ws, def.headerRow) : [];
    const pack = {ws, def, rowKeyMap, numericCols};
    targetCache.set(key, pack);
    return pack;
  };

  const out = [];
  eachCell(S, (cell, A1, R, C)=>{
    const excelRow = R+1;
    if (excelRow <= Sdef.headerRow+1) return;

    // 재/노/경 단가 열만 검사
    if (
      C!==Sdef.colMap['재료비단가'] &&
      C!==Sdef.colMap['노무비단가'] &&
      C!==Sdef.colMap['경비단가']
    ) return;

    const myName = getMyName(R);
    if (!myName) return;

    const f = safeGet(cell,'f');
    const refs = collectExternalRefs(f, selfName); // 수식에서 (시트,행) 목록
    const rep  = mostFrequentRef(refs);           // 대표 참조 1개

    let status='일치', refKey='', refSheet='', refRow='';
    if (rep){
      refSheet = rep.sheet; refRow = rep.row;

      const pack = getTargetDef(rep.sheet);
      if (pack && pack.def){
        // ---- 강화된 헤더 탐색: 합계류 변형 배제 + 숫자성 컬럼 비어있음 보장 ----
        const headerRowLike = nearestHeaderLikeUp_Strict(pack.ws, rep.row-1, pack.def, pack.numericCols, MAX_SCAN);

        const tKey = (headerRowLike!=null)
          ? buildHeaderKey(pack.ws, headerRowLike, pack.def.colMap)      // 헤더행 자체로 키 구성
          : (pack.rowKeyMap ? (pack.rowKeyMap.get(rep.row) || '') : ''); // 폴백: 동일행 키

        refKey = tKey || '';
        const tName = String(refKey).split('|')[0]?.trim() || '';
        if (!tName || normWS(tName) !== normWS(myName)) status='불일치';
      } else {
        // 대상 시트에서 헤더를 못 찾았거나 구조를 못 읽음
        status='불일치';
      }
    }

    out.push({
      시트:'공종별집계표',
      행:excelRow,
      품명:myName,
      단가열:(C===Sdef.colMap['재료비단가']?'재료비':C===Sdef.colMap['노무비단가']?'노무비':'경비'),
      참조시트:refSheet,
      참조행:refRow,
      참조키:refKey||'',
      결과:status
    });
  });

  return summarize('D', out);
}

/* ===================== 보조 유틸 (강화된 헤더 탐색) ===================== */

// 합계/소계/집계/계 등 변형 문자열(공백·괄호 포함) 배제
  function isSumLikeName(name){
    const s = String(name || '')
      .replace(/\u3000/g,' ')            // 전각공백 -> 보통 공백
      .replace(/[\s\[\]\(\)\{\}]/g,'');  // 공백/괄호류 제거
    return s.includes('합계') || s.includes('소계') || s.includes('집계') || s === '계';
  }
  
  // 숫자/빈값 판정
  function isNumericLike(v){
    if (v == null) return false;
    if (typeof v === 'number') return Number.isFinite(v);
    const s = String(v).replace(/[, \t]/g,'');
    return /^-?\d+(\.\d+)?$/.test(s);
  }
  function asNumber(v){
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (v == null) return null;
    const s = String(v).replace(/[, \t]/g,'');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : null;
  }
  function isEmptyOrDash(v){
    const s = normWS(v);
    return !s || s === '-' || s === '—';
  }
  
  // 라벨행에서 숫자성 컬럼 후보 추출(금액/단가/합계/재·노·경 등)
  function detectNumericColsAtHeader(ws, headerRow0){
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const cols = [];
    const KW = ['금액','단가','합계','소계','재료비','노무비','경비','직접재료비','노무비합계','경비합계'];
    for (let c = range.s.c; c <= range.e.c; c++){
      const t = normWS(safeGet(ws[XLSX.utils.encode_cell({r: headerRow0, c})], 'v'));
      if (!t) continue;
      if (t === '품명' || t === '규격' || t === '단위' || t === '수량') continue;
      if (KW.some(k => t.includes(k))) cols.push(c);
    }
    return cols;
  }
  
  // 강화된 후보행 판정 로직
  function headerCandidateOK(ws, row0, colMap, numericCols){
    const nameCol = colMap['품명'];
    const unitCol = colMap['단위'];
    const qtyCol  = colMap['수량'];
  
    // 1) 품명 조건
    const name = normWS(safeGet(ws[XLSX.utils.encode_cell({r: row0, c: nameCol})], 'v'));
    if (!name || isSumLikeName(name)) return false;
  
    // 2) 단위/수량 조건 (단위는 비거나 '-', '—'; 수량은 비거나 '-', '—' 또는 숫자 0만 허용)
    if (unitCol != null){
      const v = safeGet(ws[XLSX.utils.encode_cell({r: row0, c: unitCol})],'v');
      if (!isEmptyOrDash(v)) return false;
    }
    if (qtyCol != null){
      const v = safeGet(ws[XLSX.utils.encode_cell({r: row0, c: qtyCol})],'v');
      const n = asNumber(v);
      const txt = normWS(v);
      const okQty = (txt==='' || txt==='-' || txt==='—' || n===0);
      if (!okQty) return false;
    }
  
    // 3) 숫자성 컬럼(금액/단가/합계/재·노·경 등)은 반드시 비어 있어야 함
    for (const c of numericCols){
      const v = safeGet(ws[XLSX.utils.encode_cell({r: row0, c})],'v');
      if (v == null || v === '' || v === '-' || v === '—') continue;
      if (isNumericLike(v)) return false;   // 숫자/수식 결과처럼 보이면 탈락
      if (normWS(v)) return false;          // 텍스트라도 값이 있으면 탈락 (SUMIF 등)
    }
  
    return true;
  }
  
  // 강화판: 참조행 바로 위부터 위로 스캔해 "진짜 헤더 같은" 행을 찾음
  function nearestHeaderLikeUp_Strict(ws, startRow0, tDef, numericCols, maxScan){
    const { headerRow, colMap } = tDef || {};
    if (!colMap || headerRow == null) return null;
  
    const floor = Math.max(headerRow + 1, startRow0 - (maxScan||120));
    for (let r = startRow0; r >= floor; r--){
      if (headerCandidateOK(ws, r, colMap, numericCols)) {
        return r;
      }
    }
    return null;
  }

  // ------------------------------
  // 검사 E: 단가대비표 — 재료비/노무비 대표 참조 기반 키 비교(장비 단가산출서 특례)
  // ------------------------------
  function checkE(wb){
    const S = getSheetCaseInsensitive(wb,'단가대비표');
    const out = [];
    if (!S) return {name:'E', rows:[], summary:{note:'단가대비표 시트 미존재'}};

    // 시트명에 "장비" "단가산출서"가 포함된 경우, 사양 우선 사용
    const useSpecFirst = wb.SheetNames.some(n => /장비/.test(n) && /단가산출서/.test(n));

    const wants = useSpecFirst ? ['품명','규격','단위','재료비단가','노무비단가'] : ['품명','규격','단위','재료비단가','노무비단가'];
    const Sdef = findHeaderRowAndCols(S, wants);
    if (!Sdef) return {name:'E', rows:[], summary:{note:'헤더 미검출'}};

    // 비교용 키는 기본 품명|규격. (특례가 있으면 규격(사양) 가중치가 사실상 커지는 의미로 해석)
    const Smap = buildRowKeyMap(S, Sdef.headerRow, Sdef.colMap, {left:'품명', right:'규격'});
    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S) || '단가대비표';

    const targetSheets = new Map(); // 시트명 -> {def,map}
    function ensureTarget(sheetName){
      const ws = getSheetCaseInsensitive(wb, sheetName);
      if (!ws) return null;
      if (targetSheets.has(sheetName)) return targetSheets.get(sheetName);
      const def = findHeaderRowAndCols(ws, ['품명','규격','단위']);
      if (!def) return null;
      const map = buildRowKeyMap(ws, def.headerRow, def.colMap, {left:'품명', right:'규격'});
      const obj = {def, map, ws};
      targetSheets.set(sheetName, obj);
      return obj;
    }

    // 대상 열: 재료비/노무비 (경비는 파일마다 유무가 달라 E검사 범위에서 제외)
    const colsToCheck = [Sdef.colMap['재료비단가'], Sdef.colMap['노무비단가']].filter(v=>typeof v==='number');

    eachCell(S, (cell, A1, R, C)=>{
      const excelRow = R+1; if (excelRow<=Sdef.headerRow+1) return;
      if (!colsToCheck.includes(C)) return;
      const myKey = Smap.get(excelRow); if (!myKey) return;

      const f = safeGet(cell,'f'); const refs = collectExternalRefs(f, selfName); const rep = mostFrequentRef(refs);
      let status='일치', refKey='', refSheet='', refRow='';
      if (rep){
        const tgt = ensureTarget(rep.sheet);
        refSheet = rep.sheet; refRow = rep.row;
        refKey = tgt? (tgt.map.get(rep.row) || '') : '';
        if (!refKey || normWS(refKey)!==normWS(myKey)) status='불일치';
      }
      out.push({시트:'단가대비표', 행:excelRow, 키:myKey, 단가열:(C===Sdef.colMap['재료비단가']?'재료비':'노무비'), 참조시트:refSheet, 참조행:refRow, 참조키:refKey||'', 결과:status});
    });

    return summarize('E', out);
  }

  // ------------------------------
  // 요약/엑셀 출력 유틸
  // ------------------------------
  function summarize(name, rows){
    const mismatches = rows.filter(r=>r.결과==='불일치');
    const matches = rows.filter(r=>r.결과==='일치');
    const summary = {검사:name, 전체:rows.length, 일치:matches.length, 불일치:mismatches.length};
    return {name, rows, summary, matches, mismatches};
  }

  function objectsToSheet(objs){
    if (!objs.length) return XLSX.utils.aoa_to_sheet([["결과 없음"]]);
    const headers = Object.keys(objs[0]);
    const aoa = [headers];
    for (const o of objs){ aoa.push(headers.map(h=>o[h])); }
    return XLSX.utils.aoa_to_sheet(aoa);
  }

  function buildResultWorkbook(results){
    const wbOut = XLSX.utils.book_new();
    // 총괄 요약
    const sumRows = results.map(r=>r.summary);
    XLSX.utils.book_append_sheet(wbOut, objectsToSheet(sumRows), 'Summary');

    for (const r of results){
      XLSX.utils.book_append_sheet(wbOut, objectsToSheet(r.rows), `${r.name}_ALL`.slice(0,31));
      XLSX.utils.book_append_sheet(wbOut, objectsToSheet(r.matches||[]), `${r.name}_MATCH`.slice(0,31));
      XLSX.utils.book_append_sheet(wbOut, objectsToSheet(r.mismatches||[]), `${r.name}_MISMATCH`.slice(0,31));
    }
    return wbOut;
  }

  function downloadWorkbook(wb, baseName){
    const ts = new Date();
    const pad = n=>String(n).padStart(2,'0');
    const stamp = `${ts.getFullYear()}${pad(ts.getMonth()+1)}${pad(ts.getDate())}_${pad(ts.getHours())}${pad(ts.getMinutes())}${pad(ts.getSeconds())}`;
    const name = `${baseName||'matchflag_result'}_${stamp}.xlsx`;
    XLSX.writeFile(wb, name);
    return name;
  }

  // ------------------------------
  // 메인 엔트리
  // ------------------------------
  window.run = async function(){
    try{
      log('페이지 준비 확인: XLSX=' + (window.XLSX?'OK':'미로드'));
      const inp = document.getElementById('mfFile');
      const file = inp && inp.files && inp.files[0];
      if (!file) throw new Error('엑셀 파일을 먼저 선택해 주세요.');

      log(`파일 읽는 중: ${file.name}`);
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array', cellFormula:true, cellText:true, cellNF:true});

      log('일위대가 검사 실행…');
      const ra = checkA(wb); log(`일위대가 검사: 전체 ${ra.summary?.전체||0}, 불일치 ${ra.summary?.불일치||0}`);
      log('일위대가 목록 검사 실행…');
      const rb = checkB(wb); log(`일위대가 목록 검사: 전체 ${rb.summary?.전체||0}, 불일치 ${rb.summary?.불일치||0}`);
      log('공종별 내역서 검사 실행…');
      const rc = checkC(wb); log(`공종별 내역서 검사: 전체 ${rc.summary?.전체||0}, 불일치 ${rc.summary?.불일치||0}`);
      log('공종별 집계표 검사 실행…');
      const rd = checkD(wb); log(`공종별 집계표 검사: 전체 ${rd.summary?.전체||0}, 불일치 ${rd.summary?.불일치||0}`);
      log('단가대비표 검사(건축) 실행…');
      const re = checkE(wb); log(`단가대비표 검사(건축): 전체 ${re.summary?.전체||0}, 불일치 ${re.summary?.불일치||0}`);

      const resWb = buildResultWorkbook([ra,rb,rc,rd,re]);
      const outName = downloadWorkbook(resWb, file.name.replace(/\.[^.]+$/, '') + '_matchflag');
      log(`저장 완료: ${outName}`);
    } catch(e){
      console.error(e);
      log('오류: ' + (e.message || e));
    }
  };
})();













































