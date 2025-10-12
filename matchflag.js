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

  function findHeaderRowAndCols(ws, wants, scanRows=12, scanCols=40){
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

  // ------------------------------
  // 검사 A: 일위대가 ↔ (단가대비표|일위대가목록)
  // - 일위대가의 각 행 수식에서 외부참조 대표(최빈) 추출 → 참조 시트에서 키 비교
  // - 참조 없음 && 수량 값 존재(0 아님) → 불일치
  // ------------------------------
  function checkA(wb){
    const S_A = getSheetCaseInsensitive(wb,'일위대가');
    const S_B = getSheetCaseInsensitive(wb,'단가대비표');
    const S_C = getSheetCaseInsensitive(wb,'일위대가목록');
    const out = [];
    if (!S_A || !(S_B||S_C)) return {name:'A', rows:[], summary:{note:'필수 시트 미존재'}};

    // 헤더 탐색(공통 최소 요구)
    const wantsA = ['품명','규격','단위','수량'];
    const wantsBC = ['품명','규격','단위'];
    const Adef = findHeaderRowAndCols(S_A, wantsA);
    const Bdef = S_B ? findHeaderRowAndCols(S_B, wantsBC) : null;
    const Cdef = S_C ? findHeaderRowAndCols(S_C, wantsBC) : null;

    if (!Adef) return {name:'A', rows:[], summary:{note:'일위대가 헤더 미검출'}};
    const Amap = buildRowKeyMap(S_A, Adef.headerRow, Adef.colMap, {left:'품명', right:'규격'});
    const Bmap = S_B && Bdef ? buildRowKeyMap(S_B, Bdef.headerRow, Bdef.colMap, {left:'품명', right:'규격'}) : new Map();
    const Cmap = S_C && Cdef ? buildRowKeyMap(S_C, Cdef.headerRow, Cdef.colMap, {left:'품명', right:'규격'}) : new Map();

    // 수량 열 A1 주소
    const qtyCol = Adef.colMap['수량'];

    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S_A) || '일위대가';

    eachCell(S_A, (cell, A1, R, C)=>{
      const excelRow = R+1; if (excelRow<=Adef.headerRow+1) return; // 데이터 영역만
      const myKey = Amap.get(excelRow); if (!myKey) return;

      const f = safeGet(cell,'f');
      const refs = collectExternalRefs(f, selfName);
      const rep = mostFrequentRef(refs);

      let status = '일치', refKey = '', refSheet = '', refRow = '';
      if (rep){
        refSheet = rep.sheet; refRow = rep.row;
        const targetMap = (refSheet.toLowerCase()==='단가대비표')? Bmap : (refSheet.toLowerCase()==='일위대가목록')? Cmap : null;
        refKey = targetMap? (targetMap.get(rep.row)||'') : '';
        if (!refKey || normWS(refKey.split('|')[0])!==normWS(myKey.split('|')[0]) || normWS(refKey.split('|')[1])!==normWS(myKey.split('|')[1])){
          status = '불일치';
        }
      } else {
        // 참조 없음: 수량 값이 있고 0이 아니면 불일치
        if (C===qtyCol){ /* 수량셀 그 자체는 스킵 */ }
        const qtyA1 = XLSX.utils.encode_cell({r:R, c:qtyCol});
        const qtyCell = S_A[qtyA1];
        const qtyVal = safeGet(qtyCell,'v');
        const qtyF = safeGet(qtyCell,'f');
        const qn = Number(qtyVal);
        if (!qtyF && Number.isFinite(qn) && Math.abs(qn) > 0){ status = '불일치'; }
      }

      out.push({시트:'일위대가', 행:excelRow, 키:myKey, 참조시트:refSheet, 참조행:refRow, 참조키:refKey, 결과:status});
    });

    return summarize('A', out);
  }

  // ------------------------------
  // 검사 B: 일위대가목록 ↔ 일위대가(헤더 역할 행 근접)
  // - 목록의 각 행 수식에서 일위대가 시트 참조 추출 → 참조 행 주변에서 "헤더 역할" 행을 찾아 대표 키
  // ------------------------------
  function isHeaderLikeRow(ws, rowIndex, pos){
    // 품명은 있고, 단위/수량은 비거나 0인 행을 헤더 유사로 간주
    const a1 = (c)=>XLSX.utils.encode_cell({r:rowIndex, c});
    const v = (c)=> safeGet(ws[a1(c)],'v');
    const f = (c)=> safeGet(ws[a1(c)],'f');
    const hasName = !!normWS(v(pos['품명']));
    const unitEmpty = !normWS(v(pos['단위'])) && !f(pos['단위']);
    const qtyEmpty = !normWS(v(pos['수량'])) && !f(pos['수량']);
    return hasName && unitEmpty && qtyEmpty;
  }

  function nearestHeaderLikeUp(ws, startRow, pos, searchLimit=20){
    for (let r=startRow; r>=Math.max(0,startRow-searchLimit); r--){
      if (isHeaderLikeRow(ws, r, pos)) return r;
    }
    return null;
  }

  function checkB(wb){
    const S_C = getSheetCaseInsensitive(wb,'일위대가목록');
    const S_A = getSheetCaseInsensitive(wb,'일위대가');
    const out = [];
    if (!S_C || !S_A) return {name:'B', rows:[], summary:{note:'필수 시트 미존재'}};

    const wantsList = ['품명','규격','단위','수량'];
    const Cdef = findHeaderRowAndCols(S_C, wantsList);
    const Adef = findHeaderRowAndCols(S_A, wantsList);
    if (!Cdef || !Adef) return {name:'B', rows:[], summary:{note:'헤더 미검출'}};

    const Cmap = buildRowKeyMap(S_C, Cdef.headerRow, Cdef.colMap, {left:'품명', right:'규격'});
    const Amap = buildRowKeyMap(S_A, Adef.headerRow, Adef.colMap, {left:'품명', right:'규격'});
    const sName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S_C) || '일위대가목록';

    eachCell(S_C, (cell, A1, R, C)=>{
      const excelRow = R+1; if (excelRow<=Cdef.headerRow+1) return;
      const myKey = Cmap.get(excelRow); if (!myKey) return;
      const f = safeGet(cell,'f'); const refs = collectExternalRefs(f, sName);
      const onlyA = refs.filter(r => r.sheet.toLowerCase()==='일위대가');
      if (!onlyA.length) return; // 일위대가 미참조 행은 스킵
      const rep = mostFrequentRef(onlyA); if (!rep) return;

      // 참조된 행의 근처에서 헤더 역할 행을 위로 탐색
      const headerRowLike = nearestHeaderLikeUp(S_A, rep.row-1, Adef.colMap, 30);
      const refKey = headerRowLike!=null ? Amap.get(headerRowLike+1) : (Amap.get(rep.row) || '');
      const status = (!refKey || normWS(refKey)!==normWS(myKey))? '불일치':'일치';
      out.push({시트:'일위대가목록', 행:excelRow, 키:myKey, 참조시트:'일위대가', 참조행:rep.row, 참조키:refKey||'', 결과:status});
    });

    return summarize('B', out);
  }

  // ------------------------------
  // 검사 C: 공종별내역서 — 외부참조/직접입력 판정
  // - 외부참조 대표(최빈) ↔ 키 비교
  // - 외부참조 전혀 없고 합계단가가 상수(수식 아님) & 0이 아니면 불일치
  // ------------------------------
  function checkC(wb){
    const S = getSheetCaseInsensitive(wb,'공종별내역서');
    const T = getSheetCaseInsensitive(wb,'단가대비표'); // 비교 대상 기본
    const out = [];
    if (!S || !T) return {name:'C', rows:[], summary:{note:'필수 시트 미존재'}};

    const wantsS = ['품명','규격','단위','합계단가'];
    const wantsT = ['품명','규격','단위'];
    const Sdef = findHeaderRowAndCols(S, wantsS);
    const Tdef = findHeaderRowAndCols(T, wantsT);
    if (!Sdef || !Tdef) return {name:'C', rows:[], summary:{note:'헤더 미검출'}};

    const Smap = buildRowKeyMap(S, Sdef.headerRow, Sdef.colMap, {left:'품명', right:'규격'});
    const Tmap = buildRowKeyMap(T, Tdef.headerRow, Tdef.colMap, {left:'품명', right:'규격'});
    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S) || '공종별내역서';

    eachCell(S, (cell, A1, R, C)=>{
      const excelRow = R+1; if (excelRow<=Sdef.headerRow+1) return;
      const myKey = Smap.get(excelRow); if (!myKey) return;
      const f = safeGet(cell,'f'); const refs = collectExternalRefs(f, selfName);
      const rep = mostFrequentRef(refs);
      let status = '일치', refKey='', refSheet='', refRow='';
      if (rep){
        refSheet = rep.sheet; refRow = rep.row;
        const tKey = (rep.sheet.toLowerCase()==='단가대비표') ? Tmap.get(rep.row) : '';
        refKey = tKey||'';
        if (!tKey || normWS(tKey)!==normWS(myKey)) status='불일치';
      } else {
        // 외부참조 없음 → 합계단가가 값(수식 아님)이고 0이 아니면 불일치
        const sumCol = Sdef.colMap['합계단가'];
        const sumA1 = XLSX.utils.encode_cell({r:R, c:sumCol});
        const scell = S[sumA1];
        const sF = safeGet(scell,'f');
        const sV = Number(safeGet(scell,'v'));
        if (!sF && Number.isFinite(sV) && Math.abs(sV)>0){ status='불일치'; }
      }
      out.push({시트:'공종별내역서', 행:excelRow, 키:myKey, 참조시트:refSheet, 참조행:refRow, 참조키:refKey||'', 결과:status});
    });

    return summarize('C', out);
  }

  // ------------------------------
  // 검사 D: 공종별집계표 — 재/노/경 단가 외부참조 대표 ↔ 품명만 느슨 비교
  // ------------------------------
  function checkD(wb){
    const S = getSheetCaseInsensitive(wb,'공종별집계표');
    const T = getSheetCaseInsensitive(wb,'단가대비표');
    const out = [];
    if (!S || !T) return {name:'D', rows:[], summary:{note:'필수 시트 미존재'}};

    const wantsS = ['품명','재료비단가','노무비단가','경비단가'];
    const wantsT = ['품명','규격'];
    const Sdef = findHeaderRowAndCols(S, wantsS);
    const Tdef = findHeaderRowAndCols(T, wantsT);
    if (!Sdef || !Tdef) return {name:'D', rows:[], summary:{note:'헤더 미검출'}};

    const Tmap = buildRowKeyMap(T, Tdef.headerRow, Tdef.colMap, {left:'품명', right:'규격'});
    const selfName = wb.SheetNames.find(n => getSheetCaseInsensitive(wb,n)===S) || '공종별집계표';

    // S 시트를 훑으며 재/노/경 단가 셀들의 외부참조 대표를 모으고 품명만 비교
    const nameCol = Sdef.colMap['품명'];
    const getNameAt = (r)=>{ const a1=XLSX.utils.encode_cell({r, c:nameCol}); return normWS(safeGet(S[a1],'v')); };

    const targetNameFromKey = k => normWS(String(k||'').split('|')[0]);

    eachCell(S, (cell, A1, R, C)=>{
      const excelRow = R+1; if (excelRow<=Sdef.headerRow+1) return;
      if (C!==Sdef.colMap['재료비단가'] && C!==Sdef.colMap['노무비단가'] && C!==Sdef.colMap['경비단가']) return;
      const myName = getNameAt(R);
      if (!myName) return;
      const f = safeGet(cell,'f'); const refs = collectExternalRefs(f, selfName); const rep=mostFrequentRef(refs);
      let status='일치', refKey='', refSheet='', refRow='';
      if (rep){
        refSheet = rep.sheet; refRow = rep.row;
        const tKey = (rep.sheet.toLowerCase()==='단가대비표') ? Tmap.get(rep.row) : '';
        refKey = tKey||'';
        const tName = targetNameFromKey(tKey);
        if (!tName || tName !== myName) status='불일치';
      }
      out.push({시트:'공종별집계표', 행:excelRow, 품명:myName, 단가열:(C===Sdef.colMap['재료비단가']?'재료비':C===Sdef.colMap['노무비단가']?'노무비':'경비'), 참조시트:refSheet, 참조행:refRow, 참조키:refKey||'', 결과:status});
    });

    return summarize('D', out);
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

      log('검사 A 실행…');
      const ra = checkA(wb); log(`A: 전체 ${ra.summary?.전체||0}, 불일치 ${ra.summary?.불일치||0}`);
      log('검사 B 실행…');
      const rb = checkB(wb); log(`B: 전체 ${rb.summary?.전체||0}, 불일치 ${rb.summary?.불일치||0}`);
      log('검사 C 실행…');
      const rc = checkC(wb); log(`C: 전체 ${rc.summary?.전체||0}, 불일치 ${rc.summary?.불일치||0}`);
      log('검사 D 실행…');
      const rd = checkD(wb); log(`D: 전체 ${rd.summary?.전체||0}, 불일치 ${rd.summary?.불일치||0}`);
      log('검사 E 실행…');
      const re = checkE(wb); log(`E: 전체 ${re.summary?.전체||0}, 불일치 ${re.summary?.불일치||0}`);

      const resWb = buildResultWorkbook([ra,rb,rc,rd,re]);
      const outName = downloadWorkbook(resWb, file.name.replace(/\.[^.]+$/, '') + '_matchflag');
      log(`저장 완료: ${outName}`);
    } catch(e){
      console.error(e);
      log('오류: ' + (e.message || e));
    }
  };
})();
