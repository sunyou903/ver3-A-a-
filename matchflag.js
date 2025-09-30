/* matchflag.js v1.2 — A검사: 파이썬 로직과 동등화
   - 로그: 요약만 출력
   - 출력: 전체 / 불일치 / 일치 3종 XLSX
   - 핵심 로직:
     * '일위대가' 각 행의 수식에서 '단가대비표' 또는 '일위대가목록' 참조 추출
     * 참조된 셀의 '행번호(rr)'를 사용하여 해당 시트의 (품명|규격) 키를 가져와 현재 행 키와 비교
     * 참조가 전혀 없으면 '수량' 값 직입 여부 판단 → 불일치(단, % 포함 시 ‘제외’)
*/

(function () {
  'use strict';

  // ===== 로그 =====
  const log = (m) => {
    const el = document.getElementById('mfLog');
    if (el) el.textContent += (el.textContent ? '\n' : '') + String(m);
  };
  window.onerror = function (msg, url, line, col, error) {
    log(`ERROR: ${msg} @ ${url}:${line}:${col}`);
    return false;
  };

  // ===== XLSX 준비 =====
  function waitForXLSX(timeoutMs = 5000) {
    return new Promise((resolve, reject) => {
      const t0 = Date.now();
      (function loop() {
        if (window.XLSX) return resolve();
        if (Date.now() - t0 > timeoutMs) return reject(new Error('XLSX not loaded'));
        setTimeout(loop, 100);
      })();
    });
  }
  async function readWorkbook(file) {
    const buf = await file.arrayBuffer();
    return XLSX.read(new Uint8Array(buf), { type: 'array', cellFormula: true, cellNF: true, cellText: true });
  }
  function sheetToAOA(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error(`시트 없음: ${name}`);
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  // ===== 정규화 =====
  const normSimple = (s) => (s == null ? null : String(s).trim());
  function normCommasWs(s) {
    if (s == null) return '';
    let t = String(s).replace(/,/g, ' ');
    t = t.replace(/\s+/g, ' ').trim();
    return t;
  }
  function normKey(key) {
    if (key == null) return '';
    const s = String(key);
    if (s.includes('|')) {
      const [a, b] = s.split('|', 1 + 1);
      return `${normCommasWs(a)}|${normCommasWs(b)}`;
    }
    return normCommasWs(s);
  }

  // ===== 헤더 탐색 (필수 라벨 모두 존재하는 행 찾기; 라벨은 공백 제거 비교) =====
  function normLabel(s) {
    if (s == null) return null;
    return String(s).replace(/[\u3000 ]/g, '').trim();
  }
  function findHeaderRowAndColsRequired(arr, required, scanRows = 40, scanCols = 200) {
    const req = new Set(required.map(normLabel));
    for (let r = 0; r < Math.min(scanRows, arr.length); r++) {
      const found = new Map();
      const row = arr[r] || [];
      for (let c = 0; c < Math.min(scanCols, row.length); c++) {
        const nv = normLabel(row[c]);
        for (const want of req) {
          if (nv === want && !found.has(want)) found.set(want, c);
        }
      }
      if (found.size === req.size) {
        const pos = {};
        // map back to original labels
        for (const lab of required) {
          const k = normLabel(lab);
          pos[lab] = found.get(k);
        }
        return { headerRow: r, pos };
      }
    }
    throw new Error(`헤더 라벨 탐색 실패: ${required.join(',')}`);
  }

  // ===== 키맵 (1-based 행번호 -> "품명|규격") =====
  function buildKeyMap(arr, headerRow, colName, colSpec) {
    const map = {};
    for (let r = headerRow + 1; r < arr.length; r++) {
      const name = normSimple(arr[r]?.[colName]);
      const spec = normSimple(arr[r]?.[colSpec]);
      if ((name == null || name === '') && (spec == null || spec === '')) continue;
      const rr = r + 1; // 1-based
      map[rr] = `${name ?? ''}|${spec ?? ''}`;
    }
    return map;
  }

  // ===== 합계/소계 등 제외 규칙 (원본 A와 동일) =====
  function isSumRow(nameCell) {
    if (!nameCell) return false;
    const t = String(nameCell);
    return t.includes('합') && t.includes('계') && t.includes('[') && t.includes(']');
  }

  // ===== 수식 참조 추출:  '단가대비표' / '일위대가목록' ! $A$123 =====
  const REF_RE = /(?:'?)((?:단가대비표)|(?:일위대가목록))(?:'?)!\$?([A-Z]{1,3})\$?(\d+)/g;

  // ===== 객체 배열→AOA =====
  function objectsToAOA(objs) {
    if (!objs.length) return [['결과 없음']];
    const headers = Object.keys(objs[0]);
    const aoa = [headers];
    for (const o of objs) aoa.push(headers.map(h => o[h]));
    return aoa;
  }

  // ===== A검사 =====
  function runCheckA(wb) {
    const srcName = pickSheetByName(wb.SheetNames, '일위대가');
    const upName  = pickSheetByName(wb.SheetNames, '단가대비표');
    const lsName  = pickSheetByName(wb.SheetNames, '일위대가목록');

    const ulArr = sheetToAOA(wb, srcName);
    const upArr = sheetToAOA(wb, upName);
    const lsArr = sheetToAOA(wb, lsName);

    const ulHdr = findHeaderRowAndColsRequired(ulArr, ['품명','규격','단위','수량']);
    const upHdr = findHeaderRowAndColsRequired(upArr, ['품명','규격','단위']);
    const lsHdr = findHeaderRowAndColsRequired(lsArr, ['품명','규격']);

    const ulPos = ulHdr.pos, upPos = upHdr.pos, lsPos = lsHdr.pos;

    const ulKey = buildKeyMap(ulArr, ulHdr.headerRow, ulPos['품명'], ulPos['규격']);
    const upKey = buildKeyMap(upArr, upHdr.headerRow, upPos['품명'], upPos['규격']);
    const lsKey = buildKeyMap(lsArr, lsHdr.headerRow, lsPos['품명'], lsPos['규격']);

    const ws = wb.Sheets[srcName];
    const rng = XLSX.utils.decode_range(ws['!ref']);

    const records = [];
    let checked = 0;

    for (let r = ulHdr.headerRow + 1; r <= rng.e.r; r++) {
      const r1 = r + 1; // 1-based
      const cur_key = ulKey[r1];
      const nameCell = ulArr[r]?.[ulPos['품명']];
      if (!cur_key || isSumRow(nameCell)) continue;

      const specCell = ulArr[r]?.[ulPos['규격']];
      const pname_cur = normSimple(nameCell);
      const gname_cur = normSimple(specCell);

      let rowHasRef = false;

      // --- 이 행의 모든 셀에서 수식 검사 ---
      for (let c = rng.s.c; c <= rng.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        const f = cell && typeof cell.f === 'string' ? cell.f : null;
        if (!f) continue;
        if (!f.includes('단가대비표') && !f.includes('일위대가목록')) continue;

        let m;
        REF_RE.lastIndex = 0;
        while ((m = REF_RE.exec(f)) !== null) {
          const sheet_name = m[1];
          const colLetters = m[2];
          const rr = parseInt(m[3], 10); // 참조된 '행번호'

          const ref_key = sheet_name.startsWith('단가대비표') ? upKey[rr] : lsKey[rr];
          if (!ref_key) continue;

          checked += 1;
          let status = (normKey(ref_key) === normKey(cur_key)) ? '일치' : '불일치';

          // % 포함 시 불일치 → 제외
          try {
            if (status === '불일치' && ((pname_cur && String(pname_cur).includes('%')) || (gname_cur && String(gname_cur).includes('%')))) {
              status = '제외';
            }
          } catch (_) {}

          const shortF = f.length > 140 ? f.slice(0, 140) + '...' : f;

          records.push({
            "일위대가_행": r1,
            "일위대가_품명|규격": cur_key,
            "참조시트": sheet_name,
            "참조셀": `${sheet_name}!${colLetters}${rr}`,
            "참조_품명|규격": ref_key,
            "수식_셀": addr,
            "수식_일부": shortF,
            "일치여부": status
          });
          rowHasRef = true;
        }
      }

      // --- 참조가 전혀 없으면: '수량' 값 직접입력 여부 검사 ---
      if (!rowHasRef) {
        const qtyCol = ulPos['수량'];
        const val = ulArr[r]?.[qtyCol];
        if (!(val == null || val === '' || val === 0)) {
          let status_di = '불일치';
          try {
            if ((pname_cur && String(pname_cur).includes('%')) || (gname_cur && String(gname_cur).includes('%'))) {
              status_di = '제외';
            }
          } catch (_) {}

          records.push({
            "일위대가_행": r1,
            "일위대가_품명|규격": cur_key,
            "참조시트": "",
            "참조셀": "",
            "참조_품명|규격": "",
            "수식_셀": XLSX.utils.encode_cell({ r, c: qtyCol }),
            "수식_일부": String(val),
            "일치여부": status_di
          });
        }
      }
    }

    const total = records.length;
    const ok = records.filter(x => x.일치여부 === '일치').length;
    const bad = records.filter(x => x.일치여부 === '불일치').length;

    return {
      summary: { "A_검사한_참조": checked, "A_일치": ok, "A_불일치": bad },
      details: records
    };
  }

  // ===== 시트명 선택 (정확/유사) =====
  function pickSheetByName(names, target) {
    if (names.includes(target)) return target;
    const t = String(target).replace(/[\u3000 ]/g, '');
    for (const n of names) {
      if (String(n).replace(/[\u3000 ]/g, '').includes(t)) return n;
    }
    return target; // 못 찾으면 그대로(나중에 시트 없음 에러)
  }
   
// ✅ 교체할 내용 (objectsToAOA 아래의 saveAsOneWorkbook 내용 전체를 아래로 교체 - B패치)
   // 기존: function saveAsOneWorkbook(baseName, A_details, B_result)
   // 변경: C_result까지 받기
   function saveAsOneWorkbook(baseName, A_details, B_result, C_result) {
     const passA = A_details.filter(r => r.일치여부 === '일치');
     const failA = A_details.filter(r => r.일치여부 === '불일치');
   
     const wb = XLSX.utils.book_new();
   
     // A
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(A_details)), 'A_전체');
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(failA)),  'A_불일치');
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(passA)),  'A_일치');
   
     // B
     if (B_result && B_result.map) {
       XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(B_result.map)), 'B_목록_전체');
       XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(B_result.mis)), 'B_목록_불일치');
     }
   
     // C (전체/불일치/일치)
     if (C_result && C_result.details) {
       const allC = C_result.details;
       const badC = allC.filter(r => r.일치여부 === '불일치');
       const okC  = allC.filter(r => r.일치여부 === '일치');
       XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(allC)), 'C_공종_전체');
       XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(badC)), 'C_공종_불일치');
       XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(okC)),  'C_공종_일치');
     }
   
     XLSX.writeFile(wb, `${baseName}_일위대가 검사.xlsx`);
   }


// ===== B: 일위대가목록 ↔ 일위대가 매핑 =====
// 참고: 파이썬 run_check_B 동등화 (일위대가 참조 스캔 → 최근접 헤더행 키 매핑)  :contentReference[oaicite:1]{index=1}
function runCheckB(wb) {
  // 시트명 정규화 선택
  const ulName = pickSheetByName(wb.SheetNames, '일위대가');
  const lsName = pickSheetByName(wb.SheetNames, '일위대가목록');

  const ulArr = sheetToAOA(wb, ulName);
  const lsArr = sheetToAOA(wb, lsName);

  // 일위대가: 필수 헤더 풀셋 (파이썬과 동일)
  const ulHdr = findHeaderRowAndColsRequired(
    ulArr,
    ['품명','규격','단위','수량','합계 단가','합계금액','재료비 단가','재료비 금액','노무비 단가','노무비 금액','경비 단가','경비 금액','비고'],
    40, 200
  );
  const COL = {
    '품명': ulHdr.pos['품명'],
    '규격': ulHdr.pos['규격'],
    '단위': ulHdr.pos['단위'],
    '수량': ulHdr.pos['수량']
  };

  // 목록: 필수 헤더
  const lsHdr = findHeaderRowAndColsRequired(lsArr, ['코드','품명','규격'], 40, 200);

  const wsUL = wb.Sheets[ulName];
  const wsLS = wb.Sheets[lsName];
  const ulRange = XLSX.utils.decode_range(wsUL['!ref']);
  const lsRange = XLSX.utils.decode_range(wsLS['!ref']);

  // 목록 키
  const key_lst = (r) => {
    const name = normSimple(lsArr[r]?.[lsHdr.pos['품명']]);
    const spec = normSimple(lsArr[r]?.[lsHdr.pos['규격']]);
    return `${name}|${spec}`.replace(/\|$/, '');
  };

  // "최근접 헤더" 정의 (파이썬과 동일한 규칙)
  function header_row_nearest(rr) {
    for (let k = rr - 1; k > ulHdr.headerRow; k--) {
      const pname = String(ulArr[k]?.[COL['품명']] ?? '').trim();
      if (!pname) continue;
      const unit = ulArr[k]?.[COL['단위']];
      const qty  = ulArr[k]?.[COL['수량']];
      const blank = new Set([null, '', 0, '-', '—']);
      if (!blank.has(unit) || !blank.has(qty)) continue;
      // '합계' 포함 행은 제외
      const s = pname.replace(/[\u3000 ]/g, '');
      if (s.includes('합계')) continue;
      return k; // AOA 인덱스(0-based)
    }
    return null;
  }

  // 헤더 키 생성: 규격 없으면 '두 칸 이상 공백'으로 분리하여 앞 2토큰
  function build_ul_header_key(r) {
    if (r == null) return null;
    let pname = normSimple(ulArr[r]?.[COL['품명']]);
    let spec  = normSimple(ulArr[r]?.[COL['규격']]);
    if (!spec && pname) {
      const toks = (pname || '').split(/[ \u3000]{2,}/).map(t => t.trim()).filter(Boolean);
      if (toks.length >= 2) { pname = toks[0]; spec = toks[1]; }
    }
    const key = `${pname ?? ''}|${spec ?? ''}`.replace(/\|$/, '');
    return key || null;
  }

  // '일위대가' 참조 정규식 (파이썬과 동일 의미)
  const UL_CELLREF_RE = /(?:'?)일위대가(?:'?)!\$?([A-Z]{1,3})\$?(\d+)/gi;

  const mappings = [];
  const mismatches = [];
  let checked = 0;

  for (let r = lsHdr.headerRow + 1; r <= lsRange.e.r; r++) {
    const list_key = key_lst(r);
    if (!list_key || list_key === 'null|null' || list_key === '|') continue;

    const headerRows = new Set();
    let fcount = 0;

    // 목록 행 전체 수식 스캔
    for (let c = lsRange.s.c; c <= lsRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = wsLS[addr];
      const fml  = cell && typeof cell.f === 'string' ? cell.f : null;
      if (!fml || !/일위대가/.test(fml)) continue;

      let m; UL_CELLREF_RE.lastIndex = 0;
      while ((m = UL_CELLREF_RE.exec(fml)) !== null) {
        const rr = parseInt(m[2], 10); // 참조된 '일위대가' 행번호(1-based)
        // AOA 인덱스로 변환
        const rrIdx = rr - 1;
        const hdr = header_row_nearest(rrIdx);
        if (hdr != null) headerRows.add(hdr);
        checked++;
      }
      fcount++;
    }

    const hdr_row = headerRows.size ? [...headerRows].sort((a,b)=>a-b)[0] : null;
    const hdr_key = build_ul_header_key(hdr_row);
   const matchStatus = (!hdr_key)
     ? "참조없음"
     : (normKey(list_key) === normKey(hdr_key) ? "일치" : "불일치");
   
   mappings.push({
     "일위대가목록_행": r + 1,
     "일위대가목록_품명|규격": list_key,
     "매핑_헤더행": hdr_row != null ? (hdr_row + 1) : null,
     "매핑_헤더_품명|규격": hdr_key,
     "참조셀_수": fcount,
     "일치여부": matchStatus
   });


     if (matchStatus === "불일치") {
        mismatches.push({
          "일위대가목록_행": r + 1,
          "일위대가목록_품명|규격": list_key,
          "매핑_헤더행": hdr_row != null ? (hdr_row + 1) : null,
          "매핑_헤더_품명|규격": hdr_key
        });
      }

  }

  const summary = {
    "B_참조셀": checked,
    "B_매핑된_행": mappings.length,
    "B_불일치": mismatches.length
  };
  return { summary, map: mappings, mis: mismatches };
}



   // ===== C: 공종별내역서 — 외부시트 참조 비교 (원본 PY 동등화) =====
   // 참고: '공종별내역서' 각 행에서 외부 시트 참조를 모아 대표(최빈 (시트,행))의 키와 비교.
   // 외부참조가 없고 '합계 단가'가 수식이 아닌 숫자(0이 아님)이면 "값 직접입력" → 불일치. :contentReference[oaicite:1]{index=1}
   function runCheckC(wb) {
     const srcName = pickSheetByName(wb.SheetNames, '공종별내역서');
     const arr = sheetToAOA(wb, srcName);
     const ws = wb.Sheets[srcName];
   
     // 필수 헤더: 품명, 규격, 합계 단가  :contentReference[oaicite:2]{index=2}
     const hdr = findHeaderRowAndColsRequired(arr, ['품명','규격','합계 단가']);
     const HR = hdr.headerRow;
     const POS = hdr.pos;
   
     // 행 키: (품명|규격)
     const keyW = (r) => {
       const name = normSimple(arr[r]?.[POS['품명']]);
       const spec = normSimple(arr[r]?.[POS['규격']]);
       return `${name}|${spec}`;
     };
   
     // 범용 외부참조 파서: '시트명'!$A$1 또는 시트명!A1  :contentReference[oaicite:3]{index=3}
     const SHEET_REF_RE = /(?:'([^']+)'|([^'!:]+))!\$?([A-Z]{1,3})\$?(\d+)/gi;
   
     // 대상 시트에서 같은 행의 (품명|규격) 키 얻기  :contentReference[oaicite:4]{index=4}
     function getKeyFromSheet(sheetName, rownum) {
       if (!wb.Sheets[sheetName]) return null;
       const a2 = sheetToAOA(wb, sheetName);
       let hr, pos;
       try {
         const got = findHeaderRowAndColsRequired(a2, ['품명','규격']);
         hr = got.headerRow; pos = got.pos;
       } catch { return null; }
       const r = Number(rownum) - 1; // AOA index
       if (r <= hr || r >= a2.length) return null;
       const name = normSimple(a2[r]?.[pos['품명']]);
       const spec = normSimple(a2[r]?.[pos['규격']]);
       return `${name}|${spec}`;
     }
   
     const rng = XLSX.utils.decode_range(ws['!ref']);
     const records = [];
     let rowsWithRefs = 0;
     let cntDirectValue = 0;
   
     for (let r = HR + 1; r <= rng.e.r; r++) {
       const wkey = keyW(r);
       if (wkey === 'null|null' || wkey === 'None|None') continue;
   
       // 행 내 모든 셀의 수식에서 외부시트 참조 수집(자기 시트 제외)  :contentReference[oaicite:5]{index=5}
       const refs = [];
       for (let c = rng.s.c; c <= rng.e.c; c++) {
         const addr = XLSX.utils.encode_cell({ r, c });
         const cell = ws[addr];
         const fml = cell && typeof cell.f === 'string' ? cell.f : null;
         if (!fml) continue;
         let m;
         SHEET_REF_RE.lastIndex = 0;
         while ((m = SHEET_REF_RE.exec(fml)) !== null) {
           const qsheet = m[1], bsheet = m[2], row = m[4];
           let target = String(qsheet || bsheet || '').replace(/^\s*[=+]/,'').trim().replace(/^'|'+$/g,'');
           target = target.replace(/.*\(/,''); // TRUNC(시트!A1) 보호  (PY와 동등 아이디어) :contentReference[oaicite:6]{index=6}
           if (!target || target === srcName) continue;
           refs.push([target, parseInt(row,10)]);
         }
       }
   
       if (refs.length) {
         // 대표 참조: 최빈 (시트,행)  :contentReference[oaicite:7]{index=7}
         rowsWithRefs += 1;
         const freq = new Map();
         for (const [sn, rr] of refs) {
           const k = `${sn}|${rr}`;
           freq.set(k, (freq.get(k) || 0) + 1);
         }
         let best = null, bestCnt = -1;
         for (const [k, v] of freq) if (v > bestCnt) { best = k; bestCnt = v; }
         const [repSheet, repRowStr] = (best || '').split('|');
         const repRow = parseInt(repRowStr,10);
         const tkey = getKeyFromSheet(repSheet, repRow) || '';
         const match = (normKey(wkey) === normKey(tkey)) ? '일치' : '불일치';  :contentReference[oaicite:8]{index=8}
   
         records.push({
           "행": r+1,  // 1-based 표시
           "참조유형": `${repSheet} 참조`,
           "공종_키(품명|규격)": wkey,
           "참조_키(품명|규격)": tkey,
           "대표참조시트": repSheet,
           "대표참조행": repRow,
           "일치여부": match
         });
         continue;
       }
   
       // 외부참조 없음 → '합계 단가' 직접입력 검사 (수식 아님 + 숫자 + 0 아님) :contentReference[oaicite:9]{index=9}
       const priceCol = POS['합계 단가'];
       const addr = XLSX.utils.encode_cell({ r, c: priceCol });
       const cell = ws[addr];
       const isFormula = cell && typeof cell.f === 'string';
       const val = cell ? cell.v : null;
       const isNum = typeof val === 'number' && Number.isFinite(val);
       const nonzero = isNum && Math.abs(val) > 1e-9;
       if (!isFormula && nonzero) {
         records.push({
           "행": r+1, "참조유형": "값 직접입력",
           "공종_키(품명|규격)": wkey,
           "참조_키(품명|규격)": "",
           "대표참조시트": "", "대표참조행": "",
           "일치여부": "불일치",
           "입력값(합계단가)": val
         });
         cntDirectValue += 1;
       }
     }
   
     // 요약  :contentReference[oaicite:10]{index=10}
     const df = records; // JS에서는 객체배열
     const sum = {
       "C_검사대상_행수(직접참조 보유)": rowsWithRefs,
       "C_일치": df.filter(x => x.일치여부 === '일치').length,
       "C_불일치": df.filter(x => x.일치여부 === '불일치').length,
       "C_값직접입력_불일치": cntDirectValue
     };
     return { summary: sum, details: records };
   }

  // ===== 실행 =====
  async function run() {
    try {
      document.getElementById('mfLog').textContent = '';
      await waitForXLSX();

      const f = document.getElementById('mfFile').files[0];
      if (!f) throw new Error('엑셀 파일을 선택하세요.');

      const wb = await readWorkbook(f);
      // A, B 동시 실행
      const A = runCheckA(wb);
      const B = runCheckB(wb);
      const C = runCheckC(wb);
      
      // 로그 요약
      // 로그 요약
      log(`A: 참조 ${A.summary.A_검사한_참조}건, 일치 ${A.summary.A_일치}, 불일치 ${A.summary.A_불일치}`);
      log(`B: 참조셀 ${B.summary.B_참조셀}, 매핑된 행 ${B.summary.B_매핑된_행}, 불일치 ${B.summary.B_불일치}`);
      log(`C: 대상(직접참조 보유) ${C.summary["C_검사대상_행수(직접참조 보유)"]}, 일치 ${C.summary.C_일치}, 불일치 ${C.summary.C_불일치}, 값직접입력 불일치 ${C.summary["C_값직접입력_불일치"]}`);
      
      const base = f.name.replace(/\.[^.]+$/, '');
      saveAsOneWorkbook(base, A.details, B, C);
      log('엑셀 저장 완료: 한 파일에 A(3) + B(2) + C(3) 시트');
 
    } catch (e) {
      console.error(e);
      log(`ERROR: ${e.message || e}`);
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    const btn = document.getElementById('mfRun');
    if (btn) btn.addEventListener('click', run);
  });
})();




