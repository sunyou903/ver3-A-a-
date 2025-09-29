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

  // ===== 저장: 전체/불일치/일치 =====
  function saveAsOneWorkbook(baseName, allRows) {
    const passRows = allRows.filter(r => r.일치여부 === '일치');
    const failRows = allRows.filter(r => r.일치여부 === '불일치');
  
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(allRows)), 'A_전체');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(failRows)), 'A_불일치');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(passRows)), 'A_일치');
  
    const outName = `${baseName}_일위대가 검사.xlsx`;
    XLSX.writeFile(wb, outName);
    
  }

  // ===== 실행 =====
  async function run() {
    try {
      document.getElementById('mfLog').textContent = '';
      await waitForXLSX();

      const f = document.getElementById('mfFile').files[0];
      if (!f) throw new Error('엑셀 파일을 선택하세요.');

      const wb = await readWorkbook(f);
      const { summary, details } = runCheckA(wb);

      // 요약만 로그
      log(`일위대가 검사: 참조 ${summary.A_검사한_참조}건, 일치 ${summary.A_일치}, 불일치 ${summary.A_불일치}`);

      const base = f.name.replace(/\.[^.]+$/, '');
      saveAsOneWorkbook(base, details);
      log('엑셀 저장 완료: (한 파일, 시트 3장) A_전체 / A_불일치 / A_일치');
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
