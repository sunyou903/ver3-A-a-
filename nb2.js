// 전역 에러 잡기 → #log에 표시
window.onerror = function (msg, url, line, col, error) {
  const logEl = document.getElementById("log");
  const text = [
    "=== 전역 에러 감지 ===",
    "메시지: " + msg,
    "파일: " + url,
    "줄/컬럼: " + line + ":" + col,
    "오브젝트: " + (error && error.stack ? error.stack : error)
  ].join("\n");
  logEl.textContent += (logEl.textContent ? "\n" : "") + text;
  return false; // 콘솔에도 그대로 출력
};


// ===== Nb2_rev2 포팅 (브라우저) =====
// 원본 논리 근거: HEADER 탐색/핵심컬럼 추출/가중유사도/임계값 30% 등:contentReference[oaicite:6]{index=6}
const LABELS = {
  "품명": "품 명",
  "규격": "규 격",
  "단위": "단위",
  "재료비적용단가": "재료비 적용단가",
  "노무비": "노 무 비",
  "경비적용단가": "경비 적용단가",
};

function log(msg) {
  const el = document.getElementById("log");
  el.textContent += (el.textContent ? "\n" : "") + msg;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const wb = XLSX.read(data, { type: "array", cellFormula: true, cellNF: true, cellText: true });
        resolve(wb);
      } catch (e) { reject(e); }
    };
    reader.readAsArrayBuffer(file);
  });
}

function sheetToDF(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`시트 없음: ${sheetName}`);
  const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  return arr; // 2D array
}

// 헤더 후보 다중행 스캔 → 가장 많이 라벨이 잡힌 행을 헤더로 확정
function findHeaderRowAndCols(arr, scanRows = 6) {
  // arr: 2D array
  const targets = Object.values(LABELS).map(s => s.replace(/\s+/g, ""));
  const found = {}; // label -> list of [r,c]
  for (const lab of targets) found[lab] = [];
  const rows = Math.min(scanRows, arr.length);

  for (let r = 0; r < rows; r++) {
    const row = arr[r] || [];
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (typeof v !== "string") continue;
      const key = v.replace(/\s+/g, "");
      for (const tgt of targets) {
        if (key === tgt) found[tgt].push([r, c]);
      }
    }
  }
  const rowCount = new Map();
  for (const hits of Object.values(found)) {
    for (const [r] of hits) rowCount.set(r, (rowCount.get(r) || 0) + 1);
  }
  if (rowCount.size === 0) throw new Error("머리글을 찾지 못함");
  const headerRow = [...rowCount.entries()].sort((a,b)=>b[1]-a[1])[0][0];

  function pickCol(labelKorean) {
    const tgt = LABELS[labelKorean].replace(/\s+/g, "");
    const hits = (found[tgt] || []).filter(([r,_c]) => r === headerRow);
    return hits.length ? hits[0][1] : null;
  }

  const colMap = {
    "품명": pickCol("품명"),
    "규격": pickCol("규격"),
    "단위": pickCol("단위"),
    "재료비적용단가": pickCol("재료비적용단가"),
    "노무비": pickCol("노무비"),
    "경비적용단가": pickCol("경비적용단가"),
  };
  return { headerRow, colMap };
}

function extractCore(arr) {
  const { headerRow, colMap } = findHeaderRowAndCols(arr);
  const cols = Object.values(colMap);
  if (cols.some(c => c == null)) throw new Error("필수 컬럼 일부 누락");

  const out = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    const row = arr[r] || [];
    const rec = {
      품명: row[colMap["품명"]],
      규격: row[colMap["규격"]],
      단위: row[colMap["단위"]],
      재료비적용단가: toNum(row[colMap["재료비적용단가"]]),
      노무비: toNum(row[colMap["노무비"]]),
      경비적용단가: toNum(row[colMap["경비적용단가"]]),
    };
    if (rec.품명 == null || String(rec.품명).trim() === "") continue; // 품명 NaN 제거
    rec.규격 = rec.규격 == null ? "" : rec.규격;
    out.push(rec);
  }
  return out;
}

function toNum(x) {
  if (x == null || x === "") return null;
  const n = Number(String(x).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

// ==== 유사도 함수 (간단 토큰정렬 + 편집거리 기반) ====

// 토큰 정렬
function tokenSort(s) {
  return String(s ?? "").trim().split(/\s+/).sort().join(" ");
}

// 편집거리
function editDistance(a, b) {
  a = String(a); b = String(b);
  const dp = Array.from({length: a.length + 1}, ()=>Array(b.length + 1).fill(0));
  for (let i=0;i<=a.length;i++) dp[i][0]=i;
  for (let j=0;j<=b.length;j++) dp[0][j]=j;
  for (let i=1;i<=a.length;i++){
    for (let j=1;j<=b.length;j++){
      const cost = a[i-1]===b[j-1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i-1][j]+1,
        dp[i][j-1]+1,
        dp[i-1][j-1]+cost
      );
    }
  }
  return dp[a.length][b.length];
}

function tokenSortRatio(a, b) {
  const A = tokenSort(a), B = tokenSort(b);
  const dist = editDistance(A, B);
  const maxLen = Math.max(A.length, B.length) || 1;
  return (1 - dist / maxLen) * 100.0;
}

function specSim(a, b) {
  if ((a===""||a==null) && (b===""||b==null)) return 100.0;
  return tokenSortRatio(a ?? "", b ?? "");
}

function normName(s){ return String(s ?? "").trim(); }

// 매칭 (품명 완전일치 우선 → 가중유사도, 임계값 기본 30%):contentReference[oaicite:7]{index=7}
function matchAndCompare(left, right, th=30, leftPrefix="왼쪽", rightPrefix="오른쪽") {
  const rightIdx = right.map(r => ({...r, 품명_norm: normName(r.품명)}));
  const rows = [];

  for (const a of left) {
    const aName = normName(a.품명);
    const aSpec = a.규격;

    // 1) 품명 완전일치 후보
    const exact = rightIdx.filter(x => x.품명_norm === aName);
    let chosen=null, matchType=null, nameSim=null, specSimV=null, totalSim=null;

    if (exact.length>0) {
      let best=null, bestS=-1;
      for (const c of exact) {
        const s = specSim(aSpec, c.규격);
        if (s>bestS){ bestS=s; best=c; }
      }
      chosen=best; matchType="품명완전일치"; nameSim=100.0; specSimV=bestS;
      totalSim=0.8*nameSim + 0.2*specSimV;
    } else {
      // 2) 가중 유사도
      let best=null, bestScore=-1, bestName=-1, bestSpec=-1;
      for (const c of rightIdx) {
        const n = tokenSortRatio(aName, c.품명_norm);
        const s = specSim(aSpec, c.규격);
        const score = 0.8*n + 0.2*s;
        if (score>bestScore){ bestScore=score; best=c; bestName=n; bestSpec=s; }
      }
      if (bestScore>=th && best){
        chosen=best; matchType="가중유사도"; nameSim=bestName; specSimV=bestSpec; totalSim=bestScore;
      }
    }

    if (!chosen) continue;

    rows.push({
      [`${leftPrefix}_품명`]: a.품명,
      [`${leftPrefix}_규격`]: a.규격,
      [`${rightPrefix}_품명`]: chosen.품명,
      [`${rightPrefix}_규격`]: chosen.규격,
      "매칭유형": matchType,
      "종합유사도(%)": Math.round(totalSim*10)/10,
      "품명유사(%)": Math.round(nameSim*10)/10,
      "규격유사(%)": Math.round(specSimV*10)/10,
      [`${leftPrefix}_재료비적용단가`]: a.재료비적용단가,
      [`${rightPrefix}_재료비적용단가`]: chosen.재료비적용단가,
      [`${leftPrefix}_노무비`]: a.노무비,
      [`${rightPrefix}_노무비`]: chosen.노무비,
      [`${leftPrefix}_경비적용단가`]: a.경비적용단가,
      [`${rightPrefix}_경비적용단가`]: chosen.경비적용단가,
    });
  }
  return rows;
}

function aoaFromObjects(objs) {
  if (objs.length===0) return [["결과 없음"]];
  const headers = Object.keys(objs[0]);
  const aoa = [headers];
  for (const o of objs) aoa.push(headers.map(h => o[h]));
  return aoa;
}

async function run() {
  document.getElementById("log").textContent = "";
  try {
    const fa = document.getElementById("fileA").files[0];
    const fb = document.getElementById("fileB").files[0];
    if (!fa || !fb) throw new Error("두 파일을 모두 선택하세요.");

    const sheetName = (document.getElementById("sheetName").value || "단가대비표").trim();
    const th = Number(document.getElementById("threshold").value || "30");

    log("파일 로딩 중...");
    const [wa, wb] = await Promise.all([readWorkbook(fa), readWorkbook(fb)]);

    log(`시트 파싱: ${sheetName}`);
    const A = extractCore(sheetToDF(wa, sheetName));
    const B = extractCore(sheetToDF(wb, sheetName));
    log(`A행=${A.length}, B행=${B.length}`);

    // 라벨 자동 추출은 파일명 키워드로 간소화 (원본 detect_label 개념):contentReference[oaicite:8]{index=8}
    const leftLabel = detectLabel(fa.name) || "왼쪽";
    const rightLabel = detectLabel(fb.name) || "오른쪽";

    const rows = matchAndCompare(A, B, th, leftLabel, rightLabel);
    log(`매칭 결과 행=${rows.length}`);

    // 엑셀로 저장
    const aoa = aoaFromObjects(rows);
    // 맨 끝 열에 비교 수식은 엑셀에서만 의미가 있으므로 간단히 TRUE/FALSE 식 추가(원본은 J~N열 기준):contentReference[oaicite:9]{index=9}
    const headers = aoa[0];
    const compareColName = "비교결과";
    headers.push(compareColName);
    for (let r=1; r<aoa.length; r++){
      // 단순 동등 비교: 재료비/노무비/경비 세 항목이 좌우 동일하면 TRUE
      const rowObj = rows[r-1];
      const Lm = rowObj[`${leftLabel}_재료비적용단가`];
      const Rm = rowObj[`${rightLabel}_재료비적용단가`];
      const Ln = rowObj[`${leftLabel}_노무비`];
      const Rn = rowObj[`${rightLabel}_노무비`];
      const Le = rowObj[`${leftLabel}_경비적용단가`];
      const Re = rowObj[`${rightLabel}_경비적용단가`];
      aoa[r].push( (Lm===Rm && Ln===Rn && Le===Re) ? true : false );
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wbOut = XLSX.utils.book_new();
    const sheetOutName = `${leftLabel}_vs_${rightLabel}`.slice(0,31); // Excel 시트명 제한
    XLSX.utils.book_append_sheet(wbOut, ws, sheetOutName);
    const outName = `${fa.name.replace(/\.[^.]+$/,'')}_vs_${fb.name.replace(/\.[^.]+$/,'')}.xlsx`;
    XLSX.writeFile(wbOut, outName);

    log(`저장 완료: ${outName}`);
  } catch (e) {
    console.error(e);
    log(`ERROR: ${e.message || e}`);
  }
}

// 파일명 기반 간단 라벨 추출(원본 detect_label와 유사):contentReference[oaicite:10]{index=10}
function detectLabel(name) {
  const base = String(name).replace(/\s+/g, "");
  const KEYS = ["건축설비","토목","조경","건축","기계","전기"];
  for (const k of KEYS) if (base.includes(k)) return k;
  return null;
}

window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("runBtn").addEventListener("click", run);
});
