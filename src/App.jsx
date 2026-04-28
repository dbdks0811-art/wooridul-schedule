import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase, isSupabaseReady } from "./supabase.js";

/* ══════════════════════════════════════════════════════════
   CONSTANTS  — Light Theme
══════════════════════════════════════════════════════════ */
const SHIFT_TYPES = ["주","석","야","휴","월","반"];
const SHIFT = {
  주:{ bg:"#dbeafe", brd:"#3b82f6", txt:"#1d4ed8", label:"주간", hours:8 },
  석:{ bg:"#ffedd5", brd:"#f97316", txt:"#c2410c", label:"석간", hours:8 },
  야:{ bg:"#ede9fe", brd:"#8b5cf6", txt:"#5b21b6", label:"야간", hours:8 },
  휴:{ bg:"#f1f5f9", brd:"#94a3b8", txt:"#64748b", label:"휴무", hours:0 },
  월:{ bg:"#dcfce7", brd:"#22c55e", txt:"#15803d", label:"월차", hours:8 },
  반:{ bg:"#ecfccb", brd:"#84cc16", txt:"#4d7c0f", label:"반차", hours:4 },
};
const PAT_COLORS = ["#2563eb","#ea580c","#7c3aed","#16a34a","#be123c","#0891b2","#b45309","#4f46e5"];
const DOW = ["일","월","화","수","목","금","토"];
const getDIM = (y,m) => new Date(y,m,0).getDate();
const getDOW = (y,m,d) => DOW[new Date(y,m-1,d).getDay()];
const mkCell   = (type, locked=false) => ({type, locked});
const cellT    = c => typeof c==="string" ? c : c?.type??"휴";
const isLocked = c => typeof c==="object" && c?.locked===true;   /* 수동 수정 = locked */
const isManual = isLocked;   /* 하위 호환 별칭 */
const cellH    = c => SHIFT[cellT(c)]?.hours??0;
const calcHours = row => row.reduce((s,c)=>s+cellH(c),0);

let _puid=5, _duid=3;
const newPid = () => `p${_puid++}`;
const newDid = () => `d${_duid++}`;

const BUILTIN_PATTERNS = [
  { id:"p1", name:"기본 균형형",  stars:5, builtin:true,
    seq:["주","주","석","야","휴","휴","주","석","야","휴"] },
  { id:"p2", name:"야간 분산형",  stars:4, builtin:true,
    seq:["주","석","야","휴","주","석","야","휴","주","휴"] },
  { id:"p3", name:"주간 강화형",  stars:3, builtin:true,
    seq:["주","주","주","석","야","휴","휴","주","석","휴"] },
  { id:"p4", name:"휴식 강화형",  stars:4, builtin:true,
    seq:["주","석","야","휴","휴","주","석","야","휴","휴"] },
];
const mkEmp = (name,pidx=0,off=0) => ({name, patternId:BUILTIN_PATTERNS[pidx%4].id, offset:off});

const INIT_STATE = {
  patterns: BUILTIN_PATTERNS.map(p=>({...p})),
  depts:[{
    id:"dept1", name:"1병동",
    emps:["가","나","다","라","마","바","사","아","자","차"].map((n,i)=>mkEmp(n,0,i)),
    cst:{se:2,ya:2,juMin:2}, targetH:144,
  }],
};

/* ══════════════════════════════════════════════════════════
   PRINT CSS
══════════════════════════════════════════════════════════ */
const PRINT_CSS=`@media print{
  body{background:#fff!important;}
  .no-print{display:none!important;}
  .print-area{overflow:visible!important;}
  table{font-size:9px!important;border-collapse:collapse;}
  th,td{border:1px solid #bbb!important;padding:2px 4px!important;}
  @page{margin:8mm;size:A3 landscape;}
}`;

/* ══════════════════════════════════════════════════════════
   ENGINE  v7 — 6-Phase Strict Constraint Solver
   ──────────────────────────────────────────────
   PHASE 1 : 패턴으로 초기값 생성 (패턴 = 초기값일 뿐)
   PHASE 2 : 야간 후 휴무 1차 강제
   PHASE 3 : autoAdjust — 일별 하드 제약 완전 강제
             (야→석→주 순으로 승격, 과잉 시 강등)
   PHASE 4 : 개인 근무시간 맞추기 (hours 기준, juMin 유지)
   PHASE 5 : 야간 후 휴무 최종 재보장
   PHASE 6 : 시간 재균형 (Phase 5 보정 후)
   PHASE 7 : reAdjustDailyConstraints — 일별 제약 최종 재검증
══════════════════════════════════════════════════════════ */

/* ── Phase 7용 헬퍼: 일별 제약 재보정 (locked 셀 절대 불변) ── */
function reAdjustDailyConstraints(raw, D, cst, lck) {
  const N = raw.length;
  const { se: seR, ya: yaR, juMin } = cst;

  for (let d = 0; d < D; d++) {

    /* 야간 과잉 제거 */
    while (raw.filter(r => r[d] === "야").length > yaR) {
      let changed = false;
      for (let i = 0; i < N; i++) {
        if (lck[i][d]) continue;                          /* ★ locked 스킵 */
        if (raw[i][d] === "야" && raw.filter(r => r[d] === "야").length > yaR) {
          raw[i][d] = "주"; changed = true; break;
        }
      }
      if (!changed) break;
    }
    /* 야간 부족 보충 */
    while (raw.filter(r => r[d] === "야").length < yaR) {
      let ok = false;
      for (let i = 0; i < N && !ok; i++) {
        if (lck[i][d]) continue;                          /* ★ locked 스킵 */
        if (raw[i][d] !== "주") continue;
        const nextOk = d+1>=D || (!lck[i][d+1] && raw[i][d+1]!=="야" && raw[i][d+1]!=="석");
        if (nextOk) {
          raw[i][d]="야";
          if (d+1<D && !lck[i][d+1]) raw[i][d+1]="휴";
          ok=true;
        }
      }
      if (!ok) break;
    }
    /* 석간 과잉 제거 */
    while (raw.filter(r => r[d] === "석").length > seR) {
      let changed = false;
      for (let i = 0; i < N; i++) {
        if (lck[i][d]) continue;                          /* ★ locked 스킵 */
        if (raw[i][d] === "석" && raw.filter(r => r[d] === "석").length > seR) {
          raw[i][d] = "주"; changed = true; break;
        }
      }
      if (!changed) break;
    }
    /* 석간 부족 보충 */
    while (raw.filter(r => r[d] === "석").length < seR) {
      let ok = false;
      for (let i = 0; i < N && !ok; i++) {
        if (lck[i][d]) continue;                          /* ★ locked 스킵 */
        if (raw[i][d] === "주") { raw[i][d] = "석"; ok = true; }
      }
      if (!ok) break;
    }
    /* 주간 최소 인원 확보 */
    while (raw.filter(r => r[d] === "주").length < juMin) {
      let ok = false;
      for (let i = 0; i < N && !ok; i++) {
        if (lck[i][d]) continue;                          /* ★ locked 스킵 */
        if (raw[i][d] === "휴" && (d===0 || raw[i][d-1]!=="야")) {
          raw[i][d] = "주"; ok = true;
        }
      }
      if (!ok) break;
    }
  }
}

function generateSchedule(emps, patterns, D, cst, targetH, prevSch=null) {
  const N = emps.length;
  const { se: seR, ya: yaR, juMin } = cst;
  const targetDays = Math.floor(targetH / 8);

  /* ── 사전 feasibility 검사 ── */
  const minRequired = D * (seR + yaR + juMin);
  const totalAvail  = N * targetDays;
  if (totalAvail < minRequired)
    return { err: `인원 부족 — 필요 ${minRequired}일, 가용 ${totalAvail}일 (${N}명×${targetDays}일)` };

  const pMap = Object.fromEntries(patterns.map(p => [p.id, p]));

  /* ─────────────────────────────────────────────
     PHASE 1 : 패턴 기반 초기값
  ───────────────────────────────────────────── */
  const raw = Array.from({ length: N }, () => Array(D).fill("휴"));
  for (let i = 0; i < N; i++) {
    const pat = pMap[emps[i].patternId] || patterns[0];
    const seq = pat.seq, L = seq.length, off = emps[i].offset % L;
    for (let d = 0; d < D; d++) raw[i][d] = seq[(d + off) % L];
  }

  /* ─────────────────────────────────────────────
     PHASE 2 : 야간 후 휴무 1차 강제
  ───────────────────────────────────────────── */
  for (let i = 0; i < N; i++)
    for (let d = 0; d < D - 1; d++)
      if (raw[i][d] === "야") raw[i][d + 1] = "휴";

  /* ─────────────────────────────────────────────
     LOCKED OVERLAY : prevSch의 locked 셀을 raw에 반영
     → generate를 다시 눌러도 수동 수정값 보존
  ───────────────────────────────────────────── */
  const lck = Array.from({ length: N }, () => Array(D).fill(false));
  if (prevSch) {
    for (let i = 0; i < N && i < prevSch.length; i++) {
      for (let d = 0; d < D && d < (prevSch[i]?.length ?? 0); d++) {
        if (isLocked(prevSch[i][d])) {
          raw[i][d] = cellT(prevSch[i][d]);   /* locked 값을 raw에 덮어씀 */
          lck[i][d] = true;                   /* 이후 모든 Phase에서 스킵 */
        }
      }
    }
    /* locked 야간 다음날도 휴 처리 (단, 다음날이 locked이면 건드리지 않음) */
    for (let i = 0; i < N; i++)
      for (let d = 0; d < D - 1; d++)
        if (lck[i][d] && raw[i][d] === "야" && !lck[i][d + 1])
          raw[i][d + 1] = "휴";
  }

  /* ─────────────────────────────────────────────
     PHASE 3 : autoAdjust — 일별 하드 제약 완전 강제
     "근무 부족은 절대 허용하지 않는다"
  ───────────────────────────────────────────── */

  /* ─────────────────────────────────────────────────────────
     부하 정렬 헬퍼 — 두 방향
     · 승격(야/석/주 추가) 시 → 오름차순 (가장 덜 일한 사람 우선, 균등 분배)
     · 강등(야/석 제거) 시   → 내림차순 (가장 많이 일한 사람 우선, 균등 분배)
  ──────────────────────────────────────────────────────────*/
  const sortedAsc  = () =>   /* 승격용: 일 적은 사람 먼저 */
    Array.from({length:N},(_,i)=>i)
      .sort((a,b)=> raw[a].filter(s=>s!=="휴").length - raw[b].filter(s=>s!=="휴").length);
  const sortedDesc = () =>   /* 강등용: 일 많은 사람 먼저 */
    Array.from({length:N},(_,i)=>i)
      .sort((a,b)=> raw[b].filter(s=>s!=="휴").length - raw[a].filter(s=>s!=="휴").length);

  for (let d = 0; d < D; d++) {

    /* ── A. 야간 과잉 제거 (야 → 주) ── */
    while (raw.filter(r => r[d] === "야").length > yaR) {
      const order = sortedDesc();
      let changed = false;
      for (let ii = 0; ii < N; ii++) {
        const i = order[ii];
        if (lck[i][d]) continue;                              /* ★ locked 스킵 */
        if (raw[i][d] === "야" && raw.filter(r => r[d] === "야").length > yaR) {
          raw[i][d] = "주"; changed = true; break;
        }
      }
      if (!changed) break;
    }

    /* ── B. 야간 부족 보충 ── */
    while (raw.filter(r => r[d] === "야").length < yaR) {
      let promoted = false;
      const order = sortedAsc();

      /* 우선순위 1: 주간 근무자 */
      for (let ii = 0; ii < N && !promoted; ii++) {
        const i = order[ii];
        if (lck[i][d]) continue;                              /* ★ locked 스킵 */
        if (raw[i][d] !== "주") continue;
        const nextOk = d+1>=D ||
          (!lck[i][d+1] && raw[i][d+1]!=="야" && raw[i][d+1]!=="석");
        if (nextOk) {
          raw[i][d] = "야";
          if (d+1<D && !lck[i][d+1]) raw[i][d+1] = "휴";
          promoted = true;
        }
      }
      /* 우선순위 2: 휴무 (야간 후 쉬는 중 + locked 제외) */
      if (!promoted) {
        for (let ii = 0; ii < N && !promoted; ii++) {
          const i = order[ii];
          if (lck[i][d]) continue;                            /* ★ locked 스킵 */
          if (raw[i][d] !== "휴") continue;
          if (d > 0 && raw[i][d-1] === "야") continue;
          const nextOk = d+1>=D ||
            (!lck[i][d+1] && raw[i][d+1]!=="야" && raw[i][d+1]!=="석");
          if (nextOk) {
            raw[i][d] = "야";
            if (d+1<D && !lck[i][d+1]) raw[i][d+1] = "휴";
            promoted = true;
          }
        }
      }
      /* 우선순위 3: 강제 */
      if (!promoted) {
        for (let i = 0; i < N && !promoted; i++) {
          if (lck[i][d]) continue;                            /* ★ locked 스킵 */
          if (raw[i][d] !== "야" && raw[i][d] !== "석") {
            raw[i][d] = "야";
            if (d+1<D && !lck[i][d+1]) raw[i][d+1] = "휴";
            promoted = true;
          }
        }
      }
      if (!promoted) return { err: `${d+1}일 야간 배정 불가 (locked 포함 인원 ${N}명 부족)` };
    }

    /* ── C. 석간 과잉 제거 (석 → 주) ── */
    while (raw.filter(r => r[d] === "석").length > seR) {
      const order = sortedDesc();
      let changed = false;
      for (let ii = 0; ii < N; ii++) {
        const i = order[ii];
        if (lck[i][d]) continue;                              /* ★ locked 스킵 */
        if (raw[i][d] === "석" && raw.filter(r => r[d] === "석").length > seR) {
          raw[i][d] = "주"; changed = true; break;
        }
      }
      if (!changed) break;
    }

    /* ── D. 석간 부족 보충 ── */
    while (raw.filter(r => r[d] === "석").length < seR) {
      let promoted = false;
      const order = sortedAsc();
      for (let ii = 0; ii < N && !promoted; ii++) {
        const i = order[ii];
        if (lck[i][d]) continue;                              /* ★ locked 스킵 */
        if (raw[i][d] === "주") { raw[i][d] = "석"; promoted = true; }
      }
      if (!promoted) {
        for (let ii = 0; ii < N && !promoted; ii++) {
          const i = order[ii];
          if (lck[i][d]) continue;                            /* ★ locked 스킵 */
          if (raw[i][d]==="휴" && (d===0||raw[i][d-1]!=="야")) {
            raw[i][d]="석"; promoted=true;
          }
        }
      }
      if (!promoted) return { err: `${d+1}일 석간 배정 불가` };
    }

    /* ── E. 주간 최소 인원 확보 ── */
    while (raw.filter(r => r[d] === "주").length < juMin) {
      let promoted = false;
      const order = sortedAsc();
      for (let ii = 0; ii < N && !promoted; ii++) {
        const i = order[ii];
        if (lck[i][d]) continue;                              /* ★ locked 스킵 */
        if (raw[i][d]==="휴" && (d===0||raw[i][d-1]!=="야")) {
          raw[i][d]="주"; promoted=true;
        }
      }
      if (!promoted) break;
    }
  }

  /* ─────────────────────────────────────────────
     PHASE 3.5 : 주간 균등 분배
     목표: 매일 주간 인원을 juMin+1 수준으로 고르게 유지
     방식: 주간 횟수가 가장 적은 직원에게 우선 배정
     제약: targetH 초과 없음, locked 불변, 야간후휴 불변
  ───────────────────────────────────────────── */
  {
    const targetJuPerDay = juMin + 1;   /* ★ 목표 주간 인원 = 최소 + 1 */

    /* 직원별 주간 횟수를 실시간으로 셀 수 있는 배열 */
    const juCount = Array.from({length: N}, (_, i) =>
      raw[i].filter(s => s === "주").length
    );

    for (let d = 0; d < D; d++) {
      const curJu = raw.filter(r => r[d] === "주").length;
      if (curJu >= targetJuPerDay) continue;   /* 이미 충족 → 스킵 */

      /* 후보: 휴무 중 + 야간후 아님 + locked 아님 + 시간 여유 있음
         정렬: 주간 횟수 오름차순 (가장 적게 일한 사람 우선)        */
      const cands = Array.from({length: N}, (_, i) => i)
        .filter(i => {
          if (lck[i][d]) return false;
          if (raw[i][d] !== "휴") return false;
          if (d > 0 && raw[i][d-1] === "야") return false;   /* 야간 후 휴무 보호 */
          /* 총 근무시간 여유 확인 (8h 추가해도 targetH 안 넘음) */
          if (calcHours(raw[i]) + 8 > targetH) return false;
          return true;
        })
        .sort((a, b) => juCount[a] - juCount[b]);   /* ★ 주간 횟수 적은 사람 먼저 */

      const need = targetJuPerDay - curJu;
      const chosen = cands.slice(0, need);

      chosen.forEach(i => {
        raw[i][d] = "주";
        juCount[i]++;           /* 실시간 카운트 갱신 */
      });
    }
  }

  /* ─────────────────────────────────────────────
     PHASE 4 : 개인 근무시간 맞추기 (주 ↔ 휴 만)
     ★ locked 셀 절대 변경 금지
  ───────────────────────────────────────────── */
  for (let i = 0; i < N; i++) {
    let hours = calcHours(raw[i]);
    const juFloor = juMin + 1;   /* ★ 균등 분배 보호: juMin+1을 기준선으로 */

    /* 초과 → 주를 휴로 (juFloor 유지, locked 스킵) */
    for (let d = D-1; d >= 0 && hours > targetH; d--) {
      if (lck[i][d]) continue;
      if (raw[i][d] !== "주") continue;
      const dayJu = raw.filter(r => r[d] === "주").length;
      if (dayJu > juFloor) { raw[i][d] = "휴"; hours -= 8; }   /* ★ juFloor 기준 */
    }
    /* juFloor로 못 줄이면 juMin 기준으로 2차 시도 */
    for (let d = D-1; d >= 0 && hours > targetH; d--) {
      if (lck[i][d]) continue;
      if (raw[i][d] !== "주") continue;
      const dayJu = raw.filter(r => r[d] === "주").length;
      if (dayJu > juMin) { raw[i][d] = "휴"; hours -= 8; }
    }
    /* 최후 수단 초과 제거 */
    if (hours > targetH) {
      const days = Array.from({length:D},(_,d)=>d)
        .filter(d => raw[i][d]==="주" && !lck[i][d])
        .sort((a,b)=>raw.filter(r=>r[b]==="주").length - raw.filter(r=>r[a]==="주").length);
      for (const d of days) {
        if (hours <= targetH) break;
        const dayJu = raw.filter(r => r[d] === "주").length;
        if (dayJu > 1) { raw[i][d] = "휴"; hours -= 8; }
      }
      for (let d = D-1; d >= 0 && hours > targetH; d--) {
        if (lck[i][d]) continue;
        if (raw[i][d] === "주") { raw[i][d] = "휴"; hours -= 8; }
      }
    }

    /* 부족 → 휴를 주로 (dayJu>=juMin 날 우선, locked/야간후 제외) */
    for (let d = 0; d < D && hours < targetH; d++) {
      if (lck[i][d]) continue;                               /* ★ locked 스킵 */
      if (raw[i][d] !== "휴") continue;
      if (d > 0 && raw[i][d-1] === "야") continue;
      const dayJu = raw.filter(r => r[d] === "주").length;
      if (dayJu >= juMin) { raw[i][d] = "주"; hours += 8; }
    }
    /* 여전히 부족하면 조건 없이 */
    for (let d = 0; d < D && hours < targetH; d++) {
      if (lck[i][d]) continue;                               /* ★ locked 스킵 */
      if (raw[i][d] !== "휴") continue;
      if (d > 0 && raw[i][d-1] === "야") continue;
      raw[i][d] = "주"; hours += 8;
    }
  }

  /* ─────────────────────────────────────────────
     PHASE 5 : 야간 후 휴무 최종 재보장 (locked 제외)
  ───────────────────────────────────────────── */
  for (let i = 0; i < N; i++)
    for (let d = 0; d < D-1; d++)
      if (raw[i][d]==="야" && raw[i][d+1]==="주" && !lck[i][d+1])  /* ★ locked 스킵 */
        raw[i][d+1] = "휴";

  /* ─────────────────────────────────────────────
     PHASE 6 : Phase 5 보정 후 시간 재균형 (locked 제외)
  ───────────────────────────────────────────── */
  for (let i = 0; i < N; i++) {
    let hours = calcHours(raw[i]);
    for (let d = 0; d < D && hours < targetH; d++) {
      if (lck[i][d]) continue;                               /* ★ locked 스킵 */
      if (raw[i][d] !== "휴") continue;
      if (d > 0 && raw[i][d-1] === "야") continue;
      if (d < D-1 && raw[i][d+1] === "야") continue;
      raw[i][d] = "주"; hours += 8;
    }
  }

  /* ─────────────────────────────────────────────
     PHASE 7 : 일별 제약 최종 재검증 / 재보정 (locked 전달)
  ───────────────────────────────────────────── */
  reAdjustDailyConstraints(raw, D, cst, lck);               /* ★ lck 전달 */

  /* locked 플래그를 최종 sch 객체에 반영 */
  return { sch: raw.map((row, i) => row.map((t, d) => mkCell(t, lck[i][d]))) };
}

function validateSchedule(sch,empNames,D,cst,targetH){
  const N=sch.length, {se:seR,ya:yaR,juMin}=cst, errs=[];
  for(let d=0;d<D;d++){
    const col=sch.map(r=>cellT(r[d])), c=s=>col.filter(x=>x===s).length;
    if(c("석")!==seR) errs.push({t:"e",m:`${d+1}일 석간: ${c("석")}명 (필요 ${seR}명)`});
    if(c("야")!==yaR) errs.push({t:"e",m:`${d+1}일 야간: ${c("야")}명 (필요 ${yaR}명)`});
    if(c("주")<juMin)  errs.push({t:"w",m:`${d+1}일 주간: ${c("주")}명 (최소 ${juMin}명)`});
  }
  for(let i=0;i<N;i++){
    for(let d=0;d<D-1;d++)
      if(cellT(sch[i][d])==="야"&&cellT(sch[i][d+1])!=="휴")
        errs.push({t:"e",m:`${empNames[i]}: ${d+1}일 야간 후 미휴무`});
    const h=calcHours(sch[i]);
    if(h!==targetH) errs.push({t:h>targetH?"e":"w",
      m:`${empNames[i]}: ${h}h (목표 ${targetH}h, ${h>targetH?"초과":"부족"} ${Math.abs(h-targetH)}h)`});
  }
  return errs;
}

/* ══════════════════════════════════════════════════════════
   PATTERN EDITOR MODAL  — Light
══════════════════════════════════════════════════════════ */
function PatternModal({pattern,onSave,onClose}){
  const [name,setName]=useState(pattern?.name??"새 패턴");
  const [seq,setSeq]=useState([...(pattern?.seq??["주","주","석","야","휴","휴","주","석","야","휴"])]);
  const setCellAt=(idx,t)=>setSeq(p=>p.map((v,i)=>i===idx?t:v));
  const stats={주:seq.filter(s=>s==="주").length,석:seq.filter(s=>s==="석").length,
    야:seq.filter(s=>s==="야").length,휴:seq.filter(s=>s==="휴").length};

  const btnStyle=(active,t)=>({
    padding:"4px 7px",borderRadius:5,cursor:"pointer",fontWeight:700,fontSize:"0.7rem",
    background:active?SHIFT[t].bg:"#f8fafc",
    color:active?SHIFT[t].txt:"#64748b",
    border:`1.5px solid ${active?SHIFT[t].brd:"#e2e8f0"}`,
    transition:"all .1s",
  });

  return(
    <div style={{position:"fixed",inset:0,background:"#0008",zIndex:2000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:16,
        padding:24,width:580,maxWidth:"95vw",boxShadow:"0 20px 60px #0002"}}>

        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <h3 style={{margin:0,color:"#0f172a",fontSize:"1rem",fontWeight:700}}>
            {pattern?"🔧 패턴 편집":"✨ 새 패턴 추가"}
          </h3>
          <button onClick={onClose} style={{background:"#f1f5f9",border:"1px solid #e2e8f0",
            borderRadius:8,width:30,height:30,cursor:"pointer",color:"#64748b",
            fontSize:"1.1rem",lineHeight:1,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>

        <div style={{marginBottom:14}}>
          <label style={{fontSize:"0.7rem",color:"#64748b",display:"block",marginBottom:4,fontWeight:600}}>패턴 이름</label>
          <input value={name} onChange={e=>setName(e.target.value)} style={{
            background:"#f8fafc",color:"#0f172a",border:"1.5px solid #e2e8f0",borderRadius:8,
            padding:"8px 11px",fontSize:"0.88rem",outline:"none",
            width:"100%",boxSizing:"border-box",
          }}/>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <label style={{fontSize:"0.7rem",color:"#64748b",fontWeight:600}}>
              근무 순서 <span style={{color:"#94a3b8",fontWeight:400}}>({seq.length}일 사이클)</span>
            </label>
            <div style={{display:"flex",gap:4}}>
              <button onClick={()=>setSeq(p=>p.length>2?p.slice(0,-1):p)} style={{
                background:"#fff1f2",color:"#e11d48",border:"1px solid #fecdd3",
                borderRadius:6,padding:"3px 9px",fontSize:"0.72rem",cursor:"pointer",fontWeight:700}}>− 칸</button>
              <button onClick={()=>setSeq(p=>[...p,"휴"])} style={{
                background:"#f0fdf4",color:"#16a34a",border:"1px solid #bbf7d0",
                borderRadius:6,padding:"3px 9px",fontSize:"0.72rem",cursor:"pointer",fontWeight:700}}>+ 칸</button>
            </div>
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,
            background:"#f8fafc",borderRadius:10,padding:12,border:"1px solid #e2e8f0"}}>
            {seq.map((s,idx)=>(
              <div key={idx} style={{textAlign:"center"}}>
                <div style={{fontSize:"0.55rem",color:"#94a3b8",marginBottom:3,fontWeight:600}}>D{idx+1}</div>
                <div style={{display:"flex",flexDirection:"column",gap:2}}>
                  {["주","석","야","휴"].map(t=>(
                    <button key={t} onClick={()=>setCellAt(idx,t)} style={btnStyle(s===t,t)}>{t}</button>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Preview */}
        <div style={{background:"#f8fafc",borderRadius:10,padding:12,marginBottom:18,border:"1px solid #e2e8f0"}}>
          <div style={{fontSize:"0.62rem",color:"#64748b",marginBottom:7,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em"}}>미리보기</div>
          <div style={{display:"flex",gap:3,flexWrap:"wrap",marginBottom:7}}>
            {seq.map((s,i)=>(
              <span key={i} style={{display:"inline-block",padding:"3px 7px",borderRadius:5,
                fontWeight:700,fontSize:"0.74rem",background:SHIFT[s].bg,
                color:SHIFT[s].txt,border:`1.5px solid ${SHIFT[s].brd}`}}>{s}</span>
            ))}
          </div>
          <div style={{fontSize:"0.65rem",color:"#64748b",display:"flex",gap:14,flexWrap:"wrap"}}>
            {Object.entries(stats).map(([k,v])=>(
              <span key={k}><span style={{color:SHIFT[k]?.txt,fontWeight:700}}>{k}</span> {v}일</span>
            ))}
            <span style={{color:"#94a3b8"}}>| {seq.length}일 사이클 · 주야비 {(stats["야"]/(stats["주"]||1)).toFixed(2)}</span>
          </div>
        </div>

        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{
            background:"#f1f5f9",color:"#475569",border:"1px solid #e2e8f0",
            borderRadius:9,padding:"9px 20px",fontSize:"0.84rem",cursor:"pointer",fontWeight:600}}>취소</button>
          <button onClick={()=>onSave({name:name.trim()||"새 패턴",seq})} style={{
            background:"linear-gradient(135deg,#2563eb,#7c3aed)",color:"#fff",border:"none",
            borderRadius:9,padding:"9px 24px",fontSize:"0.84rem",cursor:"pointer",fontWeight:700,
            boxShadow:"0 4px 12px #2563eb30"}}>저장</button>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   CELL POPUP  — Light
══════════════════════════════════════════════════════════ */
function CellPopup({cell,onSelect,onClose}){
  const ref=useRef(null);
  useEffect(()=>{
    const h=e=>{if(ref.current&&!ref.current.contains(e.target))onClose();};
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[onClose]);
  const locked = isLocked(cell);
  return(
    <div ref={ref} style={{position:"absolute",zIndex:999,top:"calc(100% + 4px)",left:"50%",
      transform:"translateX(-50%)",
      background:"#fff",border:"1px solid #e2e8f0",borderRadius:11,
      padding:8,display:"flex",flexDirection:"column",gap:3,
      boxShadow:"0 8px 24px #00000018",minWidth:102}}>
      <div style={{fontSize:"0.6rem",color:"#94a3b8",textAlign:"center",
        marginBottom:3,fontWeight:700,letterSpacing:"0.04em"}}>
        {locked ? "🔒 잠긴 셀" : "근무 선택"}
      </div>
      {SHIFT_TYPES.map(t=>{
        const s=SHIFT[t], active=cellT(cell)===t;
        return(
          <button key={t} onClick={()=>onSelect(t)} style={{
            display:"flex",alignItems:"center",gap:6,cursor:"pointer",
            fontWeight:active?700:500,fontSize:"0.74rem",
            background:active?s.bg:"#f8fafc",
            color:active?s.txt:"#475569",
            border:`1.5px solid ${active?s.brd:"#e2e8f0"}`,
            borderRadius:7,padding:"5px 8px",transition:"all .1s",
          }}>
            <span style={{fontWeight:800,minWidth:14,color:s.txt}}>{t}</span>
            <span style={{fontSize:"0.64rem",color:"#94a3b8"}}>{s.label}</span>
            <span style={{fontSize:"0.59rem",marginLeft:"auto",color:"#cbd5e1"}}>{s.hours>0?`${s.hours}h`:"—"}</span>
          </button>
        );
      })}
      {/* ★ 잠금 해제 버튼 — locked 셀에만 표시 */}
      {locked && (
        <button onClick={()=>onSelect("__unlock__")} style={{
          marginTop:3,background:"#fff7ed",color:"#c2410c",
          border:"1.5px solid #fed7aa",borderRadius:7,
          padding:"5px 8px",cursor:"pointer",fontSize:"0.72rem",fontWeight:700,
          display:"flex",alignItems:"center",gap:5,
        }}>
          <span>🔓</span><span>잠금 해제 (자동 조정 허용)</span>
        </button>
      )}
      <button onClick={onClose} style={{marginTop:3,background:"#f1f5f9",border:"1px solid #e2e8f0",
        borderRadius:7,padding:"4px",cursor:"pointer",color:"#94a3b8",fontSize:"0.65rem",fontWeight:600}}>닫기</button>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   초기 사용자 가이드 모달 — 첫 접속 시 자동 표시
══════════════════════════════════════════════════════════ */
function InitGuideModal({ onClose, onNeverShow }) {
  const steps = [
    { step:"1단계", icon:"👥", title:"직원 등록", desc:"사이드바 '직원' 탭에서 직원 이름을 추가하세요." },
    { step:"2단계", icon:"⚙️", title:"인원 설정", desc:"'설정' 탭에서 석간·야간 필요 인원과 월 목표 근무시간을 설정하세요." },
    { step:"3단계", icon:"⚡", title:"근무표 생성", desc:"상단 '전체 근무표 생성' 버튼을 클릭하면 자동으로 근무표가 만들어져요." },
    { step:"4단계", icon:"✏️", title:"수동 수정", desc:"셀을 클릭해서 개별 근무를 변경할 수 있어요. 수정된 셀은 🔒로 표시돼요." },
    { step:"5단계", icon:"💾", title:"저장 및 출력", desc:"'저장' 버튼으로 데이터를 저장하고, 인쇄·PDF·엑셀로 출력하세요." },
  ];
  return (
    <div style={{position:"fixed",inset:0,background:"#0008",zIndex:4000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:16,padding:28,
        width:520,maxWidth:"95vw",boxShadow:"0 20px 60px #0003"}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:"2.4rem",marginBottom:8}}>🏥</div>
          <h2 style={{margin:0,fontSize:"1.1rem",fontWeight:800,color:"#0f172a"}}>
            요양원 근무표 시스템에 오신 것을 환영해요!
          </h2>
          <p style={{margin:"6px 0 0",fontSize:"0.78rem",color:"#64748b"}}>
            아래 순서대로 따라하면 바로 사용할 수 있어요 😊
          </p>
        </div>

        <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
          {steps.map((s,i)=>(
            <div key={i} style={{
              display:"flex",alignItems:"center",gap:12,
              background:"#f8fafc",borderRadius:10,padding:"10px 14px",
              border:"1px solid #e2e8f0",
            }}>
              <div style={{
                background:"linear-gradient(135deg,#2563eb,#4f46e5)",
                color:"#fff",borderRadius:8,
                padding:"4px 8px",fontSize:"0.66rem",fontWeight:700,whiteSpace:"nowrap",
              }}>{s.step}</div>
              <span style={{fontSize:"1.2rem"}}>{s.icon}</span>
              <div>
                <div style={{fontWeight:700,color:"#0f172a",fontSize:"0.84rem"}}>{s.title}</div>
                <div style={{fontSize:"0.72rem",color:"#475569"}}>{s.desc}</div>
              </div>
            </div>
          ))}
        </div>

        <button onClick={onClose} style={{
          width:"100%",background:"linear-gradient(135deg,#2563eb,#4f46e5)",
          color:"#fff",border:"none",borderRadius:10,padding:"11px",
          fontSize:"0.88rem",fontWeight:700,cursor:"pointer",marginBottom:8,
        }}>시작하기 🚀</button>

        <button onClick={onNeverShow} style={{
          width:"100%",background:"transparent",color:"#94a3b8",
          border:"none",fontSize:"0.74rem",cursor:"pointer",padding:"4px",
        }}>다시 보지 않기</button>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   코드 저장 결과 모달 — 생성된 코드 표시 + 복사
══════════════════════════════════════════════════════════ */
function CodeResultModal({ code, onClose }) {
  const [copied, setCopied] = useState(false);
  const copy = () => {
    navigator.clipboard.writeText(code).then(()=>{
      setCopied(true); setTimeout(()=>setCopied(false),2000);
    });
  };
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",zIndex:3000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
      onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:"#fff",borderRadius:16,padding:30,
        width:380,maxWidth:"95vw",textAlign:"center",
        boxShadow:"0 20px 60px #0003",
      }}>
        <div style={{fontSize:"2.2rem",marginBottom:10}}>🎉</div>
        <h3 style={{margin:"0 0 6px",color:"#0f172a",fontSize:"1rem",fontWeight:800}}>
          저장 완료!
        </h3>
        <p style={{margin:"0 0 18px",fontSize:"0.8rem",color:"#64748b"}}>
          아래 코드를 메모해두세요.<br/>다른 PC에서 이 코드로 불러올 수 있어요.
        </p>

        {/* 코드 표시 */}
        <div style={{
          background:"#f8fafc",border:"2px solid #2563eb",borderRadius:12,
          padding:"16px 20px",marginBottom:16,
          display:"flex",alignItems:"center",justifyContent:"space-between",
          gap:10,
        }}>
          <span style={{
            fontFamily:"monospace",fontSize:"1.8rem",fontWeight:900,
            color:"#2563eb",letterSpacing:"0.15em",
          }}>{code}</span>
          <button onClick={copy} style={{
            background: copied ? "#dcfce7":"#eff6ff",
            color: copied ? "#15803d":"#2563eb",
            border: `1.5px solid ${copied?"#bbf7d0":"#bfdbfe"}`,
            borderRadius:8,padding:"6px 12px",fontSize:"0.78rem",
            fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",
          }}>{copied ? "✅ 복사됨":"📋 복사"}</button>
        </div>

        <p style={{fontSize:"0.72rem",color:"#94a3b8",marginBottom:16}}>
          💡 코드는 대소문자 구분 없이 입력 가능해요
        </p>

        <button onClick={onClose} style={{
          width:"100%",background:"linear-gradient(135deg,#2563eb,#4f46e5)",
          color:"#fff",border:"none",borderRadius:10,padding:"11px",
          fontSize:"0.88rem",fontWeight:700,cursor:"pointer",
        }}>확인</button>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   코드 불러오기 모달 — 코드 입력 후 복원
══════════════════════════════════════════════════════════ */
function CodeLoadModal({ onLoad, onClose }) {
  const [code,    setCode]    = useState("");
  const [loading, setLoading] = useState(false);
  const [err,     setErr]     = useState("");

  const doLoad = async () => {
    const c = code.trim().toUpperCase();
    if (!c) { setErr("코드를 입력해주세요."); return; }
    setLoading(true); setErr("");
    try {
      const { data, error } = await supabase
        .from("schedules")
        .select("data")
        .eq("code", c)
        .single();
      if (error || !data) { setErr("코드를 찾을 수 없습니다. 다시 확인해주세요."); return; }
      onLoad(data.data);
    } catch(e) {
      setErr("불러오기 실패: " + e.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{position:"fixed",inset:0,background:"#0007",zIndex:3000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
      onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:"#fff",borderRadius:16,padding:28,
        width:380,maxWidth:"95vw",
        boxShadow:"0 20px 60px #0003",
      }}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <h3 style={{margin:0,fontSize:"1rem",fontWeight:800,color:"#0f172a"}}>
            📂 코드로 불러오기
          </h3>
          <button onClick={onClose} style={{
            background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:8,
            width:30,height:30,cursor:"pointer",fontSize:"1.1rem",color:"#64748b",
          }}>×</button>
        </div>

        <p style={{fontSize:"0.8rem",color:"#64748b",margin:"0 0 14px"}}>
          저장 시 발급받은 6자리 코드를 입력하세요.
        </p>

        <input
          value={code}
          onChange={e=>{ setCode(e.target.value.toUpperCase()); setErr(""); }}
          onKeyDown={e=>e.key==="Enter"&&doLoad()}
          placeholder="예: ABC123"
          maxLength={6}
          style={{
            width:"100%",boxSizing:"border-box",
            padding:"12px 16px",fontSize:"1.4rem",fontWeight:800,
            letterSpacing:"0.2em",textAlign:"center",
            border:"2px solid #e2e8f0",borderRadius:10,
            fontFamily:"monospace",color:"#0f172a",outline:"none",
            marginBottom:8,
          }}
        />
        {err && <p style={{color:"#dc2626",fontSize:"0.76rem",margin:"0 0 10px"}}>{err}</p>}

        <button onClick={doLoad} disabled={loading} style={{
          width:"100%",background: loading?"#94a3b8":"linear-gradient(135deg,#2563eb,#4f46e5)",
          color:"#fff",border:"none",borderRadius:10,padding:"11px",
          fontSize:"0.88rem",fontWeight:700,cursor: loading?"not-allowed":"pointer",
          marginBottom:8,
        }}>{loading ? "불러오는 중...":"🔍 불러오기"}</button>

        <button onClick={onClose} style={{
          width:"100%",background:"#f1f5f9",color:"#475569",
          border:"1px solid #e2e8f0",borderRadius:10,padding:"9px",
          fontSize:"0.84rem",cursor:"pointer",
        }}>취소</button>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   사용 설명서 모달
══════════════════════════════════════════════════════════ */
function HelpModal({ onClose }) {
  const steps = [
    { icon:"👥", title:"1단계 — 직원 추가", desc:'사이드바 "직원" 탭에서 이름을 입력하고 추가 버튼을 누르세요. 직원별로 사용할 패턴과 시작 오프셋을 설정할 수 있어요.' },
    { icon:"⚙️", title:"2단계 — 인원 설정", desc:'"설정" 탭에서 석간/야간 필요 인원과 주간 최소 인원, 월 목표 근무시간을 설정하세요. 기본값은 석간2·야간2·주간2·144h 입니다.' },
    { icon:"🎯", title:"3단계 — 패턴 설정", desc:'"패턴" 탭에서 근무 패턴을 확인하거나 새로 추가하세요. 기본 균형형(주주석야휴휴주석야휴)이 가장 안정적입니다.' },
    { icon:"⚡", title:"4단계 — 근무표 생성", desc:'상단의 "⚡ 전체 근무표 생성" 버튼을 클릭하세요. 모든 제약조건(석간2명, 야간2명, 야간후휴무)을 자동으로 맞춰줍니다.' },
    { icon:"✏️", title:"5단계 — 수동 수정", desc:'생성된 근무표의 셀을 클릭하면 근무를 변경할 수 있어요. 수정된 셀은 🔒 황색 테두리로 표시되며, 다시 생성해도 유지됩니다.' },
    { icon:"🖨️", title:"6단계 — 인쇄 / PDF / 엑셀", desc:'"인쇄" 버튼으로 바로 출력, "PDF" 버튼으로 PDF 저장, "엑셀" 버튼으로 .xlsx 파일 다운로드 할 수 있어요.' },
    { icon:"💾", title:"저장 & 불러오기", desc:'"💾 저장" 버튼으로 현재 상태를 브라우저에 저장하세요. 다음에 다시 열면 자동으로 불러옵니다. "📂 불러오기"로 수동 복원도 가능해요.' },
  ];
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",zIndex:3000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
      onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:"#fff",borderRadius:16,padding:28,
        width:560,maxWidth:"95vw",maxHeight:"88vh",overflowY:"auto",
        boxShadow:"0 20px 60px #0003",
      }}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <h2 style={{margin:0,fontSize:"1.1rem",fontWeight:800,color:"#0f172a"}}>
            📘 사용 설명서
          </h2>
          <button onClick={onClose} style={{
            background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:8,
            width:32,height:32,cursor:"pointer",fontSize:"1.1rem",color:"#64748b",
            display:"flex",alignItems:"center",justifyContent:"center",
          }}>×</button>
        </div>

        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          {steps.map((s,i)=>(
            <div key={i} style={{
              background:"#f8fafc",borderRadius:10,padding:"12px 14px",
              border:"1px solid #e2e8f0",display:"flex",gap:12,alignItems:"flex-start",
            }}>
              <span style={{fontSize:"1.4rem",flexShrink:0}}>{s.icon}</span>
              <div>
                <div style={{fontWeight:700,color:"#0f172a",fontSize:"0.88rem",marginBottom:4}}>{s.title}</div>
                <div style={{fontSize:"0.78rem",color:"#475569",lineHeight:1.6}}>{s.desc}</div>
              </div>
            </div>
          ))}
        </div>

        <div style={{
          marginTop:18,padding:"10px 14px",background:"#eff6ff",
          borderRadius:9,border:"1px solid #bfdbfe",fontSize:"0.76rem",color:"#1d4ed8",
        }}>
          💡 <strong>팁:</strong> 부서 탭을 여러 개 만들어 병동별로 따로 관리할 수 있어요!
        </div>

        <button onClick={onClose} style={{
          marginTop:16,width:"100%",background:"linear-gradient(135deg,#2563eb,#4f46e5)",
          color:"#fff",border:"none",borderRadius:10,padding:"11px",
          fontSize:"0.88rem",fontWeight:700,cursor:"pointer",
        }}>확인, 시작할게요!</button>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   MAIN APP
══════════════════════════════════════════════════════════ */
export default function App(){
  useEffect(()=>{
    if(document.getElementById("nss-print"))return;
    const s=document.createElement("style");
    s.id="nss-print"; s.textContent=PRINT_CSS;
    document.head.appendChild(s);
  },[]);

  const now=new Date();

  /* ── localStorage에서 초기값 읽기 (렌더 전 동기 실행) ── */
  const _init = (() => {
    try {
      const raw = localStorage.getItem("scheduleData");
      console.log("[초기로드] localStorage 확인:", raw ? raw.length+"bytes 존재" : "없음");
      if (!raw) return null;
      const d = JSON.parse(raw);
      console.log("[초기로드] 파싱 성공 ✅", d);
      return d;
    } catch(e) {
      console.error("[초기로드] 실패:", e);
      return null;
    }
  })();  /* ← 즉시 실행 함수 (IIFE) */

  const [year,     setYear]     = useState(_init?.year     ?? now.getFullYear());
  const [month,    setMonth]    = useState(_init?.month    ?? now.getMonth()+1);
  const [rate,     setRate]     = useState(_init?.rate     ?? 10000);
  const [patterns, setPatterns] = useState(_init?.patterns?.length ? _init.patterns : INIT_STATE.patterns);
  const [depts,    setDepts]    = useState(_init?.depts?.length    ? _init.depts    : INIT_STATE.depts);
  const [tab,      setTab]      = useState(_init?.tab      ?? "dept1");
  const [results,  setResults]  = useState(_init?.results  ?? {});
  const [popup,    setPopup]    = useState(null);
  const [patModal, setPatModal] = useState(null);
  const [sideTab,  setSideTab]  = useState("설정");
  const [newEmpName,setNewEmpName] = useState("");
  const [empSearch, setEmpSearch]  = useState("");
  const [showHelp,  setShowHelp]   = useState(false);
  const [saveMsg,   setSaveMsg]    = useState("");
  const [codeResult,   setCodeResult]   = useState(null);
  const [showLoadModal,setShowLoadModal]= useState(false);
  const [locked,    setLocked]     = useState(false);          /* ① 수정 잠금 */
  const [showGuide, setShowGuide]  = useState(false);          /* ③ 초기 가이드 */
  const tableRef    = useRef(null);
  const hasMounted  = useRef(false);
  const saveTimer   = useRef(null);                            /* ② debounce 타이머 */

  /* ── 첫 접속 시 가이드 자동 표시 ── */
  useEffect(()=>{
    if(!localStorage.getItem("guideShown")) setShowGuide(true);
  },[]);

  /* ── 자동저장 — debounce 1.5초 적용 ── */
  useEffect(()=>{
    if(!hasMounted.current){ hasMounted.current=true; return; }
    if(saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(()=>{
      try{
        const data = { year, month, rate, patterns, depts, tab, results };
        const json = JSON.stringify(data, (_k,v)=>v===undefined?null:v);
        localStorage.setItem("scheduleData", json);
        console.log("[자동저장] 완료:", json.length, "bytes");
      }catch(e){ console.error("[자동저장] 실패:", e); }
    }, 1500);
    return ()=>clearTimeout(saveTimer.current);
  },[year,month,rate,patterns,depts,tab,results]);

  /* ══════════════════════════════════════════════════════
     💾 수동 저장
  ══════════════════════════════════════════════════════ */
  const handleSave=()=>{
    try{
      const data = { year, month, rate, patterns, depts, tab, results };
      const json = JSON.stringify(data, (_k,v)=>v===undefined?null:v);
      localStorage.setItem("scheduleData", json);
      console.log("[수동저장] 저장 완료", data);
      console.log("[수동저장] 확인:", localStorage.getItem("scheduleData")?.length, "bytes");
      setSaveMsg("✅ 저장 완료!");
      setTimeout(()=>setSaveMsg(""),2500);
    }catch(e){
      console.error("[수동저장] 실패:", e);
      alert("저장 실패: "+e.message);
    }
  };

  /* ══════════════════════════════════════════════════════
     📂 수동 불러오기
  ══════════════════════════════════════════════════════ */
  const handleLoad = () => {
    console.log("=== [불러오기 버튼 클릭] ===");   /* ★ 맨 첫 줄 — 클릭 확인용 */

    const saved = localStorage.getItem("scheduleData");
    console.log("[불러오기] localStorage 원본:", saved ? saved.length+"bytes" : "없음");

    if (!saved) {
      alert("저장된 데이터가 없습니다.\n먼저 💾 저장 버튼을 눌러주세요.");
      return;
    }

    try {
      const parsed = JSON.parse(saved);
      console.log("[불러오기] 파싱 성공:", parsed);

      /* state 반영 */
      if (parsed.year     != null)     setYear(parsed.year);
      if (parsed.month    != null)     setMonth(parsed.month);
      if (parsed.rate     != null)     setRate(parsed.rate);
      if (parsed.patterns?.length)     setPatterns(parsed.patterns);
      if (parsed.depts?.length)        setDepts(parsed.depts);
      if (parsed.tab)                  setTab(parsed.tab);
      if (parsed.results)              setResults(parsed.results);

      console.log("[불러오기] 수동 불러오기 완료 ✅", parsed);
      setSaveMsg("📂 불러오기 완료!");
      setTimeout(() => setSaveMsg(""), 2500);
    } catch (e) {
      console.error("[불러오기] 실패:", e);
      alert("불러오기 실패: " + e.message);
    }
  };

  /* ══════════════════════════════════════════════════════
     📥  JSON 파일로 내보내기 (백업)
  ══════════════════════════════════════════════════════ */
  const handleExport = () => {
    try {
      const data = { year, month, rate, patterns, depts, tab, results };
      const json = JSON.stringify(data, (_k,v)=>v===undefined?null:v, 2);
      const blob = new Blob([json], { type:"application/json" });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href = url;
      a.download = `근무표_백업_${year}년${month}월.json`;
      a.click();
      URL.revokeObjectURL(url);
      setSaveMsg("📥 JSON 백업 완료!");
      setTimeout(()=>setSaveMsg(""),2500);
      console.log("[JSON내보내기] 완료");
    } catch(e) {
      console.error("[JSON내보내기] 실패:", e);
      alert("내보내기 실패: "+e.message);
    }
  };

  /* ══════════════════════════════════════════════════════
     📤  JSON 파일 가져오기 (복구)
  ══════════════════════════════════════════════════════ */
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!window.confirm(`"${file.name}" 파일로 복구하시겠습니까?\n현재 데이터가 덮어써집니다.`)) {
      e.target.value=""; return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (parsed.year     != null)  setYear(parsed.year);
        if (parsed.month    != null)  setMonth(parsed.month);
        if (parsed.rate     != null)  setRate(parsed.rate);
        if (parsed.patterns?.length)  setPatterns(parsed.patterns);
        if (parsed.depts?.length)     setDepts(parsed.depts);
        if (parsed.tab)               setTab(parsed.tab);
        if (parsed.results)           setResults(parsed.results);
        setSaveMsg("📤 JSON 불러오기 완료!");
        setTimeout(()=>setSaveMsg(""),2500);
        console.log("[JSON가져오기] 완료", parsed);
      } catch {
        alert("파일 오류: 올바른 JSON 파일이 아닙니다.");
      }
    };
    reader.readAsText(file);
    e.target.value="";
  };

  const fileInputRef = useRef(null);
  const handleCloudSave = async () => {
    if (!isSupabaseReady) {
      alert("Supabase가 설정되지 않았습니다.\n.env 파일에 VITE_SUPABASE_URL과 VITE_SUPABASE_ANON_KEY를 입력해주세요.");
      return;
    }

    /* 6자리 랜덤 코드 생성 (대문자+숫자) */
    const code = Array.from({length:6}, ()=>"ABCDEFGHJKLMNPQRSTUVWXYZ23456789"[
      Math.floor(Math.random()*32)
    ]).join("");

    const data = { year, month, rate, patterns, depts, tab, results };

    try {
      setSaveMsg("☁️ 저장 중...");
      const { error } = await supabase.from("schedules").insert({ code, data });
      if (error) throw error;

      /* localStorage에도 백업 */
      localStorage.setItem("scheduleData", JSON.stringify(data));
      localStorage.setItem("lastCode", code);

      console.log("[클라우드저장] 완료 — 코드:", code);
      setSaveMsg("");
      setCodeResult(code);   /* 코드 표시 모달 열기 */
    } catch(e) {
      console.error("[클라우드저장] 실패:", e);
      setSaveMsg("");
      alert("클라우드 저장 실패: " + e.message);
    }
  };

  /* ══════════════════════════════════════════════════════
     ☁️  클라우드 불러오기 (Supabase) — 코드로 조회
  ══════════════════════════════════════════════════════ */
  const handleCloudLoad = (parsed) => {
    console.log("[클라우드불러오기] 데이터:", parsed);
    if (parsed.year     != null)  setYear(parsed.year);
    if (parsed.month    != null)  setMonth(parsed.month);
    if (parsed.rate     != null)  setRate(parsed.rate);
    if (parsed.patterns?.length)  setPatterns(parsed.patterns);
    if (parsed.depts?.length)     setDepts(parsed.depts);
    if (parsed.tab)               setTab(parsed.tab);
    if (parsed.results)           setResults(parsed.results);

    /* localStorage 백업도 갱신 */
    localStorage.setItem("scheduleData", JSON.stringify(parsed));

    setShowLoadModal(false);
    setSaveMsg("✅ 클라우드에서 불러오기 완료!");
    setTimeout(()=>setSaveMsg(""),3000);
  };

  /* ★ D는 모든 핸들러보다 먼저 선언 */
  const D    = getDIM(year,month);
  const dept = depts.find(d=>d.id===tab)||depts[0];
  const res  = results[dept?.id];

  const patColor=id=>PAT_COLORS[patterns.findIndex(p=>p.id===id)%PAT_COLORS.length];
  const patName =id=>patterns.find(p=>p.id===id)?.name||"?";

  /* Dept helpers */
  const updDept=(id,patch)=>{
    setDepts(p=>p.map(d=>d.id===id?{...d,...patch}:d));
    setResults(p=>{const r={...p};delete r[id];return r;});
  };
  const addDept=()=>{
    const id=newDid();
    setDepts(p=>[...p,{id,name:`${p.length+1}병동`,emps:[],cst:{se:2,ya:2,juMin:2},targetH:144}]);
    setTab(id);
  };
  const delDept=id=>{
    if(depts.length<=1)return;
    const rem=depts.filter(d=>d.id!==id);
    setDepts(rem); setResults(p=>{const r={...p};delete r[id];return r;});
    if(tab===id)setTab(rem[0]?.id||"");
  };

  /* Emp helpers */
  const updEmp=(deptId,idx,patch)=>{
    setDepts(p=>p.map(d=>d.id!==deptId?d:{...d,emps:d.emps.map((e,i)=>i===idx?{...e,...patch}:e)}));
    setResults(p=>{const r={...p};delete r[deptId];return r;});
  };
  const addEmp=()=>{
    const n=newEmpName.trim();
    if(!n||dept.emps.find(e=>e.name===n))return;
    updDept(dept.id,{emps:[...dept.emps,{name:n,patternId:patterns[0]?.id||"p1",offset:dept.emps.length}]});
    setNewEmpName("");
  };
  const delEmp=idx=>updDept(dept.id,{emps:dept.emps.filter((_,i)=>i!==idx)});

  /* Pattern CRUD */
  const savePattern=data=>{
    if(patModal==="new"){ setPatterns(p=>[...p,{id:newPid(),...data,builtin:false,stars:3}]); }
    else { setPatterns(p=>p.map(pat=>pat.id===patModal.id?{...pat,...data}:pat)); setResults({}); }
    setPatModal(null);
  };
  const copyPattern=pat=>setPatterns(p=>[...p,{...pat,id:newPid(),name:`${pat.name} (복사)`,builtin:false}]);
  const delPattern=id=>{
    if(patterns.find(p=>p.id===id)?.builtin)return;
    setDepts(prev=>prev.map(d=>({...d,emps:d.emps.map(e=>e.patternId===id?{...e,patternId:"p1"}:e)})));
    setPatterns(p=>p.filter(pat=>pat.id!==id)); setResults({});
  };

  /* Generate */
  const generate=useCallback(()=>{
    const nRes={};
    for(const d of depts){
      const prevSch = results[d.id]?.sch ?? null;          /* ★ 기존 locked 셀 전달 */
      const r=generateSchedule(d.emps,patterns,D,d.cst,d.targetH, prevSch);
      if(r.err) nRes[d.id]={err:r.err};
      else nRes[d.id]={sch:r.sch,errs:validateSchedule(r.sch,d.emps.map(e=>e.name),D,d.cst,d.targetH)};
    }
    setResults(nRes); setPopup(null);
  },[depts,patterns,D,results]);

  /* Manual edit — locked 상태이면 차단 */
  const applyEdit=useCallback((empIdx,dayIdx,newType)=>{
    if(!res||res.err) return;
    if(locked) return;   /* ★ 잠금 상태 차단 */
    const newSch=res.sch.map((row,i)=>i!==empIdx?row:row.map((cell,d)=>{
      if(d!==dayIdx) return cell;
      if(newType==="__unlock__") return mkCell(cellT(cell), false);  /* 잠금 해제 */
      return mkCell(newType, true);                                   /* ★ locked=true */
    }));
    setResults(p=>({...p,[dept.id]:{...p[dept.id],sch:newSch,
      errs:validateSchedule(newSch,dept.emps.map(e=>e.name),D,dept.cst,dept.targetH)}}));
    setPopup(null);
  },[res,dept,D]);

  /* ══════════════════════════════════════════════════════
     🖨️  인쇄
  ══════════════════════════════════════════════════════ */
  const handlePrint = () => {
    console.log("[인쇄] 버튼 클릭");
    setPopup(null);
    setTimeout(() => window.print(), 100);
  };

  /* ══════════════════════════════════════════════════════
     📄  PDF — HTML Blob → 새 창 → print()
  ══════════════════════════════════════════════════════ */
  const handlePDF = () => {
    console.log("[PDF] 버튼 클릭");
    if (!res || res.err) { alert("근무표를 먼저 생성해주세요."); return; }

    const SHIFT_CSS = { 주:"#dbeafe", 석:"#ffedd5", 야:"#ede9fe", 휴:"#f1f5f9", 월:"#dcfce7", 반:"#ecfccb" };
    const SHIFT_COL = { 주:"#1d4ed8", 석:"#c2410c", 야:"#5b21b6", 휴:"#64748b", 월:"#15803d", 반:"#4d7c0f" };

    const headCols = ["직원/패턴",
      ...Array.from({length:D}, (_,d) => {
        const dow = getDOW(year,month,d+1);
        const color = dow==="일"?"#dc2626": dow==="토"?"#2563eb":"#475569";
        return `<span style="color:${color}">${d+1}<br/><small>${dow}</small></span>`;
      }),
      "근무일","총시간","석간","야간","야간수당"
    ].map(h=>`<th>${h}</th>`).join("");

    const dataRows = dept.emps.map((emp,i)=>{
      const st = stats[i];
      const cells = res.sch[i].map(cell=>{
        const t = cellT(cell), lk = isLocked(cell);
        const bg = SHIFT_CSS[t]||"#f1f5f9", fg = SHIFT_COL[t]||"#64748b";
        return `<td><span style="background:${bg};color:${fg};border:${lk?"2px solid #f59e0b":"1px solid "+fg+"44"};
          border-radius:3px;padding:1px 4px;font-weight:700;font-size:9px">${t}</span></td>`;
      }).join("");
      const hColor = st.hours>dept.targetH?"#dc2626": st.hours<dept.targetH?"#d97706":"#15803d";
      return `<tr>
        <td style="font-weight:700;white-space:nowrap">${emp.name}</td>
        ${cells}
        <td>${st.work}일</td>
        <td style="font-weight:700;color:${hColor}">${st.hours}h</td>
        <td style="color:#ea580c">${st.se}회</td>
        <td style="color:#7c3aed">${st.ya}회</td>
        <td style="color:#be123c">${st.bonus.toLocaleString()}원</td>
      </tr>`;
    }).join("");

    const dailyRows = ["석","야","주"].map(sh=>{
      const cells = Array.from({length:D},(_,d)=>{
        const cnt = res.sch.filter(r=>cellT(r[d])===sh).length;
        const req = sh==="주"?dept.cst.juMin: sh==="석"?dept.cst.se:dept.cst.ya;
        const ok  = sh==="주"?cnt>=req:cnt===req;
        return `<td style="font-size:8px;color:${ok?"#94a3b8":"#dc2626"};font-weight:${ok?400:700}">${cnt}</td>`;
      }).join("");
      return `<tr style="background:#f8fafc">
        <td style="font-size:8px;font-weight:700;color:${sh==="석"?"#ea580c":sh==="야"?"#7c3aed":"#2563eb"}">${sh}간=</td>
        ${cells}
        <td colspan="5" style="font-size:8px;color:#94a3b8">월계 ${res.sch.flat().filter(c=>cellT(c)===sh).length}회</td>
      </tr>`;
    }).join("");

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/>
      <title>${dept.name} ${year}년 ${month}월 근무표</title>
      <style>
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:'Malgun Gothic','Noto Sans KR',sans-serif;font-size:9px;padding:8mm}
        h2{font-size:13px;margin-bottom:4px;color:#0f172a}
        .sub{font-size:8px;color:#64748b;margin-bottom:8px}
        table{border-collapse:collapse;width:100%;table-layout:fixed}
        th,td{border:1px solid #e2e8f0;padding:2px 3px;text-align:center;font-size:8px;overflow:hidden}
        th{background:#f8fafc;font-weight:700;font-size:7px}
        th:first-child,td:first-child{text-align:left;width:60px;position:sticky;left:0;background:#fff}
        @page{margin:8mm;size:A3 landscape}
      </style></head><body>
      <h2>📋 ${dept.name} — ${year}년 ${month}월 근무표</h2>
      <div class="sub">목표 ${dept.targetH}h · ${dept.emps.length}명 · 석간 ${dept.cst.se}명/일 · 야간 ${dept.cst.ya}명/일 · 생성일 ${new Date().toLocaleDateString("ko-KR")}</div>
      <table>
        <thead><tr>${headCols}</tr></thead>
        <tbody>${dataRows}${dailyRows}</tbody>
      </table>
      <script>window.onload=()=>window.print();</script>
    </body></html>`;

    const blob = new Blob([html], {type:"text/html;charset=utf-8"});
    const url  = URL.createObjectURL(blob);
    const win  = window.open(url,"_blank","width=1200,height=800");
    if (!win) alert("팝업 차단을 해제한 후 다시 시도해주세요.");
    setTimeout(()=>URL.revokeObjectURL(url), 10000);
  };

  /* ══════════════════════════════════════════════════════
     📊  엑셀 다운로드 — SheetJS (import * as XLSX)
  ══════════════════════════════════════════════════════ */
  const handleExcel = () => {
    console.log("[엑셀] 버튼 클릭");
    if (!res || res.err) { alert("근무표를 먼저 생성해주세요."); return; }

    try {
      /* ── 워크북 & 시트 데이터 구성 ── */
      const header = [
        "직원",
        ...Array.from({length:D}, (_,d) => `${d+1}일`),
        "총근무일","총시간(h)","주간","석간","야간","월차","반차","야간수당(원)"
      ];
      const rows = [header];

      dept.emps.forEach((emp,i) => {
        const st = stats[i];
        rows.push([
          emp.name,
          ...res.sch[i].map(c => cellT(c)),
          st.work, st.hours, st.ju, st.se, st.ya, st.mo, st.ha, st.bonus
        ]);
      });

      /* 빈 구분 행 + 일별 집계 */
      rows.push(Array(header.length).fill(""));
      ["석","야","주"].forEach(sh => {
        rows.push([
          `${sh}간 일별`,
          ...Array.from({length:D}, (_,d) => res.sch.filter(r=>cellT(r[d])===sh).length),
          "","","","","","","",""
        ]);
      });

      /* ── 셀 스타일 (SheetJS write with cellStyles) ── */
      const ws = XLSX.utils.aoa_to_sheet(rows);

      /* 컬럼 너비 */
      ws["!cols"] = [
        {wch:8},
        ...Array(D).fill({wch:4.5}),
        {wch:8},{wch:8},{wch:5},{wch:5},{wch:5},{wch:5},{wch:5},{wch:12}
      ];

      /* 근무 셀 색상 */
      const BG_MAP = { 주:"DBEAFE", 석:"FFEDD5", 야:"EDE9FE", 휴:"F1F5F9", 월:"DCFCE7", 반:"ECFCCB" };
      const FG_MAP = { 주:"1D4ED8", 석:"C2410C", 야:"5B21B6", 휴:"64748B", 월:"15803D", 반:"4D7C0F" };

      rows.forEach((row,ri) => {
        if (ri === 0) return;
        row.forEach((val,ci) => {
          if (ci === 0 || ci > D) return;
          const addr = XLSX.utils.encode_cell({r:ri, c:ci});
          const bg = BG_MAP[val], fg = FG_MAP[val];
          if (bg && ws[addr]) {
            ws[addr].s = {
              fill:  { patternType:"solid", fgColor:{rgb:bg} },
              font:  { bold:true, color:{rgb:fg} },
              alignment: { horizontal:"center", vertical:"center" }
            };
          }
        });
      });

      /* 헤더 행 스타일 */
      header.forEach((_,ci) => {
        const addr = XLSX.utils.encode_cell({r:0, c:ci});
        if (ws[addr]) {
          ws[addr].s = {
            fill: { patternType:"solid", fgColor:{rgb:"E2E8F0"} },
            font: { bold:true, color:{rgb:"0F172A"} },
            alignment: { horizontal:"center", vertical:"center" }
          };
        }
      });

      /* ── 워크북 생성 & 다운로드 ── */
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, `${dept.name}_${month}월`);
      XLSX.writeFile(wb, `${dept.name}_${year}년${month}월_근무표.xlsx`);
      console.log("[엑셀] 다운로드 완료");

    } catch(e) {
      console.error("[엑셀] 오류:", e);
      alert("엑셀 다운로드 실패: " + e.message);
    }
  };
  /* ══════════════════════════════════════════════════════
     📊 직원별 통계 계산
  ══════════════════════════════════════════════════════ */
  const stats = (res&&!res.err) ? dept.emps.map((_,i) => {
    const row = res.sch[i];
    const ya  = row.filter(c=>cellT(c)==="야").length;
    const se  = row.filter(c=>cellT(c)==="석").length;
    const ju  = row.filter(c=>cellT(c)==="주").length;
    const mo  = row.filter(c=>cellT(c)==="월").length;
    const ha  = row.filter(c=>cellT(c)==="반").length;
    const hu  = row.filter(c=>cellT(c)==="휴").length;
    const hours  = calcHours(row);
    const manCnt = row.filter(c=>isLocked(c)).length;
    return { ya, se, ju, mo, ha, hu, work:D-hu, hours, manCnt, bonus:(se*1+ya*7)*rate };
  }) : [];

  /* ── Design tokens ── */
  const BG    = "#f0f4f8";
  const CARD  = "#ffffff";
  const ALT   = "#f8fafc";
  const BORD  = "#e2e8f0";
  const TXTH  = "#0f172a";
  const TXT   = "#475569";
  const TXTD  = "#94a3b8";
  const ACCENT= "#2563eb";

  const baseInput={background:"#fff",color:TXTH,border:`1.5px solid ${BORD}`,
    borderRadius:8,padding:"7px 11px",fontSize:"0.82rem",outline:"none",
    width:"100%",boxSizing:"border-box",
    boxShadow:"inset 0 1px 3px #0000000a",transition:"border-color .15s"};
  const tabBtn=active=>({
    background:active?`linear-gradient(135deg,${ACCENT},#4f46e5)`:CARD,
    color:active?"#fff":TXT, border:`1.5px solid ${active?ACCENT:BORD}`,
    borderRadius:9,padding:"7px 14px",fontSize:"0.8rem",
    fontWeight:active?700:500,cursor:"pointer",
    display:"flex",alignItems:"center",gap:5,transition:"all .12s",
    boxShadow:active?"0 3px 10px #2563eb25":"0 1px 3px #0000000d",
  });
  const stepBtn={width:26,height:26,background:"#eff6ff",color:ACCENT,
    border:"1.5px solid #bfdbfe",borderRadius:7,cursor:"pointer",
    fontWeight:700,fontSize:"0.9rem",display:"flex",alignItems:"center",justifyContent:"center",
    transition:"background .1s"};
  const sTh=col=>({padding:"7px 8px",textAlign:"center",fontSize:"0.66rem",fontWeight:700,
    borderBottom:`1px solid ${BORD}`,borderLeft:`1px solid ${BORD}`,
    whiteSpace:"nowrap",minWidth:60,color:col,background:"#f8fafc"});
  const sTd={padding:"5px 8px",textAlign:"center",fontSize:"0.71rem",
    borderLeft:`1px solid ${BORD}`,whiteSpace:"nowrap"};
  const sideCard={background:CARD,border:`1px solid ${BORD}`,borderRadius:12,
    padding:14,marginBottom:8,boxShadow:"0 1px 4px #0000000a"};

  const hasErr=res?.errs?.some(e=>e.t==="e");

  const tabDot=id=>{
    const r=results[id]; if(!r||r.err)return r?.err?"#ef4444":null;
    return r.errs?.some(e=>e.t==="e")?"#ef4444":r.errs?.length?"#f59e0b":"#22c55e";
  };

  const sLabel={fontSize:"0.61rem",color:TXTD,textTransform:"uppercase",
    letterSpacing:"0.07em",fontWeight:700,marginBottom:10};

  return(
    <div style={{fontFamily:"'Pretendard','Noto Sans KR',sans-serif",
      minHeight:"100vh",background:BG,color:TXT}}>

      {patModal!==null&&(
        <PatternModal pattern={patModal==="new"?null:patModal}
          onSave={savePattern} onClose={()=>setPatModal(null)}/>
      )}

      {/* 초기 가이드 모달 */}
      {showGuide&&(
        <InitGuideModal
          onClose={()=>setShowGuide(false)}
          onNeverShow={()=>{ localStorage.setItem("guideShown","1"); setShowGuide(false); }}
        />
      )}

      {/* 사용 설명서 모달 */}
      {showHelp&&<HelpModal onClose={()=>setShowHelp(false)}/>}

      {/* 저장 코드 결과 모달 */}
      {codeResult&&<CodeResultModal code={codeResult} onClose={()=>setCodeResult(null)}/>}

      {/* 코드 불러오기 모달 */}
      {showLoadModal&&<CodeLoadModal onLoad={handleCloudLoad} onClose={()=>setShowLoadModal(false)}/>}

      {/* ══ HEADER ══ */}
      <header className="no-print" style={{
        background:`linear-gradient(135deg, ${ACCENT} 0%, #4f46e5 100%)`,
        padding:"14px 22px",boxShadow:"0 4px 16px #2563eb30",
      }}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:42,height:42,borderRadius:11,background:"#ffffff25",
              display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.4rem",
              boxShadow:"0 2px 8px #0002"}}>🏥</div>
            <div>
              <h1 style={{margin:0,fontSize:"1.15rem",fontWeight:800,color:"#fff",letterSpacing:"-0.02em"}}>
                요양원 근무표 관리 시스템
                <span style={{fontSize:"0.6rem",fontWeight:400,marginLeft:8,opacity:.7}}>v5 · 다중 패턴</span>
              </h1>
              <p style={{margin:"2px 0 0",fontSize:"0.63rem",color:"#bfdbfe"}}>
                다중 패턴 · 직원별 배정 · 수동 수정 · 인쇄/PDF · 부서별 독립 운영
              </p>
            </div>
          </div>
          {/* Shift legend */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            {Object.entries(SHIFT).map(([k,v])=>(
              <div key={k} style={{display:"flex",alignItems:"center",gap:4,
                background:"#ffffff20",borderRadius:6,padding:"3px 8px"}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:v.brd,display:"inline-block"}}/>
                <span style={{color:"#fff",fontWeight:700,fontSize:"0.72rem"}}>{k}</span>
                <span style={{color:"#bfdbfe",fontSize:"0.62rem"}}>{v.label}</span>
              </div>
            ))}
            <div style={{background:"#ffffff15",borderRadius:6,padding:"3px 8px",
              fontSize:"0.62rem",color:"#fde68a"}}>
              🔒 황색 테두리=잠금(수동수정)
            </div>
          </div>
        </div>
      </header>

      <div style={{padding:"16px 18px"}}>

        {/* ══ CONTROLS BAR ══ */}
        <div className="no-print" style={{
          background:CARD,border:`1px solid ${BORD}`,borderRadius:12,
          padding:"11px 16px",marginBottom:13,
          display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",
          boxShadow:"0 1px 4px #0000000a",
        }}>
          <select style={{...baseInput,width:"auto",padding:"6px 10px"}} value={year}
            onChange={e=>{setYear(+e.target.value);setResults({});}}>
            {[2024,2025,2026,2027].map(y=><option key={y}>{y}</option>)}
          </select>
          <span style={{color:TXTD,fontSize:"0.78rem"}}>년</span>
          <select style={{...baseInput,width:"auto",padding:"6px 10px"}} value={month}
            onChange={e=>{setMonth(+e.target.value);setResults({});}}>
            {Array.from({length:12},(_,m)=><option key={m+1} value={m+1}>{m+1}월</option>)}
          </select>
          <span style={{color:TXTD,fontSize:"0.78rem"}}>월</span>
          <span style={{background:"#eff6ff",color:ACCENT,border:"1px solid #bfdbfe",
            borderRadius:6,padding:"3px 9px",fontSize:"0.7rem",fontWeight:700}}>{D}일</span>

          <div style={{width:1,height:22,background:BORD}}/>

          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{fontSize:"0.7rem",color:TXT,whiteSpace:"nowrap"}}>야간수당 시급</span>
            <input type="number" step={500} value={rate} onChange={e=>setRate(+e.target.value)}
              style={{...baseInput,width:90}}/>
            <span style={{fontSize:"0.7rem",color:TXTD}}>원</span>
            <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:6,
              padding:"3px 9px",fontSize:"0.65rem",color:"#c2410c"}}>
              석×<span style={{fontWeight:700}}>1h</span> + 야×<span style={{fontWeight:700}}>7h</span>
            </div>
          </div>

          <button onClick={generate} style={{
            background:`linear-gradient(135deg,${ACCENT},#7c3aed)`,color:"#fff",
            border:"none",borderRadius:10,padding:"9px 22px",fontSize:"0.86rem",
            fontWeight:700,cursor:"pointer",boxShadow:"0 4px 14px #2563eb30",
            marginLeft:"auto",
          }}>⚡ 전체 근무표 생성</button>

          {res&&!res.err&&(<>
            <button onClick={handlePrint} style={{
              background:"#eff6ff",color:ACCENT,border:"1.5px solid #bfdbfe",
              borderRadius:9,padding:"8px 16px",fontSize:"0.8rem",fontWeight:600,cursor:"pointer",
              boxShadow:"0 1px 3px #0000000d"}}>🖨️ 인쇄</button>
            <button onClick={handlePDF} style={{
              background:"#f5f3ff",color:"#7c3aed",border:"1.5px solid #ddd6fe",
              borderRadius:9,padding:"8px 16px",fontSize:"0.8rem",fontWeight:600,cursor:"pointer",
              boxShadow:"0 1px 3px #0000000d"}}>📄 PDF</button>
            <button onClick={handleExcel} style={{
              background:"#f0fdf4",color:"#15803d",border:"1.5px solid #bbf7d0",
              borderRadius:9,padding:"8px 16px",fontSize:"0.8rem",fontWeight:600,cursor:"pointer",
              boxShadow:"0 1px 3px #0000000d"}}>📊 엑셀</button>
          </>)}

          {/* ── 저장 / 불러오기 / 사용방법 ── */}
          <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
            {/* 클라우드 저장 (Supabase 설정 시 활성화) */}
            {isSupabaseReady ? (
              <>
                <button onClick={handleCloudSave} style={{
                  background:"linear-gradient(135deg,#2563eb,#4f46e5)",color:"#fff",
                  border:"none",borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",
                  fontWeight:700,cursor:"pointer",boxShadow:"0 2px 8px #2563eb25",
                  display:"flex",alignItems:"center",gap:4,
                }}>☁️ 코드 저장</button>
                <button onClick={()=>setShowLoadModal(true)} style={{
                  background:"#eff6ff",color:"#2563eb",border:"1.5px solid #bfdbfe",
                  borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
                  cursor:"pointer",display:"flex",alignItems:"center",gap:4,
                }}>🔍 코드 불러오기</button>
                <div style={{width:1,height:22,background:"#e2e8f0"}}/>
              </>
            ) : (
              <span style={{
                fontSize:"0.68rem",color:"#94a3b8",background:"#f8fafc",
                border:"1px solid #e2e8f0",borderRadius:7,padding:"4px 9px",
              }}>☁️ Supabase 미설정 (로컬 전용)</span>
            )}

            {/* 로컬 저장 / 불러오기 */}
            <button onClick={handleSave} style={{
              background:"#fefce8",color:"#854d0e",border:"1.5px solid #fde68a",
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer"}}>💾 저장</button>
            <button onClick={()=>handleLoad()} style={{
              background:"#f0f9ff",color:"#0369a1",border:"1.5px solid #bae6fd",
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer"}}>📂 불러오기</button>

            {/* JSON 내보내기/가져오기 */}
            <button onClick={handleExport} style={{
              background:"#f0fdf4",color:"#15803d",border:"1.5px solid #bbf7d0",
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer"}} title="JSON 파일로 백업">📥 백업</button>
            <button onClick={()=>fileInputRef.current?.click()} style={{
              background:"#faf5ff",color:"#7c3aed",border:"1.5px solid #e9d5ff",
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer"}} title="JSON 파일에서 복구">📤 복구</button>
            <input ref={fileInputRef} type="file" accept=".json"
              onChange={handleImport} style={{display:"none"}}/>

            {/* 수정 잠금 토글 */}
            <button onClick={()=>setLocked(p=>!p)} style={{
              background: locked ? "#fef2f2" : "#f8fafc",
              color:      locked ? "#dc2626" : "#64748b",
              border:     `1.5px solid ${locked?"#fecaca":"#e2e8f0"}`,
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer",
            }} title={locked?"잠금 해제":"수정 잠금"}>
              {locked ? "🔒 잠금 중" : "🔓 잠금"}
            </button>

            <button onClick={()=>setShowHelp(true)} style={{
              background:"#f0fdf4",color:"#166534",border:"1.5px solid #bbf7d0",
              borderRadius:9,padding:"8px 14px",fontSize:"0.8rem",fontWeight:600,
              cursor:"pointer"}}>📘 사용방법</button>
          </div>

          {/* 저장 알림 토스트 */}
          {saveMsg&&(
            <span style={{
              background:"#0f172a",color:"#4ade80",borderRadius:8,
              padding:"6px 12px",fontSize:"0.78rem",fontWeight:600,
              animation:"fadeIn .2s",
            }}>{saveMsg}</span>
          )}
        </div>

        {/* ══ 직원 검색 바 ══ */}
        {res&&!res.err&&(
          <div className="no-print" style={{
            marginBottom:11,display:"flex",alignItems:"center",gap:8,
          }}>
            <div style={{position:"relative",flex:"0 0 260px"}}>
              <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",
                fontSize:"0.82rem",color:TXTD,pointerEvents:"none"}}>🔍</span>
              <input
                value={empSearch}
                onChange={e=>setEmpSearch(e.target.value)}
                placeholder="직원 이름 검색..."
                style={{...baseInput,paddingLeft:30,width:"100%"}}
              />
            </div>
            {empSearch&&(
              <span style={{fontSize:"0.72rem",color:TXT}}>
                <strong style={{color:ACCENT}}>{dept.emps.filter(e=>e.name.includes(empSearch)).length}</strong>명 표시 중
              </span>
            )}
            {empSearch&&(
              <button onClick={()=>setEmpSearch("")} style={{
                background:"#f1f5f9",color:TXTD,border:`1px solid ${BORD}`,
                borderRadius:7,padding:"5px 10px",fontSize:"0.74rem",cursor:"pointer",
              }}>✕ 초기화</button>
            )}
          </div>
        )}

        {/* ══ DEPT TABS ══ */}
        <div className="no-print" style={{display:"flex",gap:5,marginBottom:13,
          alignItems:"center",flexWrap:"wrap"}}>
          {depts.map(d=>{
            const dot=tabDot(d.id);
            return(
              <button key={d.id} onClick={()=>setTab(d.id)} style={tabBtn(tab===d.id)}>
                <span>{d.name}</span>
                <span style={{
                  background:tab===d.id?"#ffffff25":"#f1f5f9",
                  color:tab===d.id?"#fff":TXTD,
                  borderRadius:10,padding:"1px 7px",fontSize:"0.62rem",fontWeight:600,
                }}>{d.emps.length}명</span>
                {dot&&<span style={{width:6,height:6,borderRadius:"50%",background:dot}}/>}
              </button>
            );
          })}
          <button onClick={addDept} style={{background:"transparent",color:ACCENT,
            border:`1.5px dashed ${BORD}`,borderRadius:9,
            padding:"7px 13px",fontSize:"0.78rem",cursor:"pointer",fontWeight:600}}>+ 부서 추가</button>
        </div>

        {dept&&(
          <div style={{display:"flex",gap:14,alignItems:"flex-start"}}>

            {/* ══ SIDEBAR ══ */}
            <div className="no-print" style={{width:272,flexShrink:0,display:"flex",flexDirection:"column"}}>

              {/* Side-tab selector */}
              <div style={{display:"flex",marginBottom:10,background:CARD,
                border:`1px solid ${BORD}`,borderRadius:11,padding:3,gap:2,
                boxShadow:"0 1px 4px #0000000a"}}>
                {["설정","패턴","직원"].map(t=>(
                  <button key={t} onClick={()=>setSideTab(t)} style={{
                    flex:1,
                    background:sideTab===t?`linear-gradient(135deg,${ACCENT},#4f46e5)`:"transparent",
                    color:sideTab===t?"#fff":TXT,border:"none",borderRadius:9,
                    padding:"7px 2px",fontSize:"0.74rem",fontWeight:sideTab===t?700:500,
                    cursor:"pointer",transition:"all .12s",
                  }}>{t}</button>
                ))}
              </div>

              {/* 설정 tab */}
              {sideTab==="설정"&&(<>
                <div style={sideCard}>
                  <div style={sLabel}>🏢 부서 설정</div>
                  <div style={{marginBottom:9}}>
                    <label style={{fontSize:"0.67rem",color:TXT,display:"block",marginBottom:4,fontWeight:600}}>부서명</label>
                    <input value={dept.name} onChange={e=>updDept(dept.id,{name:e.target.value})} style={baseInput}/>
                  </div>
                  <div style={{marginBottom:6}}>
                    <label style={{fontSize:"0.67rem",color:TXT,display:"block",marginBottom:4,fontWeight:600}}>월 목표 근무시간</label>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <input type="number" step={8} value={dept.targetH}
                        onChange={e=>updDept(dept.id,{targetH:+e.target.value})}
                        style={{...baseInput,width:78}}/>
                      <span style={{fontSize:"0.7rem",color:TXTD}}>h</span>
                      <span style={{background:"#eff6ff",color:ACCENT,border:"1px solid #bfdbfe",
                        borderRadius:6,padding:"3px 8px",fontSize:"0.65rem",fontWeight:700}}>
                        {Math.round(dept.targetH/8)}일
                      </span>
                    </div>
                  </div>
                  {depts.length>1&&(
                    <button onClick={()=>delDept(dept.id)} style={{
                      background:"#fff1f2",color:"#e11d48",border:"1px solid #fecdd3",
                      borderRadius:7,padding:"6px 11px",fontSize:"0.72rem",
                      cursor:"pointer",width:"100%",marginTop:6,fontWeight:600}}>부서 삭제</button>
                  )}
                </div>

                <div style={sideCard}>
                  <div style={sLabel}>⚙️ 인원 제약조건</div>
                  {[{key:"se",label:"석간 필요",color:"#ea580c",bg:"#fff7ed",bord:"#fed7aa"},
                    {key:"ya",label:"야간 필요",color:"#7c3aed",bg:"#f5f3ff",bord:"#ddd6fe"},
                    {key:"juMin",label:"주간 최소",color:ACCENT,bg:"#eff6ff",bord:"#bfdbfe"}].map(({key,label,color,bg,bord})=>(
                    <div key={key} style={{display:"flex",justifyContent:"space-between",
                      alignItems:"center",marginBottom:9}}>
                      <span style={{fontSize:"0.74rem",color,fontWeight:600}}>{label}</span>
                      <div style={{display:"flex",alignItems:"center",gap:4}}>
                        <button style={{...stepBtn,background:bg,color,border:`1.5px solid ${bord}`}}
                          onClick={()=>updDept(dept.id,{cst:{...dept.cst,[key]:Math.max(0,dept.cst[key]-1)}})}>−</button>
                        <span style={{background:bg,border:`1.5px solid ${bord}`,borderRadius:7,
                          padding:"3px 11px",fontSize:"0.92rem",fontWeight:700,
                          color,minWidth:34,textAlign:"center"}}>{dept.cst[key]}</span>
                        <button style={{...stepBtn,background:bg,color,border:`1.5px solid ${bord}`}}
                          onClick={()=>updDept(dept.id,{cst:{...dept.cst,[key]:dept.cst[key]+1}})}>+</button>
                      </div>
                    </div>
                  ))}
                  <div style={{marginTop:4,padding:"8px 10px",background:"#f8fafc",
                    borderRadius:8,fontSize:"0.63rem",color:TXT,lineHeight:1.9,border:`1px solid ${BORD}`}}>
                    <div>현재 인원: <strong style={{color:TXTH}}>{dept.emps.length}명</strong></div>
                    <div>권장 최소: <strong style={{color:ACCENT}}>
                      {dept.cst.se+dept.cst.ya+dept.cst.juMin+1}명 이상
                    </strong></div>
                  </div>
                </div>
              </>)}

              {/* 패턴 tab */}
              {sideTab==="패턴"&&(
                <div style={sideCard}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <div style={sLabel}>🎯 패턴 관리 <span style={{color:TXT,textTransform:"none"}}>({patterns.length}개)</span></div>
                    <button onClick={()=>setPatModal("new")} style={{
                      background:"#eff6ff",color:ACCENT,border:"1.5px solid #bfdbfe",
                      borderRadius:7,padding:"4px 10px",fontSize:"0.7rem",cursor:"pointer",fontWeight:700}}>
                      + 추가
                    </button>
                  </div>

                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {patterns.map((pat,pi)=>{
                      const col=PAT_COLORS[pi%PAT_COLORS.length];
                      const usedBy=dept.emps.filter(e=>e.patternId===pat.id).length;
                      return(
                        <div key={pat.id} style={{background:ALT,border:`1px solid ${BORD}`,
                          borderRadius:10,padding:"10px 11px",borderLeft:`3px solid ${col}`,
                          boxShadow:"0 1px 3px #0000000a"}}>
                          <div style={{display:"flex",justifyContent:"space-between",
                            alignItems:"flex-start",marginBottom:6}}>
                            <div>
                              <div style={{fontSize:"0.8rem",fontWeight:700,color:TXTH}}>{pat.name}</div>
                              <div style={{fontSize:"0.62rem",color:TXTD,marginTop:2}}>
                                {"★".repeat(pat.stars)+"☆".repeat(5-pat.stars)} · {pat.seq.length}일
                                {usedBy>0&&<span style={{color:col,marginLeft:5,fontWeight:600}}>· {usedBy}명</span>}
                              </div>
                            </div>
                            <div style={{display:"flex",gap:3}}>
                              <button title="편집" onClick={()=>setPatModal(pat)} style={{
                                background:"#eff6ff",color:ACCENT,border:"1px solid #bfdbfe",
                                borderRadius:5,padding:"3px 7px",fontSize:"0.68rem",cursor:"pointer",fontWeight:600}}>편집</button>
                              <button title="복사" onClick={()=>copyPattern(pat)} style={{
                                background:"#f5f3ff",color:"#7c3aed",border:"1px solid #ddd6fe",
                                borderRadius:5,padding:"3px 7px",fontSize:"0.68rem",cursor:"pointer",fontWeight:600}}>복사</button>
                              {!pat.builtin&&(
                                <button title="삭제" onClick={()=>delPattern(pat.id)} style={{
                                  background:"#fff1f2",color:"#e11d48",border:"1px solid #fecdd3",
                                  borderRadius:5,padding:"3px 7px",fontSize:"0.68rem",cursor:"pointer",fontWeight:600}}>삭제</button>
                              )}
                            </div>
                          </div>
                          <div style={{display:"flex",gap:2,flexWrap:"wrap"}}>
                            {pat.seq.map((s,i)=>(
                              <span key={i} style={{display:"inline-block",padding:"2px 5px",borderRadius:4,
                                fontSize:"0.64rem",fontWeight:700,background:SHIFT[s].bg,
                                color:SHIFT[s].txt,border:`1.5px solid ${SHIFT[s].brd}`}}>{s}</span>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  <div style={{marginTop:10,padding:"8px 10px",background:"#f8fafc",
                    borderRadius:8,fontSize:"0.62rem",color:TXT,lineHeight:1.9,border:`1px solid ${BORD}`}}>
                    <div>• 내장 패턴은 삭제 불가, 편집·복사 가능</div>
                    <div>• 편집 시 기존 근무표 초기화됨</div>
                    <div>• 삭제 시 해당 직원 p1 자동 재배정</div>
                  </div>

                  {/* ⭐ 추천 패턴 빠른 추가 */}
                  <div style={{marginTop:10,padding:"9px 10px",background:"#fffbeb",
                    borderRadius:9,border:"1px solid #fde68a"}}>
                    <div style={{fontSize:"0.63rem",color:"#92400e",fontWeight:700,marginBottom:7}}>
                      ⭐ 추천 패턴 빠른 추가
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:5}}>
                      {[
                        {name:"균형형 (10일)", seq:["주","주","석","야","휴","휴","주","석","야","휴"], tag:"⭐⭐⭐⭐⭐ 추천"},
                        {name:"안정형 (7일)",  seq:["주","주","주","석","야","휴","휴"], tag:"⭐⭐⭐⭐"},
                        {name:"휴식형 (5일)",  seq:["주","석","야","휴","휴"], tag:"⭐⭐⭐"},
                      ].map((rec,ri)=>(
                        <button key={ri} onClick={()=>{
                          const id=newPid();
                          setPatterns(p=>[...p,{id,name:rec.name,seq:rec.seq,builtin:false,stars:5-ri}]);
                          setSaveMsg(`✅ "${rec.name}" 추가됨!`);
                          setTimeout(()=>setSaveMsg(""),2000);
                        }} style={{
                          background:"#fff",border:"1px solid #fde68a",borderRadius:8,
                          padding:"7px 10px",cursor:"pointer",textAlign:"left",
                        }}>
                          <div style={{fontSize:"0.72rem",fontWeight:700,color:"#92400e",marginBottom:4}}>
                            {rec.name} <span style={{fontSize:"0.6rem",color:"#b45309"}}>{rec.tag}</span>
                          </div>
                          <div style={{display:"flex",gap:2,flexWrap:"wrap"}}>
                            {rec.seq.map((s,k)=>(
                              <span key={k} style={{display:"inline-block",padding:"1px 4px",
                                borderRadius:3,fontSize:"0.6rem",fontWeight:700,
                                background:SHIFT[s].bg,color:SHIFT[s].txt,
                                border:`1px solid ${SHIFT[s].brd}`}}>{s}</span>
                            ))}
                          </div>
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* 직원 tab */}
              {sideTab==="직원"&&(
                <div style={sideCard}>
                  <div style={sLabel}>👥 직원 관리 <span style={{color:TXT,textTransform:"none"}}>({dept.emps.length}명)</span></div>

                  <div style={{display:"flex",gap:5,marginBottom:10}}>
                    <input value={newEmpName} onChange={e=>setNewEmpName(e.target.value)}
                      onKeyDown={e=>e.key==="Enter"&&addEmp()}
                      placeholder="이름 입력 후 Enter" style={{...baseInput,flex:1}}/>
                    <button onClick={addEmp} style={{
                      background:ACCENT,color:"#fff",border:"none",
                      borderRadius:8,padding:"0 13px",cursor:"pointer",
                      fontWeight:700,fontSize:"0.82rem",whiteSpace:"nowrap",
                      boxShadow:"0 2px 8px #2563eb30"}}>추가</button>
                  </div>

                  <div style={{display:"flex",flexDirection:"column",gap:6,maxHeight:340,overflowY:"auto"}}>
                    {dept.emps.map((emp,i)=>{
                      const col=patColor(emp.patternId);
                      return(
                        <div key={i} style={{background:ALT,border:`1px solid ${BORD}`,
                          borderRadius:9,padding:"9px 11px",borderLeft:`3px solid ${col}`,
                          boxShadow:"0 1px 3px #0000000a"}}>
                          <div style={{display:"flex",justifyContent:"space-between",
                            alignItems:"center",marginBottom:7}}>
                            <span style={{fontWeight:700,color:TXTH,fontSize:"0.84rem"}}>{emp.name}</span>
                            <button onClick={()=>delEmp(i)} style={{
                              background:"#fff1f2",border:"1px solid #fecdd3",color:"#e11d48",
                              cursor:"pointer",fontSize:"0.7rem",borderRadius:5,
                              padding:"2px 7px",fontWeight:700,lineHeight:1}}>삭제</button>
                          </div>
                          <div style={{marginBottom:6}}>
                            <label style={{fontSize:"0.62rem",color:TXTD,display:"block",marginBottom:3,fontWeight:600}}>패턴</label>
                            <select value={emp.patternId}
                              onChange={e=>updEmp(dept.id,i,{patternId:e.target.value})}
                              style={{...baseInput,padding:"5px 8px",fontSize:"0.74rem"}}>
                              {patterns.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                            </select>
                          </div>
                          <div>
                            <label style={{fontSize:"0.62rem",color:TXTD,display:"block",marginBottom:3,fontWeight:600}}>오프셋</label>
                            <div style={{display:"flex",alignItems:"center",gap:5}}>
                              <button style={stepBtn}
                                onClick={()=>updEmp(dept.id,i,{offset:Math.max(0,emp.offset-1)})}>−</button>
                              <span style={{background:"#eff6ff",border:"1.5px solid #bfdbfe",borderRadius:7,
                                padding:"3px 10px",fontSize:"0.88rem",fontWeight:700,
                                color:ACCENT,minWidth:34,textAlign:"center"}}>{emp.offset}</span>
                              <button style={stepBtn} onClick={()=>{
                                const pat=patterns.find(p=>p.id===emp.patternId);
                                updEmp(dept.id,i,{offset:(emp.offset+1)%(pat?.seq.length||10)});
                              }}>+</button>
                              <span style={{fontSize:"0.62rem",color:TXTD}}>D{emp.offset+1}시작</span>
                            </div>
                          </div>
                          {/* Mini preview */}
                          <div style={{marginTop:7,display:"flex",gap:2,flexWrap:"wrap"}}>
                            {(()=>{
                              const pat=patterns.find(p=>p.id===emp.patternId);
                              if(!pat)return null;
                              const seq=pat.seq, L=seq.length, off=emp.offset%L;
                              return seq.map((_,k)=>seq[(k+off)%L]).slice(0,10).map((s,k)=>(
                                <span key={k} style={{display:"inline-block",padding:"2px 5px",
                                  borderRadius:4,fontSize:"0.62rem",fontWeight:700,
                                  background:SHIFT[s].bg,color:SHIFT[s].txt,
                                  border:`1.5px solid ${SHIFT[s].brd}`}}>{s}</span>
                              ));
                            })()}
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {/* Bulk assign */}
                  <div style={{marginTop:10,padding:"9px 10px",background:"#f8fafc",
                    borderRadius:9,border:`1px solid ${BORD}`}}>
                    <div style={{fontSize:"0.63rem",color:TXT,marginBottom:7,fontWeight:700}}>전체 패턴 일괄 배정</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                      {patterns.map((pat,pi)=>(
                        <button key={pat.id}
                          onClick={()=>updDept(dept.id,{
                            emps:dept.emps.map((e,i)=>({...e,patternId:pat.id,offset:i%pat.seq.length}))
                          })}
                          style={{background:"#fff",color:PAT_COLORS[pi%PAT_COLORS.length],
                            border:`1.5px solid ${PAT_COLORS[pi%PAT_COLORS.length]}50`,
                            borderRadius:7,padding:"4px 9px",fontSize:"0.68rem",
                            cursor:"pointer",fontWeight:600,
                            boxShadow:"0 1px 3px #0000000a"}}>
                          {pat.name}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* ══ MAIN PANEL ══ */}
            <div style={{flex:1,minWidth:0}}>

              {/* Validation banner */}
              {res&&(
                <div className="no-print" style={{
                  marginBottom:10,borderRadius:10,padding:"10px 14px",
                  background: res.err||hasErr ? "#fff1f2" : res.errs?.length ? "#fffbeb" : "#f0fdf4",
                  border: `1.5px solid ${res.err||hasErr?"#fecdd3":res.errs?.length?"#fde68a":"#bbf7d0"}`,
                  boxShadow:"0 1px 4px #0000000a",
                }}>
                  {res.err?(
                    <div style={{color:"#be123c",fontWeight:700,fontSize:"0.8rem"}}>🔴 생성 불가: {res.err}</div>
                  ):res.errs.length===0?(
                    <div style={{color:"#15803d",fontWeight:700,fontSize:"0.8rem"}}>
                      ✅ 검증 완료 — 모든 제약조건 충족
                    </div>
                  ):(
                    <>
                      <div style={{fontWeight:700,fontSize:"0.78rem",marginBottom:5,
                        color:hasErr?"#be123c":"#b45309"}}>
                        {hasErr?"🔴 오류":"🟡 경고"} {res.errs.length}건
                      </div>
                      <div style={{display:"flex",flexWrap:"wrap",gap:"3px 14px"}}>
                        {res.errs.map((e,i)=>(
                          <div key={i} style={{fontSize:"0.7rem",color:e.t==="e"?"#be123c":"#b45309"}}>
                            {e.t==="e"?"● ":"○ "}{e.m}
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              )}

              {/* Schedule table */}
              {res&&!res.err?(
                <div style={{background:CARD,border:`1px solid ${BORD}`,borderRadius:12,
                  overflow:"hidden",boxShadow:"0 2px 8px #0000000d"}}>
                  {/* Table header bar */}
                  <div className="no-print" style={{
                    padding:"10px 14px",borderBottom:`1px solid ${BORD}`,
                    display:"flex",justifyContent:"space-between",alignItems:"center",
                    flexWrap:"wrap",gap:6,background:"#f8fafc",
                  }}>
                    <span style={{fontWeight:700,color:TXTH,fontSize:"0.9rem"}}>
                      📋 {dept.name} — {year}년 {month}월 근무표
                    </span>
                    <div style={{display:"flex",gap:10,fontSize:"0.66rem",color:TXTD,flexWrap:"wrap"}}>
                      <span>목표 <strong style={{color:TXTH}}>{dept.targetH}h</strong> · {dept.emps.length}명</span>
                      <span>석간 <strong style={{color:"#ea580c"}}>{dept.cst.se}명</strong>/일 · 야간 <strong style={{color:"#7c3aed"}}>{dept.cst.ya}명</strong>/일</span>
                      <span>패턴 <strong style={{color:ACCENT}}>{[...new Set(dept.emps.map(e=>e.patternId))].length}종</strong> 혼용</span>
                      {res.sch.flat().filter(c=>isLocked(c)).length>0&&(
                        <span style={{color:"#b45309",fontWeight:600}}>
                          🔒 잠금 {res.sch.flat().filter(c=>isLocked(c)).length}셀
                        </span>
                      )}
                    </div>
                  </div>

                  <div id="print-area" style={{overflowX:"auto",WebkitOverflowScrolling:"touch",
                    position:"relative"}} className="print-area">
                    {locked&&(
                      <div style={{
                        position:"absolute",inset:0,zIndex:10,
                        background:"rgba(241,245,249,0.6)",
                        display:"flex",alignItems:"flex-start",justifyContent:"center",
                        paddingTop:20,pointerEvents:"none",borderRadius:8,
                      }}>
                        <div style={{
                          background:"#fff",border:"2px solid #fecaca",borderRadius:12,
                          padding:"10px 20px",display:"flex",alignItems:"center",gap:8,
                          boxShadow:"0 4px 16px #0001",
                        }}>
                          <span style={{fontSize:"1.2rem"}}>🔒</span>
                          <span style={{fontWeight:700,color:"#dc2626",fontSize:"0.84rem"}}>
                            수정 잠금 중 — 🔓 잠금 버튼을 눌러 해제하세요
                          </span>
                        </div>
                      </div>
                    )}
                    <table ref={tableRef} style={{borderCollapse:"collapse",fontSize:"0.72rem",whiteSpace:"nowrap"}}>
                      <thead>
                        <tr style={{background:"#f8fafc"}}>
                          <th style={{
                            padding:"7px 12px",textAlign:"left",color:TXT,fontWeight:700,
                            borderBottom:`1.5px solid ${BORD}`,position:"sticky",left:0,
                            background:"#f8fafc",zIndex:3,minWidth:86,
                            borderRight:`1.5px solid ${BORD}`,
                          }}>직원 / 패턴</th>
                          {Array.from({length:D},(_,d)=>{
                            const dow=getDOW(year,month,d+1);
                            const isSu=dow==="일",isSa=dow==="토";
                            return(
                              <th key={d} style={{
                                padding:"4px 1px",textAlign:"center",minWidth:29,fontWeight:600,
                                borderBottom:`1.5px solid ${BORD}`,
                                color:isSu?"#dc2626":isSa?"#2563eb":TXTD,
                                background:isSu?"#fff1f2":isSa?"#eff6ff":"#f8fafc",
                              }}>
                                <div style={{fontSize:"0.65rem"}}>{d+1}</div>
                                <div style={{fontSize:"0.54rem",opacity:.8}}>{dow}</div>
                              </th>
                            );
                          })}
                          {/* Stats headers */}
                          <th style={sTh(TXT)}>근무일</th>
                          <th style={sTh("#16a34a")}>총시간</th>
                          <th style={sTh(ACCENT)}>주간</th>
                          <th style={sTh("#ea580c")}>석간</th>
                          <th style={sTh("#7c3aed")}>야간</th>
                          <th style={sTh("#15803d")}>월차</th>
                          <th style={sTh("#4d7c0f")}>반차</th>
                          <th style={sTh("#be123c")}>야간수당</th>
                        </tr>
                      </thead>
                      <tbody>
                        {dept.emps
                          .map((emp,i)=>({emp,i}))
                          .filter(({emp})=>!empSearch||emp.name.includes(empSearch))
                          .map(({emp,i})=>{
                          const st=stats[i];
                          const over=st.hours>dept.targetH, under=st.hours<dept.targetH;
                          const bg=i%2===0?CARD:ALT;
                          const col=patColor(emp.patternId);
                          return(
                            <tr key={i} style={{borderBottom:`1px solid ${BORD}`,background:bg}}>
                              <td style={{
                                padding:"5px 12px",fontWeight:700,
                                position:"sticky",left:0,zIndex:1,background:bg,
                                borderRight:`1.5px solid ${BORD}`,borderLeft:`3px solid ${col}`,
                              }}>
                                <div style={{display:"flex",alignItems:"center",gap:5}}>
                                  <span style={{color:TXTH}}>{emp.name}</span>
                                  {st.manCnt>0&&<span style={{fontSize:"0.55rem",
                                    background:"#fffbeb",color:"#b45309",border:"1px solid #fde68a",
                                    borderRadius:4,padding:"1px 4px",fontWeight:700}}>🔒{st.manCnt}</span>}
                                </div>
                                <div style={{fontSize:"0.57rem",color:col,marginTop:2,fontWeight:600}}>
                                  {patName(emp.patternId)} +{emp.offset}
                                </div>
                              </td>
                              {res.sch[i].map((cell,d)=>{
                                const t=cellT(cell), locked=isLocked(cell);
                                const isOpen=popup?.empIdx===i&&popup?.dayIdx===d;
                                return(
                                  <td key={d} style={{padding:"2px 1px",textAlign:"center",position:"relative"}}>
                                    <span
                                      onClick={()=>setPopup(p=>(p?.empIdx===i&&p?.dayIdx===d)?null:{empIdx:i,dayIdx:d})}
                                      title={locked
                                        ? `${emp.name} ${d+1}일 🔒 잠금 — 클릭하여 변경/잠금해제`
                                        : `${emp.name} ${d+1}일 — 클릭하여 수정`}
                                      style={{
                                        display:"inline-block",padding:"3px 4px",borderRadius:5,
                                        fontWeight:700,fontSize:"0.69rem",
                                        background:SHIFT[t]?.bg,color:SHIFT[t]?.txt,
                                        border:locked
                                          ? `2px solid #f59e0b`          /* 🔒 잠금=황색 */
                                          : `1.5px solid ${SHIFT[t]?.brd}`,
                                        cursor:"pointer",minWidth:22,textAlign:"center",
                                        outline:isOpen?`2px solid ${ACCENT}`:"none",
                                        boxShadow:locked?"0 0 0 2px #fde68a88":"none",
                                        position:"relative",
                                      }}>
                                      {t}
                                      {locked&&(
                                        <span style={{
                                          position:"absolute",top:-4,right:-4,
                                          fontSize:"0.45rem",lineHeight:1,
                                          background:"#f59e0b",color:"#fff",
                                          borderRadius:"50%",width:9,height:9,
                                          display:"flex",alignItems:"center",justifyContent:"center",
                                        }}>🔒</span>
                                      )}
                                    </span>
                                    {isOpen&&(
                                      <CellPopup cell={cell}
                                        onSelect={nt=>applyEdit(i,d,nt)}
                                        onClose={()=>setPopup(null)}/>
                                    )}
                                  </td>
                                );
                              })}
                              {/* Stats */}
                              <td style={{...sTd,color:TXTH,fontWeight:600}}>{st.work}일</td>
                              <td style={{...sTd,fontWeight:700,
                                color:over?"#dc2626":under?"#d97706":"#16a34a"}}>
                                {st.hours}h
                                {(over||under)&&<div style={{fontSize:"0.55rem",opacity:.8}}>
                                  {over?"▲":"▼"}{Math.abs(st.hours-dept.targetH)}h</div>}
                              </td>
                              <td style={{...sTd,color:ACCENT}}>{st.ju}회</td>
                              <td style={{...sTd,color:"#ea580c",fontWeight:600}}>
                                {st.se}회<div style={{fontSize:"0.55rem",color:"#fdba74"}}>(×1h)</div>
                              </td>
                              <td style={{...sTd,color:"#7c3aed",fontWeight:600}}>
                                {st.ya}회<div style={{fontSize:"0.55rem",color:"#c4b5fd"}}>(×7h)</div>
                              </td>
                              <td style={{...sTd,color:"#15803d"}}>{st.mo}회</td>
                              <td style={{...sTd,color:"#4d7c0f"}}>{st.ha}회</td>
                              <td style={{...sTd,color:"#be123c",fontSize:"0.69rem"}}>
                                <div style={{fontWeight:700}}>{st.bonus.toLocaleString()}</div>
                                <div style={{fontSize:"0.55rem",color:"#fca5a5"}}>원</div>
                              </td>
                            </tr>
                          );
                        })}
                        {/* 검색 결과 없음 표시 */}
                        {empSearch && dept.emps.filter(e=>e.name.includes(empSearch)).length===0 && (
                          <tr>
                            <td colSpan={D+9} style={{
                              padding:"24px",textAlign:"center",
                              color:TXTD,fontSize:"0.82rem",fontStyle:"italic",
                            }}>
                              🔍 "<strong style={{color:TXTH}}>{empSearch}</strong>" 검색 결과가 없습니다
                            </td>
                          </tr>
                        )}
                      </tbody>
                      <tfoot>
                        {[{sh:"석",req:dept.cst.se,col:"#ea580c",gte:false},
                          {sh:"야",req:dept.cst.ya,col:"#7c3aed",gte:false},
                          {sh:"주",req:dept.cst.juMin,col:ACCENT,gte:true}].map(({sh,req,col,gte},ri)=>(
                          <tr key={sh} style={{background:"#f8fafc",
                            borderTop:ri===0?`2px solid ${BORD}`:`1px solid #f1f5f9`}}>
                            <td style={{padding:"4px 12px",fontSize:"0.65rem",color:col,fontWeight:700,
                              position:"sticky",left:0,background:"#f8fafc",
                              borderRight:`1.5px solid ${BORD}`}}>
                              {sh}간{gte?"≥":"="}
                            </td>
                            {Array.from({length:D},(_,d)=>{
                              const cnt=res.sch.filter(r=>cellT(r[d])===sh).length;
                              const ok=gte?cnt>=req:cnt===req;
                              return(
                                <td key={d} style={{padding:"2px 1px",textAlign:"center",fontSize:"0.63rem",
                                  color:ok?TXTD:"#dc2626",fontWeight:ok?400:800}}>{cnt}</td>
                              );
                            })}
                            <td colSpan={8} style={{padding:"4px 9px",fontSize:"0.62rem",color:TXTD,
                              borderLeft:`1px solid ${BORD}`}}>
                              {gte?"≥":""}{req}명/일 · 월계 {res.sch.flat().filter(c=>cellT(c)===sh).length}회
                            </td>
                          </tr>
                        ))}
                        <tr style={{borderTop:`2px solid ${BORD}`,background:"#f8fafc"}}>
                          <td style={{padding:"7px 12px",fontWeight:700,color:TXT,fontSize:"0.72rem",
                            position:"sticky",left:0,background:"#f8fafc",
                            borderRight:`1.5px solid ${BORD}`}}>합계</td>
                          <td colSpan={D} style={{padding:"7px 8px",fontSize:"0.68rem",color:TXTD}}>
                            목표 {dept.targetH}h × {dept.emps.length}명 = {(dept.targetH*dept.emps.length).toLocaleString()}h
                          </td>
                          <td style={{...sTd,fontWeight:700,color:TXTH}}>{stats.reduce((a,s)=>a+s.work,0)}일</td>
                          <td style={{...sTd,fontWeight:700,color:"#16a34a"}}>{stats.reduce((a,s)=>a+s.hours,0)}h</td>
                          <td style={{...sTd,fontWeight:700,color:ACCENT}}>{stats.reduce((a,s)=>a+s.ju,0)}회</td>
                          <td style={{...sTd,fontWeight:700,color:"#ea580c"}}>{stats.reduce((a,s)=>a+s.se,0)}회</td>
                          <td style={{...sTd,fontWeight:700,color:"#7c3aed"}}>{stats.reduce((a,s)=>a+s.ya,0)}회</td>
                          <td style={{...sTd,fontWeight:700,color:"#15803d"}}>{stats.reduce((a,s)=>a+s.mo,0)}회</td>
                          <td style={{...sTd,fontWeight:700,color:"#4d7c0f"}}>{stats.reduce((a,s)=>a+s.ha,0)}회</td>
                          <td style={{...sTd,fontWeight:700,color:"#be123c",fontSize:"0.69rem"}}>
                            {stats.reduce((a,s)=>a+s.bonus,0).toLocaleString()}원
                          </td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  {/* Footer bar */}
                  <div className="no-print" style={{padding:"9px 14px",borderTop:`1px solid ${BORD}`,
                    background:"#f8fafc",display:"flex",gap:14,flexWrap:"wrap",alignItems:"center"}}>
                    <span style={{fontSize:"0.65rem",color:TXT,fontWeight:700}}>💰 야간수당 계산</span>
                    <span style={{fontSize:"0.67rem",color:TXTD}}>
                      (석×<span style={{color:"#ea580c",fontWeight:700}}>1h</span> + 야×<span style={{color:"#7c3aed",fontWeight:700}}>7h</span>) × {rate.toLocaleString()}원/h
                    </span>
                    <div style={{fontSize:"0.63rem",color:TXTD,marginLeft:"auto",
                      display:"flex",gap:5,alignItems:"center"}}>
                      <span style={{color:"#b45309",fontWeight:600}}>✏️ 셀 클릭</span>
                      <span>→ 근무 변경 가능 · 수동 수정된 셀은 황색 테두리 표시 · 수정 후 자동 재검증</span>
                    </div>
                  </div>
                </div>

              ):!res?(
                <div style={{background:CARD,border:`1.5px dashed ${BORD}`,borderRadius:12,
                  padding:"54px 20px",textAlign:"center",boxShadow:"0 1px 4px #0000000a"}}>
                  <div style={{fontSize:"2.8rem",marginBottom:12,opacity:.15}}>📋</div>
                  <div style={{color:TXT,fontWeight:600,marginBottom:6,fontSize:"0.95rem"}}>근무표 미생성</div>
                  <div style={{fontSize:"0.76rem",color:TXTD}}>
                    직원과 패턴을 설정한 후 <strong style={{color:ACCENT}}>⚡ 전체 근무표 생성</strong> 버튼을 클릭하세요
                  </div>
                </div>
              ):null}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
