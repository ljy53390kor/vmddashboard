import { useState, useEffect, useRef, useCallback } from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx-js-style";
import { dataClient } from "./lib/dataClient";
import { isVmddashboard } from "./lib/target";
import { loadHolidays } from "./lib/holidays";

const C = {
  bg0:         "var(--c-bg0)",
  bg1:         "var(--c-bg1)",
  bg2:         "var(--c-bg2)",
  bg3:         "var(--c-bg3)",
  border:      "var(--c-border)",
  borderAccent:"var(--c-border-accent)",
  orange:      "var(--c-orange)",
  orangeDim:   "var(--c-orange-dim)",
  blue:        "var(--c-blue)",
  green:       "var(--c-green)",
  text0:       "var(--c-text0)",
  text1:       "var(--c-text1)",
  text2:       "var(--c-text2)",
  surface:     "var(--c-surface)",
  overlay:     "var(--c-overlay)",
  surface2:    "var(--c-surface2)",
  mono:        "'JetBrains Mono', monospace",
};

// ─── 공휴일 Set (앱 시작 시 API로 채워짐, 로드 전에는 빈 Set) ──────────────
// isHoliday() 등 모듈 레벨 함수들이 이 Set을 참조함
// loadHolidays()가 완료되면 clear() + add()로 in-place 업데이트 → 리렌더 트리거
const KR_HOLIDAYS = new Set();

const ROLE_CONFIG = {
  admin:   { label: "관리자",  color: "#ff6b35", bg:"rgba(255,107,53,0.15)", icon:"⚙" },
  skn:     { label: "SKN",    color: "#3b82f6", bg:"rgba(59,130,246,0.15)",  icon:"✦" },
  regional:{ label: "지역본부", color: "#10b981", bg:"rgba(16,185,129,0.15)", icon:"◈" },
};


function toKey(y, m, d) {
  return `${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
}
function isHoliday(y, m, d) { return KR_HOLIDAYS.has(toKey(y,m,d)); }
function isWeekend(dow) { return dow === 0 || dow === 6; }

// 달력 셀이 어떤 배송 기간에라도 속하면 음영 처리
// (해당 날짜 기준 당월 or 전월의 21~익월20 구간 중 하나라도 포함되면 true)
function inPeriod(y, m, d) {
  const target = new Date(y, m, d);
  // m월 기준: m/21 ~ m+1/20
  if (target >= new Date(y, m, 21) && target <= new Date(y, m+1, 20)) return true;
  // m-1월 기준: m-1/21 ~ m/20
  if (target >= new Date(y, m-1, 21) && target <= new Date(y, m, 20)) return true;
  return false;
}

// AI 추천: 5회 배송일 자동 산정
// ① 기간(당월21~익월20) 내 매주 월요일 → 공휴일/주말이면 다음 평일로
// ② 5회 미만이면 마지막 주 금요일 → 공휴일이면 그 이전 가장 가까운 평일로 채움
function getRecommendedDates(baseYear, baseMonth) {
  // 배송 기간: 전월 21일 ~ 당월 20일
  // 예) 5월 달력 → 4/21 ~ 5/20
  const start = new Date(baseYear, baseMonth-1, 21);
  const end   = new Date(baseYear, baseMonth, 20);
  const MAX = 5;

  // 유효한 평일인지 (기간 내 + 주말/공휴일 아님)
  function isValidDay(dt) {
    const dow = dt.getDay();
    if (dow === 0 || dow === 6) return false;
    if (isHoliday(dt.getFullYear(), dt.getMonth(), dt.getDate())) return false;
    return dt >= start && dt <= end;
  }

  // 월요일 기반 추천: 공휴일이면 다음 평일
  const dates = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate()+1)) {
    if (d.getDay() === 1) {
      let rec = new Date(d);
      let safety = 0;
      while ((!isValidDay(rec) || isHoliday(rec.getFullYear(), rec.getMonth(), rec.getDate())) && safety < 7) {
        rec.setDate(rec.getDate()+1);
        safety++;
      }
      if (isValidDay(rec)) {
        const k = toKey(rec.getFullYear(), rec.getMonth(), rec.getDate());
        if (!dates.includes(k)) dates.push(k);
      }
    }
  }

  // 5회 미만이면 마지막 주 금요일(공휴일이면 이전 평일)로 채움
  if (dates.length < MAX) {
    // end(익월 20일) 기준으로 가장 가까운 금요일 탐색
    let fri = new Date(end);
    while (fri.getDay() !== 5) fri.setDate(fri.getDate()-1);

    let candidate = new Date(fri);
    let safety = 0;
    while (safety < 10) {
      if (isValidDay(candidate)) {
        const k = toKey(candidate.getFullYear(), candidate.getMonth(), candidate.getDate());
        if (!dates.includes(k)) {
          dates.push(k);
          break;
        }
      }
      candidate.setDate(candidate.getDate()-1);
      safety++;
    }
  }

  return dates.sort();
}

// ─── 초기 확정 배송일 (예시) ────────────────────────────────────────
const INITIAL_CONFIRMED = {
  "2025-05-26": { dow:"화", note:"", status:"confirmed" },
  "2025-06-01": { dow:"월", note:"", status:"confirmed" },
  "2025-06-08": { dow:"월", note:"", status:"confirmed" },
  "2025-06-15": { dow:"월", note:"", status:"confirmed" },
  "2025-06-19": { dow:"금", note:"", status:"confirmed" },
};
// status: "pending" = 회색 원 (임시 지정, SK 컨펌 대기)
//         "confirmed" = 빨간 원 (SK 컨펌 완료)
const DOW_KR = ["일","월","화","수","목","금","토"];

// ═══════════════════════════════════════════════════════════════════════
export default function App() {
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [theme, setTheme] = useState(() => localStorage.getItem('vmd-theme') || 'light');
  // 공휴일 API 로드 완료 여부 (true가 되면 달력 리렌더)
  const [holidaysReady, setHolidaysReady] = useState(false);

  // ── 공휴일 API 로드 ──────────────────────────────────────────────
  useEffect(() => {
    const curYear = new Date().getFullYear();
    const years   = [curYear - 1, curYear, curYear + 1, curYear + 2];
    loadHolidays(years)
      .then(set => {
        // KR_HOLIDAYS를 in-place 업데이트 → isHoliday() 등이 자동으로 새 데이터 사용
        KR_HOLIDAYS.clear();
        set.forEach(d => KR_HOLIDAYS.add(d));
        setHolidaysReady(true);
      })
      .catch(err => {
        console.error('[holidays] 로드 실패, 빈 Set으로 계속:', err);
        setHolidaysReady(true); // 실패해도 앱은 계속 동작
      });
  }, []);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('vmd-theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(t => t === 'dark' ? 'light' : 'dark');
  // confirmed, tempSelected를 최상위에서 관리 → 로그아웃해도 유지
  const [confirmed, setConfirmed] = useState(INITIAL_CONFIRMED);
  const [tempSelected, setTempSelected] = useState(new Set());
  const [showConfirmPopup, setShowConfirmPopup] = useState(false);
  // SKN이 컨펌한 시점의 날짜 스냅샷 { "기간키": [{key, dow}] }
  // SKN 컨펌 스냅샷 초기값: INITIAL_CONFIRMED 기반
  // periodKey 규칙: 날짜가 21일 이후면 "다음달" key, 20일 이하면 "당월" key
  // 예) 5/26 → key="2025-06", 6/1 → key="2025-06" (둘 다 6월 기준 기간)
  const buildInitialSnap = () => {
    const snap = {};
    Object.entries(INITIAL_CONFIRMED).forEach(([k,v]) => {
      const [ys,ms,ds] = k.split('-').map(Number);
      // 날짜가 21 이상이면 다음달 기준, 20 이하면 당월 기준
      let pk;
      if (ds >= 21) {
        // 당월 21~ → 익월 key
        const nm = ms === 12 ? 1 : ms + 1;
        const ny = ms === 12 ? ys + 1 : ys;
        pk = `${ny}-${String(nm).padStart(2,'0')}`;
      } else {
        // ~당월 20 → 당월 key
        pk = `${ys}-${String(ms).padStart(2,'0')}`;
      }
      if (!snap[pk]) snap[pk] = [];
      snap[pk].push({key:k, dow:v.dow});
    });
    Object.keys(snap).forEach(pk => snap[pk].sort((a,b)=>a.key.localeCompare(b.key)));
    return snap;
  };
  const [sknConfirmedSnap, setSknConfirmedSnap] = useState(buildInitialSnap);
  // 수정된 날짜 추적: {dateKey: true} - SKN 컨펌 후 추가/변경된 날짜
  const [updatedDates, setUpdatedDates] = useState({});
  // 수정 전 확정 날짜 보관: {periodKey: [{key, dow}]}
  const [preEditSnap, setPreEditSnap] = useState({});
  // SKN이 마지막으로 컨펌한 날짜 보관 (절대 지우지 않음, 수정 비교 before용)
  const [lastSKNConfirmed, setLastSKNConfirmed] = useState({});
  // SKN/공지 메일 발송 시점의 수정 내용 저장: {periodKey: {before:[...], after:[...]}}
  // "수정 배송일 공지" 버튼에서 이 정보를 사용
  const [lastRevisionData, setLastRevisionData] = useState({});
  const [mailRecipients, setMailRecipients] = useState([
    { id:1, name:"홍길동", email:"hong@telecom.co.kr", registeredAt:"2026-01-10 09:00:00" },
    { id:2, name:"김지수", email:"kim@telecom.co.kr",  registeredAt:"2026-01-10 09:00:00" },
  ]);
  const [sknRecipients, setSknRecipients] = useState([
    { id:1, name:"이민준", email:"minjun@sknetworks.co.kr", registeredAt:"2026-01-10 09:00:00" },
  ]);
  const [showMailPopup, setShowMailPopup] = useState(false);

  // ── 배송 수량 설정 (설정 메뉴) ──────────────────────────────────────────
  // rows: 구분/본부/기본값/그룹1~4(배수)/커스텀값/활성화여부
  const INITIAL_GROUPS = [
    // 지역본부 (g_dm: 도매포함 그룹 = 기본값+도매)
    { id:1,  구분:"지역본부", 본부:"수도권", 기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:2,  구분:"지역본부", 본부:"부산",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:3,  구분:"지역본부", 본부:"대구",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:4,  구분:"지역본부", 본부:"서부",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:5,  구분:"지역본부", 본부:"제주",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:6,  구분:"지역본부", 본부:"중부",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    // PS&M
    { id:7,  구분:"PS&M", 본부:"수도권", 기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:8,  구분:"PS&M", 본부:"부산",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:9,  구분:"PS&M", 본부:"대구",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:10, 구분:"PS&M", 본부:"서부",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:11, 구분:"PS&M", 본부:"제주",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:12, 구분:"PS&M", 본부:"중부",   기본값:0, 도매값:0, g1:1, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    // Biz (g1 비활성화, g_dm은 기본값만)
    { id:13, 구분:"Biz", 본부:"수도권", 기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:14, 구분:"Biz", 본부:"부산",   기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:15, 구분:"Biz", 본부:"대구",   기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:16, 구분:"Biz", 본부:"서부",   기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:17, 구분:"Biz", 본부:"제주",   기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    { id:18, 구분:"Biz", 본부:"중부",   기본값:0, 도매값:0, g1:null, g2:1, g3:1, g_dm:1, g4:1, custom:null, active:true },
    // 대형 (수도권만, g1/g2 비활성화, g_dm은 기본값)
    { id:19, 구분:"대형", 본부:"수도권", 기본값:0, 도매값:0, g1:null, g2:null, g3:1, g_dm:1, g4:1, custom:null, active:true },
    // 택배배송 (g1~g3/g_dm 비활성화, g4만 활성)
    { id:20, 구분:"택배배송", 본부:"-", 기본값:0, 도매값:0, g1:null, g2:null, g3:null, g_dm:null, g4:1, custom:null, active:true },
  ];
  const [shippingGroups, setShippingGroups] = useState(INITIAL_GROUPS);

  // ── 본사 배송 물품 (hqItems) ────────────────────────────────────────────
  const [hqItems, setHqItems] = useState([]);
  // ── GTM 취합 ────────────────────────────────────────────────────────────
  const [gtmWideColorData, setGtmWideColorData] = useState([]);
  const [gtmHangingData, setGtmHangingData] = useState([]);
  // 본부별 제출 현황: { widecolor: { [본부]: {submitted, submittedAt} }, hanging: {...} }
  const [gtmSubmissions, setGtmSubmissions] = useState({});
  // 와이드컬러 누락 매장 비교용 — 신설 매장 리스트 (전체 매장리스트가 아님)
  const [gtmNewStoreList, setGtmNewStoreList] = useState([]);
  // 라지그래픽 확정된 라운드 히스토리: [{id,name,label,data,locked,finalizedAt}] — S25/폴더블7/아이폰16/S26은 기본 제공(업로드 가능), 이후 "최종 확정"으로 추가되는 라운드는 locked:true
  const [gtmLargeGfxRounds, setGtmLargeGfxRounds] = useState(
    ["아이폰16","S25","폴더블7","아이폰17","S26"].map(name => ({ id:name, name, label:`${name} 선정매장`, data:[], locked:false, finalizedAt:null }))
  );
  // 라지그래픽 "신규 라지그래픽 선정" 작업용 초안(수정/삭제 가능, 최종 확정 전까지)
  const [gtmLargeGfxDraft, setGtmLargeGfxDraft] = useState([]);
  // 라지그래픽 매장별 전면사진: { [매장코드]: [{id,name,dataUrl}] } — 라운드 공통
  const [gtmLargeGfxPhotos, setGtmLargeGfxPhotos] = useState({});
  // 라지그래픽 탭별 O/X 비교 라운드 수동 선택: { [탭이름]: [name|null, name|null, name|null] } — 없으면 자동(최신 3개)
  const [gtmLargeGfxCompareOverrides, setGtmLargeGfxCompareOverrides] = useState({});
  // 설정에서 업로드한 매장 개별 목록 (누락 매장 비교용)
  const [storeList, setStoreList] = useState([]); // [{매장코드, 매장명, 대리점코드, 구분, 본부}]
  // ── 배송 수량 설정 표1/표2 상태 (메뉴 전환 시 유지) ─────────────────────
  const [shippingTable1, setShippingTable1] = useState([]);
  const [shippingStep, setShippingStep] = useState("upload");
  const [shippingCustomCols, setShippingCustomCols] = useState([]);
  const [shippingNextColId, setShippingNextColId] = useState(1);

  useEffect(() => {
    dataClient.auth.getSession().then(({ data: { session } }) => {
      if (session) fetchProfile(session.user.id);
      else setAuthLoading(false);
    });
    const { data: { subscription } } = dataClient.auth.onAuthStateChange((_event, session) => {
      if (session) fetchProfile(session.user.id);
      else { setUser(null); setAuthLoading(false); }
    });
    return () => subscription.unsubscribe();
  }, []);


  const fetchProfile = async (userId) => {
    const { data } = await dataClient.members.fetchProfile(userId);
    if (data) setUser({ id: data.employee_id, role: data.role, name: data.name });
    setAuthLoading(false);
  };

  const handleLogout = async () => {
    setUpdatedDates({});
    await dataClient.appData.saveKey('updated_dates', {});
    await dataClient.auth.signOut();
    setUser(null);
  };

  // ── Supabase 영속 저장 ────────────────────────────────────────────────
  const [appDataLoaded, setAppDataLoaded] = useState(false);

  useEffect(() => {
    if (!user) { setAppDataLoaded(false); return; }
    (async () => {
      const { data } = await dataClient.appData.loadAll();
      if (data) {
        const m = Object.fromEntries(data.map(r => [r.key, r.value]));
        if (m.confirmed)            setConfirmed(m.confirmed);
        if (m.temp_selected)        setTempSelected(new Set(m.temp_selected));
        if (m.skn_confirmed_snap)   setSknConfirmedSnap(m.skn_confirmed_snap);
        if (m.updated_dates)        setUpdatedDates(m.updated_dates);
        if (m.pre_edit_snap)        setPreEditSnap(m.pre_edit_snap);
        if (m.last_skn_confirmed)   setLastSKNConfirmed(m.last_skn_confirmed);
        // last_revision_data는 메모리 전용 — 로그인마다 초기화되도록 로드하지 않음
        if (m.mail_recipients)      setMailRecipients(m.mail_recipients);
        if (m.skn_recipients)       setSknRecipients(m.skn_recipients);
        if (m.shipping_groups)      setShippingGroups(m.shipping_groups);
        if (m.hq_items)             setHqItems(m.hq_items);
        if (m.store_list)           setStoreList(m.store_list);
        if (m.shipping_table1)      setShippingTable1(m.shipping_table1);
        if (m.shipping_step)        setShippingStep(m.shipping_step);
        if (m.shipping_custom_cols) setShippingCustomCols(m.shipping_custom_cols);
        if (m.shipping_next_col_id !== undefined) setShippingNextColId(m.shipping_next_col_id);
        if (m.gtm_widecolor_data)   setGtmWideColorData(m.gtm_widecolor_data);
        if (m.gtm_hanging_data)     setGtmHangingData(m.gtm_hanging_data);
        if (m.gtm_submissions)      setGtmSubmissions(m.gtm_submissions);
        if (m.gtm_new_store_list)   setGtmNewStoreList(m.gtm_new_store_list);
        // gtm_largegfx_rounds/draft는 예전 버전(객체 형태)으로 저장된 값이 남아있을 수 있어 배열인 경우만 반영
        if (Array.isArray(m.gtm_largegfx_rounds)) {
          const filtered = m.gtm_largegfx_rounds.filter(r => r && typeof r.name==='string' && Array.isArray(r.data));
          // 이름 중복(이전 마이그레이션 버그로 생긴 중복 탭)은 데이터가 더 많은 쪽만 남기고 정리
          const byName = new Map();
          filtered.forEach(r => {
            const prev = byName.get(r.name);
            if (!prev || r.data.length > prev.data.length) byName.set(r.name, r);
          });
          let validRounds = filtered.map(r=>r.name).filter((n,i,arr)=>arr.indexOf(n)===i).map(n=>byName.get(n));
          // 기본 라운드 5개는 항상 정해진 시간순(오래된→최신)으로 고정 — 저장된 순서가 꼬여있어도 여기서 바로잡음.
          // 이후 "최종 확정"으로 추가된 라운드는 기존 상대 순서 그대로 뒤에 붙임.
          const CANONICAL_ORDER = ["아이폰16","S25","폴더블7","아이폰17","S26"];
          const known = CANONICAL_ORDER.map(n => validRounds.find(r=>r.name===n)).filter(Boolean);
          const extra = validRounds.filter(r => !CANONICAL_ORDER.includes(r.name));
          validRounds = [...known, ...extra];
          if (validRounds.length > 0) setGtmLargeGfxRounds(validRounds);
        }
        if (Array.isArray(m.gtm_largegfx_draft)) setGtmLargeGfxDraft(m.gtm_largegfx_draft);
        if (m.gtm_largegfx_photos)  setGtmLargeGfxPhotos(m.gtm_largegfx_photos);
        if (m.gtm_largegfx_compare) setGtmLargeGfxCompareOverrides(m.gtm_largegfx_compare);
      }
      setAppDataLoaded(true);
    })();
  }, [user?.id]);

  const saveKey = useCallback(async (key, value) => {
    const { error } = await dataClient.appData.saveKey(key, value);
    if (error) console.error('[saveKey] upsert failed:', key, error);
  }, []);

  useEffect(() => { if (appDataLoaded) saveKey('confirmed', confirmed); }, [confirmed, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('temp_selected', [...tempSelected]); }, [tempSelected, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('skn_confirmed_snap', sknConfirmedSnap); }, [sknConfirmedSnap, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('updated_dates', updatedDates); }, [updatedDates, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('pre_edit_snap', preEditSnap); }, [preEditSnap, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('last_skn_confirmed', lastSKNConfirmed); }, [lastSKNConfirmed, appDataLoaded]);
  // last_revision_data: 메모리 전용이므로 DB 저장 없음
  useEffect(() => { if (appDataLoaded) saveKey('mail_recipients', mailRecipients); }, [mailRecipients, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('skn_recipients', sknRecipients); }, [sknRecipients, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('shipping_groups', shippingGroups); }, [shippingGroups, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('hq_items', hqItems); }, [hqItems, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('store_list', storeList); }, [storeList, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('shipping_table1', shippingTable1); }, [shippingTable1, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('shipping_step', shippingStep); }, [shippingStep, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('shipping_custom_cols', shippingCustomCols); }, [shippingCustomCols, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('shipping_next_col_id', shippingNextColId); }, [shippingNextColId, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_widecolor_data', gtmWideColorData); }, [gtmWideColorData, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_hanging_data', gtmHangingData); }, [gtmHangingData, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_submissions', gtmSubmissions); }, [gtmSubmissions, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_new_store_list', gtmNewStoreList); }, [gtmNewStoreList, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_largegfx_rounds', gtmLargeGfxRounds); }, [gtmLargeGfxRounds, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_largegfx_draft', gtmLargeGfxDraft); }, [gtmLargeGfxDraft, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_largegfx_photos', gtmLargeGfxPhotos); }, [gtmLargeGfxPhotos, appDataLoaded]);
  useEffect(() => { if (appDataLoaded) saveKey('gtm_largegfx_compare', gtmLargeGfxCompareOverrides); }, [gtmLargeGfxCompareOverrides, appDataLoaded]);

  // 예전에 base64로 hq_items JSON에 통째로 박혀있던 시안 이미지를 Storage로 옮기고 URL만 남긴다 (1회성, admin만).
  const sianMigratedRef = useRef(false);
  useEffect(() => {
    if (!appDataLoaded || user?.role !== "admin" || sianMigratedRef.current) return;
    const hasLegacy = hqItems.some(it => (it.sianImages||[]).some(img => img.dataUrl && !img.url));
    if (!hasLegacy) { sianMigratedRef.current = true; return; }
    sianMigratedRef.current = true;
    (async () => {
      const migrated = await Promise.all(hqItems.map(async (item) => {
        if (!(item.sianImages||[]).some(img => img.dataUrl && !img.url)) return item;
        const sianImages = await Promise.all(item.sianImages.map(async (img) => {
          if (img.url || !img.dataUrl) return img;
          try {
            const blob = await (await fetch(img.dataUrl)).blob();
            const file = new File([blob], img.name || "시안.png", { type: blob.type });
            const uploaded = await dataClient.storage.uploadImage(file);
            if (uploaded) return { id: img.id, name: img.name, url: uploaded.url, path: uploaded.path };
          } catch (err) {
            console.error("[sian migration] failed for", img.name, err);
          }
          return img;
        }));
        return { ...item, sianImages };
      }));
      setHqItems(migrated);
    })();
  }, [appDataLoaded, hqItems, user?.role]);

  if (authLoading) return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"center", height:"100vh", background:"#f4f6fb" }}>
      <div style={{ fontSize:16, color:"#888" }}>불러오는 중...</div>
    </div>
  );
  if (!user) return <AuthPage theme={theme} onToggle={toggleTheme} />;
  return (
    <>
      <Dashboard
        user={user}
        onLogout={handleLogout}
        theme={theme}
        onToggleTheme={toggleTheme}
        confirmed={confirmed}
        setConfirmed={setConfirmed}
        tempSelected={tempSelected}
        setTempSelected={setTempSelected}
        mailRecipients={mailRecipients}
        setMailRecipients={setMailRecipients}
        sknRecipients={sknRecipients}
        setSknRecipients={setSknRecipients}
        updatedDates={updatedDates}
        setUpdatedDates={setUpdatedDates}
        preEditSnap={preEditSnap}
        setPreEditSnap={setPreEditSnap}
        lastSKNConfirmed={lastSKNConfirmed}
        lastRevisionData={lastRevisionData}
        setLastRevisionData={setLastRevisionData}
        shippingGroups={shippingGroups}
        setShippingGroups={setShippingGroups}
        hqItems={hqItems}
        setHqItems={setHqItems}
        gtmWideColorData={gtmWideColorData}
        setGtmWideColorData={setGtmWideColorData}
        gtmHangingData={gtmHangingData}
        setGtmHangingData={setGtmHangingData}
        gtmSubmissions={gtmSubmissions}
        setGtmSubmissions={setGtmSubmissions}
        gtmNewStoreList={gtmNewStoreList}
        setGtmNewStoreList={setGtmNewStoreList}
        gtmLargeGfxRounds={gtmLargeGfxRounds}
        setGtmLargeGfxRounds={setGtmLargeGfxRounds}
        gtmLargeGfxDraft={gtmLargeGfxDraft}
        setGtmLargeGfxDraft={setGtmLargeGfxDraft}
        gtmLargeGfxPhotos={gtmLargeGfxPhotos}
        setGtmLargeGfxPhotos={setGtmLargeGfxPhotos}
        gtmLargeGfxCompareOverrides={gtmLargeGfxCompareOverrides}
        setGtmLargeGfxCompareOverrides={setGtmLargeGfxCompareOverrides}
        storeList={storeList}
        setStoreList={setStoreList}
        shippingTable1={shippingTable1}
        setShippingTable1={setShippingTable1}
        shippingStep={shippingStep}
        setShippingStep={setShippingStep}
        shippingCustomCols={shippingCustomCols}
        setShippingCustomCols={setShippingCustomCols}
        shippingNextColId={shippingNextColId}
        setShippingNextColId={setShippingNextColId}
        showMailPopup={showMailPopup}
        setShowMailPopup={setShowMailPopup}
        onSKConfirm={(snapEntries) => {
          // pending → confirmed
          setConfirmed(prev => {
            const next = {...prev};
            Object.keys(next).forEach(k => {
              if (next[k].status === "pending") next[k] = {...next[k], status:"confirmed"};
            });
            return next;
          });
          // 컨펌 시점 스냅샷 저장
          if (snapEntries && snapEntries.length > 0) {
            // 먼저 새로 추가된 날짜 계산 (현재 snap 기반)
            const newUpdated = {};
            setSknConfirmedSnap(prev => {
              const next = {...prev};
              snapEntries.forEach(({periodKey, entries}) => {
                const prevEntries = prev[periodKey] || [];
                const prevKeys = new Set(prevEntries.map(e=>e.key));
                if (prevKeys.size > 0) {
                  entries.forEach(e => {
                    if (!prevKeys.has(e.key)) newUpdated[e.key] = true;
                  });
                  // 삭제된 날짜도 추적 (before에 있었는데 after에 없는 것)
                  // → 추가 목적은 아니지만 수정 표시 목적으로 제거된 날짜 위치에 새 날짜 표시
                }
                next[periodKey] = entries;
              });
              return next;
            });
            // snap 업데이트 후 updatedDates 반영
            if (Object.keys(newUpdated).length > 0) {
              setUpdatedDates(prev => ({...prev, ...newUpdated}));
            }
          }
          // SKN 컨펌 완료 → preEditSnap, lastRevisionData 해당 기간 제거
          if (snapEntries && snapEntries.length > 0) {
            setPreEditSnap(prev => {
              const next = {...prev};
              snapEntries.forEach(({periodKey}) => { delete next[periodKey]; });
              return next;
            });
            // lastRevisionData는 유지 (수정 배송일 공지용)
          // SKN 컨펌 시점 날짜 저장 (수정 비교 before로 사용)
          if (snapEntries && snapEntries.length > 0) {
            setLastSKNConfirmed(prev => {
              const next = {...prev};
              snapEntries.forEach(({periodKey, entries}) => { next[periodKey] = entries; });
              return next;
            });
          }
          }
          setShowConfirmPopup(true);
        }}
      />
      {showConfirmPopup && (
        <div style={styles2.popupOverlay}>
          <div style={styles2.popupBox}>
            <div style={styles2.popupIcon}>✅</div>
            <div style={styles2.popupTitle}>배송일 컨펌이 완료되었습니다</div>
            <div style={styles2.popupSub}>확정된 배송일이 빨간 원으로 표시됩니다.</div>
            <button style={styles2.popupBtn} onClick={()=>setShowConfirmPopup(false)}>확인</button>
          </div>
        </div>
      )}
      {showMailPopup && (
        <MailSendPopup
          recipients={mailRecipients}
          mailType={typeof showMailPopup==="object" ? showMailPopup.type : showMailPopup}
          mailData={typeof showMailPopup==="object" ? showMailPopup : null}
          onClose={()=>setShowMailPopup(false)}
          onSent={(type)=>addNotification?.(
            type==="revision"
              ? `📧 [배송 공지] 수정 배송일 공지 메일 발송됨 (${mailRecipients.length}명)`
              : `📧 [배송 공지] 최초 배송일 공지 메일 발송됨 (${mailRecipients.length}명)`,
            'mail'
          )}
        />
      )}
    </>
  );
}

// ─── 로그인 ────────────────────────────────────────────────────────────
function ThemeToggle({ theme, onToggle, style={} }) {
  return (
    <button onClick={onToggle} style={{
      background:"var(--c-toggle-bg)",
      border:`1px solid ${C.border}`,
      color:"var(--c-toggle-color)",
      borderRadius:50, padding:"6px 14px",
      fontSize:13, cursor:"pointer", fontWeight:600,
      display:"flex", alignItems:"center", gap:6,
      transition:"all 0.2s",
      ...style,
    }}>
      {theme === 'dark' ? '☀️ 라이트' : '🌙 다크'}
    </button>
  );
}

function AuthPage({ theme, onToggle }) {
  const [tab, setTab] = useState("login");

  return (
    <div style={styles.loginBg}>
      {/* decorative blobs */}
      <div style={{position:"absolute",width:600,height:600,borderRadius:"50%",background:"radial-gradient(circle,rgba(255,107,53,0.07) 0%,transparent 70%)",top:"-15%",left:"-10%",pointerEvents:"none"}}/>
      <div style={{position:"absolute",width:400,height:400,borderRadius:"50%",background:"radial-gradient(circle,rgba(59,130,246,0.06) 0%,transparent 70%)",bottom:"-10%",right:"5%",pointerEvents:"none"}}/>

      {/* 테마 토글 — 우측 상단 */}
      <div style={{position:"absolute", top:20, right:24, zIndex:10}}>
        <ThemeToggle theme={theme} onToggle={onToggle} />
      </div>

      <div style={styles.loginCard} className="auth-card">
        {/* Logo */}
        <div style={styles.loginLogo}>📦</div>
        <div style={styles.loginTitle}>VMD 배송 관리</div>
        <div style={styles.loginSub}>전국 SK텔레콤 매장 배송 플랫폼</div>

        {tab !== "verify" && !isVmddashboard && (
          <div style={authStyles.tabRow}>
            <button style={{...authStyles.tab, ...(tab==="login"?authStyles.tabActive:{})}} onClick={()=>setTab("login")}>로그인</button>
            <button style={{...authStyles.tab, ...(tab==="signup"?authStyles.tabActive:{})}} onClick={()=>setTab("signup")}>회원가입</button>
          </div>
        )}

        {tab === "login"  && <LoginForm />}
        {tab === "signup" && <SignupForm onDone={()=>setTab("verify")} />}
        {tab === "verify" && <VerifyNotice onBack={()=>setTab("login")} />}

        <div style={{marginTop:20, fontSize:11, color:"rgba(148,163,184,0.4)", letterSpacing:0.5}}>VMD DELIVERY PLATFORM © SK Telecom</div>
      </div>
    </div>
  );
}

function LoginForm() {
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    if (!email || !pw) { setErr("이메일과 비밀번호를 입력하세요."); return; }
    setLoading(true); setErr("");
    const { error } = await dataClient.auth.signInWithPassword({ email, password: pw });
    if (error) {
      if (error.message.includes("Email not confirmed")) setErr("이메일 인증이 완료되지 않았습니다. 메일함을 확인해주세요.");
      else setErr("이메일 또는 비밀번호가 올바르지 않습니다.");
    }
    setLoading(false);
  };

  return (
    <>
      <div style={styles.loginField}>
        <label style={styles.loginLabel}>이메일</label>
        <input style={styles.loginInput} type="email" value={email}
          onChange={e=>{setEmail(e.target.value);setErr("");}}
          onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="이메일 입력" />
      </div>
      <div style={styles.loginField}>
        <label style={styles.loginLabel}>비밀번호</label>
        <input style={styles.loginInput} type="password" value={pw}
          onChange={e=>{setPw(e.target.value);setErr("");}}
          onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="비밀번호 입력" />
      </div>
      {err && <div style={styles.loginErr}>{err}</div>}
      <button style={{...styles.loginBtn, opacity:loading?0.7:1}} onClick={handleLogin} disabled={loading}>
        {loading ? "로그인 중..." : "로그인"}
      </button>
    </>
  );
}

function SignupForm({ onDone }) {
  const [form, setForm] = useState({ name:"", employee_id:"", phone:"", email:"", password:"", password2:"" });
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  const set = (k) => (e) => { setForm(f=>({...f,[k]:e.target.value})); setErr(""); };

  const handleSignup = async () => {
    const { name, employee_id, phone, email, password, password2 } = form;
    if (!name||!employee_id||!phone||!email||!password) { setErr("모든 항목을 입력하세요."); return; }
    if (password !== password2) { setErr("비밀번호가 일치하지 않습니다."); return; }
    if (password.length < 6) { setErr("비밀번호는 6자 이상이어야 합니다."); return; }
    setLoading(true); setErr("");
    const { error } = await dataClient.auth.signUp({
      email, password,
      options: { data: { name, employee_id, phone, role: "regional" } },
    });
    if (error) { setErr(error.message); setLoading(false); return; }
    setLoading(false);
    onDone();
  };

  const fields = [
    { key:"name",        label:"이름",   type:"text",     placeholder:"이름 입력" },
    { key:"employee_id", label:"사번",   type:"text",     placeholder:"사번 입력" },
    { key:"phone",       label:"연락처", type:"tel",      placeholder:"연락처 입력 (예: 010-1234-5678)" },
    { key:"email",       label:"이메일", type:"email",    placeholder:"이메일 입력" },
    { key:"password",    label:"비밀번호",type:"password", placeholder:"비밀번호 (6자 이상)" },
    { key:"password2",   label:"비밀번호 확인",type:"password",placeholder:"비밀번호 재입력" },
  ];

  return (
    <>
      {fields.map(f=>(
        <div key={f.key} style={styles.loginField}>
          <label style={styles.loginLabel}>{f.label}</label>
          <input style={styles.loginInput} type={f.type} value={form[f.key]}
            onChange={set(f.key)} placeholder={f.placeholder} />
        </div>
      ))}
      {err && <div style={styles.loginErr}>{err}</div>}
      <button style={{...styles.loginBtn, opacity:loading?0.7:1}} onClick={handleSignup} disabled={loading}>
        {loading ? "처리 중..." : "회원가입"}
      </button>
    </>
  );
}

function VerifyNotice({ onBack }) {
  return (
    <div style={{textAlign:"center", padding:"10px 0 6px", width:"100%"}}>
      <div style={{
        width:64, height:64, borderRadius:"50%",
        background:"rgba(16,185,129,0.12)", border:"1px solid rgba(16,185,129,0.3)",
        display:"flex", alignItems:"center", justifyContent:"center",
        fontSize:28, margin:"0 auto 16px",
      }}>✉️</div>
      <div style={{fontSize:17, fontWeight:700, color:C.text0, marginBottom:8}}>이메일 인증을 완료해주세요</div>
      <div style={{fontSize:13, color:C.text1, lineHeight:1.8, marginBottom:22}}>
        입력하신 이메일로 인증 메일을 발송했습니다.<br/>
        메일함에서 인증 링크를 클릭하면<br/>계정이 활성화됩니다.
      </div>
      <button style={{...styles.loginBtn, background:C.surface2, color:C.text1, boxShadow:"none", border:`1px solid ${C.border}`}} onClick={onBack}>
        ← 로그인으로 돌아가기
      </button>
    </div>
  );
}

const authStyles = {
  tabRow: {
    display:"flex", gap:4, marginBottom:22, width:"100%",
    background:C.surface, padding:4, borderRadius:10,
    border:`1px solid ${C.border}`,
  },
  tab: {
    flex:1, padding:"9px 0", background:"transparent", border:"none",
    cursor:"pointer", fontSize:13.5, color:C.text1, fontWeight:500,
    borderRadius:7, transition:"all 0.18s",
  },
  tabActive: {
    background:"linear-gradient(135deg,#ff6b35,#e8420a)",
    color:"#fff", fontWeight:700,
    boxShadow:"none",
  },
};

// ─── 대쉬보드 ───────────────────────────────────────────────────────────
function Dashboard({ user, onLogout, theme, onToggleTheme, confirmed, setConfirmed, tempSelected, setTempSelected, onSKConfirm, sknConfirmedSnap, mailRecipients, setMailRecipients, sknRecipients, setSknRecipients, showMailPopup, setShowMailPopup, shippingGroups, setShippingGroups, hqItems, setHqItems, gtmWideColorData, setGtmWideColorData, gtmHangingData, setGtmHangingData, gtmSubmissions, setGtmSubmissions, gtmNewStoreList, setGtmNewStoreList, gtmLargeGfxRounds, setGtmLargeGfxRounds, gtmLargeGfxDraft, setGtmLargeGfxDraft, gtmLargeGfxPhotos, setGtmLargeGfxPhotos, gtmLargeGfxCompareOverrides, setGtmLargeGfxCompareOverrides, storeList, setStoreList, updatedDates, setUpdatedDates, preEditSnap, setPreEditSnap, lastSKNConfirmed, lastRevisionData, setLastRevisionData, shippingTable1, setShippingTable1, shippingStep, setShippingStep, shippingCustomCols, setShippingCustomCols, shippingNextColId, setShippingNextColId }) {
  const [menu, setMenu] = useState("schedule");
  const role = user.role;
  const cfg  = ROLE_CONFIG[role];

  // ── 알림 시스템 (Supabase 공유 — 모든 계정에서 동일하게 표시) ──────────────
  const [notifications, setNotifications] = useState([]);
  const [notifOpen,     setNotifOpen]     = useState(false);
  const [toasts,        setToasts]        = useState([]);
  const notifPanelRef = useRef(null);

  // 최근 알림 로드 + Realtime 구독 (마운트 시 1회)
  useEffect(() => {
    // 최근 100개 로드
    dataClient.notifications.list(100).then(({ data }) => {
      if (data) {
        setNotifications(data.map(n => ({
          id: n.id, message: n.message, type: n.type,
          time: new Date(n.created_at), read: false,
        })));
      }
    });

    // 새 알림 구독 (Vercel: Supabase Realtime, Playground: 10초 폴링) → 모든 접속 클라이언트에 전파
    const unsubscribe = dataClient.notifications.subscribeInsert(({ new: n }) => {
      const notif = { id: n.id, message: n.message, type: n.type,
                      time: new Date(n.created_at), read: false };
      setNotifications(prev => {
        if (prev.some(p => p.id === n.id)) return prev; // 중복 방지
        return [notif, ...prev].slice(0, 100);
      });
      setToasts(prev => [...prev, { id: n.id, message: n.message, type: n.type }]);
      setTimeout(() => setToasts(prev => prev.filter(t => t.id !== n.id)), 3500);
    });

    return () => { unsubscribe(); };
  }, []);

  // 알림 추가 (Vercel: Supabase INSERT+Realtime 전파, Playground: 백엔드 저장+폴링 전파)
  const addNotification = useCallback(async (message, type = 'info') => {
    const { error } = await dataClient.notifications.insert(message, type);
    if (error) {
      // Supabase 실패 시 로컬 폴백
      const id = Date.now() + Math.random();
      const notif = { id, message, type, time: new Date(), read: false };
      setNotifications(prev => [notif, ...prev].slice(0, 100));
      setToasts(prev => [...prev, { id, message, type }]);
      setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 3500);
    }
  }, []);

  const markAllRead = () => setNotifications(prev => prev.map(n => ({...n, read:true})));
  const clearAll    = async () => {
    setNotifications([]);
    await dataClient.notifications.deleteAll();
  };
  const dismiss     = async (id) => {
    setNotifications(prev => prev.filter(n => n.id !== id));
    await dataClient.notifications.deleteOne(id);
  };
  const unread      = notifications.filter(n => !n.read).length;

  // 패널 외부 클릭 시 닫기
  useEffect(() => {
    if (!notifOpen) return;
    const handler = e => {
      if (notifPanelRef.current && !notifPanelRef.current.contains(e.target)) setNotifOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [notifOpen]);

  // 시간 포맷 (방금 전 / N분 전 / N시간 전 / 날짜)
  const fmtTime = (date) => {
    const diff = Math.floor((Date.now() - date) / 1000);
    if (diff < 60)   return '방금 전';
    if (diff < 3600) return `${Math.floor(diff/60)}분 전`;
    if (diff < 86400)return `${Math.floor(diff/3600)}시간 전`;
    return `${date.getMonth()+1}/${date.getDate()}`;
  };

  const typeStyle = {
    info:    { icon:'📋', color:'#3b82f6' },
    success: { icon:'✅', color:'#10b981' },
    warning: { icon:'⚠️', color:'#f59e0b' },
    delete:  { icon:'🗑', color:'#e85d26' },
    mail:    { icon:'📧', color:'#6366f1' },
    confirm: { icon:'✔️', color:'#10b981' },
    edit:    { icon:'✏️', color:'#8b5cf6' },
  };

  const menus = [
    { id:"schedule",  label:"배송 일정",    icon:"📅", roles: ["admin","skn","regional"] },
    { id:"hq",        label:"본사 배송 물품", icon:"🏢", roles: ["admin","skn","regional"] },
    { id:"regional",  label:"지역 배송 물품", icon:"📍", roles: ["admin","skn","regional"] },
    { id:"gtm",       label:"GTM 취합",    icon:"📊", roles: ["admin","regional"] },
    { id:"incoming",  label:"입고처",       icon:"📦", roles: ["admin","skn","regional"] },
    { id:"settings",  label:"설정",         icon:"⚙️", roles: ["admin"] },
  ].filter(m => m.roles.includes(role));

  return (
    <div style={styles.app}>
      {/* 사이드바 */}
      <aside style={styles.sidebar}>
        <div style={styles.sidebarLogo}>
          <div style={styles.sidebarLogoIcon}>📦</div>
          <div>
            <div style={styles.sidebarLogoText}>VMD</div>
            <div style={{fontSize:10, color:"rgba(148,163,184,0.5)", letterSpacing:1}}>DELIVERY</div>
          </div>
        </div>
        <nav style={styles.nav}>
          {menus.map(m=>(
            <button key={m.id}
              className={menu===m.id ? "" : "nav-btn-hover"}
              style={{...styles.navBtn, ...(menu===m.id?styles.navBtnActive:{})}}
              onClick={()=>setMenu(m.id)}>
              <span style={{fontSize:15}}>{m.icon}</span>
              <span>{m.label}</span>
            </button>
          ))}
        </nav>
        <div style={styles.sidebarFooter}>
          <div style={{display:"flex", alignItems:"center", gap:8, padding:"8px 10px", background:cfg.bg, borderRadius:10, border:`1px solid ${cfg.color}22`}}>
            <span style={{fontSize:18}}>{cfg.icon}</span>
            <div>
              <div style={{fontSize:11, color:cfg.color, fontWeight:700}}>{cfg.label}</div>
              <div style={{fontSize:11, color:"rgba(148,163,184,0.6)", fontFamily:"'JetBrains Mono', monospace"}}>{user.id}</div>
            </div>
          </div>
          <ThemeToggle theme={theme} onToggle={onToggleTheme} style={{width:"100%", justifyContent:"center"}} />
          <button style={styles.logoutBtn} onClick={onLogout}>로그아웃</button>
        </div>
      </aside>

      {/* ── 알림 벨 버튼 (우상단 고정) ─────────────────────────────────── */}
      <div ref={notifPanelRef}>
        <button onClick={()=>{ setNotifOpen(o=>!o); if(!notifOpen) markAllRead(); }}
          style={{ position:"fixed", top:16, right:20, zIndex:10000,
            width:38, height:38, borderRadius:"50%",
            background: unread>0 ? C.orange : "var(--c-surface2)",
            border:`1px solid ${unread>0 ? C.orange : C.border}`,
            cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
            fontSize:17, transition:"all 0.2s", boxShadow: unread>0 ? `0 0 0 3px ${C.orangeDim}` : "none",
          }}>
          🔔
          {unread > 0 && (
            <span style={{ position:"absolute", top:-4, right:-4,
              background:"#e85d26", color:"#fff", borderRadius:"50%",
              width:18, height:18, fontSize:10, fontWeight:800,
              display:"flex", alignItems:"center", justifyContent:"center",
              border:"2px solid var(--c-bg0)",
            }}>{unread > 99 ? '99+' : unread}</span>
          )}
        </button>

        {/* ── 알림 패널 ─────────────────────────────────────────────────── */}
        {notifOpen && (
          <div style={{ position:"fixed", top:62, right:16, zIndex:9999,
            width:340, maxHeight:"70vh", borderRadius:16,
            background:"var(--c-popup-bg)", border:`1px solid ${C.border}`,
            boxShadow:"0 16px 48px rgba(0,0,0,0.2)",
            display:"flex", flexDirection:"column", overflow:"hidden",
            animation:"fadeUp 0.2s cubic-bezier(.22,1,.36,1) both",
          }}>
            {/* 헤더 */}
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
              padding:"14px 16px 12px", borderBottom:`1px solid ${C.border}`, flexShrink:0 }}>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontWeight:800, fontSize:15, color:C.text0 }}>알림</span>
                {notifications.length > 0 && (
                  <span style={{ background:C.orangeDim, color:C.orange, borderRadius:20,
                    padding:"1px 8px", fontSize:11, fontWeight:700 }}>{notifications.length}</span>
                )}
              </div>
              <div style={{ display:"flex", gap:6 }}>
                {notifications.length > 0 && (
                  <button onClick={clearAll}
                    style={{ background:"none", border:`1px solid ${C.border}`, borderRadius:6,
                      padding:"3px 8px", fontSize:11, color:C.text2, cursor:"pointer" }}>전체 삭제</button>
                )}
              </div>
            </div>
            {/* 목록 */}
            <div style={{ overflowY:"auto", flex:1 }}>
              {notifications.length === 0 ? (
                <div style={{ padding:32, textAlign:"center", color:C.text2, fontSize:13 }}>
                  알림이 없습니다
                </div>
              ) : notifications.map(n => {
                const ts = typeStyle[n.type] || typeStyle.info;
                return (
                  <div key={n.id} style={{ display:"flex", gap:10, padding:"12px 16px",
                    borderBottom:`1px solid ${C.border}`,
                    background: n.read ? "transparent" : `${ts.color}08`,
                    transition:"background 0.3s",
                  }}>
                    <span style={{ fontSize:16, flexShrink:0, marginTop:1 }}>{ts.icon}</span>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontSize:13, color:C.text0, lineHeight:1.5,
                        wordBreak:"break-word" }}>{n.message}</div>
                      <div style={{ fontSize:11, color:C.text2, marginTop:3 }}>{fmtTime(n.time)}</div>
                    </div>
                    <button onClick={()=>dismiss(n.id)}
                      style={{ background:"none", border:"none", color:C.text2,
                        cursor:"pointer", fontSize:15, flexShrink:0, padding:"0 2px",
                        lineHeight:1, alignSelf:"flex-start" }}>×</button>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* ── 토스트 알림 ──────────────────────────────────────────────────── */}
      <div style={{ position:"fixed", top:62, right:16, zIndex:9998,
        display:"flex", flexDirection:"column", gap:8, pointerEvents:"none" }}>
        {toasts.map(t => {
          const ts = typeStyle[t.type] || typeStyle.info;
          return (
            <div key={t.id} style={{ background:"var(--c-popup-bg)",
              border:`1px solid ${C.border}`, borderLeft:`3px solid ${ts.color}`,
              borderRadius:10, padding:"10px 14px",
              display:"flex", alignItems:"center", gap:8, minWidth:240, maxWidth:320,
              boxShadow:"0 4px 16px rgba(0,0,0,0.15)",
              animation:"slideInRight 0.3s cubic-bezier(.22,1,.36,1) both",
            }}>
              <span style={{ fontSize:15 }}>{ts.icon}</span>
              <span style={{ fontSize:12, color:C.text0, lineHeight:1.4 }}>{t.message}</span>
            </div>
          );
        })}
      </div>

      {/* 메인 */}
      <main style={styles.main}>
        {menu === "schedule"
          ? <SchedulePage
              role={role}
              confirmed={confirmed}
              setConfirmed={setConfirmed}
              tempSelected={tempSelected}
              setTempSelected={setTempSelected}
              onSKConfirm={onSKConfirm}
              sknConfirmedSnap={sknConfirmedSnap}
              updatedDates={updatedDates}
              setUpdatedDates={setUpdatedDates}
              preEditSnap={preEditSnap}
              setPreEditSnap={setPreEditSnap}
              lastSKNConfirmed={lastSKNConfirmed}
              lastRevisionData={lastRevisionData}
              setLastRevisionData={setLastRevisionData}
              mailRecipients={mailRecipients}
              sknRecipients={sknRecipients}
              showMailPopup={showMailPopup}
              setShowMailPopup={setShowMailPopup}
              addNotification={addNotification}
            />
          : menu === "hq"
          ? <HQItemsPage
              hqItems={hqItems} setHqItems={setHqItems}
              shippingGroups={shippingGroups}
              shippingCustomCols={shippingCustomCols}
              confirmed={confirmed}
              role={role}
              addNotification={addNotification}
            />
          : menu === "settings"
          ? <SettingsPage
              mailRecipients={mailRecipients} setMailRecipients={setMailRecipients}
              sknRecipients={sknRecipients} setSknRecipients={setSknRecipients}
              shippingGroups={shippingGroups} setShippingGroups={setShippingGroups}
              shippingTable1={shippingTable1} setShippingTable1={setShippingTable1}
              shippingStep={shippingStep} setShippingStep={setShippingStep}
              shippingCustomCols={shippingCustomCols} setShippingCustomCols={setShippingCustomCols}
              shippingNextColId={shippingNextColId} setShippingNextColId={setShippingNextColId}
              setStoreList={setStoreList}
            />
          : menu === "gtm"
          ? <GTMPage
              gtmWideColorData={gtmWideColorData}
              setGtmWideColorData={setGtmWideColorData}
              gtmHangingData={gtmHangingData}
              setGtmHangingData={setGtmHangingData}
              gtmSubmissions={gtmSubmissions}
              setGtmSubmissions={setGtmSubmissions}
              gtmNewStoreList={gtmNewStoreList}
              setGtmNewStoreList={setGtmNewStoreList}
              gtmLargeGfxRounds={gtmLargeGfxRounds}
              setGtmLargeGfxRounds={setGtmLargeGfxRounds}
              gtmLargeGfxDraft={gtmLargeGfxDraft}
              setGtmLargeGfxDraft={setGtmLargeGfxDraft}
              gtmLargeGfxPhotos={gtmLargeGfxPhotos}
              setGtmLargeGfxPhotos={setGtmLargeGfxPhotos}
              gtmLargeGfxCompareOverrides={gtmLargeGfxCompareOverrides}
              setGtmLargeGfxCompareOverrides={setGtmLargeGfxCompareOverrides}
              role={role}
            />
          : <Placeholder label={menus.find(m=>m.id===menu)?.label} />}
      </main>
    </div>
  );
}

function Placeholder({ label }) {
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:14}}>
      <div style={{fontSize:36,opacity:0.3}}>🚧</div>
      <div style={{color:C.text2,fontSize:15,fontWeight:600,letterSpacing:0.3}}>{label} — 준비 중</div>
    </div>
  );
}

// ─── 배송 일정 페이지 ────────────────────────────────────────────────────
function SchedulePage({ role, confirmed, setConfirmed, tempSelected, setTempSelected, onSKConfirm, sknConfirmedSnap, updatedDates, setUpdatedDates, preEditSnap, setPreEditSnap, lastSKNConfirmed, lastRevisionData, setLastRevisionData, mailRecipients, sknRecipients, showMailPopup, setShowMailPopup, addNotification }) {
  const today = new Date();
  const [viewYear,  setViewYear]  = useState(today.getFullYear());
  const [viewMonth, setViewMonth] = useState(today.getMonth()); // 0-indexed
  // confirmed, tempSelected는 App에서 props로 받음 (로그아웃 후 재로그인해도 유지)
  // 배송일 수정 모드
  const [editMode, setEditMode]         = useState(false);
  const [editReleased, setEditReleased] = useState(new Set());
  const [editNewSel, setEditNewSel]     = useState(new Set());
  // 5회 초과 경고 팝업: 확인 시 추가할 key를 보관
  const [overLimitPopup, setOverLimitPopup] = useState(null);

  const isAdmin = role === "admin";
  const isSK    = role === "skn";

  // 기준 기간 (당월 21일~익월 20일)을 표시 중인 달 기준으로 계산
  const periodKeys = new Set();
  for (let y=viewYear, m=viewMonth; ;) {
    const start = new Date(y, m, 21);
    const end   = new Date(y, m+1, 20);
    for (let d=new Date(start); d<=end; d.setDate(d.getDate()+1)) {
      periodKeys.add(toKey(d.getFullYear(), d.getMonth(), d.getDate()));
    }
    break;
  }

  // 달력 계산
  const firstDay = new Date(viewYear, viewMonth, 1).getDay();
  const daysInMonth = new Date(viewYear, viewMonth+1, 0).getDate();

  const prevMonth = () => {
    if (viewMonth === 0) { setViewYear(y=>y-1); setViewMonth(11); }
    else setViewMonth(m=>m-1);
    // tempSelected 유지 (달 넘겨도 파란 원 보존)
  };
  const nextMonth = () => {
    if (viewMonth === 11) { setViewYear(y=>y+1); setViewMonth(0); }
    else setViewMonth(m=>m+1);
    // tempSelected 유지 (달 넘겨도 파란 원 보존)
  };

  // 관리자 날짜 클릭 (임시 지정 모드)
  const handleDayClick = (d) => {
    if (!isAdmin || editMode) return;
    const key = toKey(viewYear, viewMonth, d);
    if (confirmed[key]) return; // 이미 확정된 날짜
    if (!inPeriod(viewYear, viewMonth, d)) return;
    if (isHoliday(viewYear, viewMonth, d) || isWeekend(new Date(viewYear,viewMonth,d).getDay())) return;
    setTempSelected(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  // 배송일 임시 지정 확정 → status: "pending" (회색 원, SK 컨펌 대기)
  const handleTempConfirm = () => {
    if (tempSelected.size === 0) return;
    const updates = {};
    tempSelected.forEach(key => {
      const dow = DOW_KR[new Date(key).getDay()];
      updates[key] = { dow, note:"", status:"pending" };
    });
    setConfirmed(prev=>({...prev,...updates}));
    setTempSelected(new Set());
  };

  // AI 추천 → tempSelected(파란 원)에 표시, 임시 지정 버튼으로 확정
  const handleAIRecommend = () => {
    const recs = getRecommendedDates(viewYear, viewMonth);
    const newTemps = new Set(tempSelected);
    let added = 0;
    recs.forEach(key => {
      if (!confirmed[key] && !newTemps.has(key)) {
        newTemps.add(key);
        added++;
      }
    });
    setTempSelected(newTemps);
    if (added > 0) alert(`AI 추천 배송일 ${added}건을 파란 원으로 표시했습니다.\n확인 후 [배송일 임시 지정] 버튼을 눌러 확정하세요.`);
    else alert("이미 모든 추천 날짜가 선택되어 있습니다.");
  };

  // 수정 전 스냅샷 (수정 모드 진입 시 현재 기간 날짜 저장)
  // 배송일 수정 모드 진입
  const handleEditMode = () => {
    // 현재 기간(viewMonth 기준: 전월21~당월20) 내 confirmed 날짜 스냅샷
    // 수정 전 확정 날짜 저장 (SKN 메일 비교용 before)
    const pStart = new Date(viewYear, viewMonth-1, 21);
    const pEnd   = new Date(viewYear, viewMonth, 20, 23,59,59);
    const pk     = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
    const snap   = Object.entries(confirmed)
      .filter(([k,v]) => { const dt=new Date(k); return dt>=pStart && dt<=pEnd && v.status==="confirmed"; })
      .sort(([a],[b])=>a.localeCompare(b))
      .map(([k,v]) => ({ key:k, dow:v.dow }));
    // 이미 이번 수정 사이클의 스냅샷이 있으면 덮어쓰지 않음.
    // (SKN 컨펌 전 여러 번 수정 시 "최초 수정 전" 상태를 before로 유지하기 위해)
    // preEditSnap은 SKN 컨펌 시에만 초기화됨.
    setPreEditSnap(prev => {
      if ((prev[pk] || []).length > 0) return prev; // 이미 존재하면 유지
      return {...prev, [pk]: snap};
    });
    setEditMode(true);
    setEditReleased(new Set());
    setEditNewSel(new Set());
  };

  // 현재 기간(전월21~당월20) 내 총 배송 횟수 계산
  // 예: 6월 달력 → 5/21~6/20 기간
  // confirmed(해제 제외) + tempSelected + editNewSel
  const countPeriodDeliveries = (baseY, baseM, releasedSet, newSelSet) => {
    const start = new Date(baseY, baseM-1, 21);          // 전월 21일
    const end   = new Date(baseY, baseM, 20, 23, 59, 59); // 당월 20일
    let count = 0;
    Object.keys(confirmed).forEach(k => {
      if (releasedSet.has(k)) return;
      const dt = new Date(k);
      if (dt >= start && dt <= end) count++;
    });
    tempSelected.forEach(k => {
      const dt = new Date(k);
      if (dt >= start && dt <= end) count++;
    });
    newSelSet.forEach(k => {
      const dt = new Date(k);
      if (dt >= start && dt <= end) count++;
    });
    return count;
  };

  // 수정 모드에서 날짜 클릭
  const handleEditClick = (d) => {
    if (!editMode) return;
    const key = toKey(viewYear, viewMonth, d);
    const dow = new Date(viewYear,viewMonth,d).getDay();
    if (isHoliday(viewYear,viewMonth,d)||isWeekend(dow)) return;

    if (confirmed[key] && !editReleased.has(key)) {
      // 빨간 원 해제
      setEditReleased(prev=>{ const n=new Set(prev); n.add(key); return n; });
    } else if (confirmed[key] && editReleased.has(key)) {
      // 다시 활성화
      setEditReleased(prev=>{ const n=new Set(prev); n.delete(key); return n; });
    } else if (!confirmed[key]) {
      // 이미 선택된 거면 해제
      if (editNewSel.has(key)) {
        setEditNewSel(prev=>{ const n=new Set(prev); n.delete(key); return n; });
        return;
      }
      // 추가 전 5회 초과 체크 (클릭한 날짜가 속하는 배송 기간 기준)
      // d >= 21이면 다음달 배송 기간, d <= 20이면 이번달 배송 기간
      const periodBaseM = d >= 21 ? viewMonth + 1 : viewMonth;
      const currentCount = countPeriodDeliveries(viewYear, periodBaseM, editReleased, editNewSel);
      if (currentCount >= 5) {
        setOverLimitPopup(key); // 팝업 열고 key 보관
      } else {
        setEditNewSel(prev=>{ const n=new Set(prev); n.add(key); return n; });
      }
    }
  };

  // 수정 저장: 새 날짜는 pending(회색) → SKN 컨펌 대기 흐름 동일
  const handleEditSave = () => {
    setConfirmed(prev => {
      const next = {...prev};
      editReleased.forEach(k=>delete next[k]);
      editNewSel.forEach(k=>{ next[k]={dow:DOW_KR[new Date(k).getDay()], note:"", status:"pending"}; });
      return next;
    });
    setEditMode(false);
    setEditReleased(new Set());
    setEditNewSel(new Set());
    const label = `${viewYear}년 ${viewMonth+1}월`;
    alert("배송일이 수정되었습니다. SKN에 메일 보내기 버튼으로 수정 내용을 전달하세요.");
  };

  // SK 컨펌 → App에서 관리 (pending → confirmed + 팝업)
  const handleSKConfirm = () => {
    // 날짜 → 기간 키 변환 (day >= 21 → 다음달, day <= 20 → 당월)
    const dateToPeriodKey = (dateKey) => {
      const dt = new Date(dateKey);
      const y  = dt.getFullYear();
      const m  = dt.getMonth(); // 0-indexed
      const d  = dt.getDate();
      if (d >= 21) {
        // 다음달 키
        const nextM = m + 2; // 1-indexed 다음달
        const nextY = nextM > 12 ? y + 1 : y;
        const nm    = nextM > 12 ? 1 : nextM;
        return `${nextY}-${String(nm).padStart(2,"0")}`;
      } else {
        return `${y}-${String(m+1).padStart(2,"0")}`;
      }
    };

    // 모든 confirmed + tempSelected 날짜를 기간 키별로 그룹핑
    const periodMap = {}; // periodKey → Set of {key, dow}

    Object.entries(confirmed).forEach(([k, v]) => {
      const pk = dateToPeriodKey(k);
      if (!periodMap[pk]) periodMap[pk] = new Map();
      periodMap[pk].set(k, { key: k, dow: v.dow });
    });
    tempSelected.forEach(k => {
      const pk = dateToPeriodKey(k);
      if (!periodMap[pk]) periodMap[pk] = new Map();
      periodMap[pk].set(k, { key: k, dow: DOW_KR[new Date(k).getDay()] });
    });

    // 각 기간별로 sorted entries 생성
    const snapEntries = Object.entries(periodMap).map(([periodKey, map]) => {
      const sorted = [...map.values()].sort((a,b) => a.key.localeCompare(b.key));
      return { periodKey, entries: sorted };
    });

    // 컨펌 직전 pending 날짜만 알림에 표시 (모든 confirmed 날짜가 아님)
    const fmtShort = k => { const [,m,d] = k.split('-'); return `${parseInt(m)}/${parseInt(d)}`; };
    const justConfirmed = Object.entries(confirmed)
      .filter(([, v]) => v.status === "pending")
      .sort(([a],[b]) => a.localeCompare(b))
      .map(([k, v]) => ({ key: k, dow: v.dow }));
    onSKConfirm(snapEntries);
    const dateList = justConfirmed.map(e => `${fmtShort(e.key)}(${e.dow})`).join(', ');
    addNotification?.(`✅ [배송 일정] 컨펌 완료 — ${dateList}`, 'confirm');
  };

  // SKN에 메일 보내기
  const [sknMailData, setSknMailData] = useState(null);
  const [sknPendingPopup, setSknPendingPopup] = useState(false);
  const [sknPendingDates, setSknPendingDates] = useState([]);

  // SKN 로그인 시 회색 원(pending) 있으면 알림 팝업
  useEffect(() => {
    if (role !== "skn") return;
    const pendingEntries = Object.entries(confirmed)
      .filter(([,v]) => v.status === "pending")
      .sort(([a],[b]) => a.localeCompare(b))
      .map(([k,v]) => ({ key:k, dow:v.dow }));
    if (pendingEntries.length > 0) {
      setSknPendingDates(pendingEntries);
      setSknPendingPopup(true);
    }
  }, [role]);

  const handleSendSKNMail = () => {
    const pStart = new Date(viewYear, viewMonth-1, 21);
    const pEnd   = new Date(viewYear, viewMonth, 20, 23,59,59);

    const periodDates = Object.entries(confirmed)
      .filter(([k]) => { const dt=new Date(k); return dt>=pStart && dt<=pEnd; })
      .sort(([a],[b])=>a.localeCompare(b))
      .map(([k,v]) => ({ key:k, dow:v.dow, status:v.status }));

    const hasPending   = periodDates.some(r => r.status === "pending");
    const hasConfirmed = periodDates.some(r => r.status === "confirmed");
    const allPending   = hasPending && !hasConfirmed;

    const pk = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
    // before = 수정 모드 진입 전 저장된 snap (없으면 현재 빨간원들)
    const storedBefore = (preEditSnap || {})[pk] || [];
    const before = storedBefore.length > 0
      ? storedBefore
      : periodDates.filter(r => r.status === "confirmed");

    if (allPending) {
      // 회색원만 → 임시 지정 검토
      setSknMailData({ type:"initial", dates:periodDates });
    } else if (hasPending && hasConfirmed) {
      // 회색원+빨간원 섞임 → 수정 검토
      const pk2 = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
      // before 우선순위:
      //  1) lastSKNConfirmed  — SKN이 직접 컨펌한 날짜 (가장 권위있는 기준)
      //  2) preEditSnap       — 수정 모드 진입 시점 스냅샷 (삭제된 날짜 포함)
      //  3) confirmedNow      — 현재 남아있는 확정 날짜 (최후 폴백)
      const sknBefore    = (lastSKNConfirmed || {})[pk2] || [];
      const preSnap      = (preEditSnap || {})[pk2] || [];
      const confirmedNow = periodDates.filter(r => r.status === "confirmed");
      const before2 = sknBefore.length > 0 ? sknBefore
                    : preSnap.length > 0    ? preSnap
                    : confirmedNow;
      setLastRevisionData(prev => ({...prev, [pk2]: {before: before2, after:periodDates}}));
      // 새로 추가/변경된 날짜를 updatedDates에 기록 (DB 저장되어 재로그인 후에도 표시)
      const beforeKeys2 = new Set(before2.map(r => r.key));
      const newlyAdded2 = periodDates.filter(r => !beforeKeys2.has(r.key));
      if (newlyAdded2.length > 0) {
        const upd = Object.fromEntries(newlyAdded2.map(r => [r.key, true]));
        setUpdatedDates(prev => ({...prev, ...upd}));
      }
      setSknMailData({ type:"revision", before: before2, after:periodDates });
    } else {
      // pending 없음 → 순수 삭제 여부 확인
      // preEditSnap 또는 lastSKNConfirmed 기준으로 사라진 날짜가 있으면 수정 검토 요청
      const pk2       = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
      const sknBefore = (lastSKNConfirmed || {})[pk2] || [];
      const preSnap   = (preEditSnap || {})[pk2] || [];
      const baseline  = sknBefore.length > 0 ? sknBefore : preSnap;
      const curKeys   = new Set(periodDates.map(r => r.key));
      const hasDeleted = baseline.some(r => !curKeys.has(r.key));

      if (hasDeleted) {
        // 삭제된 날짜 있음 → 수정 검토 (after = 현재 확정 날짜들)
        setLastRevisionData(prev => ({...prev, [pk2]: {before: baseline, after:periodDates}}));
        setSknMailData({ type:"revision", before: baseline, after:periodDates });
      } else {
        // 실제로 변경 없음 → 전달 없음
        setSknMailData({ type:"none" });
      }
    }
  };

  // 배송 기간(전월21일~당월20일) 내 확정일 목록
  const periodStart = new Date(viewYear, viewMonth-1, 21);
  const periodEnd   = new Date(viewYear, viewMonth, 20, 23, 59, 59);

  // SKN용: 모든 기간 중 5회 초과 기간 탐지
  // confirmed + tempSelected 기준으로 각 기준월별 카운트
  const overLimitPeriods = (() => {
    if (role !== "skn") return [];
    // 기준월 범위: confirmed 키에서 추출
    const baseMonths = new Set();
    [...Object.keys(confirmed), ...Array.from(tempSelected)].forEach(k => {
      const dt = new Date(k);
      // 이 날짜가 어느 기준월 기간에 속하는지 (당월 21일 이후면 당월, 이하면 전월)
      const y = dt.getFullYear(), m = dt.getMonth(), d = dt.getDate();
      if (d >= 21) baseMonths.add(`${y}-${m}`);
      else         baseMonths.add(m === 0 ? `${y-1}-11` : `${y}-${m-1}`);
    });
    const over = [];
    baseMonths.forEach(bk => {
      const [by, bm] = bk.split("-").map(Number);
      const s = new Date(by, bm, 21), e = new Date(by, bm+1, 20, 23,59,59);
      let cnt = 0;
      [...Object.keys(confirmed), ...Array.from(tempSelected)].forEach(k => {
        const dt = new Date(k); if (dt >= s && dt <= e) cnt++;
      });
      if (cnt > 5) over.push({year:by, month:bm, count:cnt});
    });
    return over;
  })();

  // 확정 일정 표: SKN 컨펌 전까지 고정 (임시저장/수정 상태와 무관)
  // 우선순위:
  //  1) lastSKNConfirmed  — SKN이 실제로 컨펌한 날짜 (가장 권위있는 기준)
  //  2) sknConfirmedSnap  — 앱 시작 시 INITIAL_CONFIRMED 기반 (또는 이전 컨펌 시점)
  //  3) preEditSnap       — 수정 모드 최초 진입 시점 스냅샷 (don't-overwrite 적용됨)
  //                         → SKN 컨펌 이력이 없을 때 "수정 전 확정 상태"를 유지해줌
  // * preEditSnap은 "don't overwrite" 규칙으로 수정 사이클 내에서 안정적으로 유지됨
  const pk4Table = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
  const sknSnap     = (lastSKNConfirmed  || {})[pk4Table] || [];
  const sknConfSnap = (sknConfirmedSnap  || {})[pk4Table] || [];
  const preSnap4Tbl = (preEditSnap       || {})[pk4Table] || [];
  const snapForTable = sknSnap.length > 0     ? sknSnap
                     : sknConfSnap.length > 0 ? sknConfSnap
                     : preSnap4Tbl;
  const monthConfirmed = snapForTable.length > 0
    ? snapForTable.map(e => [e.key, { dow:e.dow, status:"confirmed", note: confirmed[e.key]?.note||"" }])
        .filter(([k]) => { const dt=new Date(k); return dt>=periodStart && dt<=periodEnd; })
        .sort(([a],[b])=>a.localeCompare(b))
    : Object.entries(confirmed)
        .filter(([k, v]) => {
          const dt = new Date(k);
          return dt >= periodStart && dt <= periodEnd && v.status === "confirmed";
        })
        .sort(([a],[b])=>a.localeCompare(b));

  // 인접 달 날짜들 (실제 year/month/day 포함)
  const prevDaysInMonth = new Date(viewYear, viewMonth, 0).getDate();
  const prevYear  = viewMonth === 0 ? viewYear-1 : viewYear;
  const prevMon   = viewMonth === 0 ? 11 : viewMonth-1;
  const nextYear  = viewMonth === 11 ? viewYear+1 : viewYear;
  const nextMon   = viewMonth === 11 ? 0 : viewMonth+1;

  const calCells = [];
  for (let i=0;i<firstDay;i++) {
    const d = prevDaysInMonth-firstDay+1+i;
    calCells.push({day:d, cur:false, y:prevYear, m:prevMon});
  }
  for (let d=1;d<=daysInMonth;d++) calCells.push({day:d, cur:true, y:viewYear, m:viewMonth});
  const remaining = 42-calCells.length;
  for (let d=1;d<=remaining;d++) calCells.push({day:d, cur:false, y:nextYear, m:nextMon});

  const monthLabel = `${viewYear}년 ${viewMonth+1}월`;

  return (
    <>
    <div style={styles.page}>
      <div style={styles.pageHeader}>
        <h1 style={styles.pageTitle}>배송 일정</h1>
        <div style={styles.periodNote}>*배송일은 당월 21일~익월 20일을 한 달로 계산합니다</div>
      </div>

      <div style={styles.scheduleLayout}>
        {/* 달력 영역 */}
        <div style={styles.calSection}>
          {/* 달력 툴바: 월 이동 + 액션 버튼 */}
          <div style={styles.calToolbar}>
            <div style={styles.monthNav}>
              <button style={styles.navArrow} onClick={prevMonth}>‹</button>
              <span style={styles.monthLabel}>{monthLabel}</span>
              <button style={styles.navArrow} onClick={nextMonth}>›</button>
            </div>
            <div style={styles.calActions}>
              {isAdmin && !editMode && (
                <>
                  <button style={styles.aiBtn} onClick={handleAIRecommend}>⚡ AI 추천</button>
                  <button style={{...styles.tempBtn, opacity:tempSelected.size===0?0.5:1}}
                    onClick={handleTempConfirm} disabled={tempSelected.size===0}>
                    임시 저장{tempSelected.size>0?` (${tempSelected.size})`:""}
                  </button>
                  <button style={styles.editBtn} onClick={handleEditMode}>수정</button>
                  <button style={{...styles.aiBtn, background:"#0066cc"}} onClick={handleSendSKNMail}>
                    📨 SKN 메일
                  </button>
                </>
              )}
              {isAdmin && editMode && (
                <>
                  <button style={{...styles.aiBtn, background:"#e85d26"}} onClick={handleEditSave}>💾 저장</button>
                  <button style={{...styles.tempBtn, background:"#aaa"}}
                    onClick={()=>{
                  setEditMode(false);setEditReleased(new Set());setEditNewSel(new Set());
                  // 취소 시 preEditSnap도 초기화
                  const pk=`${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
                  setPreEditSnap(prev=>{const n={...prev};delete n[pk];return n;});
                }}>취소</button>
                </>
              )}
              {isSK && (
                <button style={styles.skBtn} onClick={handleSKConfirm}>✓ 배송일 컨펌</button>
              )}
            </div>
          </div>

          {/* 달력 */}
          <div style={styles.calendar}>
            <div style={styles.calHead}>
              {["일","월","화","수","목","금","토"].map((d,i)=>(
                <div key={d} style={{...styles.calHeadCell, color:i===0?"#fb7185":i===6?"#60a5fa":C.text2}}>{d}</div>
              ))}
            </div>
            <div style={styles.calBody}>
              {calCells.map((cell, idx) => {
                const dow = idx % 7;
                const isCurrent = cell.cur;
                const cellY = cell.y, cellM = cell.m;
                const key    = toKey(cellY, cellM, cell.day);
                const isConf     = confirmed[key] && !editReleased.has(key);
                const isTemp     = tempSelected.has(key);
                const isEditNew  = editMode && editNewSel.has(key);
                const isReleased = editReleased.has(key);
                const isHol      = isHoliday(cellY, cellM, cell.day);
                const isWknd     = isWeekend(dow);
                // 배송 기간: 현재 보는 달 기준 21일 ~ 다음달 20일 (인접달 포함)
                const isPeriod   = inPeriod(cellY, cellM, cell.day);
                const isToday    = key === toKey(today.getFullYear(), today.getMonth(), today.getDate());

                let circleColor = null;
                if (isConf) {
                  const st = confirmed[key]?.status;
                  if (st === "confirmed") {
                    circleColor = "#ff6b35";
                  } else if (st === "pending") {
                    circleColor = (role === "admin" || role === "skn") ? "#475569" : null;
                  }
                }
                if ((isTemp || isEditNew) && role === "admin") circleColor = "#3b82f6";
                if (isReleased) circleColor = "transparent";

                const clickable = isCurrent && isAdmin && !isHol && !isWknd;
                return (
                  <div key={idx} style={{
                    ...styles.calCell,
                    opacity: isCurrent ? 1 : 0.3,
                    background: isPeriod && role === "admin" ? "rgba(59,130,246,0.05)" : "transparent",
                    cursor: clickable ? "pointer" : "default",
                  }}
                    onClick={() => {
                      if (!isCurrent || !isAdmin) return;
                      if (editMode) handleEditClick(cell.day);
                      else handleDayClick(cell.day);
                    }}
                  >
                    {isToday && <div style={styles.todayUnderline} />}
                    {circleColor && !isReleased && (
                      <div style={{
                        ...styles.circle,
                        background: circleColor,
                        border: "none",
                      }} />
                    )}
                    <span style={{
                      ...styles.dayNum,
                      color: isHol||dow===0?"#fb7185": dow===6?"#60a5fa":C.text0,
                      fontWeight: isConf||isTemp?"700":"400",
                      position:"relative", zIndex:2,
                    }}>
                      {cell.day}
                    </span>
                    {isHol && <div style={styles.holDot} title="공휴일">🔴</div>}
                    {isCurrent && cell.day === 20 && (
                      <svg style={{position:"absolute",bottom:3,right:3,zIndex:5}} width="10" height="10" viewBox="0 0 10 10">
                        <path d="M8,2 L8,8 L2,8" stroke="#1d6fa4" strokeWidth="1.8" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    )}
                    {isCurrent && cell.day === 21 && (
                      <svg style={{position:"absolute",top:3,left:3,zIndex:5}} width="10" height="10" viewBox="0 0 10 10">
                        <path d="M2,8 L2,2 L8,2" stroke="#1d6fa4" strokeWidth="1.8" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    )}

                  </div>
                );
              })}
            </div>
          </div>

          {/* 범례 */}
          <div style={styles.legend}>
            <LegendItem color="#ff6b35" label="SKN 컨펌 완료" />
            {(role === "admin" || role === "skn") && <LegendItem color="#64748b" label="임시 지정 (SKN 대기)" />}
            {role === "admin" && <LegendItem color="#3b82f6" label="선택 중" />}
            {role === "admin" && <span style={{color:"#888",fontSize:12}}>음영: 배송 기간 (21일~익월20일) &nbsp;
              <svg style={{display:"inline",verticalAlign:"middle"}} width="12" height="12" viewBox="0 0 12 12">
                <path d="M10,2 L10,10 L2,10" stroke="#1d6fa4" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
              </svg> 마감 &nbsp;
              <svg style={{display:"inline",verticalAlign:"middle"}} width="12" height="12" viewBox="0 0 12 12">
                <path d="M2,10 L2,2 L10,2" stroke="#1d6fa4" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
              </svg> 시작
            </span>}
          </div>

          {/* 수정 모드 안내 */}
          {editMode && (
            <div style={{marginTop:10, padding:"10px 14px", background:"#fffbe6",
              borderRadius:8, fontSize:12, color:"#7a6000", border:"1px solid #ffe066"}}>
              ✏️ <b>수정 모드</b>: 기존 원 클릭 시 해제 / 빈 날짜 클릭 시 파란 원(신규 추가) → 저장 시 회색 원(SKN 컨펌 대기)
            </div>
          )}

          {/* SKN: 5회 초과 경고 */}
          {isSK && overLimitPeriods.length > 0 && (
            <div style={{marginTop:10, padding:"12px 14px", background:"#fff0f0",
              borderRadius:8, fontSize:13, color:"#c00", border:"1px solid #fcc",
              display:"flex", alignItems:"center", gap:8}}>
              ⚠️ <b>5회 초과하는 배송 일정이 있습니다</b>
              <span style={{fontSize:12, color:"#888"}}>
                ({overLimitPeriods.map(p=>`${p.year}년 ${p.month+1}월 21일 ~ ${p.month+2}월 20일 기준 ${p.count}회`).join(", ")})
              </span>
            </div>
          )}
        </div>

        {/* 확정 일정 테이블 */}
        <div style={styles.tableSection}>
          <div style={styles.tableTitleRow}>
            <div style={{marginBottom:12}}>
              <span style={styles.tableTitle}>
                {viewMonth === 0
                  ? `${viewYear-1}년 12월 21일 ~ ${viewYear}년 1월 20일 확정 일정`
                  : `${viewYear}년 ${viewMonth}월 21일 ~ ${viewMonth+1}월 20일 확정 일정`}
              </span>
            </div>
            {isAdmin && (() => {
              const periodEntries = Object.entries(confirmed).filter(([k]) => {
                const dt = new Date(k);
                return dt >= new Date(viewYear, viewMonth-1, 21) && dt <= new Date(viewYear, viewMonth, 20, 23,59,59);
              });
              const allConfirmed = periodEntries.length > 0 && periodEntries.every(([,v]) => v.status === "confirmed");
              return (
                <div style={{display:"flex", flexDirection:"column", gap:8, marginBottom:16}}>
                  {(() => {
                    const pStart2 = new Date(viewYear, viewMonth-1, 21);
                    const pEnd2   = new Date(viewYear, viewMonth, 20, 23,59,59);
                    const pKey2   = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
                    const snapB   = (sknConfirmedSnap||{})[pKey2]||[];
                    const curDates = periodEntries.map(([k,v])=>({key:k,dow:v.dow,status:v.status}));
                    return (<>
                      <button style={{
                        ...styles.aiBtn,
                        background: allConfirmed ? "#2e8b57" : "#cccccc",
                        color: allConfirmed ? "#fff" : "#777",
                        cursor: allConfirmed ? "pointer" : "not-allowed",
                        width:"100%", padding:"9px", fontSize:12, borderRadius:8,
                      }}
                        disabled={!allConfirmed}
                        onClick={() => allConfirmed && setShowMailPopup({
                          type:"initial", dates:curDates
                        })}>
                        📧 최초 배송일 공지
                      </button>
                      {(() => {
                        const pk5r = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
                        const hasRevData = !!(lastRevisionData||{})[pk5r];
                        const revActive = allConfirmed || hasRevData;
                        return (
                      <button style={{
                        ...styles.aiBtn,
                        background: revActive ? "#e85d26" : "#cccccc",
                        color: revActive ? "#fff" : "#777",
                        cursor: revActive ? "pointer" : "not-allowed",
                        width:"100%", padding:"9px", fontSize:12, borderRadius:8,
                      }}
                        disabled={!revActive}
                        onClick={() => {
                          if (!revActive) return;
                          // lastRevisionData: SKN 메일 발송 시점의 수정 내용 사용
                          const pk5 = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
                          const rev = (lastRevisionData||{})[pk5];
                          if (rev) {
                            setShowMailPopup({ type:"revision", before:rev.before, after:rev.after });
                          } else {
                            // lastRevisionData 없으면 preEditSnap 기반
                            const snapB2 = (preEditSnap||{})[pk5] || [];
                            const before2 = snapB2.length > 0 ? snapB2 : curDates.filter(r=>r.status==="confirmed");
                            setShowMailPopup({ type:"revision", before:before2, after:curDates });
                          }
                        }}>
                        📧 수정 배송일 공지
                      </button>
                        );
                      })()}
                    </>);
                  })()}
                </div>
              );
            })()}
          </div>
          <table style={styles.table}>
            <thead>
              <tr>
                <th style={styles.th}>날짜</th>
                <th style={styles.th}>요일</th>
                <th style={{...styles.th,flex:2}}>비고</th>
              </tr>
            </thead>
            <tbody>
              {monthConfirmed.length === 0 && (
                <tr><td colSpan={3} style={{...styles.td,textAlign:"center",color:"#aaa"}}>확정된 배송일 없음</td></tr>
              )}
              {monthConfirmed.map(([k,v])=>{
                const [,m,d] = k.split("-");
                // preEditSnap에 없던 날짜 = 이번 수정으로 새로 추가된 날짜
                // UPDATED: lastRevisionData에서 after에만 있는 날짜 (새로 추가/변경된 날짜)
                const pk6 = `${viewYear}-${String(viewMonth+1).padStart(2,"0")}`;
                const rev6 = (lastRevisionData||{})[pk6];
                const isUpdated = (() => {
                  if (rev6) {
                    const beforeKeys6 = new Set(rev6.before.map(r=>r.key));
                    return !beforeKeys6.has(k);
                  }
                  // lastRevisionData 없으면 DB에 저장된 updatedDates 사용
                  return !!(updatedDates && updatedDates[k]);
                })();
                return (
                  <tr key={k} style={{background: isUpdated ? "#fffbe6" : "transparent"}}>
                    <td style={styles.td}>
                      <div style={{display:"flex", alignItems:"center", gap:6}}>
                        <span>{parseInt(m)}월 {parseInt(d)}일</span>
                        {isUpdated && (
                          <span style={{
                            fontSize:10, fontWeight:800, color:"#ff6b35",
                            background:"rgba(255,107,53,0.12)", border:"1px solid rgba(255,107,53,0.4)",
                            borderRadius:4, padding:"1px 5px", letterSpacing:0.5,
                          }}>UPDATED</span>
                        )}
                      </div>
                    </td>
                    <td style={{...styles.td,fontWeight:700,color:"#ff6b35"}}>{v.dow}</td>
                    <td style={{...styles.td,flex:2}}>
                      {isAdmin
                        ? <input style={styles.noteInput} value={v.note} placeholder="비고 입력"
                            onChange={e=>setConfirmed(prev=>({...prev,[k]:{...prev[k],note:e.target.value}}))} />
                        : <span>{v.note||"-"}</span>
                      }
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

        </div>
      </div>
    </div>

    {/* SKN 메일 발송 팝업 */}
    {sknMailData && (
      <SKNMailPopup
        recipients={sknRecipients}
        mailData={sknMailData}
        onClose={()=>setSknMailData(null)}
        onSent={()=>{ setSknMailData(null); }}
      />
    )}

    {/* 5회 초과 확인 팝업 */}
    {overLimitPopup && (
      <div style={styles2.popupOverlay}>
        <div style={{...styles2.popupBox, minWidth:340, gap:14}}>
          <div style={{fontSize:36}}>⚠️</div>
          <div style={{fontSize:17, fontWeight:800, color:"#1a1a2e", textAlign:"center"}}>배송 5회가 초과됩니다</div>
          <div style={{fontSize:13, color:"#666", textAlign:"center", lineHeight:1.7}}>
            이번 기간 배송이 5회를 초과합니다.<br/>
            추가 배송은 유료로 진행됩니다.<br/>
            그래도 진행하시겠습니까?
          </div>
          <div style={{display:"flex", gap:10, marginTop:4}}>
            <button style={{...styles2.popupBtn, background:"#aaa", color:"#333"}} onClick={()=>setOverLimitPopup(null)}>취소</button>
            <button style={styles2.popupBtn} onClick={()=>{
              setEditNewSel(prev=>{ const n=new Set(prev); n.add(overLimitPopup); return n; });
              setOverLimitPopup(null);
            }}>확인 (파란 원으로 추가)</button>
          </div>
        </div>
      </div>
    )}
    {/* SKN: 컨펌 필요 날짜 알림 팝업 */}
    {isSK && sknPendingPopup && (
      <div style={styles2.popupOverlay}>
        <div style={{...styles2.popupBox, minWidth:360, gap:14}}>
          <div style={{fontSize:40}}>🔔</div>
          <div style={{fontSize:18, fontWeight:800, color:"#1a1a2e"}}>컨펌이 필요한 날짜가 있습니다</div>
          <div style={{width:"100%", borderRadius:8, border:"1px solid #eee", overflow:"hidden"}}>
            {sknPendingDates.map((r,i) => {
              const [y,m,d] = r.key.split("-");
              return (
                <div key={i} style={{padding:"9px 16px", borderBottom:"1px solid #f5f5f5",
                  fontSize:13, display:"flex", justifyContent:"space-between"}}>
                  <span style={{fontWeight:600}}>{parseInt(m)}월 {parseInt(d)}일</span>
                  <span style={{color:"#888"}}>{r.dow}요일</span>
                </div>
              );
            })}
          </div>
          <button style={styles2.popupBtn} onClick={()=>setSknPendingPopup(false)}>확인</button>
        </div>
      </div>
    )}
    </>
  );
}

function LegendItem({ color, label }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <div style={{width:12,height:12,borderRadius:"50%",background:color,boxShadow:`0 0 6px ${color}88`}} />
      <span style={{fontSize:11.5,color:C.text1}}>{label}</span>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// STYLES
// ═══════════════════════════════════════════════════════════════════════
// ─── 수정/신규 diff 계산 (MailSendPopup + SKNMailPopup 공용) ─────────────────
/**
 * before/after 배열을 받아 { before, after, type } 행 목록을 반환.
 * type: "same" | "changed" | "removed" | "added"
 *
 * 변경 판정 규칙:
 *  - 같은 key → "same"
 *  - before에만 있고, 날짜 차이 14일 이내인 after 날짜와 짝지을 수 있으면 → "changed"
 *  - 짝 없는 before → "removed"
 *  - 짝 없는 after  → "added" (신규)
 */
function computeDiffRows(before, after) {
  const beforeKeys = new Set(before.map(r=>r.key));
  const afterKeys  = new Set(after.map(r=>r.key));

  const kept    = before.filter(r => afterKeys.has(r.key));
  const removed = [...before.filter(r => !afterKeys.has(r.key))].sort((a,b)=>a.key.localeCompare(b.key));
  const added   = [...after.filter(r => !beforeKeys.has(r.key))].sort((a,b)=>a.key.localeCompare(b.key));

  const rows = [];
  kept.forEach(b => {
    const a = after.find(r => r.key === b.key);
    rows.push({ before:b, after:a, type:"same" });
  });

  // ── 순서 보존 최적 매칭 ──────────────────────────────────────────────
  // 정렬된 두 배열에서 순서를 유지하며 매칭하는 것이 전체 거리 합 최소임이 보장됨.
  // (rearrangement inequality: a₁≤a₂, b₁≤b₂ → |a₁-b₁|+|a₂-b₂| ≤ |a₁-b₂|+|a₂-b₁|)
  const usedRemoved = new Set();
  const usedAdded   = new Set();
  const changedRows = [];

  let ri = 0, ai = 0;
  while (ri < removed.length && ai < added.length) {
    const diff = Math.abs(new Date(added[ai].key) - new Date(removed[ri].key)) / 86400000;
    if (diff <= 14) {
      changedRows.push({ before: removed[ri], after: added[ai], type: "changed" });
      usedRemoved.add(ri); usedAdded.add(ai);
      ri++; ai++;
    } else if (new Date(removed[ri].key) < new Date(added[ai].key)) {
      ri++;
    } else {
      ai++;
    }
  }

  const remainRemoved = removed.filter((_, i) => !usedRemoved.has(i));
  const remainAdded   = added.filter((_, i) => !usedAdded.has(i));

  changedRows.forEach(row => rows.push(row));
  remainRemoved.forEach(r => rows.push({ before:r, after:null, type:"removed" }));
  remainAdded.forEach(a   => rows.push({ before:null, after:a, type:"added"   }));

  rows.sort((a,b) => ((a.before||a.after)?.key||"").localeCompare((b.before||b.after)?.key||""));
  return rows;
}

// ─── 이메일 본문 미리보기 모달 ─────────────────────────────────────────────
function EmailPreviewModal({ onClose, subject, htmlBody }) {
  return (
    <div style={{
      position:"fixed", inset:0, zIndex:9999,
      background:"rgba(0,0,0,0.75)", display:"flex", alignItems:"center", justifyContent:"center",
    }} onClick={onClose}>
      <div style={{
        background:"#fff", borderRadius:12, width:"min(680px, 95vw)", maxHeight:"90vh",
        display:"flex", flexDirection:"column", overflow:"hidden",
        boxShadow:"0 24px 80px rgba(0,0,0,0.5)",
      }} onClick={e=>e.stopPropagation()}>
        {/* 모달 헤더 */}
        <div style={{
          background:"#f8f9fa", borderBottom:"1px solid #e0e0e0",
          padding:"14px 20px", display:"flex", alignItems:"center", justifyContent:"space-between",
        }}>
          <div style={{display:"flex", alignItems:"center", gap:10}}>
            <span style={{fontSize:18}}>📧</span>
            <span style={{fontWeight:700, fontSize:15, color:"#222"}}>메일 본문 미리보기</span>
          </div>
          <button onClick={onClose} style={{
            background:"none", border:"none", cursor:"pointer", fontSize:20, color:"#888",
            lineHeight:1, padding:"2px 6px",
          }}>✕</button>
        </div>
        {/* 이메일 클라이언트 느낌 메타 */}
        <div style={{borderBottom:"1px solid #eee", padding:"10px 20px", background:"#fafafa"}}>
          <div style={{fontSize:12, color:"#999", marginBottom:3}}>제목</div>
          <div style={{fontSize:14, fontWeight:600, color:"#222"}}>{subject}</div>
        </div>
        {/* 본문 스크롤 영역 */}
        <div style={{overflowY:"auto", flex:1, padding:"0"}}>
          <div dangerouslySetInnerHTML={{__html: htmlBody}} />
        </div>
        <div style={{borderTop:"1px solid #eee", padding:"12px 20px", background:"#f8f9fa", textAlign:"right"}}>
          <button onClick={onClose} style={{
            padding:"8px 24px", borderRadius:8, border:"none", cursor:"pointer",
            background:"#555", color:"#fff", fontSize:13, fontWeight:600,
          }}>닫기</button>
        </div>
      </div>
    </div>
  );
}

// 날짜 키("YYYY-MM-DD")에서 배송 기간의 연/월 계산 (21일 이후면 익월 기준)
function getPeriodLabel(dateKey) {
  if (!dateKey) { const n = new Date(); return `${n.getFullYear()}년 ${n.getMonth()+1}월 배송 일정`; }
  const [y, m, d] = dateKey.split('-').map(Number);
  const pm = d >= 21 ? (m === 12 ? 1 : m + 1) : m;
  const py = d >= 21 ? (m === 12 ? y + 1 : y) : y;
  return `${py}년 ${pm}월 배송 일정`;
}

/** HTML 이메일 본문 생성 (최초/수정 배송일 공지 - 점주/지역본부 대상) */
function buildDeliveryMailHtml({ isRevision, diffRows, dates, fmtDate }) {
  const refKey = isRevision
    ? (diffRows?.find(r => r.after)?.after?.key || diffRows?.find(r => r.before)?.before?.key)
    : dates?.[0]?.key;
  const periodLabel = getPeriodLabel(refKey);
  const accentColor = isRevision ? "#ff6b35" : "#0066cc";
  const headerTitle = isRevision ? "📝 배송일 수정 공지" : "📦 배송일 공지";
  const subTitle    = isRevision
    ? "아래와 같이 배송 일정이 수정되었습니다."
    : "이번 달 배송 일정을 안내드립니다.";
  const FF = "'Malgun Gothic','Apple SD Gothic Neo','나눔고딕',sans-serif";

  let tableHtml = "";

  if (isRevision && diffRows.length > 0) {
    const rows = diffRows.map(row => {
      const isChanged = row.type==="changed";
      const isAdded   = row.type==="added";
      const isRemoved = row.type==="removed";
      const rowBg     = (isChanged||isAdded||isRemoved) ? "#fff8f5" : "#ffffff";
      const beforeTd  = row.before
        ? `<span style="font-family:${FF};color:${isRemoved?"#e8420a":"#333333"};text-decoration:${isRemoved?"line-through":"none"};font-weight:${isChanged||isRemoved?700:400}">${fmtDate(row.before.key)} (${row.before.dow})</span>`
        : `<span style="font-family:${FF};color:#cccccc">-</span>`;
      let afterContent = "";
      if (row.after) {
        const color = isAdded ? "#059669" : isChanged ? "#d97706" : "#333333";
        const badge = isChanged
          ? `<span style="font-family:${FF};font-size:10px;margin-left:6px;background:#fffbe6;color:#b8860b;padding:1px 6px;font-weight:600">변경</span>`
          : isAdded
          ? `<span style="font-family:${FF};font-size:10px;margin-left:6px;background:#e8f8ee;color:#059669;padding:1px 6px;font-weight:600">신규</span>` : "";
        afterContent = `<span style="font-family:${FF};color:${color};font-weight:${isChanged||isAdded?700:400}">${fmtDate(row.after.key)} (${row.after.dow})</span>${badge}`;
      } else {
        afterContent = `<span style="font-family:${FF};color:#cccccc">-</span>`;
      }
      return `<tr bgcolor="${rowBg}"><td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0">${beforeTd}</td><td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0">${afterContent}</td></tr>`;
    }).join("");
    tableHtml = `
      <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:14px;font-family:${FF}">
        <thead><tr bgcolor="#f5f5f5">
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#666666;font-weight:600;border-bottom:2px solid #e0e0e0;width:50%">수정 전</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#666666;font-weight:600;border-bottom:2px solid #e0e0e0">수정 후</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  } else if (!isRevision && dates?.length > 0) {
    const rows = dates.map((r, i) =>
      `<tr bgcolor="${i%2===0?"#ffffff":"#fafafa"}">
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0;font-weight:600;color:#222222">${i+1}차</td>
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0;color:#333333">${fmtDate(r.key)}</td>
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0;color:#888888">${r.dow}요일</td>
      </tr>`).join("");
    tableHtml = `
      <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:14px;font-family:${FF}">
        <thead><tr bgcolor="#f5f5f5">
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#666666;font-weight:600;border-bottom:2px solid #e0e0e0;width:60px">회차</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#666666;font-weight:600;border-bottom:2px solid #e0e0e0">날짜</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#666666;font-weight:600;border-bottom:2px solid #e0e0e0">요일</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  }

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background-color:#f0f4f8;font-family:${FF}">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f4f8">
<tr><td align="center" style="padding:32px 16px">
<table width="560" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border:1px solid #e0e0e0">
  <tr><td bgcolor="${accentColor}" style="background-color:${accentColor};padding:28px 36px 24px">
    <p style="font-family:${FF};font-size:11px;color:#ffe0d0;letter-spacing:2px;margin:0 0 10px 0">SK Telecom VMD</p>
    <p style="font-family:${FF};font-size:22px;font-weight:700;color:#ffffff;margin:0 0 6px 0">${headerTitle}</p>
    <p style="font-family:${FF};font-size:13px;color:#ffe0d0;margin:0">${periodLabel}</p>
  </td></tr>
  <tr><td style="padding:28px 36px 20px">
    <p style="font-family:${FF};font-size:14px;color:#444444;line-height:1.7;margin:0 0 8px 0">안녕하세요.</p>
    <p style="font-family:${FF};font-size:14px;color:#444444;line-height:1.7;margin:0">VMD 담당Bot입니다.<br>${subTitle}</p>
  </td></tr>
  <tr><td style="padding:0 36px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="border-top:1px solid #f0f0f0;font-size:0">&nbsp;</td></tr></table></td></tr>
  <tr><td style="padding:20px 36px 8px">
    <p style="font-family:${FF};font-size:11px;color:#999999;font-weight:700;letter-spacing:1px;margin:0 0 12px 0">배송 일정</p>
    ${tableHtml}
  </td></tr>
  <tr><td style="padding:16px 36px 28px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="4" bgcolor="${accentColor}" style="background-color:${accentColor}">&nbsp;</td>
      <td bgcolor="#fff8f5" style="background-color:#fff8f5;padding:12px 16px;font-family:${FF};font-size:12px;color:#666666;line-height:1.7">
        ※ 배송 일정은 상황에 따라 변경될 수 있으며, 변경 시 별도 공지됩니다.<br>
        ※ 문의사항은 VMD 담당자에게 연락해 주시기 바랍니다.
      </td>
    </tr></table>
  </td></tr>
  <tr><td bgcolor="#f8f9fa" style="background-color:#f8f9fa;border-top:1px solid #eeeeee;padding:20px 36px">
    <p style="font-family:${FF};font-size:11px;color:#aaaaaa;line-height:1.8;margin:0">
      <strong style="font-family:${FF};color:#888888">SK Telecom VMD 배송 관리 시스템</strong><br>
      본 메일은 자동 발송된 메일입니다. 회신하지 마세요.
    </p>
  </td></tr>
</table>
</td></tr></table>
</body></html>`;
}

/** HTML 이메일 본문 생성 (SKN 메일 - 배송일 검토 요청) */
function buildSknMailHtml({ isRevision, diffRows, dates, fmtDate, fmtRow }) {
  const refKey = isRevision
    ? (diffRows?.find(r => r.after)?.after?.key || diffRows?.find(r => r.before)?.before?.key)
    : dates?.[0]?.key;
  const periodLabel = getPeriodLabel(refKey);
  const accentColor = isRevision ? "#ff6b35" : "#0066cc";
  const noteBg      = isRevision ? "#fff8f5" : "#f0f4ff";
  const headerSub   = isRevision ? "#ffe0d0" : "#cce0ff";
  const headerTitle = isRevision ? "🔄 배송일 수정 검토 요청" : "📋 배송일 임시 지정 검토 요청";
  const subTitle    = isRevision
    ? "아래와 같이 배송일 수정 내용에 대한 검토를 요청드립니다."
    : "다음 달 배송일 임시 지정 내용에 대한 검토를 요청드립니다. 배송일 5개 확정 후 회신 부탁드립니다.";
  const FF = "'Malgun Gothic','Apple SD Gothic Neo','나눔고딕',sans-serif";

  let tableHtml = "";

  if (isRevision && diffRows.length > 0) {
    const rows = diffRows.map(row => {
      const isChanged = row.type==="changed";
      const isAdded   = row.type==="added";
      const isRemoved = row.type==="removed";
      const rowBg = (isChanged||isAdded||isRemoved) ? "#fff8f0" : "#ffffff";
      const beforeTd = row.before
        ? `<span style="font-family:${FF};color:${isRemoved?"#e85d26":"#333333"};text-decoration:${isRemoved?"line-through":"none"};font-weight:${isChanged||isRemoved?700:400}">${fmtRow(row.before)}</span>`
        : `<span style="font-family:${FF};color:#cccccc">-</span>`;
      let afterContent = "";
      if (row.after) {
        const color = isAdded ? "#2e8b57" : isChanged ? "#e07800" : "#333333";
        const badge = isChanged
          ? `<span style="font-family:${FF};font-size:10px;margin-left:6px;background:#fffbe6;color:#b8860b;padding:1px 6px;font-weight:600">변경</span>`
          : isAdded
          ? `<span style="font-family:${FF};font-size:10px;margin-left:6px;background:#e8f8ee;color:#2e8b57;padding:1px 6px;font-weight:600">신규</span>` : "";
        afterContent = `<span style="font-family:${FF};color:${color};font-weight:${isChanged||isAdded?700:400}">${fmtRow(row.after)}</span>${badge}`;
      } else {
        afterContent = `<span style="font-family:${FF};color:#cccccc">-</span>`;
      }
      return `<tr bgcolor="${rowBg}"><td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0">${beforeTd}</td><td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #f0f0f0">${afterContent}</td></tr>`;
    }).join("");
    tableHtml = `
      <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:14px;font-family:${FF}">
        <thead><tr bgcolor="#f5f5f5">
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#555555;font-weight:600;border-bottom:2px solid #e0e0e0;width:50%">수정 전</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#555555;font-weight:600;border-bottom:2px solid #e0e0e0">수정 후</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  } else if (!isRevision && dates?.length > 0) {
    const rows = dates.map((r, i) =>
      `<tr bgcolor="${i%2===0?"#ffffff":"#f9fbff"}">
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #eef2ff;font-weight:700;color:#0066cc">${i+1}차</td>
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #eef2ff;font-weight:600;color:#222222">${fmtDate(r.key)}</td>
        <td style="font-family:${FF};padding:10px 16px;border-bottom:1px solid #eef2ff;color:#888888">${r.dow}요일</td>
      </tr>`).join("");
    tableHtml = `
      <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:14px;font-family:${FF}">
        <thead><tr bgcolor="#f0f4ff">
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#555555;font-weight:600;border-bottom:2px solid #dce6ff;width:60px">회차</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#555555;font-weight:600;border-bottom:2px solid #dce6ff">날짜</th>
          <th style="font-family:${FF};padding:9px 16px;text-align:left;color:#555555;font-weight:600;border-bottom:2px solid #dce6ff">요일</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  }

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background-color:#f0f4f8;font-family:${FF}">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f4f8">
<tr><td align="center" style="padding:32px 16px">
<table width="560" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border:1px solid #e0e0e0">
  <tr><td bgcolor="${accentColor}" style="background-color:${accentColor};padding:28px 36px 24px">
    <p style="font-family:${FF};font-size:11px;color:${headerSub};letter-spacing:2px;margin:0 0 10px 0">SK Telecom VMD — SKN 협업</p>
    <p style="font-family:${FF};font-size:22px;font-weight:700;color:#ffffff;margin:0 0 6px 0">${headerTitle}</p>
    <p style="font-family:${FF};font-size:13px;color:${headerSub};margin:0">${periodLabel}</p>
  </td></tr>
  <tr><td style="padding:28px 36px 20px">
    <p style="font-family:${FF};font-size:14px;color:#444444;line-height:1.7;margin:0 0 8px 0">SKN 담당자님께,</p>
    <p style="font-family:${FF};font-size:14px;color:#444444;line-height:1.7;margin:0">안녕하세요. SK Telecom VMD 담당Bot입니다.<br>${subTitle}</p>
  </td></tr>
  <tr><td style="padding:0 36px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="border-top:1px solid #f0f0f0;font-size:0">&nbsp;</td></tr></table></td></tr>
  <tr><td style="padding:20px 36px 8px">
    <p style="font-family:${FF};font-size:11px;color:#999999;font-weight:700;letter-spacing:1px;margin:0 0 12px 0">${isRevision ? "수정 내용" : "임시 배송일"}</p>
    ${tableHtml}
  </td></tr>
  <tr><td style="padding:16px 36px 28px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="4" bgcolor="${accentColor}" style="background-color:${accentColor}">&nbsp;</td>
      <td bgcolor="${noteBg}" style="background-color:${noteBg};padding:12px 16px;font-family:${FF};font-size:12px;color:#555555;line-height:1.7">
        ※ 일정 검토 후 확정 여부를 회신해 주시기 바랍니다.<br>
        ※ 수정이 필요한 경우 수정 날짜와 사유를 함께 알려주세요.
      </td>
    </tr></table>
  </td></tr>
  <tr><td bgcolor="#f8f9fa" style="background-color:#f8f9fa;border-top:1px solid #eeeeee;padding:20px 36px">
    <p style="font-family:${FF};font-size:11px;color:#aaaaaa;line-height:1.8;margin:0">
      <strong style="font-family:${FF};color:#888888">SK Telecom VMD 배송 관리 시스템</strong><br>
      본 메일은 자동 발송된 메일입니다. 문의사항은 담당자에게 직접 연락 바랍니다.
    </p>
  </td></tr>
</table>
</td></tr></table>
</body></html>`;
}

// ─── 메일 발송 팝업 ─────────────────────────────────────────────────────
function MailSendPopup({ recipients, mailType, mailData, onClose, onSent }) {
  const [sent, setSent] = useState(false);
  const [sending, setSending] = useState(false);
  const [sendError, setSendError] = useState("");
  const [showPreview, setShowPreview] = useState(false);
  const isRevision = mailType === "revision";
  const title = isRevision ? "수정 배송일 공지 메일 발송" : "최초 배송일 공지 메일 발송";
  const desc  = isRevision
    ? "수정된 배송 일정을 아래 수신자들에게 공지합니다."
    : "확정된 배송 일정을 아래 수신자들에게 최초 공지합니다.";
  const fmtDate = (key) => {
    if (!key) return "";
    const [,m,d] = key.split("-");
    return `${parseInt(m)}월 ${parseInt(d)}일`;
  };
  // 비교 diff (revision용) — 공용 함수 사용
  const diffRows = isRevision && mailData
    ? computeDiffRows(mailData.before || [], mailData.after || [])
    : [];
  return (
    <div style={styles2.popupOverlay}>
      <div style={{...styles2.popupBox, minWidth:400, gap:16}}>
        {!sent ? (
          <>
            <div style={styles2.popupIcon}>{isRevision?"📝":"📧"}</div>
            <div style={styles2.popupTitle}>{title}</div>
            <div style={{fontSize:12, color:C.text1}}>{desc}</div>

            {!isRevision && mailData?.dates && (
              <div style={{width:"100%", borderRadius:10, border:`1px solid ${C.border}`, overflow:"hidden", maxHeight:180, overflowY:"auto"}}>
                {mailData.dates.map((r,i)=>(
                  <div key={i} style={{padding:"8px 14px", borderBottom:`1px solid ${C.border}`,
                    fontSize:13, display:"flex", justifyContent:"space-between", background:i%2===0?"rgba(255,255,255,0.02)":"transparent"}}>
                    <span style={{fontWeight:600, color:C.text0}}>{fmtDate(r.key)}</span>
                    <span style={{color:C.text1}}>{r.dow}요일</span>
                  </div>
                ))}
              </div>
            )}
            {isRevision && diffRows.length > 0 && (
              <div style={{width:"100%", borderRadius:10, border:`1px solid ${C.border}`, overflow:"hidden", maxHeight:260, overflowY:"auto"}}>
                <table style={{width:"100%", borderCollapse:"collapse", fontSize:13}}>
                  <thead>
                    <tr style={{background:"rgba(255,255,255,0.04)"}}>
                      <th style={{padding:"7px 12px", textAlign:"left", color:C.text2, fontWeight:600, borderBottom:`1px solid ${C.border}`}}>수정 전</th>
                      <th style={{padding:"7px 12px", textAlign:"left", color:C.text2, fontWeight:600, borderBottom:`1px solid ${C.border}`}}>수정 후</th>
                    </tr>
                  </thead>
                  <tbody>
                    {diffRows.map((row,i)=>{
                      const isChanged = row.type==="changed";
                      const isAdded   = row.type==="added";
                      const isRemoved = row.type==="removed";
                      const highlight = isChanged||isAdded||isRemoved;
                      return (
                        <tr key={i} style={{borderBottom:`1px solid ${C.border}`,
                          background:highlight?"rgba(255,107,53,0.05)":"transparent"}}>
                          <td style={{padding:"8px 12px", color:isRemoved?"#fb7185":C.text1,
                            textDecoration:isRemoved?"line-through":"none", fontWeight:isChanged||isRemoved?700:400}}>
                            {row.before
                              ? `${fmtDate(row.before.key)} (${row.before.dow})`
                              : <span style={{color:C.text2}}>-</span>}
                          </td>
                          <td style={{padding:"8px 12px",
                            color:isAdded?"#34d399":isChanged?"#fbbf24":C.text0,
                            fontWeight:isChanged||isAdded?700:400}}>
                            {row.after
                              ? `${fmtDate(row.after.key)} (${row.after.dow})`
                              : <span style={{color:C.text2}}>-</span>}
                            {isChanged && <span style={{fontSize:10,marginLeft:6,background:"rgba(251,191,36,0.12)",color:"#fbbf24",padding:"1px 6px",borderRadius:6}}>변경</span>}
                            {isAdded   && <span style={{fontSize:10,marginLeft:6,background:"rgba(52,211,153,0.12)",color:"#34d399",padding:"1px 6px",borderRadius:6}}>신규</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            <div style={{width:"100%", fontSize:11, color:C.text2, marginBottom:-8, letterSpacing:0.5, textTransform:"uppercase"}}>수신자</div>
            <div style={{width:"100%", maxHeight:150, overflowY:"auto", borderRadius:10, border:`1px solid ${C.border}`}}>
              {recipients.map(r=>(
                <div key={r.id} style={{padding:"8px 14px", borderBottom:`1px solid ${C.border}`, fontSize:13}}>
                  <span style={{fontWeight:600, color:C.text0}}>{r.name}</span>
                  <span style={{color:C.text1, marginLeft:10}}>{r.email}</span>
                </div>
              ))}
            </div>
            {/* 미리보기 링크 */}
            <div style={{width:"100%", textAlign:"center"}}>
              <button
                style={{background:"none", border:"none", cursor:"pointer", fontSize:12,
                  color:C.blue, textDecoration:"underline", padding:"2px 0"}}
                onClick={()=>setShowPreview(true)}>
                📄 메일 본문 미리보기
              </button>
            </div>
            {sendError && <div style={{color:"#f87171", fontSize:12, textAlign:"center"}}>{sendError}</div>}
            <div style={{display:"flex", gap:10}}>
              <button style={{...styles2.popupBtn, background:C.surface2, color:C.text1, boxShadow:"none", border:`1px solid ${C.border}`}} onClick={onClose} disabled={sending}>취소</button>
              <button style={{...styles2.popupBtn, background: isRevision?"linear-gradient(135deg,#ff6b35,#e8420a)":"linear-gradient(135deg,#10b981,#059669)", opacity:sending?0.6:1}}
                disabled={sending}
                onClick={async () => {
                  setSending(true); setSendError("");
                  const subject = isRevision ? "[VMD] 배송일 수정 공지" : "[VMD] 배송일 공지";
                  const html = buildDeliveryMailHtml({ isRevision, diffRows, dates:mailData?.dates, fmtDate });
                  const to = recipients.map(r => r.email);
                  const { error } = await dataClient.mail.sendEmail({ to, subject, html });
                  setSending(false);
                  if (error) { setSendError("발송 실패: " + (error.message || "서버 오류")); return; }
                  setSent(true); onSent?.(mailType);
                }}>{sending ? "발송 중..." : "발송"}</button>
            </div>
          </>
        ) : (
          <>
            <div style={styles2.popupIcon}>✅</div>
            <div style={styles2.popupTitle}>메일 발송이 완료되었습니다</div>
            <div style={{fontSize:13, color:C.text1}}>
              {recipients.length}명에게 {isRevision?"수정된 ":""}배송 일정이 공지되었습니다.
            </div>
            <button style={styles2.popupBtn} onClick={onClose}>확인</button>
          </>
        )}
      </div>
      {showPreview && (
        <EmailPreviewModal
          onClose={()=>setShowPreview(false)}
          subject={isRevision ? "[VMD] 배송일 수정 공지" : "[VMD] 배송일 공지"}
          htmlBody={buildDeliveryMailHtml({ isRevision, diffRows, dates:mailData?.dates, fmtDate })}
        />
      )}
    </div>
  );
}

// ─── SKN 메일 발송 팝업 ──────────────────────────────────────────────────
function SKNMailPopup({ recipients, mailData, onClose, onSent }) {
  const [sent, setSent] = useState(false);
  const [sending, setSending] = useState(false);
  const [sendError, setSendError] = useState("");
  const [showPreview, setShowPreview] = useState(false);
  const isRevision = mailData.type === "revision";
  const isNone     = mailData.type === "none";

  const fmtDate = (key) => {
    if (!key) return "";
    const [,m,d] = key.split("-");
    return `${parseInt(m)}월 ${parseInt(d)}일`;
  };
  const fmtRow = (r) => `${fmtDate(r.key)} (${r.dow})`;

  // key-based diff — 공용 함수 사용
  const diffRows = isRevision
    ? computeDiffRows(mailData.before || [], mailData.after || [])
    : [];

  return (
    <div style={styles2.popupOverlay}>
      <div style={{...styles2.popupBox, minWidth:460, gap:14, maxHeight:"92vh", overflowY:"auto"}}>
        {isNone ? (
          <>
            <div style={{fontSize:48}}>📭</div>
            <div style={styles2.popupTitle}>전달할 내용이 없습니다</div>
            <div style={{fontSize:13, color:"#888", textAlign:"center"}}>
              현재 기간의 배송일이 이미 모두 확정되어 있거나<br/>변경된 내용이 없습니다.
            </div>
            <button style={styles2.popupBtn} onClick={onClose}>확인</button>
          </>
        ) : sent ? (
          <>
            <div style={styles2.popupIcon}>✅</div>
            <div style={styles2.popupTitle}>메일 발송 완료</div>
            <div style={{fontSize:13, color:"#666", textAlign:"center"}}>
              {recipients.length}명에게 {isRevision?"수정 내용이":"배송일 임시 지정 내용이"} 발송되었습니다.
            </div>
            <button style={styles2.popupBtn} onClick={onSent || onClose}>확인</button>
          </>
        ) : (
          <>
            <div style={styles2.popupIcon}>📨</div>
            <div style={styles2.popupTitle}>
              {isRevision ? "배송일 수정 검토 요청" : "배송일 임시 지정 검토 요청"}
            </div>
            <div style={{fontSize:12, color:"#888", textAlign:"center"}}>
              {isRevision
                ? "아래 수정 내용을 SKN 담당자에게 검토 요청합니다."
                : "다음 달 배송일 5개를 SKN 담당자에게 검토 요청합니다."}
            </div>

            {/* 날짜 내용 */}
            <div style={{width:"100%", borderRadius:8, border:"1px solid #eee", overflow:"hidden"}}>
              {isRevision ? (
                <table style={{width:"100%", borderCollapse:"collapse", fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f8f9fa"}}>
                      <th style={{padding:"8px 14px", textAlign:"left", color:"#666", fontWeight:600, borderBottom:"1px solid #eee", width:"50%"}}>수정 전</th>
                      <th style={{padding:"8px 14px", textAlign:"left", color:"#666", fontWeight:600, borderBottom:"1px solid #eee"}}>수정 후</th>
                    </tr>
                  </thead>
                  <tbody>
                    {diffRows.map((row,i)=>{
                      const isChanged = row.type==="changed";
                      const isAdded   = row.type==="added";
                      const isRemoved = row.type==="removed";
                      return (
                        <tr key={i} style={{borderBottom:"1px solid #f5f5f5",
                          background:isChanged||isAdded||isRemoved?"#fff8f8":"transparent"}}>
                          <td style={{padding:"9px 14px", color:isRemoved?"#e85d26":"#555",
                            textDecoration:isRemoved?"line-through":"none",
                            fontWeight:isChanged||isRemoved?700:400}}>
                            {row.before ? fmtRow(row.before) : <span style={{color:"#ccc"}}>-</span>}
                          </td>
                          <td style={{padding:"9px 14px",
                            color:isAdded?"#2e8b57":isChanged?"#e07800":"#555",
                            fontWeight:isChanged||isAdded?700:400}}>
                            {row.after ? fmtRow(row.after) : <span style={{color:"#ccc"}}>-</span>}
                            {isChanged && <span style={{fontSize:10,marginLeft:6,background:"#fffbe6",color:"#b8860b",padding:"1px 6px",borderRadius:8}}>변경</span>}
                            {isAdded   && <span style={{fontSize:10,marginLeft:6,background:"#e8f8ee",color:"#2e8b57",padding:"1px 6px",borderRadius:8}}>신규</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              ) : (
                mailData.dates?.map((r,i)=>(
                  <div key={i} style={{padding:"9px 14px", borderBottom:"1px solid #f5f5f5",
                    fontSize:13, display:"flex", justifyContent:"space-between"}}>
                    <span style={{fontWeight:600}}>{fmtDate(r.key)}</span>
                    <span style={{color:"#888"}}>{r.dow}요일</span>
                  </div>
                ))
              )}
            </div>

            {/* 수신자 */}
            <div style={{width:"100%"}}>
              <div style={{fontSize:12, color:"#aaa", marginBottom:6}}>수신자</div>
              <div style={{borderRadius:8, border:"1px solid #eee"}}>
                {recipients.length === 0
                  ? <div style={{padding:14, textAlign:"center", color:"#aaa", fontSize:13}}>
                      설정 {">"} SKN 담당자를 먼저 등록해주세요.
                    </div>
                  : recipients.map(r=>(
                    <div key={r.id} style={{padding:"8px 14px", borderBottom:"1px solid #f5f5f5", fontSize:13}}>
                      <span style={{fontWeight:600}}>{r.name}</span>
                      <span style={{color:"#888", marginLeft:10}}>{r.email}</span>
                    </div>
                  ))
                }
              </div>
            </div>

            {/* 미리보기 링크 */}
            <div style={{width:"100%", textAlign:"center"}}>
              <button
                style={{background:"none", border:"none", cursor:"pointer", fontSize:12,
                  color:"#0066cc", textDecoration:"underline", padding:"2px 0"}}
                onClick={()=>setShowPreview(true)}>
                📄 메일 본문 미리보기
              </button>
            </div>
            {sendError && <div style={{color:"#e85d26", fontSize:12, textAlign:"center"}}>{sendError}</div>}
            <div style={{display:"flex", gap:10, marginTop:4, width:"100%"}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}} onClick={onClose} disabled={sending}>취소</button>
              <button style={{...styles2.popupBtn, background:"#0066cc", flex:1, opacity:sending?0.6:1}}
                disabled={recipients.length === 0 || sending}
                onClick={async () => {
                  if (!recipients.length) return;
                  setSending(true); setSendError("");
                  const subject = isRevision ? "[VMD] SKN 배송일 수정 검토 요청" : "[VMD] SKN 배송일 임시 지정 검토 요청";
                  const html = buildSknMailHtml({ isRevision, diffRows, dates:mailData?.dates, fmtDate, fmtRow });
                  const to = recipients.map(r => r.email);
                  const { error } = await dataClient.mail.sendEmail({ to, subject, html });
                  setSending(false);
                  if (error) { setSendError("발송 실패: " + (error.message || "서버 오류")); return; }
                  setSent(true);
                }}>
                {sending ? "발송 중..." : isRevision ? "수정 내용 발송" : "임시 지정 내용 발송"}
              </button>
            </div>
          </>
        )}
      </div>
      {showPreview && (
        <EmailPreviewModal
          onClose={()=>setShowPreview(false)}
          subject={isRevision ? "[VMD] SKN 배송일 수정 검토 요청" : "[VMD] SKN 배송일 임시 지정 검토 요청"}
          htmlBody={buildSknMailHtml({ isRevision, diffRows, dates:mailData?.dates, fmtDate, fmtRow })}
        />
      )}
    </div>
  );
}

// ─── 설정 페이지 ─────────────────────────────────────────────────────────
function SettingsPage({ mailRecipients, setMailRecipients, sknRecipients, setSknRecipients, shippingGroups, setShippingGroups, shippingTable1, setShippingTable1, shippingStep, setShippingStep, shippingCustomCols, setShippingCustomCols, shippingNextColId, setShippingNextColId, setStoreList }) {
  const [activeTab, setActiveTab] = useState("quantity"); // "quantity" | "recipients"
  const [editing, setEditing]           = useState(null);
  const [adding, setAdding]             = useState(false);
  const [form, setForm]                 = useState({name:"", email:""});
  const [uploadStatus, setUploadStatus] = useState(null);
  const [previewRows, setPreviewRows]   = useState(null);
  const [dupPopup, setDupPopup]         = useState(null);
  const fileInputRef = useRef(null);

  const nowStr = () => {
    const d = new Date();
    const pad = n => String(n).padStart(2,"0");
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  };
  const nextId = (list) => Math.max(0, ...(list.length ? list.map(r=>r.id) : [0])) + 1;
  const isDuplicate = (email, list) => list.some(r => r.email.toLowerCase() === email.toLowerCase());

  const startEdit = (r) => { setEditing(r); setForm({name:r.name, email:r.email}); setAdding(false); setPreviewRows(null); };
  const startAdd  = () => { setAdding(true); setEditing(null); setForm({name:"", email:""}); setPreviewRows(null); };
  const cancel    = () => { setEditing(null); setAdding(false); };

  // 직접 추가 저장 (중복 체크)
  const save = () => {
    if (!form.name.trim() || !form.email.trim()) return;
    if (editing) {
      setMailRecipients(prev => prev.map(r => r.id===editing.id ? {...r, name:form.name, email:form.email} : r));
      cancel();
    } else {
      if (isDuplicate(form.email, mailRecipients)) {
        setDupPopup(`이미 등록된 이메일입니다.
(${form.email})`);
        return;
      }
      setMailRecipients(prev => [...prev, {id:nextId(prev), name:form.name, email:form.email, registeredAt:nowStr()}]);
      cancel();
    }
  };
  const remove = (id) => setMailRecipients(prev => prev.filter(r=>r.id!==id));

  // 엑셀 템플릿 다운로드
  const downloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["이름","이메일"],
      ["홍길동","hong@example.com"],
      ["김지수","kim@example.com"],
    ]);
    ws["!cols"] = [{wch:15},{wch:30}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "수신자");
    XLSX.writeFile(wb, "메일수신자_템플릿.xlsx");
  };

  // 엑셀 업로드 파싱 + 중복 표시
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploadStatus(null); setPreviewRows(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        const parsed = rows.slice(1)
          .filter(r => r[0] || r[1])
          .map(r => ({ name:String(r[0]||"").trim(), email:String(r[1]||"").trim() }))
          .filter(r => r.name && r.email);
        if (parsed.length === 0) {
          setUploadStatus({type:"error", msg:"유효한 데이터가 없습니다. 이름과 이메일 열을 확인하세요."});
        } else {
          // 1) 엑셀 내부 중복 제거: 먼저 등장한 행 기준으로 seenInFile로 추적
          const seenInFile = new Set();
          const data = parsed.map(r => {
            const emailLower = r.email.toLowerCase();
            const dupInFile = seenInFile.has(emailLower);   // 파일 내 중복
            const dupInList = isDuplicate(r.email, mailRecipients); // 기존 목록 중복
            if (!dupInFile) seenInFile.add(emailLower);
            return { ...r, isDup: dupInFile || dupInList, dupReason: dupInFile ? "파일내중복" : dupInList ? "기존중복" : null };
          });
          const dupCount = data.filter(r=>r.isDup).length;
          const newCount = data.filter(r=>!r.isDup).length;
          setPreviewRows(data);
          setUploadStatus({type:"preview", msg:`총 ${data.length}건 중 신규 ${newCount}건, 중복 ${dupCount}건 (파일 내 중복 포함)`});
        }
      } catch(err) {
        setUploadStatus({type:"error", msg:"파일을 읽는 중 오류가 발생했습니다."});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  // 추가 적용 (중복 제외하고 신규만)
  const applyUpload = (mode) => {
    if (!previewRows) return;
    const now = nowStr();
    setMailRecipients(prev => {
      const base = mode === "replace" ? [] : [...prev];
      let id = nextId(base);
      // replace일 땐 모두 추가, append일 땐 중복 제외
      const toAdd = mode === "replace"
        ? previewRows.map(r => ({id:id++, name:r.name, email:r.email, registeredAt:now}))
        : previewRows.filter(r=>!r.isDup).map(r => ({id:id++, name:r.name, email:r.email, registeredAt:now}));
      return [...base, ...toAdd];
    });
    const added = mode === "replace" ? previewRows.length : previewRows.filter(r=>!r.isDup).length;
    const skipped = mode === "replace" ? 0 : previewRows.filter(r=>r.isDup).length;
    setPreviewRows(null);
    setUploadStatus({
      type:"success",
      msg: mode === "replace"
        ? `${added}건으로 전체 교체되었습니다.`
        : `${added}건 추가 완료. ${skipped>0?`중복 ${skipped}건 제외됨.`:""}`
    });
  };

  return (
    <div style={styles.page}>
      <div style={styles.pageHeader}>
        <h1 style={styles.pageTitle}>설정</h1>
      </div>
      {/* 탭 네비게이션 */}
      <div style={{display:"flex", gap:0, marginBottom:24, borderBottom:"2px solid #eee"}}>
        {[
          {id:"quantity", label:"📦 수량 관련"},
          {id:"recipients", label:"📧 공지 수신자 관련"},
        ].map(tab=>(
          <button key={tab.id}
            style={{
              padding:"10px 28px", border:"none", background:"transparent",
              fontSize:14, fontWeight:700, cursor:"pointer",
              color: activeTab===tab.id ? "#1d6fa4" : "#aaa",
              borderBottom: activeTab===tab.id ? "3px solid #1d6fa4" : "3px solid transparent",
              marginBottom:-2, transition:"all 0.15s",
            }}
            onClick={()=>setActiveTab(tab.id)}>{tab.label}</button>
        ))}
      </div>

      {activeTab === "quantity" && (
      <div style={{maxWidth:960, display:"flex", flexDirection:"column", gap:20}}>
        {/* 배송 수량 설정 섹션 */}
        <ShippingGroupsSection
          shippingGroups={shippingGroups} setShippingGroups={setShippingGroups}
          table1={shippingTable1} setTable1={setShippingTable1}
          step={shippingStep} setStep={setShippingStep}
          customCols={shippingCustomCols} setCustomCols={setShippingCustomCols}
          nextColId={shippingNextColId} setNextColId={setShippingNextColId}
          setStoreList={setStoreList}
        />
      </div>
      )}

      {activeTab === "recipients" && (
      <div style={{maxWidth:900, display:"flex", flexDirection:"column", gap:20}}>

        {/* 엑셀 업로드 섹션 */}
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:28}}>
          <div style={{fontSize:16, fontWeight:800, color:"#1a1a2e", marginBottom:4}}>📂 엑셀로 수신자 일괄 업로드</div>
          <div style={{fontSize:12, color:"#888", marginBottom:18}}>A열: 이름 &nbsp;|&nbsp; B열: 이메일 — 중복 이메일은 자동으로 걸러집니다.</div>

          <div style={{display:"flex", gap:10, flexWrap:"wrap", marginBottom:16}}>
            <button style={{...settingsBtn("#2e8b57"), padding:"8px 18px", fontSize:13}} onClick={downloadTemplate}>⬇ 템플릿 다운로드</button>
            <button style={{...settingsBtn("#1d6fa4"), padding:"8px 18px", fontSize:13}} onClick={()=>fileInputRef.current?.click()}>📂 엑셀 파일 선택</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleFileChange} />
          </div>

          {uploadStatus && (
            <div style={{
              padding:"10px 14px", borderRadius:8, fontSize:13, marginBottom: previewRows ? 12 : 0,
              background: uploadStatus.type==="error"?"#fff0f0": uploadStatus.type==="success"?"#f0fff4":"#fffbe6",
              color: uploadStatus.type==="error"?"#c00": uploadStatus.type==="success"?"#2e8b57":"#7a6000",
              border:`1px solid ${uploadStatus.type==="error"?"#fcc": uploadStatus.type==="success"?"#9be":"#ffe066"}`,
            }}>
              {uploadStatus.type==="error"?"❌ ": uploadStatus.type==="success"?"✅ ":"👁 "}{uploadStatus.msg}
            </div>
          )}

          {previewRows && (
            <div style={{marginTop:8}}>
              <div style={{maxHeight:220, overflowY:"auto", border:"1px solid #eee", borderRadius:8, marginBottom:12}}>
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={{...styles.th, width:36}}>#</th>
                      <th style={styles.th}>이름</th>
                      <th style={styles.th}>이메일</th>
                      <th style={{...styles.th, width:80, textAlign:"center"}}>상태</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.map((r,i)=>(
                      <tr key={i} style={{background: r.isDup ? "#fff8f8" : "transparent"}}>
                        <td style={{...styles.td, color:"#aaa", fontSize:12}}>{i+1}</td>
                        <td style={{...styles.td, color: r.isDup?"#bbb":"inherit"}}>{r.name}</td>
                        <td style={{...styles.td, color: r.isDup?"#bbb":"inherit"}}>{r.email}</td>
                        <td style={{...styles.td, textAlign:"center"}}>
                          {r.isDup
                            ? <span style={{fontSize:11, color:"#e85d26", fontWeight:700, background:"#ffeee8", padding:"2px 8px", borderRadius:10}}>
                                {r.dupReason === "파일내중복" ? "파일내중복" : "기존중복"}
                              </span>
                            : <span style={{fontSize:11, color:"#2e8b57", fontWeight:700, background:"#e8f8ee", padding:"2px 8px", borderRadius:10}}>신규</span>
                          }
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{display:"flex", gap:10, flexWrap:"wrap"}}>
                <button style={{...settingsBtn("#1d6fa4"), padding:"8px 18px", fontSize:13}} onClick={()=>applyUpload("append")}>
                  + 신규만 추가 ({previewRows.filter(r=>!r.isDup).length}건)
                </button>
                <button style={{...settingsBtn("#e85d26"), padding:"8px 18px", fontSize:13}} onClick={()=>applyUpload("replace")}>
                  🔄 전체 교체 ({previewRows.length}건)
                </button>
                <button style={{...settingsBtn("#aaa", "#333"), padding:"8px 18px", fontSize:13}} onClick={()=>{setPreviewRows(null);setUploadStatus(null);}}>취소</button>
              </div>
            </div>
          )}
        </div>

        {/* 수신자 목록 */}
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:28}}>
          <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20}}>
            <div style={{fontSize:16, fontWeight:800, color:"#1a1a2e"}}>
              📧 메일 수신자 목록
              <span style={{fontSize:13, fontWeight:400, color:"#888", marginLeft:8}}>({mailRecipients.length}명)</span>
            </div>
            <button style={styles.aiBtn} onClick={startAdd}>+ 직접 추가</button>
          </div>

          <div style={{overflowX:"auto"}}>
            <table style={{...styles.table, marginBottom:20, minWidth:560}}>
              <thead>
                <tr>
                  <th style={{...styles.th, width:36}}>#</th>
                  <th style={styles.th}>이름</th>
                  <th style={styles.th}>이메일</th>
                  <th style={{...styles.th, width:160}}>등록일시</th>
                  <th style={{...styles.th, width:110}}>관리</th>
                </tr>
              </thead>
              <tbody>
                {mailRecipients.length === 0 && (
                  <tr><td colSpan={5} style={{...styles.td, textAlign:"center", color:"#aaa", padding:24}}>
                    수신자가 없습니다. 엑셀 업로드 또는 직접 추가해주세요.
                  </td></tr>
                )}
                {mailRecipients.map((r,i)=>(
                  <tr key={r.id}>
                    <td style={{...styles.td, color:"#aaa", fontSize:12}}>{i+1}</td>
                    <td style={styles.td}>{r.name}</td>
                    <td style={styles.td}>{r.email}</td>
                    <td style={{...styles.td, fontSize:12, color:"#888"}}>{r.registeredAt || "-"}</td>
                    <td style={styles.td}>
                      <button style={settingsBtn("#1d6fa4")} onClick={()=>startEdit(r)}>수정</button>
                      {" "}
                      <button style={settingsBtn("#e85d26")} onClick={()=>remove(r.id)}>삭제</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {(adding || editing) && (
            <div style={{background:"#f8f9fa", borderRadius:10, padding:20, display:"flex", flexDirection:"column", gap:12}}>
              <div style={{fontWeight:700, fontSize:14, color:"#333"}}>{adding?"수신자 직접 추가":"수신자 수정"}</div>
              <div style={{display:"flex", gap:10}}>
                <input style={{...styles.loginInput, flex:1}} placeholder="이름"
                  value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} />
                <input style={{...styles.loginInput, flex:2}} placeholder="이메일"
                  value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} />
              </div>
              <div style={{display:"flex", gap:8, justifyContent:"flex-end"}}>
                <button style={settingsBtn("#aaa", "#333")} onClick={cancel}>취소</button>
                <button style={settingsBtn("#e85d26")} onClick={save}>저장</button>
              </div>
            </div>
          )}
        </div>
        {/* SKN 담당자 섹션 */}
        <SKNRecipientsSection sknRecipients={sknRecipients} setSknRecipients={setSknRecipients} />

      </div>
      )}

      {/* 중복 팝업 */}
      {dupPopup && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:300, gap:14}}>
            <div style={{fontSize:40}}>⚠️</div>
            <div style={{fontSize:18, fontWeight:800, color:"#1a1a2e"}}>중복된 수신자입니다</div>
            <div style={{fontSize:13, color:"#666", textAlign:"center", whiteSpace:"pre-line"}}>{dupPopup}</div>
            <button style={styles2.popupBtn} onClick={()=>setDupPopup(null)}>확인</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── 배송 수량 설정 섹션 ─────────────────────────────────────────────────
const MULTIPLIERS = [2, 3];

// 본부 → 물류센터 매핑 (표2 커스텀 열의 "입고 수량으로 조정하기" 기능에서 사용)
const LOGISTICS_GROUPS = [
  { key:"이천", label:"이천물류",     본부s:["수도권","제주"] },
  { key:"부산", label:"부산물류",     본부s:["부산"] },
  { key:"대구", label:"대구물류",     본부s:["대구"] },
  { key:"광주", label:"광주물류",     본부s:["서부"] },
  { key:"중부", label:"중부물류(대전)", 본부s:["중부"] },
];

function ShippingGroupsSection({ shippingGroups, setShippingGroups, table1, setTable1, step, setStep, customCols, setCustomCols, nextColId, setNextColId, setStoreList }) {
  const storeFileRef = useRef(null);
  const [uploadStatus, setUploadStatus] = useState(null);
  const [logisticsPopup, setLogisticsPopup] = useState(null); // {colId, inputs:{이천,부산,대구,광주,중부}}
  const [totalQtyPopup, setTotalQtyPopup] = useState(null); // {colId, colLabel, total}
  // step, table1, customCols, nextColId는 App에서 props로 받음 (메뉴 전환 시 유지)

  // 물류센터 그룹별 입고 수량(총량)을, 체크박스 활성화된 행들끼리의 기존 비중대로 재분배
  const applyLogisticsAdjust = (colId, inputs) => {
    setShippingGroups(prev => {
      const next = prev.map(r => ({...r}));
      LOGISTICS_GROUPS.forEach(grp => {
        const total = Number(inputs[grp.key]) || 0;
        const activeRows = next.filter(r => grp.본부s.includes(r.본부) && r[`c${colId}`] !== false);
        if (activeRows.length === 0) return;
        const curSum = activeRows.reduce((s,r)=>s+(r[`c${colId}Val`]||0),0);
        activeRows.forEach(r => {
          r[`c${colId}Val`] = curSum > 0
            ? Math.round((r[`c${colId}Val`]||0) / curSum * total)
            : Math.round(total / activeRows.length);
        });
      });
      return next;
    });
  };

  // 총 수량을 체크박스 활성화된 행들끼리의 기존 비중대로 재분배 (물류센터 구분 없이 전체)
  const applyTotalQtyAdjust = (colId, total) => {
    setShippingGroups(prev => {
      const activeRows = prev.filter(r => r[`c${colId}`] !== false);
      if (activeRows.length === 0) return prev;
      const curSum = activeRows.reduce((s,r)=>s+(r[`c${colId}Val`]||0),0);
      return prev.map(r => {
        if (r[`c${colId}`] === false) return r;
        const newVal = curSum > 0
          ? Math.round((r[`c${colId}Val`]||0) / curSum * total)
          : Math.round(total / activeRows.length);
        return {...r, [`c${colId}Val`]: newVal};
      });
    });
  };

  // 시도명 → 지역본부 매핑
  const sidoToRegion = (sido) => {
    if (['서울','경기','인천','강원'].includes(sido)) return '수도권';
    if (['충남','충북','대전','세종'].includes(sido)) return '중부';
    if (['광주','전남','전북'].includes(sido)) return '서부';
    if (['대구','경북'].includes(sido)) return '대구';
    if (['부산','울산','경남'].includes(sido)) return '부산';
    if (['제주'].includes(sido)) return '제주';
    return null;
  };

  // 매장 엑셀 업로드 → 표1 생성
  const handleStoreUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploadStatus(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        // 집계: {구분_지역: count}
        // 헤더 기반 열 인덱스 자동 감지
        const header = (rows[0] || []).map(h => String(h||"").trim());
        const colIdx = (name) => { const i = header.indexOf(name); return i >= 0 ? i : -1; };
        const mktgCol      = colIdx("마케팅본부");
        const sidoCol      = colIdx("시도명");
        const storeCodeCol = colIdx("매장코드");
        const storeNameCol = colIdx("매장명");
        const agentCodeCol = colIdx("대리점코드");
        const agentNameCol = colIdx("대리점명");

        if (mktgCol < 0 || sidoCol < 0) {
          setUploadStatus({type:"error", msg:"열 이름을 찾을 수 없습니다. '마케팅본부', '시도명' 열이 있는지 확인해주세요."});
          return;
        }

        const counts = {};
        rows.slice(1).forEach(row => {
          const mktg = String(row[mktgCol]||"");
          const sido  = String(row[sidoCol]||"");
          const region = sidoToRegion(sido);
          if (!region) return;
          const 구분 = mktg.includes('유통사업부') ? 'PS&M' : '지역본부';
          const key = `${구분}_${region}`;
          counts[key] = (counts[key]||0) + 1;
        });
        // 총 소매 매장 수
        const totalSoMae = Object.values(counts).reduce((s,v)=>s+v,0);
        // 표1 생성
        const REGIONS = ['수도권','부산','대구','서부','제주','중부'];
        const newRows = [];
        let id = 1;
        for (const 구분 of ['지역본부','PS&M']) {
          for (const 본부 of REGIONS) {
            const key = `${구분}_${본부}`;
            const 소매매장 = counts[key] || 0;
            const adj102 = Math.round(소매매장 * 1.02);
            const 매장비중 = totalSoMae > 0 ? ((소매매장/totalSoMae)*100).toFixed(1)+"%" : "0%";
            newRows.push({
              id: id++, 구분, 본부,
              소매매장, adj102, 매장비중,
              biz: 구분==='지역본부' ? 0 : null,
              // 대형유통: 지역본부-수도권만 활성화
              대형유통: (구분==='지역본부' && 본부==='수도권') ? 0 : null,
              도매: 0,
            });
          }
        }
        // Biz / 대형 / 택배 행
        for (const 본부 of REGIONS) {
          newRows.push({ id:id++, 구분:"Biz", 본부, 소매매장:null, adj102:null, 매장비중:null, biz:0, 대형유통:null, 도매:null });
        }
        newRows.push({ id:id++, 구분:"대형", 본부:"-", 소매매장:null, adj102:null, 매장비중:null, biz:null, 대형유통:0, 도매:null });
        newRows.push({ id:id++, 구분:"택배배송", 본부:"-", 소매매장:null, adj102:null, 매장비중:null, biz:null, 대형유통:null, 도매:0 });
        setTable1(newRows);
        // 개별 매장 목록 저장 (GTM 누락 매장 비교용)
        const storeListParsed = rows.slice(1).filter(r => storeCodeCol >= 0 ? r[storeCodeCol] : r[5]).map(r=>({
          매장코드: String(storeCodeCol >= 0 ? r[storeCodeCol] : r[5]||"").trim(),
          매장명: String(storeNameCol >= 0 ? r[storeNameCol] : r[4]||""),
          대리점코드: String(agentCodeCol >= 0 ? r[agentCodeCol] : r[2]||"").trim(),
          대리점명: String(agentNameCol >= 0 ? r[agentNameCol] : r[3]||""),
          구분: String(r[mktgCol]||"").includes('유통사업부') ? 'PS&M' : '지역본부',
          본부: sidoToRegion(String(r[sidoCol]||"")) || "-",
        }));
        if (setStoreList) setStoreList(storeListParsed);
        setStep("upload");
        setUploadStatus({type:"success", msg:`${rows.length-1}개 매장 집계 완료`});
      } catch(err) {
        setUploadStatus({type:"error", msg:"오류: "+err.message});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const downloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["매장코드","마케팅본부","대리점코드","대리점명","매장명","우편번호","주소","시도명","시군구명","읍면동면"],
      ["D1","유통사업부","43","PS&M","PS&M (삼척점)","25920","강원 삼척시 중앙로 215","강원","삼척시","남양동"],
      ["A1","수도권마케팅담당","1","직영","직영 (강남)","06100","서울 강남구 테헤란로 1","서울","강남구","역삼동"],
    ]);
    ws["!cols"] = Array(10).fill({wch:16});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "매장리스트");
    XLSX.writeFile(wb, "매장리스트_템플릿.xlsx");
  };

  const updateT1 = (id, field, val) => {
    setTable1(prev => prev.map(r => r.id===id ? {...r, [field]:val} : r));
  };

  // 기본값으로 컨펌 → 표2 생성 (shippingGroups 업데이트)
  const confirmToTable2 = () => {
    setShippingGroups(prev => {
      // 대형 row (표1 기준): 수도권 행의 대형유통 값
      const t1Daehyung = table1.find(r => r.구분==="대형" || (r.본부==="수도권" && r.구분==="지역본부"));
      const daehyungVal = table1.find(r=>r.구분==="지역본부"&&r.본부==="수도권")?.대형유통 || 0;

      return prev.map(row => {
        const t1row = table1.find(r => r.구분===row.구분 && r.본부===row.본부);
        let next = row;
        if (!t1row) {
          // 대형 행: 표1 지역본부-수도권의 대형유통 값 사용
          if (row.구분==="대형") {
            const dv = table1.find(r=>r.본부==="수도권"&&r.구분==="지역본부")?.대형유통 ?? 0;
            next = {...row, 기본값: dv, g1:null, g2:null, g3:1, g4:1};
          }
          // 택배배송: 수기 입력 유지
        } else if (row.구분==="지역본부" || row.구분==="PS&M") {
          // 지역본부/PS&M: 102% 조정값 + 도매값
          next = {...row, 기본값: t1row.adj102 ?? t1row.소매매장 ?? row.기본값,
                         도매값: t1row.도매 ?? 0};
        } else if (row.구분==="Biz") {
          // Biz: 표1 지역본부 행의 biz 수기 입력값 (같은 본부)
          const bizT1 = table1.find(r => r.본부===row.본부 && r.구분==="지역본부");
          next = {...row, 기본값: bizT1?.biz ?? 0};
        }
        // 잠금 해제(🔓)된 커스텀 열은 새 기본값×배수로 자동 재계산.
        // 잠긴(🔒) 열은 손대지 않아 표1 변경의 영향을 받지 않는다.
        for (const col of customCols) {
          if (col.locked === false) {
            next = next===row ? {...row} : next;
            next[`c${col.id}Val`] = Math.round((next.기본값||0) * (col.mult||1));
          }
        }
        return next;
      });
    });
    setStep("table2");
  };

  // 표2 커스텀 열 추가/삭제
  const addCustomCol = () => {
    const label = `커스텀${nextColId}`;
    setCustomCols(prev => [...prev, {id:nextColId, label}]);
    setNextColId(n=>n+1);
    // shippingGroups에 새 열 추가
    setShippingGroups(prev => prev.map(r => ({...r, [`c${nextColId}`]:null})));
  };
  const removeCustomCol = (colId) => {
    setCustomCols(prev => prev.filter(c=>c.id!==colId));
  };

  // 표1 섹션별 집계
  // 표1에는 지역본부/PS&M만 (Biz/대형/택배배송은 표1에서 제외)
  const t1Sections = table1.length > 0 ? ['지역본부','PS&M'] : [];
  const totalSoMae = table1.filter(r=>r.소매매장!==null).reduce((s,r)=>s+(r.소매매장||0),0);
  const totalAdj  = table1.filter(r=>r.adj102!==null).reduce((s,r)=>s+(r.adj102||0),0);

  // 표2
  const GCOLS = ["g1","g2","g3","g_dm","g4"];
  const GLABELS = ["그룹1(소매Only)","그룹2(소매+Biz)","그룹3(소매+Biz+대형)","그룹3+도매","그룹4"];
  const MULTIPLIERS = [2,3];
  // 도매포함 그룹 값 계산 (기본값 + 도매값)
  const calcDmVal = (row) => {
    if (row.g_dm === null) return null;
    return (row.기본값||0) + (row.도매값||0);
  };
  const updateRow2 = (id, field, val) => setShippingGroups(prev=>prev.map(r=>r.id===id?{...r,[field]:val}:r));
  const calcVal = (row, gk) => {
    if (row[gk]===null) return null;
    // custom은 더 이상 사용 안 함, 배수만
    return Math.round((row.기본값||0)*(typeof row[gk]==="number" ? row[gk] : 1));
  };
  const sections2 = [...new Set(shippingGroups.map(r=>r.구분))];

  return (
    <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:28}}>
      <div style={{fontSize:16, fontWeight:800, color:"#1a1a2e", marginBottom:4}}>📦 배송 수량 설정</div>

      {/* 매장 업로드 영역 */}
      <div style={{background:"#f8f9fa", borderRadius:10, padding:"14px 18px", marginBottom:16}}>
        <div style={{fontSize:13, fontWeight:700, marginBottom:8}}>📂 매장 리스트 업로드</div>
        <div style={{fontSize:11, color:"#888", marginBottom:10}}>
          시도명 기준으로 지역본부/PS&M별 매장 수가 자동 집계됩니다.
        </div>
        <div style={{display:"flex", gap:8, flexWrap:"wrap", alignItems:"center"}}>
          <button style={{...settingsBtn("#1d6fa4"), padding:"7px 16px", fontSize:12}}
            onClick={()=>storeFileRef.current?.click()}>📂 파일선택</button>
          <button style={{...settingsBtn("#2e8b57"), padding:"7px 16px", fontSize:12}}
            onClick={downloadTemplate}>템플릿</button>
          <input ref={storeFileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
            onChange={handleStoreUpload} />
          {uploadStatus && (
            <span style={{fontSize:12, padding:"4px 10px", borderRadius:8,
              background:uploadStatus.type==="success"?"#e8f8ee":"#fff0f0",
              color:uploadStatus.type==="success"?"#2e8b57":"#c00"}}>
              {uploadStatus.type==="success"?"✅ ":"❌ "}{uploadStatus.msg}
            </span>
          )}
        </div>
      </div>

      {/* 탭 전환 */}
      {table1.length > 0 && (
        <div style={{display:"flex", gap:0, marginBottom:20, borderBottom:"2px solid #eee"}}>
          {[{id:"upload",label:"표1 · 매장 현황"},{id:"table2",label:"표2 · 수량 그룹 설정"}].map(t=>(
            <button key={t.id} onClick={()=>setStep(t.id)} style={{
              padding:"8px 24px", border:"none", background:"transparent",
              fontSize:13, fontWeight:700, cursor:"pointer",
              color:step===t.id?"#1d6fa4":"#aaa",
              borderBottom:step===t.id?"3px solid #1d6fa4":"3px solid transparent",
              marginBottom:-2,
            }}>{t.label}</button>
          ))}
        </div>
      )}

      {/* ── 표1 ── */}
      {(step==="upload" || table1.length===0) && (
        <>
          {table1.length === 0 ? (
            <div style={{textAlign:"center", padding:"40px 0", color:"#aaa", fontSize:14}}>
              매장 리스트 파일을 업로드하면 표1이 자동 생성됩니다.
            </div>
          ) : (
            <>
              <div style={{overflowX:"auto", marginBottom:16}}>
                <table style={{borderCollapse:"collapse", width:"100%", minWidth:680, fontSize:12}}>
                  <thead>
                    <tr style={{background:"#f0f4f8"}}>
                      <th style={tH}>구분</th><th style={tH}>본부</th>
                      <th style={tH}>소매 매장</th><th style={tH}>102% 조정</th>
                      <th style={tH}>매장 비중</th><th style={tH}>Biz</th>
                      <th style={tH}>대형유통</th><th style={tH}>도매</th>
                    </tr>
                  </thead>
                  <tbody>
                    {t1Sections.map(sec => {
                      const rows = table1.filter(r=>r.구분===sec);
                      return rows.map((row,ri)=>(
                        <tr key={row.id} style={{background:ri%2===0?"#fafafa":"#fff"}}>
                          {ri===0 && (
                            <td style={{...tD,fontWeight:700,textAlign:"center",
                              borderLeft:"3px solid #1d6fa4",background:"#f0f4f8"}}
                              rowSpan={rows.length}>{sec}</td>
                          )}
                          <td style={{...tD,textAlign:"center"}}>{row.본부}</td>
                          <td style={{...tD,textAlign:"center",fontWeight:600}}>
                            {row.소매매장!==null ? row.소매매장 : <span style={{color:"#ccc"}}>-</span>}
                          </td>
                          <td style={{...tD,textAlign:"center"}}>
                            {row.adj102!==null ? (
                              <div style={{display:"flex",alignItems:"center",gap:4,justifyContent:"center"}}>
                                <span style={{color:"#aaa",fontSize:11}}>={row.소매매장}×102%→</span>
                                <input type="number" min="0" style={{...numInput,width:60,color:"#1d6fa4",fontWeight:700}}
                                  value={row.adj102 ?? ""}
                                  onChange={e=>updateT1(row.id,"adj102", e.target.value==="" ? 0 : parseInt(e.target.value)||0)} />
                              </div>
                            ) : <span style={{color:"#ccc"}}>-</span>}
                          </td>
                          <td style={{...tD,textAlign:"center"}}>
                            {row.매장비중!==null ? row.매장비중 : <span style={{color:"#ccc"}}>-</span>}
                          </td>
                          {/* Biz */}
                          <td style={{...tD,textAlign:"center",background:row.biz===null?"#eee":"transparent"}}>
                            {row.biz===null
                              ? <span style={{color:"#aaa",fontSize:11}}>비활성화</span>
                              : <input type="number" min="0" style={numInput}
                                  value={row.biz ?? ""} placeholder="0"
                                  onChange={e=>updateT1(row.id,"biz", e.target.value==="" ? 0 : parseInt(e.target.value)||0)} />
                            }
                          </td>
                          {/* 대형유통: 지역본부-수도권만 활성 */}
                          <td style={{...tD,textAlign:"center",background:row.대형유통===null?"#f0f0f0":"transparent"}}>
                            {row.대형유통===null
                              ? <span style={{color:"#aaa",fontSize:11}}>비활성화</span>
                              : <input type="number" min="0" style={numInput}
                                  value={row.대형유통 ?? ""} placeholder="0"
                                  onChange={e=>updateT1(row.id,"대형유통", e.target.value==="" ? 0 : parseInt(e.target.value)||0)} />
                            }
                          </td>
                          {/* 도매 */}
                          <td style={{...tD,textAlign:"center"}}>
                            <input type="number" min="0" style={numInput}
                              value={row.도매 ?? ""} placeholder="0"
                              onChange={e=>updateT1(row.id,"도매", e.target.value==="" ? 0 : parseInt(e.target.value)||0)} />
                          </td>
                        </tr>
                      ));
                    })}
                    {/* 계 행 */}
                    <tr style={{background:"#f0f4f8",fontWeight:700}}>
                      <td style={{...tD,textAlign:"center"}} colSpan={2}>계</td>
                      <td style={{...tD,textAlign:"center"}}>{totalSoMae}</td>
                      <td style={{...tD,textAlign:"center",color:"#1d6fa4"}}>{totalAdj}</td>
                      <td style={{...tD,textAlign:"center"}}>100%</td>
                      <td style={{...tD,textAlign:"center"}}>
                        {table1.filter(r=>r.biz!==null).reduce((s,r)=>s+(r.biz||0),0)}
                      </td>
                      <td style={{...tD,textAlign:"center"}}>
                        {table1.filter(r=>r.대형유통!==null&&(r.구분==='지역본부'&&r.본부==='수도권')).reduce((s,r)=>s+(r.대형유통||0),0)}
                      </td>
                      <td style={{...tD,textAlign:"center"}}>
                        {table1.filter(r=>r.도매!==null).reduce((s,r)=>s+(r.도매||0),0)}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div style={{textAlign:"center"}}>
                <button style={{
                  background:"#1d6fa4", color:"#fff", border:"none", borderRadius:50,
                  padding:"12px 40px", fontSize:15, fontWeight:700, cursor:"pointer",
                  boxShadow:"0 2px 12px rgba(29,111,164,0.25)",
                }} onClick={confirmToTable2}>기본값으로 컨펌하기</button>
              </div>
            </>
          )}
        </>
      )}

      {/* ── 표2 ── */}
      {step==="table2" && (
        <>
          <div style={{display:"flex", gap:10, justifyContent:"flex-end", marginBottom:12, flexWrap:"wrap"}}>
            <button style={{...settingsBtn("#1d6fa4"), padding:"7px 18px", fontSize:12}}
              onClick={()=>{
                // 현재 남아있는 열들의 이름을 기준으로 비어있는 가장 작은 번호를 붙인다
                // (예: 커스텀1을 지웠다가 다시 만들면 커스텀1부터 다시 채워짐)
                const usedLabels = new Set(customCols.map(c=>c.label));
                let n = 1;
                while (usedLabels.has(`커스텀${n}`)) n++;
                const label = `커스텀${n}`;
                const newId = nextColId;
                setCustomCols(prev=>[...prev,{id:newId,label,mult:1,locked:true}]);
                setNextColId(n=>n+1);
                setShippingGroups(prev=>prev.map(r=>({...r,[`c${newId}`]:true,[`c${newId}Val`]:Math.round((r.기본값||0)*1)})));
              }}>+ 열 추가 커스텀</button>
            {customCols.length > 0 && (
              <button style={{...settingsBtn("#e85d26"), padding:"7px 18px", fontSize:12}}
                onClick={()=>{
                  const last=customCols[customCols.length-1];
                  setCustomCols(prev=>prev.filter(c=>c.id!==last.id));
                }}>열 삭제</button>
            )}
          </div>

          <div style={{overflowX:"auto", borderRadius:10, border:"1px solid #dde"}}>
            <table style={{borderCollapse:"collapse", fontSize:12, width:"100%", minWidth:600}}>
              <thead>
                {/* 헤더 행 */}
                <tr style={{background:"#f0f4f8"}}>
                  <th style={{...tH,width:70,borderLeft:"4px solid #1d6fa4"}}>구분</th>
                  <th style={{...tH,width:60}}>본부</th>
                  <th style={{...tH,width:90}}>기본 값</th>
                  {/* 그룹1~3: 라벨만 */}
                  {[["그룹1","그룹1(소매Only)"],["그룹2","그룹2(소매+Biz)"],["그룹3","그룹3(소매+Biz+대형)"]].map(([k,label])=>(
                    <th key={k} style={tH}>{label}</th>
                  ))}
                  {/* 그룹3+도매 */}
                  <th style={{...tH, background:"#f0ecff", color:"#7b52d9"}}>그룹3+도매</th>
                  {/* 그룹4: X2 X3 버튼 */}
                  <th style={tH}>
                    그룹4
                    <div style={{display:"flex",gap:3,justifyContent:"center",marginTop:4}}>
                      {[2,3].map(m=>(
                        <button key={m} style={{background:"#1d6fa4",color:"#fff",border:"none",
                          borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}
                          onClick={()=>setShippingGroups(prev=>prev.map(r=>
                            r.g4!==null?{...r,g4:m}:r))}>×{m}</button>
                      ))}
                    </div>
                  </th>
                  {/* 커스텀 열: X1 X2 X3 일괄 버튼 */}
                  {customCols.map(col=>(
                    <th key={col.id} style={tH}>
                      <button
                        title={col.locked!==false
                          ? "잠김: 표1에서 \"기본값으로 컨펌하기\"를 다시 눌러도 이 열의 수량은 그대로 유지됩니다 (직접 수정은 언제든 가능). 클릭하면 잠금 해제"
                          : "해제됨: 표1에서 다시 컨펌하면 기본값×배수로 자동 재계산됩니다 (그 전까지는 직접 수정 가능). 클릭하면 다시 잠금"}
                        style={{border:"none",background:"transparent",cursor:"pointer",marginRight:3,fontSize:13,padding:0}}
                        onClick={()=>{
                          const willLock = col.locked===false; // 현재 해제 상태면 → 잠그는 전환
                          setCustomCols(prev=>prev.map(c=>c.id===col.id?{...c,locked:willLock}:c));
                          if (willLock) {
                            // 잠그는 순간의 계산값을 스냅샷으로 저장해서 그 이후로는 표1과 무관하게 유지
                            setShippingGroups(prev=>prev.map(r=>({...r,[`c${col.id}Val`]:Math.round((r.기본값||0)*(col.mult||1))})));
                          }
                        }}>{col.locked!==false ? "🔒" : "🔓"}</button>
                      <input style={{width:64,border:"1px solid #ddd",borderRadius:6,
                        padding:"2px 4px",fontSize:11,textAlign:"center"}}
                        value={col.label}
                        onChange={e=>setCustomCols(prev=>prev.map(c=>c.id===col.id?{...c,label:e.target.value}:c))} />
                      <button style={{...settingsBtn("#e85d26"),marginLeft:4,padding:"1px 5px",fontSize:10}}
                        onClick={()=>setCustomCols(prev=>prev.filter(c=>c.id!==col.id))}>×</button>
                      <div style={{display:"flex",gap:2,justifyContent:"center",marginTop:4}}>
                        {[1,2,3].map(m=>(
                          <button key={m} title="현재 기본값 기준으로 전체 행에 일괄 적용" style={{
                            background:col.mult===m?"#1d6fa4":"#e8eef5",
                            color:col.mult===m?"#fff":"#333",
                            border:"none",borderRadius:10,padding:"2px 7px",fontSize:10,fontWeight:700,cursor:"pointer"}}
                            onClick={()=>{
                              setCustomCols(prev=>prev.map(c=>c.id===col.id?{...c,mult:m}:c));
                              setShippingGroups(prev=>prev.map(r=>({...r,[`c${col.id}Val`]:Math.round((r.기본값||0)*m)})));
                            }}>×{m}</button>
                        ))}
                      </div>
                      <button title="물류센터별 입고 수량을 입력하면, 활성화(체크)된 행들끼리 기존 비중대로 나눠 반영합니다"
                        style={{...settingsBtn("#2e8b57"),marginTop:4,padding:"2px 6px",fontSize:10,width:"100%"}}
                        onClick={()=>{
                          const init = {};
                          LOGISTICS_GROUPS.forEach(g=>{
                            init[g.key] = shippingGroups
                              .filter(r=>g.본부s.includes(r.본부) && r[`c${col.id}`]!==false)
                              .reduce((s,r)=>s+(r[`c${col.id}Val`]||0),0);
                          });
                          setLogisticsPopup({colId:col.id, colLabel:col.label, inputs:init});
                        }}>📦 입고수량 조정</button>
                      <button title="전체 수량을 입력하면, 활성화(체크)된 행들끼리의 기존 비중대로 나눠 반영합니다"
                        style={{...settingsBtn("#1d6fa4"),marginTop:4,padding:"2px 6px",fontSize:10,width:"100%"}}
                        onClick={()=>{
                          const curTotal = shippingGroups
                            .filter(r=>r[`c${col.id}`]!==false)
                            .reduce((s,r)=>s+(r[`c${col.id}Val`]||0),0);
                          setTotalQtyPopup({colId:col.id, colLabel:col.label, total:curTotal});
                        }}>📦 총 수량으로 조정</button>
                    </th>
                  ))}
                </tr>
                {/* 총계 행 */}
                <tr style={{background:"#dce8f5",fontWeight:700,borderTop:"3px solid #6b8cba"}}>
                  <td style={{...tD,textAlign:"center",borderLeft:"4px solid #1d6fa4"}} colSpan={2}>총 계</td>
                  <td style={{...tD,textAlign:"center",color:"#1d6fa4"}}>
                    {shippingGroups.filter(r=>r.active).reduce((s,r)=>s+(r.기본값||0),0)}
                  </td>
                  {["g1","g2","g3"].map(gk=>(
                    <td key={gk} style={{...tD,textAlign:"center",color:"#1d6fa4"}}>
                      {shippingGroups.filter(r=>r.active&&r[gk]!==null).reduce((s,r)=>s+(r.기본값||0),0)}
                    </td>
                  ))}
                  <td style={{...tD,textAlign:"center",color:"#7b52d9",background:"#f0ecff"}}>
                    {shippingGroups.filter(r=>r.active&&r.g_dm!==null).reduce((s,r)=>s+((r.기본값||0)+(r.도매값||0)),0)}
                  </td>
                  <td style={{...tD,textAlign:"center",color:"#1d6fa4"}}>
                    {shippingGroups.filter(r=>r.active&&r.g4!==null&&r.g4Active!==false).reduce((s,r)=>s+Math.round((r.기본값||0)*(r.g4||1)),0)}
                  </td>
                  {customCols.map(col=>(
                    <td key={col.id} style={{...tD,textAlign:"center",color:"#1d6fa4"}}>
                      {shippingGroups.filter(r=>r.active&&r[`c${col.id}`]!==false).reduce((s,r)=>s+(r[`c${col.id}Val`]||0),0)}
                    </td>
                  ))}
                </tr>
              </thead>
              <tbody>
                {[...new Set(shippingGroups.map(r=>r.구분))].map(sec=>{
                  const rows=shippingGroups.filter(r=>r.구분===sec);
                  return rows.map((row,ri)=>(
                    <tr key={row.id} style={{
                      background:ri%2===0?"#fafafa":"#fff",
                      opacity:row.active?1:0.4,
                      borderTop:ri===0?"3px solid #6b8cba":"none",
                    }}>
                      {ri===0&&(
                        <td style={{...tD,fontWeight:700,textAlign:"center",
                          borderLeft:"4px solid #1d6fa4",background:"#e8f0fb",
                          borderTop:"3px solid #6b8cba"}} rowSpan={rows.length}>{sec}</td>
                      )}
                      <td style={{...tD,textAlign:"center"}}>{row.본부}</td>
                      {/* 기본값 */}
                      <td style={{...tD,textAlign:"center"}}>
                        {row.구분==="택배배송"
                          ? <input type="number" min="0" style={{...numInput,width:60}}
                              value={row.기본값 ?? ""} placeholder="0"
                              onChange={e=>setShippingGroups(prev=>prev.map(r=>r.id===row.id?{...r,기본값:e.target.value===""?0:parseInt(e.target.value)||0}:r))} />
                          : <><span style={{fontWeight:700,color:"#1d6fa4",fontSize:13}}>{row.기본값||0}</span>
                            <div style={{fontSize:10,color:"#aaa"}}>102% 조정값</div></>
                        }
                      </td>
                      {/* 그룹1~3: 기본값 그대로, null이면 비활성 */}
                      {["g1","g2","g3"].map(gk=>(
                        <td key={gk} style={{...tD,textAlign:"center"}}>
                          {row[gk]===null
                            ? <span style={{color:"#ccc",fontSize:11,background:"#f0f0f0",display:"block",padding:"4px 0",borderRadius:4}}>비활성화</span>
                            : <span style={{fontWeight:700,color:"#333",fontSize:13}}>{row.기본값||0}</span>
                          }
                        </td>
                      ))}
                      {/* 그룹3+도매 */}
                      <td style={{...tD,textAlign:"center",background:"#f8f5ff"}}>
                        {row.g_dm===null
                          ? <span style={{color:"#ccc",fontSize:11,background:"#f0f0f0",display:"block",padding:"4px 0",borderRadius:4}}>비활성화</span>
                          : <span style={{fontWeight:700,color:"#7b52d9",fontSize:13}}>
                              {(row.기본값||0)+(row.도매값||0)}
                              {row.도매값>0&&<div style={{fontSize:10,color:"#aaa"}}>{row.기본값||0}+{row.도매값||0}</div>}
                            </span>
                        }
                      </td>
                      {/* 그룹4: 배수 적용 + 체크박스 */}
                      <td style={{...tD,textAlign:"center"}}>
                        {row.g4===null
                          ? <span style={{color:"#ccc",fontSize:11,background:"#f0f0f0",display:"block",padding:"4px 0",borderRadius:4}}>비활성화</span>
                          : <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                              <span style={{fontWeight:700,color:row.g4Active===false?"#ccc":"#1d6fa4",fontSize:13,
                                textDecoration:row.g4Active===false?"line-through":"none"}}>
                                {Math.round((row.기본값||0)*(row.g4||1))}
                              </span>
                              <input type="checkbox" checked={row.g4Active!==false} style={{width:16,height:16,accentColor:"#2e8b57",cursor:"pointer"}}
                                onChange={e=>setShippingGroups(prev=>prev.map(r=>r.id===row.id?{...r,g4Active:e.target.checked}:r))} />
                            </div>
                        }
                      </td>
                      {/* 커스텀 열: 잠금 상태와 무관하게 항상 수량 직접 입력 가능.
                          🔒 잠김 = 표1을 다시 컨펌해도 이 값 유지. 🔓 해제 = 표1 재컨펌 시 기본값×배수로 자동 재계산(그 전까지는 자유 수정 가능). */}
                      {customCols.map(col=>(
                        <td key={col.id} style={{...tD,textAlign:"center"}}>
                          <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
                            <input type="number" min="0" style={{...numInput,width:56,
                              color:row[`c${col.id}`]===false?"#ccc":"#1d6fa4",fontWeight:700,
                              textDecoration:row[`c${col.id}`]===false?"line-through":"none"}}
                              value={row[`c${col.id}Val`] ?? 0}
                              onChange={e=>{
                                const v = parseInt(e.target.value)||0;
                                setShippingGroups(prev=>prev.map(r=>r.id===row.id?{...r,[`c${col.id}Val`]:v}:r));
                              }} />
                            <input type="checkbox" checked={row[`c${col.id}`]!==false} style={{width:16,height:16,accentColor:"#2e8b57",cursor:"pointer"}}
                              onChange={e=>setShippingGroups(prev=>prev.map(r=>r.id===row.id?{...r,[`c${col.id}`]:e.target.checked}:r))} />
                          </div>
                        </td>
                      ))}
                    </tr>
                  ));
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {/* 입고 수량으로 조정하기 팝업 */}
      {logisticsPopup && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:340, gap:14, alignItems:"stretch"}}>
            <div style={{fontSize:16, fontWeight:800}}>📦 입고 수량으로 조정 — {logisticsPopup.colLabel}</div>
            <div style={{fontSize:11.5, color:"#888"}}>
              물류센터별 입고 수량을 입력하면, 해당 물류센터로 가는 본부들 중 체크박스가 활성화된 행들끼리
              현재 수량의 비중대로 나눠서 반영됩니다.
            </div>
            {LOGISTICS_GROUPS.map(g=>(
              <div key={g.key} style={{display:"flex", alignItems:"center", justifyContent:"space-between", gap:10}}>
                <label style={{fontSize:13, fontWeight:600, color:"#333"}}>{g.label}</label>
                <input type="number" min="0" style={{...numInput, width:100}}
                  value={logisticsPopup.inputs[g.key] ?? ""}
                  onChange={e=>{
                    // 입력 중에는 원본 문자열을 그대로 보관 (즉시 숫자로 강제 변환하면
                    // 지웠을 때 0이 남아있다가 뒤에 이어 쓰는 값과 붙어버리는 문제가 생김)
                    const raw = e.target.value;
                    setLogisticsPopup(prev=>({...prev, inputs:{...prev.inputs, [g.key]: raw===""?"":raw}}));
                  }} />
              </div>
            ))}
            <div style={{display:"flex", gap:10, marginTop:4}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>setLogisticsPopup(null)}>취소</button>
              <button style={{...styles2.popupBtn, flex:1}}
                onClick={()=>{
                  applyLogisticsAdjust(logisticsPopup.colId, logisticsPopup.inputs);
                  setLogisticsPopup(null);
                }}>적용</button>
            </div>
          </div>
        </div>
      )}

      {/* 총 수량으로 조정 팝업 */}
      {totalQtyPopup && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:320, gap:14, alignItems:"stretch"}}>
            <div style={{fontSize:16, fontWeight:800}}>📦 총 수량으로 조정 — {totalQtyPopup.colLabel}</div>
            <div style={{fontSize:11.5, color:"#888"}}>
              입력한 총 수량을, 이 열에서 체크박스가 활성화된 행들끼리 현재 수량의 비중대로 나눠서 반영합니다.
            </div>
            <input type="number" min="0" style={{...numInput, width:"100%"}}
              value={totalQtyPopup.total ?? ""}
              onChange={e=>{
                const raw = e.target.value;
                setTotalQtyPopup(prev=>({...prev, total: raw===""?"":raw}));
              }} />
            <div style={{display:"flex", gap:10, marginTop:4}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>setTotalQtyPopup(null)}>취소</button>
              <button style={{...styles2.popupBtn, flex:1}}
                onClick={()=>{
                  applyTotalQtyAdjust(totalQtyPopup.colId, Number(totalQtyPopup.total)||0);
                  setTotalQtyPopup(null);
                }}>적용</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


const tH = { padding:"8px 10px", border:"1px solid #e0e6ed", textAlign:"center",
  fontWeight:700, color:"#444", fontSize:12 };
const tD = { padding:"6px 8px", border:"1px solid #eee", fontSize:12 };
const numInput = { width:60, border:"1px solid #ddd", borderRadius:6, padding:"3px 6px",
  fontSize:12, textAlign:"center", outline:"none" };
const multBtn = { border:"none", borderRadius:8, padding:"2px 7px", fontSize:11,
  fontWeight:700, cursor:"pointer" };

// ─── SKN 담당자 섹션 서브컴포넌트 ───────────────────────────────────────
function SKNRecipientsSection({ sknRecipients, setSknRecipients }) {
  const [adding, setAdding]   = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm]       = useState({name:"", email:""});
  const [dupPopup, setDupPopup] = useState(null);

  const nowStr = () => {
    const d = new Date(), p = n=>String(n).padStart(2,"0");
    return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  };
  const nextId = (list) => Math.max(0, ...(list.length?list.map(r=>r.id):[0]))+1;
  const isDup  = (email, list) => list.some(r=>r.email.toLowerCase()===email.toLowerCase());

  const startAdd  = () => { setAdding(true); setEditing(null); setForm({name:"",email:""}); };
  const startEdit = (r) => { setEditing(r); setAdding(false); setForm({name:r.name,email:r.email}); };
  const cancel    = () => { setAdding(false); setEditing(null); };

  const save = () => {
    if (!form.name.trim() || !form.email.trim()) return;
    if (editing) {
      setSknRecipients(prev=>prev.map(r=>r.id===editing.id?{...r,...form}:r));
      cancel();
    } else {
      if (isDup(form.email, sknRecipients)) { setDupPopup(`이미 등록된 이메일입니다.\n(${form.email})`); return; }
      setSknRecipients(prev=>[...prev,{id:nextId(prev), name:form.name, email:form.email, registeredAt:nowStr()}]);
      cancel();
    }
  };
  const remove = (id) => setSknRecipients(prev=>prev.filter(r=>r.id!==id));

  return (
    <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:28}}>
      <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20}}>
        <div>
          <div style={{fontSize:16, fontWeight:800, color:"#1a1a2e"}}>
            📨 SKN 담당자 메일 목록
            <span style={{fontSize:13, fontWeight:400, color:"#888", marginLeft:8}}>({sknRecipients.length}명)</span>
          </div>
          <div style={{fontSize:12, color:"#888", marginTop:4}}>'SKN에 메일 보내기' 버튼의 수신처입니다.</div>
        </div>
        <button style={styles.aiBtn} onClick={startAdd}>+ 추가</button>
      </div>

      <div style={{overflowX:"auto"}}>
        <table style={{...styles.table, marginBottom:20, minWidth:500}}>
          <thead>
            <tr>
              <th style={{...styles.th, width:36}}>#</th>
              <th style={styles.th}>이름</th>
              <th style={styles.th}>이메일</th>
              <th style={{...styles.th, width:160}}>등록일시</th>
              <th style={{...styles.th, width:110}}>관리</th>
            </tr>
          </thead>
          <tbody>
            {sknRecipients.length === 0 && (
              <tr><td colSpan={5} style={{...styles.td, textAlign:"center", color:"#aaa", padding:24}}>
                SKN 담당자를 추가해주세요.
              </td></tr>
            )}
            {sknRecipients.map((r,i)=>(
              <tr key={r.id}>
                <td style={{...styles.td, color:"#aaa", fontSize:12}}>{i+1}</td>
                <td style={styles.td}>{r.name}</td>
                <td style={styles.td}>{r.email}</td>
                <td style={{...styles.td, fontSize:12, color:"#888"}}>{r.registeredAt||"-"}</td>
                <td style={styles.td}>
                  <button style={settingsBtn("#1d6fa4")} onClick={()=>startEdit(r)}>수정</button>
                  {" "}
                  <button style={settingsBtn("#e85d26")} onClick={()=>remove(r.id)}>삭제</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {(adding || editing) && (
        <div style={{background:"#f8f9fa", borderRadius:10, padding:20, display:"flex", flexDirection:"column", gap:12}}>
          <div style={{fontWeight:700, fontSize:14, color:"#333"}}>{adding?"SKN 담당자 추가":"SKN 담당자 수정"}</div>
          <div style={{display:"flex", gap:10}}>
            <input style={{...styles.loginInput, flex:1}} placeholder="이름"
              value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} />
            <input style={{...styles.loginInput, flex:2}} placeholder="이메일"
              value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} />
          </div>
          <div style={{display:"flex", gap:8, justifyContent:"flex-end"}}>
            <button style={settingsBtn("#aaa", "#333")} onClick={cancel}>취소</button>
            <button style={settingsBtn("#e85d26")} onClick={save}>저장</button>
          </div>
        </div>
      )}

      {dupPopup && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:300, gap:14}}>
            <div style={{fontSize:40}}>⚠️</div>
            <div style={{fontSize:18, fontWeight:800, color:"#1a1a2e"}}>중복된 담당자입니다</div>
            <div style={{fontSize:13, color:"#666", textAlign:"center", whiteSpace:"pre-line"}}>{dupPopup}</div>
            <button style={styles2.popupBtn} onClick={()=>setDupPopup(null)}>확인</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── 본사 배송 물품 페이지 ───────────────────────────────────────────────
function HQItemsPage({ hqItems, setHqItems, shippingGroups, confirmed, role, shippingCustomCols, addNotification }) {
  const isAdmin = role === "admin";
  const today = new Date();
  // 21일 이후면 다음달 기준 기간(당월21~익월20)으로 초기화
  const initMonth = (() => {
    if (today.getDate() >= 21) {
      return today.getMonth() === 11 ? 0 : today.getMonth() + 1;
    }
    return today.getMonth();
  })();
  const initYear = today.getDate() >= 21 && today.getMonth() === 11
    ? today.getFullYear() + 1 : today.getFullYear();
  const [viewYear,  setViewYear]  = useState(initYear);
  const [viewMonth, setViewMonth] = useState(initMonth);
  const [activeTab, setActiveTab] = useState(null); // dateKey or "pending"
  const [addingItem, setAddingItem] = useState(false);
  const [newItemName, setNewItemName] = useState("");
  const [reassignPopup, setReassignPopup] = useState(null);
  const [qtyPopup, setQtyPopup] = useState(null);
  const [sianPopup, setSianPopup] = useState(null); // {itemId, itemName, images:[]}
  const [noticeDatePopup, setNoticeDatePopup] = useState(null); // {itemId}

  const periodStart = new Date(viewYear, viewMonth-1, 21);
  const periodEnd   = new Date(viewYear, viewMonth, 20, 23,59,59);

  // ① 확정된 배송일만 (status==="confirmed") 탭으로 사용
  const deliveryDates = Object.entries(confirmed)
    .filter(([k,v]) => {
      const dt = new Date(k);
      return dt >= periodStart && dt <= periodEnd && v.status === "confirmed";
    })
    .sort(([a],[b])=>a.localeCompare(b));

  // 이전에 confirmed였다가 현재는 없어진(수정된) 날짜 → Pending 탭
  // hqItems 중 assignedDate가 현재 deliveryDates에 없는 것들
  const confirmedKeys = new Set(deliveryDates.map(([k])=>k));

  // 현재 activeTab이 없으면 오늘 기준 가장 가까운(지나지 않은) 배송일로 초기화, 없으면 첫 배송일
  const todayKey = toKey(today.getFullYear(), today.getMonth(), today.getDate());
  const effectiveTab = activeTab
    || (deliveryDates.find(([k]) => k >= todayKey)?.[0] ?? deliveryDates[0]?.[0] ?? "pending");

  const fmtDate = (k) => {
    if (!k) return "-";
    const [,m,d] = k.split("-");
    return `${parseInt(m)}월 ${parseInt(d)}일`;
  };

  // 영업일 기준 D-1 계산 (주말·공휴일 건너뜀)
  const prevBizDay = (dateKey) => {
    if (!dateKey) return null;
    let d = new Date(dateKey);
    d.setDate(d.getDate() - 1);
    while (d.getDay()===0 || d.getDay()===6 || KR_HOLIDAYS.has(toKey(d.getFullYear(), d.getMonth(), d.getDate()))) {
      d.setDate(d.getDate() - 1);
    }
    return toKey(d.getFullYear(), d.getMonth(), d.getDate());
  };

  const addItem = () => {
    if (!newItemName.trim()) return;
    const assignedDate = effectiveTab === "pending" ? null : effectiveTab;
    const ingoDate     = prevBizDay(assignedDate);
    const noticeDate   = prevBizDay(ingoDate);
    const name = newItemName.trim();
    setHqItems(prev => [...prev, {
      id: Date.now(), name, qty: 0,
      group: null, assignedDate,
      ingoSudo: ingoDate, ingoJibang: ingoDate, noticeDate, publicNotice:false, note:"",
    }]);
    addNotification?.(
      `📦 [본사 배송 물품] "${name}" 추가됨${assignedDate ? ` — ${fmtDate(assignedDate)}` : ""}`,
      'success'
    );
    setNewItemName(""); setAddingItem(false);
  };

  const fieldLabel = {
    ingoSudo:    'SKN 입고 (수도권)',
    ingoJibang:  'SKN 입고 (지방)',
    noticeDate:  '배송공지 날짜',
    assignedDate:'배정 날짜',
    qty:         '수량',
  };

  const updateItem = (id, field, val) => {
    const item = hqItems.find(it => it.id === id);
    if (item && fieldLabel[field] && val) {
      const valStr = (field === 'qty') ? `${val}개` : fmtDate(val);
      addNotification?.(
        `✏️ [${item.name}] ${fieldLabel[field]} → ${valStr}`,
        'edit'
      );
    }
    setHqItems(prev => prev.map(it => it.id===id ? {...it, [field]:val} : it));
  };
  const removeItem = (id) => {
    const item = hqItems.find(it => it.id === id);
    addNotification?.(`🗑 [본사 배송 물품] "${item?.name}" 삭제됨`, 'delete');
    setHqItems(prev => prev.filter(it => it.id!==id));
  };
  const groupOptions = [...new Set(shippingGroups.map(r=>r.구분))];

  // 배송표 엑셀 내보내기
  const exportDeliveryExcel = (item) => {
    const REGIONS = ['수도권','제주','부산','대구','서부','중부'];
    const GROUP_KEY_MAP = { "그룹1(소매Only)":"g1","그룹2(소매+Biz)":"g2","그룹3(소매+Biz+대형)":"g3","그룹3+도매":"g_dm","그룹4":"g4" };
    const gKey = GROUP_KEY_MAP[item.qtyGroup];
    // qtyGroup이 커스텀 열 이름일 수도 있음 — 이 경우 기본값×배수가 아니라
    // 행마다 저장된 실제 수량(c{id}Val)을 그대로 써야 함
    const customCol = shippingCustomCols.find(c => c.label === item.qtyGroup);
    const active = shippingGroups.filter(r => r.active);
    const fmtD = (k) => { if(!k) return ""; const[,m,d]=k.split("-"); return `${parseInt(m)}월 ${parseInt(d)}일`; };
    const calcV = (row,gk) => { if(!gk||row[gk]===null) return null; return Math.round((row.기본값||0)*(typeof row[gk]==="number"?row[gk]:1)); };
    const getVal = (구분,본부) => {
      const row = active.find(r=>r.구분===구분&&r.본부===본부);
      if (!row) return null;
      if (customCol) return row[`c${customCol.id}`]===false ? null : (row[`c${customCol.id}Val`] ?? 0);
      return calcV(row,gKey);
    };

    const jb  = REGIONS.map(r => getVal("지역본부",r) ?? 0);
    const psm = REGIONS.map(r => getVal("PS&M",r) ?? 0);
    const bizRaw = REGIONS.map(r => getVal("Biz",r));
    const biz = bizRaw.map(v => (v!==null&&v>0)?v:null);
    const daehyung = getVal("대형","수도권") ?? 0;
    const hasBiz = biz.some(v=>v!==null);

    const jbTotal  = jb.reduce((s,v)=>s+v,0) + daehyung;
    const psmTotal = psm.reduce((s,v)=>s+v,0);
    const bizTotal = biz.reduce((s,v)=>s+(v||0),0);
    const icheon   = jb[0]+jb[1]+psm[0]+psm[1]+(biz[0]||0)+daehyung;
    const busanC   = jb[2]+psm[2]+(biz[2]||0);
    const daeguC   = jb[3]+psm[3]+(biz[3]||0);
    const gwangju  = jb[4]+psm[4]+(biz[4]||0);
    const daejeon  = jb[5]+psm[5]+(biz[5]||0);
    const dataRows = hasBiz ? 3 : 2;
    const nm = `(본) ${item.name}`;

    // ── 스타일 정의 ──
    const thin = { style:"thin", color:{rgb:"FF000000"} };
    const bd   = { top:thin, bottom:thin, left:thin, right:thin };
    const C1 = "FFD9D9D9"; // 라벨행 회색
    const C2 = "FFBFBFBF"; // 헤더 회색
    const C3 = "FFDCE6F1"; // 데이터 연파랑
    const mkS = (fill, bold=true, align="center", wrap=true) => ({
      font: { bold, name:"맑은 고딕", sz:10 },
      fill: fill ? { fgColor:{rgb:fill} } : undefined,
      alignment: { horizontal:align, vertical:"center", wrapText:wrap },
      border: bd,
    });
    const mkN = (fill, bold=true) => ({ ...mkS(fill,bold,"center",false), numFmt:"#,##0" });

    const ws = {};
    const sc = (r,c,v,s) => {
      const addr = XLSX.utils.encode_cell({r,c});
      const t = typeof v==="number" ? "n" : "s";
      ws[addr] = { v: v??""  , t: (v===null||v===undefined)?"s":t, s };
    };

    // Row 0 (Excel 1): 타이틀
    sc(0,0,"입 고 정 보",{ font:{bold:true,sz:14,name:"맑은 고딕"}, alignment:{horizontal:"center",vertical:"center"} });
    for(let c=1;c<13;c++) sc(0,c,"",{});

    // Row 1 (Excel 2): 빈 행
    for(let c=0;c<13;c++) sc(1,c,"",{});

    // Row 2 (Excel 3): 물류 주관부서
    sc(2,0,"물류 주관부서",mkS(C1,true,"left"));
    sc(2,1,"채널지원팀",mkS(null,false,"left"));
    sc(2,2,"",{}); sc(2,3,"",{});
    for(let c=4;c<13;c++) sc(2,c,"",{});

    // Row 3 (Excel 4): 담당자
    sc(3,0,"담당자 (연락처)",mkS(C1,true,"left"));
    sc(3,1,"",mkS(null,false,"left")); sc(3,2,"",{}); sc(3,3,"",{});
    for(let c=4;c<13;c++) sc(3,c,"",{});

    // Row 4 (Excel 5): 빈 행
    for(let c=0;c<13;c++) sc(4,c,"",{});

    // Row 5 (Excel 6): 물류센터별 입고수량
    sc(5,0,"■ 물류 센터별 입고수량",{ font:{bold:true,sz:10,name:"맑은 고딕"}, alignment:{horizontal:"left",vertical:"center"} });
    sc(5,1,icheon,mkN(null)); sc(5,2,"",mkS(null));
    sc(5,3,busanC,mkN(null)); sc(5,4,daeguC,mkN(null));
    sc(5,5,gwangju,mkN(null)); sc(5,6,daejeon,mkN(null));
    for(let c=7;c<13;c++) sc(5,c,"",{});

    // Rows 6-8 (Excel 7-9): 헤더
    sc(6,0,"구   분",mkS(C2)); sc(6,1,"수     량",mkS(C2));
    for(let c=2;c<9;c++) sc(6,c,"",mkS(C2));
    sc(6,9,"입고예정일",mkS(C2)); sc(6,10,"출고요청일",mkS(C2));
    sc(6,11,"배포완료일",mkS(C2)); sc(6,12,"업체명",mkS(C2));

    sc(7,0,"",mkS(C2)); sc(7,1,"이천물류",mkS(C2)); sc(7,2,"",mkS(C2));
    sc(7,3,"부산물류",mkS(C2)); sc(7,4,"대구물류",mkS(C2));
    sc(7,5,"광주물류",mkS(C2)); sc(7,6,"대전물류",mkS(C2));
    sc(7,7,"이천물류",mkS(C2)); sc(7,8,"총계",mkS(C2));
    for(let c=9;c<13;c++) sc(7,c,"",mkS(C2));

    sc(8,0,"",mkS(C2)); sc(8,1,"수도권",mkS(C2)); sc(8,2,"제주",mkS(C2));
    sc(8,3,"부산",mkS(C2)); sc(8,4,"대구",mkS(C2));
    sc(8,5,"서부",mkS(C2)); sc(8,6,"중부",mkS(C2));
    sc(8,7,"대형양판점",mkS(C2)); sc(8,8,"",mkS(C2));
    for(let c=9;c<13;c++) sc(8,c,"",mkS(C2));

    // 데이터 행
    const dataRowDefs = [
      { label: nm+"\n(지역본부_소매)", vals: jb },
      { label: nm+"\n(PS&M_소매)",    vals: psm },
      ...(hasBiz ? [{ label: nm+"\n(비즈)", vals: biz }] : []),
    ];
    dataRowDefs.forEach((row,ri) => {
      const r = 9+ri;
      sc(r,0,row.label,mkS(C3,true,"left"));
      row.vals.forEach((v,ci) => sc(r,ci+1,(v!==null?v:""),mkN(C3)));
      // 병합 컬럼들 (첫 행에만 값)
      if(ri===0) {
        sc(r,7,daehyung||"",mkN(C3));
        sc(r,8,jbTotal+psmTotal+bizTotal,mkN(C3));
        sc(r,9,fmtD(item.ingoSudo),mkS(C3,false));
        sc(r,10,fmtD(item.assignedDate),mkS(C3,false));
        sc(r,11,fmtD(item.assignedDate),mkS(C3,false));
        sc(r,12,"",mkS(C3,false));
      } else {
        for(let c=7;c<13;c++) sc(r,c,"",mkS(C3,false));
      }
    });

    // 푸터
    const fs = 9+dataRows+1;
    sc(fs,0,"■ 특이사항",{ font:{bold:true,sz:10,name:"맑은 고딕"}, alignment:{horizontal:"left"} });
    ["1. 용도 : ","2. 상품관련 : ","3. 배포대상 : ","4. 참고사항 : "].forEach((t,i)=>
      sc(fs+1+i,0,t,{ font:{sz:10,name:"맑은 고딕"}, alignment:{horizontal:"left"} })
    );
    sc(fs+6,0,' ※ 상품 발송시 "납품확인서"를 첨부하여 보내도록 제작업체에 전달 바랍니다!',{ font:{sz:10,name:"맑은 고딕"}, alignment:{horizontal:"left"} });
    sc(fs+7,0,' ※ SKN 물류센터 입고시 입고확인증 지참 부탁드립니다.',{ font:{sz:10,name:"맑은 고딕"}, alignment:{horizontal:"left"} });

    // 병합
    ws["!merges"] = [
      {s:{r:0,c:0},e:{r:0,c:12}},           // A1:M1 타이틀
      {s:{r:2,c:1},e:{r:2,c:3}},            // B3:D3 채널지원팀
      {s:{r:3,c:1},e:{r:3,c:3}},            // B4:D4 담당자값
      {s:{r:5,c:1},e:{r:5,c:2}},            // B6:C6 이천합계
      {s:{r:6,c:0},e:{r:8,c:0}},            // A7:A9 구분
      {s:{r:6,c:1},e:{r:6,c:8}},            // B7:I7 수량
      {s:{r:6,c:9},e:{r:8,c:9}},            // J7:J9 입고예정일
      {s:{r:6,c:10},e:{r:8,c:10}},          // K7:K9 출고요청일
      {s:{r:6,c:11},e:{r:8,c:11}},          // L7:L9 배포완료일
      {s:{r:6,c:12},e:{r:8,c:12}},          // M7:M9 업체명
      {s:{r:7,c:1},e:{r:7,c:2}},            // B8:C8 이천물류
      {s:{r:7,c:8},e:{r:8,c:8}},            // I8:I9 총계
      {s:{r:9,c:7},e:{r:9+dataRows-1,c:7}}, // H 대형 병합
      {s:{r:9,c:8},e:{r:9+dataRows-1,c:8}}, // I 총계 병합
      {s:{r:9,c:9},e:{r:9+dataRows-1,c:9}}, // J 입고예정일 병합
      {s:{r:9,c:10},e:{r:9+dataRows-1,c:10}},
      {s:{r:9,c:11},e:{r:9+dataRows-1,c:11}},
      {s:{r:9,c:12},e:{r:9+dataRows-1,c:12}},
    ];
    ws["!ref"] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:fs+7,c:12}});
    ws["!cols"] = [{wch:34},{wch:10},{wch:7},{wch:10},{wch:10},{wch:10},{wch:10},{wch:12},{wch:9},{wch:13},{wch:13},{wch:13},{wch:16}];
    ws["!rows"] = [
      {hpt:24},{hpt:6},{hpt:18},{hpt:18},{hpt:6},{hpt:18},
      {hpt:18},{hpt:18},{hpt:18},
      ...Array(dataRows).fill({hpt:48}),
    ];

    const wb = { SheetNames:["SKN입고정보"], Sheets:{"SKN입고정보":ws} };
    XLSX.writeFile(wb, `배송표_${item.name}.xlsx`);
  };

  // 탭별 아이템 필터
  const tabItems = hqItems.filter(it => {
    if (effectiveTab === "pending") {
      return !it.assignedDate || !confirmedKeys.has(it.assignedDate);
    }
    return it.assignedDate === effectiveTab;
  });

  // 월 표시: "5월 기준 (4월 21일~5월 20일)"
  const monthLabel = `${viewYear}년 ${viewMonth+1}월 기준 (${viewMonth}월 21일~${viewMonth+1}월 20일)`;

  return (
    <div style={styles.page}>
      {/* 헤더 */}
      <div style={{display:"flex", alignItems:"center", gap:12, marginBottom:20, flexWrap:"wrap"}}>
        <h1 style={{...styles.pageTitle, margin:0}}>본사 배송 물품</h1>
        <div style={{display:"flex", alignItems:"center", gap:6}}>
          <button style={styles.navArrow} onClick={()=>{
            setActiveTab(null);
            if(viewMonth===0){setViewYear(y=>y-1);setViewMonth(11);}else setViewMonth(m=>m-1);
          }}>‹</button>
          <span style={{fontSize:13, fontWeight:700, color:"#555", whiteSpace:"nowrap"}}>{monthLabel}</span>
          <button style={styles.navArrow} onClick={()=>{
            setActiveTab(null);
            if(viewMonth===11){setViewYear(y=>y+1);setViewMonth(0);}else setViewMonth(m=>m+1);
          }}>›</button>
        </div>
        {isAdmin && (
          <button style={{...styles.aiBtn, marginLeft:"auto"}} onClick={()=>setAddingItem(true)}>
            + 배송물품 추가하기
          </button>
        )}
      </div>

      {/* 날짜 탭 */}
      <div style={{display:"flex", gap:8, marginBottom:20, flexWrap:"wrap", alignItems:"center"}}>
        {deliveryDates.length === 0 && (
          <div style={{fontSize:13, color:"#aaa"}}>이 기간에 확정된 배송일(빨간 원)이 없습니다.</div>
        )}
        {deliveryDates.map(([k]) => (
          <button key={k}
            style={{
              padding:"8px 18px", borderRadius:50, border:"none", cursor:"pointer",
              background: effectiveTab===k ? "#1d6fa4" : "#e0e6ed",
              color: effectiveTab===k ? "#fff" : "#444",
              fontSize:13, fontWeight:700, transition:"all 0.15s",
            }}
            onClick={()=>setActiveTab(k)}>
            {fmtDate(k)}
          </button>
        ))}
        {/* 재배정 대기 탭 */}
        <button
          style={{
            padding:"8px 18px", borderRadius:50, border:"none", cursor:"pointer",
            background: effectiveTab==="pending" ? "#e85d26" : "#ffeee8",
            color: effectiveTab==="pending" ? "#fff" : "#e85d26",
            fontSize:13, fontWeight:700, transition:"all 0.15s",
          }}
          onClick={()=>setActiveTab("pending")}>
          ⚠️ 날짜 재배정 필요
          {hqItems.filter(it=>!it.assignedDate||!confirmedKeys.has(it.assignedDate)).length > 0 &&
            <span style={{marginLeft:6, background:"#c00", color:"#fff",
              borderRadius:10, padding:"0 6px", fontSize:11}}>
              {hqItems.filter(it=>!it.assignedDate||!confirmedKeys.has(it.assignedDate)).length}
            </span>
          }
        </button>
      </div>

      {/* 물품 테이블 */}
      <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse", width:"100%", minWidth:780, fontSize:13}}>
            <thead>
              <tr style={{background:"#f0f4f8"}}>
                <th style={tH}>#</th>
                <th style={{...tH, minWidth:120}}>Item</th>
                <th style={{...tH, width:70}}>시안</th>
                <th style={{...tH, width:80}}>수량</th>
                <th style={{...tH, minWidth:120}}>SKN 입고 (수도권)</th>
                <th style={{...tH, minWidth:120}}>SKN 입고 (지방)</th>
                <th style={{...tH, width:70}}>배송공지</th>
                <th style={{...tH, minWidth:110}}>배송공지 날짜</th>
                <th style={{...tH, minWidth:130}}>비고</th>
                <th style={{...tH, width:100}}>날짜 배정</th>
                <th style={{...tH, width:90}}>배송표</th>
                {isAdmin && <th style={{...tH, width:50}}>삭제</th>}
              </tr>
            </thead>
            <tbody>
              {tabItems.length === 0 && (
                <tr><td colSpan={isAdmin?11:10} style={{...tD, textAlign:"center", padding:28, color:"#aaa"}}>
                  {effectiveTab==="pending" ? "재배정이 필요한 물품이 없습니다." : "이 날짜에 배정된 물품이 없습니다."}
                </td></tr>
              )}
              {tabItems.map((item,idx)=>(
                <tr key={item.id} style={{background:idx%2===0?"#fafafa":"#fff"}}>
                  <td style={{...tD, textAlign:"center", color:"#aaa"}}>{idx+1}</td>
                  <td style={tD}>
                    <span style={{fontWeight:600}}>{item.name}</span>
                  </td>
                  {/* 시안 */}
                  <td style={{...tD, textAlign:"center"}}>
                    <button style={{
                      background: item.sianImages?.length>0?"#e8f8ee":"#f5f5f5",
                      border:`1px solid ${item.sianImages?.length>0?"#2e8b57":"#dde"}`,
                      borderRadius:8, padding:"4px 10px", fontSize:11, cursor:"pointer",
                      color: item.sianImages?.length>0?"#2e8b57":"#888",
                    }}
                      onClick={()=>setSianPopup({itemId:item.id, itemName:item.name, images:item.sianImages||[]})}>
                      {item.sianImages?.length>0 ? `🖼 ${item.sianImages.length}개` : "시안"}
                    </button>
                  </td>
                  {/* 수량 */}
                  <td style={{...tD, textAlign:"center"}}>
                    <button style={{
                      background: item.qty>0?"#e8f0fb":"#f5f5f5",
                      border:`1px solid ${item.qty>0?"#1d6fa4":"#dde"}`,
                      borderRadius:8, padding:"4px 10px",
                      fontSize:12, cursor:"pointer", color:"#1d6fa4", fontWeight:item.qty>0?700:400,
                    }}
                      onClick={()=>setQtyPopup({itemId:item.id, itemName:item.name})}>
                      {item.qty>0 ? item.qty : "선택"}
                    </button>
                  </td>
                  <td style={{...tD, textAlign:"center"}}>
                    <CalendarPickerCell
                      value={item.ingoSudo}
                      onChange={val=>updateItem(item.id,"ingoSudo",val)} />
                  </td>
                  <td style={{...tD, textAlign:"center"}}>
                    <CalendarPickerCell
                      value={item.ingoJibang}
                      onChange={val=>updateItem(item.id,"ingoJibang",val)} />
                  </td>
                  {/* 배송공지 */}
                  <td style={{...tD, textAlign:"center"}}>
                    <input type="checkbox" checked={!!item.publicNotice}
                      onChange={e=>updateItem(item.id,"publicNotice",e.target.checked)} disabled={!isAdmin} />
                  </td>
                  {/* 배송공지 날짜 */}
                  <td style={{...tD, textAlign:"center"}}>
                    <CalendarPickerCell
                      value={item.noticeDate}
                      onChange={val=>updateItem(item.id,"noticeDate",val)} />
                  </td>
                  <td style={tD}>
                    {isAdmin
                      ? <input style={{...numInput, width:"100%", textAlign:"left"}} value={item.note||""} placeholder="비고"
                          onChange={e=>updateItem(item.id,"note",e.target.value)} />
                      : item.note||"-"}
                  </td>
                  <td style={{...tD, textAlign:"center"}}>
                    <button style={{...settingsBtn("#1d6fa4"), padding:"4px 10px", fontSize:11}}
                      onClick={()=>setReassignPopup({itemId:item.id, name:item.name})}>
                      {effectiveTab === "pending" ? "날짜 배정" : "재배정"}
                    </button>
                  </td>
                  <td style={{...tD, textAlign:"center"}}>
                    <button
                      style={{...settingsBtn("#2e8b57"), padding:"4px 10px", fontSize:11,
                        opacity: item.qty > 0 ? 1 : 0.4,
                        cursor: item.qty > 0 ? "pointer" : "not-allowed"}}
                      onClick={()=>{ if(item.qty > 0) exportDeliveryExcel(item); }}
                      title={item.qty > 0 ? "배송표 엑셀 다운로드" : "수량 그룹을 먼저 선택하세요"}>
                      📊 배송표
                    </button>
                  </td>
                  {isAdmin && (
                    <td style={{...tD, textAlign:"center"}}>
                      <button style={{background:"#e85d26", border:"none", cursor:"pointer", color:"#fff",
                        width:22, height:22, borderRadius:"50%", fontSize:16, fontWeight:700,
                        display:"inline-flex", alignItems:"center", justifyContent:"center", lineHeight:1, padding:0}}
                        onClick={()=>removeItem(item.id)} title="삭제">−</button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* 물품 추가 팝업 */}
      {addingItem && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:320, gap:14}}>
            <div style={{fontSize:18, fontWeight:800}}>배송 물품 추가</div>
            <input style={{...styles.loginInput, width:"100%"}} placeholder="물품명 입력"
              value={newItemName} onChange={e=>setNewItemName(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter"){e.preventDefault();addItem();}}} autoFocus />
            <div style={{display:"flex", gap:10, width:"100%"}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1, padding:"10px 12px", marginTop:0, fontSize:13}}
                onClick={()=>{setAddingItem(false);setNewItemName("");}}>취소</button>
              <button style={{...styles2.popupBtn, flex:1, padding:"10px 12px", marginTop:0, fontSize:13}} onClick={addItem}>추가</button>
            </div>
          </div>
        </div>
      )}

      {/* 시안 팝업 */}
      {sianPopup && (
        <SianPopup
          itemId={sianPopup.itemId}
          itemName={sianPopup.itemName}
          images={sianPopup.images}
          isAdmin={isAdmin}
          onSave={(imgs)=>{ updateItem(sianPopup.itemId,"sianImages",imgs); setSianPopup(null); }}
          onClose={()=>setSianPopup(null)}
        />
      )}

      {/* 수량 선택 팝업 */}
      {qtyPopup && (() => {
        // shippingGroups에서 그룹별 합계 계산
        const calcV2 = (row, gk) => {
          if (row[gk]===null) return null;
          return Math.round((row.기본값||0)*(typeof row[gk]==="number"?row[gk]:1));
        };
        const calcDm2 = (row) => row.g_dm!==null ? (row.기본값||0)+(row.도매값||0) : null;
        const activeRows = shippingGroups.filter(r=>r.active);
        const groups = [
          { label:"그룹1(소매Only)", key:"g1",
            total: activeRows.filter(r=>r.g1!==null).reduce((s,r)=>s+(calcV2(r,"g1")||0),0) },
          { label:"그룹2(소매+Biz)", key:"g2",
            total: activeRows.filter(r=>r.g2!==null).reduce((s,r)=>s+(calcV2(r,"g2")||0),0) },
          { label:"그룹3(소매+Biz+대형)", key:"g3",
            total: activeRows.filter(r=>r.g3!==null).reduce((s,r)=>s+(calcV2(r,"g3")||0),0) },
          { label:"그룹3+도매", key:"g_dm",
            total: activeRows.filter(r=>r.g_dm!==null).reduce((s,r)=>s+(calcDm2(r)||0),0) },
          { label:"그룹4", key:"g4",
            // g4Active 체크박스 반영: false면 해당 행 제외
            total: activeRows.filter(r=>r.g4!==null&&r.g4Active!==false).reduce((s,r)=>s+Math.round((r.기본값||0)*(r.g4||1)),0) },
          ...shippingCustomCols.map(col=>({
            label: col.label,
            key: `c${col.id}`,
            total: activeRows.filter(r=>r[`c${col.id}`]!==false).reduce((s,r)=>s+(r[`c${col.id}Val`]||0),0),
          })),
        ].filter(g=>g.total>0);
        return (
          <div style={styles2.popupOverlay}>
            <div style={{...styles2.popupBox, minWidth:320, gap:14}}>
              <div style={{fontSize:18, fontWeight:800}}>수량 그룹 선택</div>
              <div style={{fontSize:13, color:"#666"}}><b>{qtyPopup.itemName}</b>의 발송 수량 그룹을 선택하세요.</div>
              <div style={{width:"100%", display:"flex", flexDirection:"column", gap:8}}>
                {groups.map(g=>(
                  <button key={g.key}
                    style={{
                      ...styles2.popupBtn,
                      background: "#1d6fa4",
                      padding:"12px 20px",
                      display:"flex", justifyContent:"space-between", alignItems:"center",
                    }}
                    onClick={()=>{
                      updateItem(qtyPopup.itemId, "qty", g.total);
                      updateItem(qtyPopup.itemId, "qtyGroup", g.label);
                      setQtyPopup(null);
                    }}>
                    <span>{g.label}</span>
                    <span style={{fontSize:18, fontWeight:800}}>{g.total.toLocaleString()}개</span>
                  </button>
                ))}
                {groups.length===0 && (
                  <div style={{textAlign:"center",color:"#aaa",fontSize:13,padding:16}}>
                    설정 → 배송 수량 설정에서 수량을 먼저 입력해주세요.
                  </div>
                )}
              </div>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", width:"100%"}}
                onClick={()=>setQtyPopup(null)}>취소</button>
            </div>
          </div>
        );
      })()}

      {/* 날짜 재배정 팝업 */}
      {reassignPopup && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:320, gap:14}}>
            <div style={{fontSize:18, fontWeight:800}}>날짜 재배정</div>
            <div style={{fontSize:13, color:"#666"}}>
              <b>{reassignPopup.name}</b>을 배정할 날짜를 선택하세요.
            </div>
            <div style={{width:"100%", display:"flex", flexDirection:"column", gap:8}}>
              {deliveryDates.map(([k])=>(
                <button key={k}
                  style={{...styles2.popupBtn, background:"#1d6fa4", padding:"10px"}}
                  onClick={()=>{
                    updateItem(reassignPopup.itemId, "assignedDate", k);
                    setReassignPopup(null);
                    setActiveTab(k);
                  }}>
                  {fmtDate(k)} 배정
                </button>
              ))}
            </div>
            <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", width:"100%"}}
              onClick={()=>setReassignPopup(null)}>취소</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── 일반 캘린더 날짜 선택 ──────────────────────────────────────────────
function CalendarPickerCell({ value, onChange }) {
  const [open, setOpen] = useState(false);
  const wrapRef  = useRef(null);
  const today    = new Date();
  const [calYear,  setCalYear]  = useState(today.getFullYear());
  const [calMonth, setCalMonth] = useState(today.getMonth());
  const pad = n => String(n).padStart(2,"0");
  const fmtDisp = v => (v && typeof v === 'string')
    ? `${parseInt(v.split('-')[1])}월 ${parseInt(v.split('-')[2])}일`
    : "날짜 선택";
  const daysInMonth = new Date(calYear, calMonth+1, 0).getDate();
  const firstDow    = new Date(calYear, calMonth, 1).getDay();
  // ▶ 빈칸과 날짜 셀에 통합 key = 배열 인덱스 사용 → key 충돌 방지
  const cells = Array.from({length:firstDow},()=>null)
                     .concat(Array.from({length:daysInMonth},(_,i)=>i+1));

  // 날짜 색상 (주말·공휴일 빨강, 토 파랑)
  const getDayColor = (d, selected) => {
    if (selected) return "#fff";
    const dKey = `${calYear}-${pad(calMonth+1)}-${pad(d)}`;
    const dow  = new Date(calYear, calMonth, d).getDay();
    if (KR_HOLIDAYS.has(dKey) || dow === 0) return "#e85d26";
    if (dow === 6) return "#0066cc";
    return C.text0;
  };

  // 외부 클릭 시 닫기
  useEffect(() => {
    if (!open) return;
    const handleClick = e => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  return (
    <div ref={wrapRef} style={{position:"relative",display:"inline-block"}}>
      <button style={{background:value?"#e8f0fb":"var(--c-surface2)",border:`1px solid ${value?"#1d6fa4":C.border}`,
        borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer",color:value?"#1d6fa4":C.text2}}
        onClick={()=>setOpen(o=>!o)}>{fmtDisp(value)}</button>
      {open && (
        <div style={{position:"fixed",zIndex:9999,background:"var(--c-popup-bg)",borderRadius:12,
          boxShadow:"0 8px 32px rgba(0,0,0,0.18)",padding:16,width:256,border:`1px solid ${C.border}`}}
          ref={el=>{if(el){const btn=wrapRef.current?.querySelector('button');if(btn){const r=btn.getBoundingClientRect();
            el.style.top=(r.bottom+4)+"px";el.style.left=Math.min(r.left,window.innerWidth-260)+"px";}}}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <button style={{background:"none",border:"none",fontSize:16,cursor:"pointer",padding:"0 6px",color:C.text0}}
              onClick={()=>{if(calMonth===0){setCalYear(y=>y-1);setCalMonth(11);}else setCalMonth(m=>m-1);}}>‹</button>
            <span style={{fontWeight:700,fontSize:13,color:C.text0}}>{calYear}. {pad(calMonth+1)}.</span>
            <button style={{background:"none",border:"none",fontSize:16,cursor:"pointer",padding:"0 6px",color:C.text0}}
              onClick={()=>{if(calMonth===11){setCalYear(y=>y+1);setCalMonth(0);}else setCalMonth(m=>m+1);}}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",marginBottom:4}}>
            {["일","월","화","수","목","금","토"].map((label,i)=>(
              <div key={label} style={{textAlign:"center",fontSize:11,fontWeight:600,padding:"2px 0",
                color:i===0?"#fb7185":i===6?"#60a5fa":C.text2}}>{label}</div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
            {cells.map((d,i)=>{
              // ▶ key=i (배열 인덱스) → 빈칸·날짜 셀 key 충돌 없음
              if(!d) return <div key={i} style={{width:30,height:30}}/>;
              const dKey=`${calYear}-${pad(calMonth+1)}-${pad(d)}`;
              const isSel=value===dKey;
              const isHol=KR_HOLIDAYS.has(dKey);
              return (
                <button key={i} onClick={()=>{onChange(dKey);setOpen(false);}} style={{
                  background:isSel?"#1d6fa4":"transparent",border:"none",borderRadius:"50%",
                  color:getDayColor(d,isSel),
                  width:30,height:30,margin:"auto",cursor:"pointer",fontSize:12,fontWeight:isSel?700:400,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  opacity: isHol && !isSel ? 0.85 : 1,
                }}>{d}</button>
              );
            })}
          </div>
          <button style={{width:"100%",marginTop:8,padding:"5px",background:"var(--c-surface2)",
            border:`1px solid ${C.border}`,borderRadius:8,cursor:"pointer",fontSize:12,color:C.text1}}
            onClick={()=>{onChange(null);setOpen(false);}}>초기화</button>
        </div>
      )}
    </div>
  );
}

// 날짜 선택 셀 컴포넌트
function DatePickerCell({ value, dates, onChange, fmtDate, confirmed }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);
  // value: {dateKey: "YYYY-MM-DD"} or {}
  const selectedKey = value?.dateKey || null;

  // 외부 클릭 시 닫기
  useEffect(() => {
    if (!open) return;
    const handleClick = e => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  // 배송일 바로 직전 평일 자동 계산
  const prevWeekday = (dateKey) => {
    if (!dateKey) return null;
    let d = new Date(dateKey);
    d.setDate(d.getDate()-1);
    while (d.getDay()===0||d.getDay()===6||KR_HOLIDAYS.has(toKey(d.getFullYear(),d.getMonth(),d.getDate()))) {
      d.setDate(d.getDate()-1);
    }
    return toKey(d.getFullYear(), d.getMonth(), d.getDate());
  };

  const autoDefault = dates.length>0 ? prevWeekday(dates[0][0]) : null;
  const displayKey = selectedKey || autoDefault;
  const [,m,d] = displayKey ? displayKey.split("-") : ["","",""];
  const displayStr = displayKey ? `${parseInt(m)}월 ${parseInt(d)}일` : "-";

  return (
    <div ref={wrapRef} style={{position:"relative"}}>
      <button style={{background:"#f0f7ff", border:"1px solid #cde", borderRadius:8,
        padding:"4px 10px", fontSize:12, cursor:"pointer", color:"#1d6fa4"}}
        onClick={()=>setOpen(o=>!o)}>
        {displayStr}
      </button>
      {open && (
        <div style={{position:"fixed", zIndex:9999, background:"#fff", borderRadius:10,
          boxShadow:"0 4px 20px rgba(0,0,0,0.15)", padding:12, minWidth:180,
          top:"auto", left:"auto", transform:"none",
          maxHeight:"70vh", overflowY:"auto"}}
          ref={el=>{
            if(el){
              const rect=el.getBoundingClientRect();
              const btn=el.previousSibling;
              if(btn){
                const bRect=btn.getBoundingClientRect();
                el.style.top=(bRect.bottom+4)+"px";
                el.style.left=Math.min(bRect.left, window.innerWidth-200)+"px";
              }
            }
          }}>
          <div style={{fontSize:11, color:"#aaa", marginBottom:6}}>날짜 선택</div>
          {dates.map(([k]) => {
            const prev = prevWeekday(k);
            return (
              <div key={k} style={{marginBottom:6}}>
                <div style={{fontSize:11, color:"#888", marginBottom:2}}>배송일: {fmtDate(k)}</div>
                <button style={{
                  width:"100%", background: selectedKey===prev?"#1d6fa4":"#f5f5f5",
                  color: selectedKey===prev?"#fff":"#333",
                  border:"none", borderRadius:6, padding:"5px 8px", fontSize:12, cursor:"pointer",
                }}
                  onClick={()=>{onChange({dateKey:prev});setOpen(false);}}>
                  {prev ? (()=>{const[,pm,pd]=prev.split("-");return `${parseInt(pm)}월 ${parseInt(pd)}일 (직전 평일)`;})() : "날짜 없음"}
                </button>
              </div>
            );
          })}
          <button style={{width:"100%", background:"#fff", border:"1px solid #ddd",
            borderRadius:6, padding:"4px", fontSize:11, cursor:"pointer", color:"#aaa",
            marginTop:4}}
            onClick={()=>{onChange({});setOpen(false);}}>초기화</button>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// GTM 취합 페이지
// ═══════════════════════════════════════════════════════════════════════
const GTM_TABS = [
  { id:"widecolor",  label:"와이드컬러" },
  { id:"hanging",    label:"행잉배너" },
  { id:"a3acrylic",  label:"A3 아크릴 삽지" },
  { id:"expdesk",    label:"체험대 백월" },
  { id:"largegfx",   label:"라지그래픽" },
];

// 신설 매장 리스트의 "마케팅본부/마케팅팀" 값을 본부(수도권/부산/대구/서부/제주/중부)로 매핑
const GTM_REGIONS = ["수도권","부산","대구","서부","제주","중부","유통사업부"];
// 와이드컬러 업로드 시 "어느 소속인가요?" 선택지 (제주 제외, 유통사업부 포함)
const WIDECOLOR_UPLOAD_SCOPES = ["수도권","부산","대구","서부","중부","유통사업부"];
const GTM_TEAM_TO_REGION = {
  "대전마케팅팀":"중부", "서대구마케팅팀":"대구", "서부소매사업팀":"서부",
  "서광주마케팅팀":"서부", "동대구마케팅팀":"대구", "남부마케팅팀":"수도권",
  "중앙마케팅팀":"수도권", "충북마케팅팀":"중부", "북부마케팅팀":"수도권",
  "동광주마케팅팀":"서부", "인천마케팅팀":"수도권", "경남마케팅팀":"부산",
  "서부마케팅팀":"서부", "중부산마케팅팀":"부산", "동부산마케팅팀":"부산",
  "충남마케팅팀":"중부", "강원마케팅팀":"수도권", "수도권2소매사업팀":"수도권",
  "경북마케팅팀":"대구",
};
function guessGtmRegion(mktgHQ, mktgTeam) {
  const hq = String(mktgHQ||"").trim();
  if (hq.includes("유통사업부")) return "유통사업부";
  const direct = GTM_REGIONS.find(r => hq.startsWith(r));
  if (direct) return direct;
  const team = String(mktgTeam||"").trim();
  if (team.includes("유통사업부")) return "유통사업부";
  if (GTM_TEAM_TO_REGION[team]) return GTM_TEAM_TO_REGION[team];
  // 부산/대구 계열을 먼저 체크 (중부산·경북 등이 다른 지역명을 부분 포함하는 것 방지)
  if (team.includes("부산") || team.includes("경남") || team.includes("울산")) return "부산";
  if (team.includes("대구") || team.includes("경북")) return "대구";
  if (team.includes("광주") || team.includes("전남") || team.includes("전북") || team.includes("서부")) return "서부";
  if (team.includes("제주")) return "제주";
  if (team.includes("충남") || team.includes("충북") || team.includes("대전") || team.includes("세종")) return "중부";
  if (team.includes("인천") || team.includes("강원") || team.includes("수도권")) return "수도권";
  return "";
}
// 이미 정규화된 본부명이면 그대로, 아니면(예: "중부마케팅담당") 다시 정규화 시도
function normalizeGtmRegion(raw) {
  const s = String(raw||"").trim();
  if (GTM_REGIONS.includes(s)) return s;
  return guessGtmRegion(s, "") || s || "-";
}

function GTMPage({ gtmWideColorData, setGtmWideColorData, gtmHangingData, setGtmHangingData, gtmSubmissions, setGtmSubmissions, gtmNewStoreList, setGtmNewStoreList, gtmLargeGfxRounds, setGtmLargeGfxRounds, gtmLargeGfxDraft, setGtmLargeGfxDraft, gtmLargeGfxPhotos, setGtmLargeGfxPhotos, gtmLargeGfxCompareOverrides, setGtmLargeGfxCompareOverrides, role }) {
  const [activeTab, setActiveTab] = useState("widecolor");
  const isAdmin = role === "admin";

  const setSectionSubmissions = (sectionKey) => (updater) => {
    setGtmSubmissions(prev => {
      const cur = prev[sectionKey] || {};
      const next = typeof updater === "function" ? updater(cur) : updater;
      return { ...prev, [sectionKey]: next };
    });
  };

  return (
    <div style={styles.page}>
      <div style={styles.pageHeader}>
        <h1 style={styles.pageTitle}>GTM 취합</h1>
      </div>
      {/* 탭 네비게이션 */}
      <div style={{display:"flex", gap:0, marginBottom:24, borderBottom:"2px solid #eee"}}>
        {GTM_TABS.map(tab=>(
          <button key={tab.id} onClick={()=>setActiveTab(tab.id)} style={{
            padding:"10px 24px", border:"none", background:"transparent",
            fontSize:14, fontWeight:700, cursor:"pointer",
            color: activeTab===tab.id ? "#1d6fa4" : "#aaa",
            borderBottom: activeTab===tab.id ? "3px solid #1d6fa4" : "3px solid transparent",
            marginBottom:-2, transition:"all 0.15s",
          }}>{tab.label}</button>
        ))}
      </div>

      {activeTab === "widecolor"
        ? <GTMCollectSection
            key="widecolor"
            variant="store"
            data={gtmWideColorData}
            setData={setGtmWideColorData}
            newStoreList={gtmNewStoreList}
            setNewStoreList={setGtmNewStoreList}
            isAdmin={isAdmin}
            submissions={gtmSubmissions.widecolor || {}}
            setSubmissions={setSectionSubmissions("widecolor")}
            sectionLabel="와이드컬러"
          />
        : activeTab === "hanging"
        ? <GTMCollectSection
            key="hanging"
            variant="hq"
            data={gtmHangingData}
            setData={setGtmHangingData}
            isAdmin={isAdmin}
            submissions={gtmSubmissions.hanging || {}}
            setSubmissions={setSectionSubmissions("hanging")}
            sectionLabel="행잉배너"
          />
        : activeTab === "largegfx"
        ? <GTMLargeGfx
            key="largegfx"
            rounds={gtmLargeGfxRounds}
            setRounds={setGtmLargeGfxRounds}
            draft={gtmLargeGfxDraft}
            setDraft={setGtmLargeGfxDraft}
            photos={gtmLargeGfxPhotos}
            setPhotos={setGtmLargeGfxPhotos}
            compareOverrides={gtmLargeGfxCompareOverrides}
            setCompareOverrides={setGtmLargeGfxCompareOverrides}
            isAdmin={isAdmin}
            submissions={gtmSubmissions.largegfx || {}}
            setSubmissions={setSectionSubmissions("largegfx")}
          />
        : <div style={{background:"#fff", borderRadius:14, padding:48,
            boxShadow:"0 2px 16px rgba(0,0,0,0.07)", textAlign:"center", color:"#aaa", fontSize:15}}>
            준비 중입니다.
          </div>
      }
    </div>
  );
}

function GTMCollectSection({ data, setData, isAdmin, submissions, setSubmissions, sectionLabel, variant="store", newStoreList, setNewStoreList }) {
  const fileRef = useRef(null);
  const reuploadFileRef = useRef(null);
  const newStoreFileRef = useRef(null);
  const [uploadStatus, setUploadStatus] = useState(null);
  const [uploadRegion, setUploadRegion] = useState("");
  const [reuploadStatus, setReuploadStatus] = useState(null);
  const [newStoreUploadStatus, setNewStoreUploadStatus] = useState(null);
  const [activeSection, setActiveSection] = useState("list"); // "list" | "missing"
  const [hqFilter, setHqFilter] = useState("전체");
  const [showAddStore, setShowAddStore] = useState(false);
  const [addStoreForm, setAddStoreForm] = useState({ 구분:"", 본부:"", 대리점코드:"", 대리점명:"", 매장코드:"", 매장명:"", 주소:"", 단면양면:"", 슬롯:"", 신청수량:"" });
  const isStore = variant === "store";

  // 예전에 업로드되어 id 필드가 없는 신설 매장 리스트(레거시 데이터)를 한 번만 채워준다.
  // id가 없으면 여러 행이 undefined===undefined로 매칭되어 체크박스 등이 한꺼번에 바뀌는 문제가 생긴다.
  useEffect(() => {
    if (isStore && setNewStoreList && newStoreList && newStoreList.some(r => r.id == null)) {
      setNewStoreList(prev => prev.map((r,i) => r.id == null ? { ...r, id: Date.now()+i } : r));
    }
  }, [isStore, newStoreList, setNewStoreList]);

  // 데이터에 있는 매장코드 목록
  const wcCodeSet = new Set(data.map(r => String(r.매장코드||"").trim()));
  // 신설 매장 리스트 중 아직 이 데이터에 반영되지 않은 매장
  const missingStores = isStore
    ? (newStoreList || [])
        .filter(r => r.매장코드 && !wcCodeSet.has(String(r.매장코드).trim()))
        .map(r => ({ ...r, 본부: normalizeGtmRegion(r.본부) }))
    : [];

  // 필터 옵션: store 변형은 본부 단위, hq 변형은 "구분 - 본부" 단위(단, 유통사업부/PS&M류는 전체 하나로 묶음)
  const filterOptions = isStore
    ? Array.from(new Set(data.map(r=>r.본부).filter(Boolean))).map(hq => ({
        key: hq, label: hq, predicate: row => row.본부 === hq,
      }))
    : (() => {
        const opts = [];
        const seen = new Set();
        data.forEach(r => {
          const gu = String(r.구분||"").trim();
          if (!gu) return;
          const isDistrib = gu.includes("유통") || gu.toUpperCase().includes("PS&M");
          if (isDistrib) {
            if (!seen.has(gu)) {
              seen.add(gu);
              opts.push({ key: gu, label: `${gu} 전체`, predicate: row => String(row.구분||"").trim()===gu });
            }
          } else {
            const key = `${gu}|${r.본부}`;
            if (!seen.has(key)) {
              seen.add(key);
              opts.push({ key, label: `${gu} - ${r.본부}`, predicate: row => String(row.구분||"").trim()===gu && row.본부===r.본부 });
            }
          }
        });
        return opts;
      })();
  const selectedOption = filterOptions.find(o => o.key === hqFilter);
  const filteredData = hqFilter==="전체" ? data : data.filter(selectedOption ? selectedOption.predicate : ()=>true);
  const filteredMissingStores = hqFilter==="전체" ? missingStores : missingStores.filter(selectedOption ? selectedOption.predicate : ()=>true);
  // 필터 드롭다운의 개수 표시는 현재 보고 있는 섹션(전체 목록 vs 누락 매장) 기준이어야 한다.
  const filterCountSource = (isStore && activeSection === "missing") ? missingStores : data;

  const handleUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (isStore && !isAdmin && !uploadRegion) {
      setUploadStatus({type:"error", msg:"먼저 어느 본부 소속 데이터인지 선택해주세요."});
      e.target.value="";
      return;
    }
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        let parsed;
        let fileHqRegions = null;
        if (isStore) {
          // 헤더 행을 텍스트로 찾아서 열 위치를 인식 (공백 유무 등 표기가 조금씩 달라도 안전하게 대응)
          const norm = (s) => String(s||"").replace(/\s+/g,"");
          const headerRowIdx = rows.findIndex(r => Array.isArray(r) && r.some(c => norm(c)==="매장코드"));
          const headerRow = headerRowIdx >= 0 ? rows[headerRowIdx] : [];
          const colIdx = (name) => headerRow.findIndex(c => norm(c)===norm(name));
          const guCol        = colIdx("구분");
          const hqCol        = colIdx("본부");
          const teamCol      = colIdx("마케팅팀");
          const agentCodeCol = colIdx("대리점코드");
          const agentNameCol = colIdx("대리점명");
          const storeCodeCol = colIdx("매장코드");
          const storeNameCol = colIdx("매장명");
          const addressCol   = colIdx("주소");
          const sideCol      = colIdx("단면형/양면형") >= 0 ? colIdx("단면형/양면형") : colIdx("단면/양면");
          const slotCol      = colIdx("도광판슬롯");
          // 수량 열: 몇 가지 흔한 이름으로 먼저 찾고, 못 찾으면 헤더의 마지막 열을 수량으로 간주(파일마다 열 개수가 달라도 대응)
          let qtyCol = ["수량","기존수량","이전수량","배송수량","신청수량","물량"].map(colIdx).find(i=>i>=0);
          if (qtyCol === undefined) qtyCol = -1;
          if (qtyCol < 0 && headerRowIdx >= 0) qtyCol = headerRow.length - 1;
          // 헤더를 못 찾으면 기존 고정 위치(2행 헤더, 11열)로 대체
          const startRow = headerRowIdx >= 0 ? headerRowIdx+1 : 2;
          const get = (r, idx, fallbackIdx) => (idx>=0 ? r[idx] : r[fallbackIdx]);
          const dataRows = rows.slice(startRow).filter(r => (storeCodeCol>=0 ? r[storeCodeCol] : r[5]) != null);
          // 파일 자체에 적힌 본부 값(검증용). uploadRegion으로 강제 덮어쓰기 전의 원본 값을 별도로 남겨둔다.
          fileHqRegions = new Set(dataRows.map(r => normalizeGtmRegion(get(r, hqCol, 1))).filter(Boolean));
          parsed = dataRows.map((r,i) => {
            const qty = get(r, qtyCol, 10);
            return {
              id: Date.now()+i,
              구분: get(r, guCol, 0),
              // 관리자는 파일 자체의 본부 값을 그대로 사용(여러 본부 한번에 업로드), 일반 계정은 선택한 본부로 강제 지정
              본부: isAdmin ? get(r, hqCol, 1) : uploadRegion,
              마케팅팀: get(r, teamCol, 2),
              대리점코드: String(get(r, agentCodeCol, 3)||"").trim(),
              대리점명: get(r, agentNameCol, 4),
              매장코드: String(get(r, storeCodeCol, 5)||"").trim(),
              매장명: get(r, storeNameCol, 6),
              주소: get(r, addressCol, 7),
              단면양면: get(r, sideCol, 8),
              슬롯: get(r, slotCol, 9),
              이전수량: qty,
              신규수량: (qty!==undefined && qty!==null && qty!=="") ? qty : "", // 이전 수량을 기본값으로 채움 (수정 가능)
            };
          });
        } else {
          // 헤더: Row1 = [구분,본부,이전 취합 수량,희망 수량] — 구분은 그룹 첫 행에만 존재(병합 셀), 총계 행은 본부가 비어있어 제외
          const dataRows = rows.slice(1).filter(r=>r[1]!=null);
          let lastGroup = "";
          parsed = dataRows.map((r,i) => {
            if (r[0]!=null && String(r[0]).trim()!=="") lastGroup = String(r[0]).trim();
            const 이전수량 = r[2];
            return {
              id: i+1,
              구분: lastGroup,
              본부: r[1],
              이전수량,
              신규수량: (이전수량!==undefined && 이전수량!==null && 이전수량!=="") ? 이전수량 : "",
            };
          });
        }
        if (isStore) {
          if (isAdmin) {
            // 관리자는 여러 본부가 섞인 전체 파일을 한 번에 올릴 수 있으므로 전체 교체
            setData(parsed);
            // 데이터 전체가 바뀌었으므로 기존 제출 현황(전체 목록/누락 매장)도 모두 초기화
            setSubmissions({});
          } else {
            // 자기 소속이 아닌 본부의 파일을 올리면 튕기기: 파일 자체에 적힌 본부 값이
            // 선택한 uploadRegion과 다른 행이 하나라도 있으면 업로드를 거부한다.
            const mismatched = [...(fileHqRegions||[])].filter(r => r !== uploadRegion);
            if (mismatched.length > 0) {
              setUploadStatus({type:"error", msg:`선택한 소속(${uploadRegion})과 다른 본부(${mismatched.join(", ")})의 데이터가 파일에 포함되어 있어 업로드를 거부했습니다. 본인 소속 파일만 업로드해주세요.`});
              return;
            }
            // 일반 계정은 선택한 본부 소속 데이터만 교체하고 다른 본부 데이터는 그대로 유지
            setData(prev => [...prev.filter(r=>r.본부 !== uploadRegion), ...parsed]);
            // 그 본부의 데이터가 바뀌었으므로 그 본부의 제출 현황만 초기화
            setSubmissions(prev => {
              const list = { ...(prev.list||{}) }; delete list[uploadRegion];
              const missing = { ...(prev.missing||{}) }; delete missing[uploadRegion];
              return { list, missing };
            });
          }
          setUploadStatus({type:"success", msg:`${isAdmin ? "전체" : uploadRegion} ${parsed.length}개 항목 로드 완료`});
          setUploadRegion(""); // 다음 업로드 때 다시 본부를 선택하도록 초기화
        } else {
          setData(parsed);
          setUploadStatus({type:"success", msg:`${parsed.length}개 항목 로드 완료`});
        }
      } catch(err) {
        setUploadStatus({type:"error", msg:"오류: "+err.message});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  const handleNewStoreUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        // 헤더: [매장코드,마케팅본부,마케팅팀,대리점코드,대리점명,매장명] — 열 이름으로 자동 인식
        const header = (rows[0]||[]).map(h=>String(h||"").trim());
        const colIdx = (name) => { const i = header.indexOf(name); return i>=0 ? i : -1; };
        const storeCodeCol = colIdx("매장코드");
        const mktgCol      = colIdx("마케팅본부");
        const teamCol      = colIdx("마케팅팀");
        const agentCodeCol = colIdx("대리점코드");
        const agentNameCol = colIdx("대리점명");
        const storeNameCol = colIdx("매장명");
        if (storeCodeCol < 0) {
          setNewStoreUploadStatus({type:"error", msg:"열 이름을 찾을 수 없습니다. '매장코드' 열이 있는지 확인해주세요."});
          return;
        }
        const dataRows = rows.slice(1).filter(r=>r[storeCodeCol]);
        const parsed = dataRows.map((r,i) => {
          const mktgHQ = String(mktgCol>=0 ? (r[mktgCol]||"") : "");
          const team   = String(teamCol>=0 ? (r[teamCol]||"") : "");
          return {
            id: Date.now()+i, // 매장코드가 중복될 수 있어 행 식별은 별도 id로 (체크박스 등 개별 수정용)
            매장코드: String(r[storeCodeCol]||"").trim(),
            매장명: String(storeNameCol>=0 ? (r[storeNameCol]||"") : ""),
            대리점코드: String(agentCodeCol>=0 ? (r[agentCodeCol]||"") : "").trim(),
            대리점명: String(agentNameCol>=0 ? (r[agentNameCol]||"") : ""),
            구분: mktgHQ.includes('유통사업부') ? 'PS&M' : '지역본부',
            본부: guessGtmRegion(mktgHQ, team) || "-",
            확인여부: false,
            비고: "",
          };
        });
        setNewStoreList(parsed);
        setNewStoreUploadStatus({type:"success", msg:`${parsed.length}개 신설 매장 로드 완료`});
      } catch(err) {
        setNewStoreUploadStatus({type:"error", msg:"오류: "+err.message});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  const updateRow = (id, field, val) => {
    setData(prev => prev.map(r => r.id===id ? {...r, [field]:val} : r));
  };
  const deleteRow = (id) => {
    setData(prev => prev.filter(r => r.id!==id));
  };

  const updateNewStore = (id, field, val) => {
    setNewStoreList(prev => prev.map(r => r.id===id ? {...r, [field]:val} : r));
  };

  // 전체 목록에 매장 한 건 추가 (누락 매장에서 "추가" 또는 "매장 추가" 버튼으로 진입)
  const addStoreRow = (fields) => {
    setData(prev => {
      const nextId = prev.reduce((m,r)=>Math.max(m, r.id||0), 0) + 1;
      return [...prev, {
        id: nextId,
        구분: fields.구분||"", 본부: fields.본부||"", 마케팅팀: "",
        대리점코드: fields.대리점코드||"", 대리점명: fields.대리점명||"",
        매장코드: fields.매장코드||"", 매장명: fields.매장명||"", 주소: fields.주소||"",
        단면양면: fields.단면양면||"", 슬롯: fields.슬롯||"",
        이전수량: 0, 신규수량: fields.신청수량 ?? "",
      }];
    });
  };

  const downloadResult = () => {
    const header = isStore
      ? ["구분","본부","마케팅팀","대리점코드","대리점명","매장코드","매장명","주소","단면형/양면형","도광판슬롯","이전수량","신규수량"]
      : ["구분","본부","이전수량","신규수량"];
    const body = isStore
      ? data.map(r=>[r.구분,r.본부,r.마케팅팀,r.대리점코드,r.대리점명,r.매장코드,r.매장명,r.주소,r.단면양면,r.슬롯,r.이전수량,r.신규수량||0])
      : data.map(r=>[r.구분,r.본부,r.이전수량,r.신규수량||0]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
    ws["!cols"] = (isStore ? [8,8,12,10,10,12,14,24,10,8,8,8] : [10,8,10,10]).map(w=>({wch:w}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sectionLabel);
    XLSX.writeFile(wb, `${sectionLabel}_수량취합.xlsx`);
  };

  // 다운로드한 수량취합 엑셀을 지역본부에서 작업 후 재업로드 → 매장코드 기준으로 덮어쓰기(신규 매장코드는 추가)
  const handleReuploadResult = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (isStore && !isAdmin && !uploadRegion) {
      setReuploadStatus({type:"error", msg:"먼저 어느 본부 소속 데이터인지 선택해주세요."});
      e.target.value="";
      return;
    }
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        const dataRows = rows.slice(1).filter(r=>r[5]);
        const updatesByCode = new Map();
        dataRows.forEach(r => {
          const code = String(r[5]||"").trim();
          if (!code) return;
          updatesByCode.set(code, {
            구분:r[0], 본부:r[1], 마케팅팀:r[2],
            대리점코드:String(r[3]||"").trim(), 대리점명:r[4],
            매장코드:code, 매장명:r[6], 주소:r[7],
            단면양면:r[8], 슬롯:r[9], 이전수량:r[10],
            신규수량: (r[11]!==undefined && r[11]!==null && r[11]!=="") ? r[11] : r[10], // 수량 열이 비어있으면 이전수량을 그대로 사용
          });
        });
        // 업로드 파일에 포함된 본부(들)는 기존 pool을 통째로 비우고 이 파일 내용으로만 다시 채움
        // (매장코드 매칭 갱신이 아니라, 그 본부 소속 전체를 새 목록으로 교체)
        const regionsInFile = new Set([...updatesByCode.values()].map(f=>f.본부));
        if (isStore && !isAdmin) {
          // 자기 소속이 아닌 본부의 파일을 재업로드하면 튕기기
          const mismatched = [...regionsInFile].filter(r => normalizeGtmRegion(r) !== uploadRegion);
          if (mismatched.length > 0) {
            setReuploadStatus({type:"error", msg:`선택한 소속(${uploadRegion})과 다른 본부(${mismatched.join(", ")})의 데이터가 파일에 포함되어 있어 업로드를 거부했습니다. 본인 소속 파일만 업로드해주세요.`});
            return;
          }
        }
        setData(prev => {
          const kept = prev.filter(r => !regionsInFile.has(r.본부));
          const newRows = [...updatesByCode.values()].map((fields,i) => ({ id: Date.now()+i, ...fields }));
          return [...kept, ...newRows];
        });
        setReuploadStatus({type:"success", msg:`${[...regionsInFile].join(", ")} ${updatesByCode.size}개 매장으로 교체 완료`});
      } catch(err) {
        setReuploadStatus({type:"error", msg:"오류: "+err.message});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  // 전체 목록 제출과 누락 매장 제출을 별개로 추적: submissions = { list: {...}, missing: {...} }
  const submissionScope = (isStore && activeSection === "missing") ? "missing" : "list";
  const scopeLabel = submissionScope === "missing" ? "누락 매장" : "전체 목록";
  const submission = hqFilter!=="전체" ? (submissions[submissionScope]||{})[hqFilter] : null;
  const isSubmitted = !!submission?.submitted;
  const inputsLocked = isSubmitted && !isAdmin;

  const handleSubmit = () => {
    if (hqFilter === "전체") return;
    setSubmissions(prev => ({ ...prev, [submissionScope]: { ...(prev[submissionScope]||{}), [hqFilter]: { submitted:true, submittedAt:new Date().toISOString() } } }));
  };
  const handleCancelSubmit = () => {
    if (hqFilter === "전체") return;
    setSubmissions(prev => ({ ...prev, [submissionScope]: { ...(prev[submissionScope]||{}), [hqFilter]: { submitted:false, submittedAt:null } } }));
  };

  return (
    <div style={{display:"flex", flexDirection:"column", gap:16}}>
      {/* 관리자 전용: 사업부별 제출 현황 (맨 위) */}
      {isAdmin && filterOptions.length > 0 && (
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:20}}>
          <div style={{fontSize:14, fontWeight:800, marginBottom:12, color:"#333"}}>✅ 사업부별 제출 현황(관리자 전용)</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse", width:"100%", minWidth: isStore?680:480, fontSize:12.5}}>
              <thead>
                <tr style={{background:"#f0f4f8"}}>
                  <th style={tH}>지역담당</th>
                  <th style={tH}>{isStore ? "전체 목록 제출" : "제출여부"}</th>
                  <th style={tH}>{isStore ? "전체 목록 제출 일시" : "제출 일시"}</th>
                  {isStore && <th style={tH}>누락 매장 제출</th>}
                  {isStore && <th style={tH}>누락 매장 제출 일시</th>}
                </tr>
              </thead>
              <tbody>
                {filterOptions.map((opt,i)=>{
                  const listSub = (submissions.list||{})[opt.key];
                  const missingSub = (submissions.missing||{})[opt.key];
                  const statusCell = (sub) => sub?.submitted
                    ? <span style={{color:"#2e8b57", fontWeight:700}}>✅ 제출완료</span>
                    : <span style={{color:"#aaa"}}>⏳ 미제출</span>;
                  return (
                    <tr key={opt.key} style={{background:i%2===0?"#fafafa":"#fff"}}>
                      <td style={{...tD, textAlign:"center", fontWeight:700}}>{opt.label}</td>
                      <td style={{...tD, textAlign:"center"}}>{statusCell(listSub)}</td>
                      <td style={{...tD, textAlign:"center", color:"#666"}}>
                        {listSub?.submittedAt ? new Date(listSub.submittedAt).toLocaleString('ko-KR') : "-"}
                      </td>
                      {isStore && <td style={{...tD, textAlign:"center"}}>{statusCell(missingSub)}</td>}
                      {isStore && (
                        <td style={{...tD, textAlign:"center", color:"#666"}}>
                          {missingSub?.submittedAt ? new Date(missingSub.submittedAt).toLocaleString('ko-KR') : "-"}
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 업로드 영역 */}
      <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:24}}>
        <div style={{fontSize:15, fontWeight:800, marginBottom:10}}>📂 {sectionLabel} 데이터 업로드</div>
        <div style={{fontSize:12, color:"#888", marginBottom:12}}>
          기존 {sectionLabel} 배송 리스트 엑셀 파일을 업로드하면 이전 수량이 자동으로 로드됩니다.
          {isStore && (isAdmin
            ? " 관리자는 본부 선택 없이 전체 데이터를 한 번에 업로드합니다."
            : " 최초 업로드는 관리자만 가능하며, 일반 계정은 아래 \"취합 결과 업로드\"로 선택한 본부 데이터만 제출합니다.")}
        </div>
        <div style={{display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
          {isStore && !isAdmin && (
            <select value={uploadRegion} onChange={e=>setUploadRegion(e.target.value)}
              style={{padding:"7px 12px", borderRadius:8, border:"1px solid #ddd",
                fontSize:12.5, fontWeight:600, color: uploadRegion?"#444":"#c00", background:"#fff", cursor:"pointer"}}>
              <option value="">어느 본부 소속인가요?</option>
              {WIDECOLOR_UPLOAD_SCOPES.map(r=>(<option key={r} value={r}>{r}</option>))}
            </select>
          )}
          {(!isStore || isAdmin) && (
            <>
              <button style={{...settingsBtn("#1d6fa4"), padding:"7px 18px", fontSize:12}}
                onClick={()=>fileRef.current?.click()}>📂 파일 선택</button>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleUpload}/>
            </>
          )}
          {data.length > 0 && (
            <button style={{...settingsBtn("#2e8b57"), padding:"7px 18px", fontSize:12}}
              onClick={downloadResult}>⬇ 수량취합 엑셀 다운로드</button>
          )}
          {isStore && data.length > 0 && (
            <>
              <button style={{...settingsBtn("#e8420a"), padding:"7px 18px", fontSize:12}}
                onClick={()=>reuploadFileRef.current?.click()}>📤 취합 결과 업로드</button>
              <input ref={reuploadFileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleReuploadResult}/>
            </>
          )}
          {uploadStatus && (
            <span style={{fontSize:12, padding:"5px 12px", borderRadius:8,
              background:uploadStatus.type==="success"?"#e8f8ee":"#fff0f0",
              color:uploadStatus.type==="success"?"#2e8b57":"#c00"}}>
              {uploadStatus.type==="success"?"✅ ":"❌ "}{uploadStatus.msg}
            </span>
          )}
          {reuploadStatus && (
            <span style={{fontSize:12, padding:"5px 12px", borderRadius:8,
              background:reuploadStatus.type==="success"?"#e8f8ee":"#fff0f0",
              color:reuploadStatus.type==="success"?"#2e8b57":"#c00"}}>
              {reuploadStatus.type==="success"?"✅ ":"❌ "}{reuploadStatus.msg}
            </span>
          )}
        </div>
        {isStore && data.length > 0 && (
          <div style={{fontSize:11, color:"#aaa", marginTop:8}}>
            "⬇ 수량취합 엑셀 다운로드"로 받은 파일을 그대로 수정해서 "📤 취합 결과 업로드"로 다시 올리면, 매장코드를 기준으로 전체 목록에 덮어써집니다 (없는 매장코드는 새로 추가됩니다).
          </div>
        )}
      </div>

      {data.length > 0 && (
        <>
          {/* 섹션 탭 */}
          <div style={{display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
            <button onClick={()=>setActiveSection("list")} style={{
              padding:"8px 20px", borderRadius:50, border:"none", cursor:"pointer", fontSize:13, fontWeight:700,
              background:activeSection==="list"?"#1d6fa4":"#e0e6ed",
              color:activeSection==="list"?"#fff":"#444",
            }}>📋 전체 목록 ({data.length})</button>
            {isStore && (
              <button onClick={()=>setActiveSection("missing")} style={{
                padding:"8px 20px", borderRadius:50, border:"none", cursor:"pointer", fontSize:13, fontWeight:700,
                background:activeSection==="missing"?"#e85d26":"#ffeee8",
                color:activeSection==="missing"?"#fff":"#e85d26",
              }}>
                ⚠️ 누락 매장
                {missingStores.length > 0 &&
                  <span style={{marginLeft:6, background:"#c00", color:"#fff", borderRadius:10,
                    padding:"0 6px", fontSize:11}}>{missingStores.length}</span>}
              </button>
            )}
            {filterOptions.length > 0 && (
              <select value={hqFilter} onChange={e=>setHqFilter(e.target.value)}
                style={{marginLeft:"auto", padding:"7px 14px", borderRadius:8, border:"1px solid #ddd",
                  fontSize:12.5, fontWeight:600, color:"#444", background:"#fff", cursor:"pointer"}}>
                <option value="전체">전체 ({filterCountSource.length})</option>
                {filterOptions.map(opt=>(
                  <option key={opt.key} value={opt.key}>{opt.label} ({filterCountSource.filter(opt.predicate).length})</option>
                ))}
              </select>
            )}
            {hqFilter !== "전체" && (
              <>
                {isSubmitted && (
                  <span style={{fontSize:12, fontWeight:700, color:"#2e8b57", background:"#e8f8ee",
                    padding:"6px 12px", borderRadius:8}}>
                    ✅ {scopeLabel} 제출완료 ({new Date(submission.submittedAt).toLocaleString('ko-KR')})
                  </span>
                )}
                <button
                  onClick={handleSubmit}
                  disabled={isSubmitted}
                  style={{...settingsBtn(isSubmitted?"#bbb":"#e8420a"), padding:"7px 18px", fontSize:12.5,
                    cursor:isSubmitted?"default":"pointer"}}>
                  {isSubmitted ? `${scopeLabel} 제출완료` : `${scopeLabel} 제출하기`}
                </button>
                {isAdmin && isSubmitted && (
                  <button onClick={handleCancelSubmit}
                    style={{...settingsBtn("#aaa"), padding:"7px 14px", fontSize:12}}>제출 취소</button>
                )}
                {isStore && activeSection === "list" && (
                  <button onClick={()=>{
                    setAddStoreForm({ 구분:"", 본부: hqFilter!=="전체" ? hqFilter : "", 대리점코드:"", 대리점명:"", 매장코드:"", 매장명:"", 주소:"", 단면양면:"", 슬롯:"", 신청수량:"" });
                    setShowAddStore(true);
                  }}
                    style={{...settingsBtn("#1d6fa4"), padding:"7px 14px", fontSize:12}}>+ 매장 추가</button>
                )}
              </>
            )}
          </div>

          {activeSection === "list" && (
            <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", overflow:"hidden"}}>
              <div style={{overflowX:"auto"}}>
                <table style={{borderCollapse:"collapse", width:"100%", minWidth: isStore?900:500, fontSize:12}}>
                  <thead>
                    <tr style={{background:"#f0f4f8"}}>
                      {isStore ? (
                        <>
                          <th style={tH}>#</th>
                          <th style={tH}>본부</th>
                          <th style={tH}>매장코드</th>
                          <th style={{...tH, minWidth:120}}>매장명</th>
                          <th style={tH}>단면/양면</th>
                          <th style={tH}>슬롯</th>
                        </>
                      ) : (
                        <>
                          <th style={tH}>구분</th>
                          <th style={tH}>본부</th>
                        </>
                      )}
                      <th style={{...tH, background:"#e8f0fb", color:"#1d6fa4"}}>이전수량</th>
                      <th style={{...tH, background:"#e8f8ee", color:"#2e8b57", minWidth:100}}>수량</th>
                      <th style={{...tH, width:50}}>삭제</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row,i)=>(
                      <tr key={row.id} style={{background:i%2===0?"#fafafa":"#fff"}}>
                        {isStore ? (
                          <>
                            <td style={{...tD, textAlign:"center", color:"#aaa"}}>{i+1}</td>
                            <td style={{...tD, textAlign:"center"}}>{row.본부}</td>
                            <td style={{...tD, textAlign:"center", fontSize:11, color:"#666"}}>{row.매장코드}</td>
                            <td style={{...tD, fontWeight:600}}>{row.매장명}</td>
                            <td style={{...tD, textAlign:"center"}}>
                              <input
                                type="text"
                                disabled={inputsLocked}
                                value={row.단면양면||""}
                                onChange={e=>updateRow(row.id,"단면양면",e.target.value)}
                                style={{...numInput, width:56, textAlign:"center", opacity: inputsLocked?0.6:1}}
                              />
                            </td>
                            <td style={{...tD, textAlign:"center"}}>
                              <input
                                type="text"
                                disabled={inputsLocked}
                                value={row.슬롯||""}
                                onChange={e=>updateRow(row.id,"슬롯",e.target.value)}
                                style={{...numInput, width:56, textAlign:"center", opacity: inputsLocked?0.6:1}}
                              />
                            </td>
                          </>
                        ) : (
                          <>
                            <td style={{...tD, textAlign:"center", fontWeight:700, color:"#555"}}>{row.구분}</td>
                            <td style={{...tD, textAlign:"center"}}>{row.본부}</td>
                          </>
                        )}
                        <td style={{...tD, textAlign:"center", background:"#f0f7ff"}}>
                          <input
                            type="number" min="0"
                            disabled={inputsLocked}
                            value={row.이전수량 ?? ""}
                            onChange={e=>updateRow(row.id,"이전수량",e.target.value)}
                            style={{...numInput, width:60, textAlign:"center", fontWeight:700, color:"#1d6fa4", opacity: inputsLocked?0.6:1}}
                          />
                        </td>
                        <td style={{...tD, textAlign:"center", background:"#f0fff4"}}>
                          <input
                            type="number" min="0"
                            disabled={inputsLocked}
                            style={{...numInput, width:70, fontWeight:700,
                              color: row.신규수량!==""?"#2e8b57":"#aaa",
                              borderColor: row.신규수량!==""?"#2e8b57":"#ddd",
                              opacity: inputsLocked?0.6:1, cursor: inputsLocked?"not-allowed":"text"}}
                            value={row.신규수량}
                            placeholder={String(row.이전수량||0)}
                            onChange={e=>updateRow(row.id,"신규수량",e.target.value)}
                          />
                        </td>
                        <td style={{...tD, textAlign:"center"}}>
                          <button onClick={()=>deleteRow(row.id)} disabled={inputsLocked}
                            style={{width:24, height:24, borderRadius:"50%", border:"none",
                              background: inputsLocked?"#ccc":"#c00",
                              color:"#fff", fontSize:13, fontWeight:800, cursor: inputsLocked?"not-allowed":"pointer", lineHeight:1}}>×</button>
                        </td>
                      </tr>
                    ))}
                    {/* 합계 행 */}
                    <tr style={{background:"#e8f0fb", fontWeight:700}}>
                      <td colSpan={isStore?6:2} style={{...tD, textAlign:"right", color:"#555"}}>합 계</td>
                      <td style={{...tD, textAlign:"center", color:"#1d6fa4"}}>
                        {filteredData.reduce((s,r)=>s+(Number(r.이전수량)||0),0)}
                      </td>
                      <td style={{...tD, textAlign:"center", color:"#2e8b57"}}>
                        {filteredData.reduce((s,r)=>s+(Number(r.신규수량)||0),0)}
                      </td>
                      <td style={tD}></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {isStore && activeSection === "missing" && (
            <div style={{display:"flex", flexDirection:"column", gap:12}}>
              {/* 신설 매장 리스트 업로드 */}
              <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:20}}>
                <div style={{fontSize:13.5, fontWeight:800, marginBottom:8}}>🆕 신설 매장 리스트 업로드</div>
                <div style={{fontSize:11.5, color:"#888", marginBottom:10}}>
                  매장코드, 마케팅본부, 마케팅팀, 대리점코드, 대리점명, 매장명 열이 포함된 엑셀을 업로드하면, 이 리스트를 기준으로 {sectionLabel}에 누락된 매장을 확인합니다. (설정 메뉴의 전체 매장 리스트가 아닌, 별도의 신설 매장 리스트입니다.)
                </div>
                <div style={{display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
                  <button style={{...settingsBtn("#e8420a"), padding:"7px 18px", fontSize:12}}
                    onClick={()=>newStoreFileRef.current?.click()}>📂 신설 매장 리스트 업로드</button>
                  <input ref={newStoreFileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleNewStoreUpload}/>
                  <span style={{fontSize:12, color:"#666"}}>현재 {newStoreList?.length||0}개 등록됨</span>
                  {newStoreUploadStatus && (
                    <span style={{fontSize:12, padding:"5px 12px", borderRadius:8,
                      background:newStoreUploadStatus.type==="success"?"#e8f8ee":"#fff0f0",
                      color:newStoreUploadStatus.type==="success"?"#2e8b57":"#c00"}}>
                      {newStoreUploadStatus.type==="success"?"✅ ":"❌ "}{newStoreUploadStatus.msg}
                    </span>
                  )}
                </div>
              </div>

              <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", overflow:"hidden"}}>
                {(!newStoreList || newStoreList.length === 0) ? (
                  <div style={{padding:40, textAlign:"center", color:"#aaa", fontSize:14}}>
                    신설 매장 리스트를 먼저 업로드해주세요.
                  </div>
                ) : filteredMissingStores.length === 0 ? (
                  <div style={{padding:40, textAlign:"center", color:"#aaa", fontSize:14}}>
                    누락된 매장이 없습니다. ✅
                  </div>
                ) : (
                  <div style={{overflowX:"auto"}}>
                    <div style={{padding:"12px 16px", background:"#fff8f0", fontSize:12, color:"#e85d26", borderBottom:"1px solid #ffd"}}>
                      ⚠️ 신설 매장 리스트에는 있지만 {sectionLabel} 데이터에 아직 없는 매장입니다.
                    </div>
                    <table style={{borderCollapse:"collapse", width:"100%", minWidth:700, fontSize:12}}>
                      <thead>
                        <tr style={{background:"#fff8f0"}}>
                          <th style={tH}>#</th>
                          <th style={tH}>본부</th>
                          <th style={tH}>매장코드</th>
                          <th style={{...tH, minWidth:120}}>매장명</th>
                          <th style={{...tH, width:100}}>확인</th>
                          <th style={{...tH, minWidth:120}}>비고</th>
                          <th style={{...tH, width:60}}>추가</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredMissingStores.map((s,i)=>(
                          <tr key={s.id ?? s.매장코드+i} style={{background:i%2===0?"#fafafa":"#fff"}}>
                            <td style={{...tD, textAlign:"center", color:"#aaa"}}>{i+1}</td>
                            <td style={{...tD, textAlign:"center"}}>{s.본부||"-"}</td>
                            <td style={{...tD, fontSize:11, color:"#666"}}>{s.매장코드}</td>
                            <td style={{...tD, fontWeight:600}}>{s.매장명||s.매장코드}</td>
                            <td style={{...tD, textAlign:"center"}}>
                              <label style={{display:"flex", alignItems:"center", justifyContent:"center", gap:6, cursor: inputsLocked?"not-allowed":"pointer"}}>
                                <input type="checkbox"
                                  disabled={inputsLocked}
                                  checked={!!s.확인여부}
                                  onChange={e=>updateNewStore(s.id,"확인여부",e.target.checked)}
                                  style={{width:16,height:16,accentColor:"#e85d26"}} />
                                <span style={{fontSize:11,color:s.확인여부?"#e85d26":"#aaa"}}>
                                  {s.확인여부?"확인":"미확인"}
                                </span>
                              </label>
                            </td>
                            <td style={tD}>
                              <input style={{...numInput, width:"100%", textAlign:"left", opacity: inputsLocked?0.6:1}}
                                disabled={inputsLocked}
                                value={s.비고||""} placeholder="비고 입력"
                                onChange={e=>updateNewStore(s.id,"비고",e.target.value)} />
                            </td>
                            <td style={{...tD, textAlign:"center"}}>
                              <button
                                title="매장리스트에 추가"
                                disabled={inputsLocked}
                                onClick={()=>addStoreRow({ 구분:s.구분, 본부:s.본부, 대리점코드:s.대리점코드, 대리점명:s.대리점명, 매장코드:s.매장코드, 매장명:s.매장명 })}
                                style={{width:26, height:26, borderRadius:"50%", border:"none",
                                  background: inputsLocked?"#bbb":"#2e8b57",
                                  color:"#fff", fontSize:15, fontWeight:800, lineHeight:1, cursor: inputsLocked?"not-allowed":"pointer",
                                  display:"inline-flex", alignItems:"center", justifyContent:"center", padding:0}}>
                                +
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          )}
        </>
      )}

      {showAddStore && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:420, maxWidth:480, alignItems:"stretch", gap:10}}>
            <div style={{fontSize:16, fontWeight:800, marginBottom:4}}>+ 매장 추가</div>
            {[
              ["본부","본부"], ["대리점코드","대리점코드"], ["대리점명","대리점명"],
              ["매장코드","매장코드"], ["매장명","매장명"], ["주소","주소"], ["단면양면","단면/양면"], ["슬롯","도광판슬롯"], ["신청수량","신청 수량"],
            ].map(([field,label])=>(
              <div key={field} style={{display:"flex", flexDirection:"column", gap:4}}>
                <label style={{fontSize:11.5, color:"#888", fontWeight:600}}>{label}</label>
                <input
                  type={field==="신청수량"?"number":"text"}
                  value={addStoreForm[field]}
                  onChange={e=>setAddStoreForm(prev=>({...prev,[field]:e.target.value}))}
                  style={{...numInput, width:"100%", textAlign:"left"}}
                />
              </div>
            ))}
            <div style={{display:"flex", gap:10, marginTop:8}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>setShowAddStore(false)}>취소</button>
              <button style={{...styles2.popupBtn, flex:1}}
                onClick={()=>{
                  if (!addStoreForm.매장코드) return;
                  addStoreRow(addStoreForm);
                  setShowAddStore(false);
                }}>추가</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const DRAFT_TAB = "__draft__";

// 라지그래픽 매장 엑셀 파싱 공용 함수 — "Summary" 시트를 건너뛰고 '매장코드' 열이 있는 시트를 찾아 파싱
// 헤더 앵커: '매장코드'가 있으면 우선 쓰고, 없으면 '매장명'만으로도 시트/헤더를 인식
function hasLargeGfxAnchor(rows) {
  return rows.some(r => Array.isArray(r) && r.some(c => { const s=String(c||"").trim(); return s==="매장코드" || s==="매장명"; }));
}

function parseLargeGfxWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, {type:"array"});
  const dataSheetName = wb.SheetNames.find(name => hasLargeGfxAnchor(XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1})))
    || wb.SheetNames[wb.SheetNames.length-1];
  const ws = wb.Sheets[dataSheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1});
  const headerRowIdx = rows.findIndex(r => Array.isArray(r) && r.some(c => { const s=String(c||"").trim(); return s==="매장코드" || s==="매장명"; }));
  if (headerRowIdx < 0) throw new Error("'매장코드' 또는 '매장명' 열을 찾을 수 없습니다. 매장 목록 시트가 맞는지 확인해주세요.");
  const headerRow = rows[headerRowIdx] || [];
  const colIdx = (name) => headerRow.findIndex(c => String(c||"").trim()===name);
  const colIdxIncludes = (needle) => headerRow.findIndex(c => String(c||"").trim().toUpperCase().includes(needle.toUpperCase()));
  const storeCodeCol = colIdx("매장코드"); // 없으면 -1 — 매장코드 없이도 업로드 진행
  const storeNameCol = colIdx("매장명");
  const addressCol   = colIdx("주소");
  const surveyCol    = colIdx("실사가능여부");
  const szoneCol     = colIdx("S.ZONE 여부");
  const noteCol      = colIdx("비고");
  const accpCol      = colIdxIncludes("ACCP");
  const mktgCol      = (storeCodeCol >= 0 ? storeCodeCol : storeNameCol) - 1; // 마케팅담당/본부 — 앵커 열 바로 앞(원본에 헤더명 없음)
  // 매장코드가 비어있는 행도 있을 수 있으므로, 매장명이나 매장코드 중 하나라도 있으면 포함
  const dataRows = rows.slice(headerRowIdx+1).filter(r=>(storeNameCol>=0 && r[storeNameCol]) || (storeCodeCol>=0 && r[storeCodeCol]));
  return dataRows.map((r,i) => ({
    id: Date.now()+i,
    순번: r[0],
    본부: guessGtmRegion(r[mktgCol], "") || String(r[mktgCol]||"").trim() || "-",
    매장코드: storeCodeCol>=0 ? String(r[storeCodeCol]||"").trim() : "",
    매장명: storeNameCol>=0 ? r[storeNameCol] : "",
    주소: addressCol>=0 ? r[addressCol] : "",
    실사가능: surveyCol>=0 ? r[surveyCol] : "",
    SZONE: szoneCol>=0 ? r[szoneCol] : "",
    비고: noteCol>=0 ? r[noteCol] : "",
    ACCP매장: accpCol>=0 ? r[accpCol] : "",
  }));
}

function GTMLargeGfx({ rounds, setRounds, draft, setDraft, photos, setPhotos, compareOverrides, setCompareOverrides, isAdmin, submissions, setSubmissions }) {
  const fileRef = useRef(null);
  const [activeRoundTab, setActiveRoundTab] = useState(DRAFT_TAB);
  const [uploadStatus, setUploadStatus] = useState(null);
  const [hqFilter, setHqFilter] = useState("전체");
  const [photoPopupCode, setPhotoPopupCode] = useState(null);
  const [showAddStore, setShowAddStore] = useState(false);
  const [addStoreForm, setAddStoreForm] = useState({ 본부:"", 매장코드:"", 매장명:"", 주소:"" });
  const [showFinalize, setShowFinalize] = useState(false);
  const [showBaseModal, setShowBaseModal] = useState(false);
  const [finalizeName, setFinalizeName] = useState("");
  const [showSubmitNote, setShowSubmitNote] = useState(false);
  const [submitNote, setSubmitNote] = useState("");

  const isDraft = activeRoundTab === DRAFT_TAB;
  const activeRound = isDraft ? null : rounds.find(r=>r.name===activeRoundTab);
  const viewData = isDraft ? draft : (activeRound?.data || []);
  const setViewData = isDraft ? setDraft : null; // 라운드 탭은 읽기 전용(재업로드로만 교체)

  // O/X 대조 컬럼: 기본값은 지금 보는 라운드 바로 이전(과거) 3개(최신순). 사용자가 3칸을 직접 골라 덮어쓸 수 있음(빈 칸 허용)
  const activeIdx = isDraft ? rounds.length : rounds.findIndex(r=>r.name===activeRoundTab);
  const autoCompareNames = activeIdx < 0 ? [] : rounds.slice(Math.max(0, activeIdx-3), activeIdx).reverse().map(r=>r.name);
  const compareSlots = (compareOverrides[activeRoundTab] || autoCompareNames).slice(0,3);
  while (compareSlots.length < 3) compareSlots.push(null);
  const compareRounds = compareSlots.filter(Boolean).map(n=>rounds.find(r=>r.name===n)).filter(Boolean);
  const compareSlotOptions = rounds.filter(r => r.name !== activeRoundTab);
  const updateCompareSlot = (slotIdx, name) => {
    const next = compareSlots.slice();
    next[slotIdx] = name || null;
    setCompareOverrides(prev => ({ ...prev, [activeRoundTab]: next }));
  };

  const hqOptions = Array.from(new Set(viewData.map(r=>r.본부).filter(Boolean)));
  // 아이폰 라운드(아이폰16/아이폰17 등)에서만 의미있는 추가 체크 항목 — 데이터가 있을 때만 컬럼 노출
  const showAccpCol = viewData.some(r => String(r.ACCP매장||"").trim());
  const filteredData = hqFilter==="전체" ? viewData : viewData.filter(r=>r.본부===hqFilter);

  const submission = (isDraft && hqFilter!=="전체") ? submissions[hqFilter] : null;
  const isSubmitted = !!submission?.submitted;
  const handleSubmitConfirm = () => {
    if (hqFilter === "전체") return;
    const note = submitNote.trim();
    setSubmissions(prev => {
      const prevSub = prev[hqFilter] || {};
      const history = [...(prevSub.history||[]), { note, submittedAt: new Date().toISOString() }];
      return { ...prev, [hqFilter]: { submitted:true, submittedAt:new Date().toISOString(), note, history } };
    });
    setSubmitNote("");
    setShowSubmitNote(false);
  };
  const handleCancelSubmit = () => {
    if (hqFilter === "전체") return;
    setSubmissions(prev => ({ ...prev, [hqFilter]: { ...(prev[hqFilter]||{}), submitted:false } }));
  };

  const openTab = (name) => {
    if (name === DRAFT_TAB && isAdmin) {
      // 이미 작업 중인 초안이 있어도 다시 눌러서 기반 리스트를 다시 고르고 덮어쓸 수 있도록 매번 확인 (관리자 전용)
      setActiveRoundTab(DRAFT_TAB);
      setHqFilter("전체");
      setShowBaseModal(true);
      return;
    }
    setHqFilter("전체");
    setActiveRoundTab(name);
  };

  const startDraftFrom = (roundName) => {
    if (roundName) {
      const base = rounds.find(r=>r.name===roundName);
      setDraft(base ? base.data.map((r,i)=>({ ...r, id: Date.now()+i })) : []);
    } else {
      setDraft([]);
    }
    setSubmissions({}); // 새 기반 리스트로 엎어쓰면 이전 제출 기록은 더 이상 유효하지 않으므로 초기화
    setShowBaseModal(false);
  };

  // 라운드 탭(비확정) 전용 업로드 — 해당 라운드 데이터를 통째로 교체
  const handleUpload = (e) => {
    const file = e.target.files[0];
    if (!file || !activeRound) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const parsed = parseLargeGfxWorkbook(ev.target.result);
        setRounds(prev => prev.map(r => r.name===activeRound.name ? { ...r, data: parsed } : r));
        setUploadStatus({type:"success", msg:`${parsed.length}개 매장 로드 완료`});
      } catch(err) {
        setUploadStatus({type:"error", msg: err.message});
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  const updateDraftRow = (id, field, val) => {
    setDraft(prev => prev.map(r => r.id===id ? {...r, [field]:val} : r));
  };
  const deleteDraftRow = (id) => {
    setDraft(prev => prev.filter(r => r.id!==id));
  };
  const addDraftRow = () => {
    if (!addStoreForm.매장코드 && !addStoreForm.매장명) return;
    setDraft(prev => [...prev, { id: Date.now(), 순번:"", ...addStoreForm, 비고:"" }]);
    setAddStoreForm({ 본부:"", 매장코드:"", 매장명:"", 주소:"" });
    setShowAddStore(false);
  };

  const handleFinalize = () => {
    const name = finalizeName.trim();
    if (!name) return;
    setRounds(prev => [...prev, {
      id: `${name}_${Date.now()}`, name, label: `${name} 선정매장`,
      data: draft.map((r,i)=>({...r, id: Date.now()+i})),
      locked: true, finalizedAt: new Date().toISOString(),
    }]);
    setFinalizeName("");
    setShowFinalize(false);
    setActiveRoundTab(name);
  };

  // 다른 라운드에 이 매장이 있었는지 O/X (매장코드가 있으면 매장코드로, 없으면 매장명으로 대조. 그 라운드가 비어있으면 "-")
  const CompareCell = ({ round, row }) => {
    if (!round.data || round.data.length === 0) return <span style={{color:"#ccc"}}>-</span>;
    const code = String(row.매장코드||"").trim();
    const name = String(row.매장명||"").trim();
    const has = code
      ? round.data.some(r=>String(r.매장코드||"").trim()===code)
      : (name ? round.data.some(r=>String(r.매장명||"").trim()===name) : false);
    return has
      ? <span style={{color:"#2e8b57", fontWeight:800}}>O</span>
      : <span style={{color:"#c00", fontWeight:800}}>X</span>;
  };

  const OXField = ({ row, field }) => {
    if (!isDraft) {
      const v = String(row[field]||"").trim();
      if (v==="O") return <span style={{color:"#2e8b57", fontWeight:800}}>O</span>;
      if (v==="X") return <span style={{color:"#c00", fontWeight:800}}>X</span>;
      return <span style={{color:"#ccc"}}>-</span>;
    }
    const v = String(row[field]||"").trim();
    return (
      <select
        value={v}
        onChange={e=>updateDraftRow(row.id, field, e.target.value)}
        style={{
          padding:"4px 6px", borderRadius:6, border:"1px solid #ddd", fontSize:12, fontWeight:800,
          background:"#fff", cursor:"pointer",
          color: v==="O" ? "#2e8b57" : v==="X" ? "#c00" : "#aaa",
        }}>
        <option value="">-</option>
        <option value="O">O</option>
        <option value="X">X</option>
      </select>
    );
  };

  const photoPopupRow = photoPopupCode ? viewData.find(r=>r.매장코드===photoPopupCode) : null;

  return (
    <div style={{display:"flex", flexDirection:"column", gap:16}}>
      {/* 라운드 탭: 신규 → 최신 → 과거 순 */}
      <div style={{display:"flex", gap:6, flexWrap:"wrap"}}>
        <button onClick={()=>openTab(DRAFT_TAB)} style={{
          padding:"8px 16px", borderRadius:50, border:"none", cursor:"pointer", fontSize:12.5, fontWeight:700,
          background: isDraft ? "#e8420a" : "#ffe8e0",
          color: isDraft ? "#fff" : "#e8420a",
        }}>
          + 신규 라지그래픽 선정
        </button>
        {[...rounds].reverse().map(r => (
          <button key={r.name} onClick={()=>openTab(r.name)} style={{
            padding:"8px 16px", borderRadius:50, border:"none", cursor:"pointer", fontSize:12.5, fontWeight:700,
            background: activeRoundTab===r.name ? "#1d6fa4" : "#e0e6ed",
            color: activeRoundTab===r.name ? "#fff" : "#444",
          }}>
            {r.label}{r.locked && " 🔒"}
          </button>
        ))}
      </div>

      {isDraft && isAdmin && hqOptions.length > 0 && (
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:20}}>
          <div style={{fontSize:14, fontWeight:800, marginBottom:12, color:"#333"}}>✅ 사업부별 제출 현황(관리자 전용)</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse", width:"100%", minWidth:480, fontSize:12.5}}>
              <thead>
                <tr style={{background:"#f0f4f8"}}>
                  <th style={tH}>지역담당</th>
                  <th style={tH}>제출여부</th>
                  <th style={tH}>제출 일시</th>
                  <th style={{...tH, minWidth:220}}>제출 히스토리</th>
                </tr>
              </thead>
              <tbody>
                {hqOptions.map((hq,i)=>{
                  const sub = submissions[hq];
                  const history = sub?.history || [];
                  return (
                    <tr key={hq} style={{background:i%2===0?"#fafafa":"#fff"}}>
                      <td style={{...tD, textAlign:"center", fontWeight:700}}>{hq}</td>
                      <td style={{...tD, textAlign:"center"}}>
                        {sub?.submitted
                          ? <span style={{color:"#2e8b57", fontWeight:700}}>✅ 제출완료</span>
                          : <span style={{color:"#aaa"}}>⏳ 미제출</span>}
                      </td>
                      <td style={{...tD, textAlign:"center", color:"#666"}}>
                        {sub?.submittedAt ? new Date(sub.submittedAt).toLocaleString('ko-KR') : "-"}
                      </td>
                      <td style={tD}>
                        {history.length === 0 ? (
                          <span style={{color:"#ccc"}}>-</span>
                        ) : (
                          <div style={{display:"flex", flexDirection:"column", gap:4, maxHeight:80, overflowY:"auto"}}>
                            {[...history].reverse().map((h,j)=>(
                              <div key={j} style={{fontSize:11}}>
                                <span style={{color:"#999"}}>{new Date(h.submittedAt).toLocaleString('ko-KR')}</span>
                                {h.note && <span style={{color:"#444"}}> — {h.note}</span>}
                              </div>
                            ))}
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {!isDraft && !activeRound?.locked && (
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:24}}>
          <div style={{fontSize:15, fontWeight:800, marginBottom:10}}>📂 {activeRoundTab} 선정 매장 업로드</div>
          <div style={{fontSize:12, color:"#888", marginBottom:12}}>
            {activeRoundTab} 때 라지그래픽 선정 매장 엑셀을 업로드하면 이 탭의 목록이 통째로 교체됩니다.
          </div>
          <div style={{display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
            <button style={{...settingsBtn("#1d6fa4"), padding:"7px 18px", fontSize:12}}
              onClick={()=>fileRef.current?.click()}>📂 파일 선택</button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleUpload}/>
            {uploadStatus && (
              <span style={{fontSize:12, padding:"5px 12px", borderRadius:8,
                background:uploadStatus.type==="success"?"#e8f8ee":"#fff0f0",
                color:uploadStatus.type==="success"?"#2e8b57":"#c00"}}>
                {uploadStatus.type==="success"?"✅ ":"❌ "}{uploadStatus.msg}
              </span>
            )}
          </div>
        </div>
      )}

      {!isDraft && activeRound?.locked && (
        <div style={{background:"#fff8f0", borderRadius:14, padding:"12px 20px", fontSize:12.5, color:"#e85d26"}}>
          🔒 {activeRoundTab} 선정 매장은 확정된 히스토리입니다 ({new Date(activeRound.finalizedAt).toLocaleString('ko-KR')} 확정) — 조회 전용이며 수정할 수 없습니다.
        </div>
      )}

      {isDraft && (
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:16,
          display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
          <span style={{fontSize:12, color:"#888"}}>현재 작업 중인 신규 선정 리스트입니다. 매장을 추가/수정/삭제한 뒤, 준비되면 관리자가 이름을 정해 확정하세요.</span>
          <button onClick={()=>setShowAddStore(true)}
            style={{...settingsBtn("#1d6fa4"), padding:"7px 14px", fontSize:12, marginLeft:"auto"}}>+ 매장 추가</button>
          {isAdmin && (
            <button onClick={()=>setShowFinalize(true)}
              style={{...settingsBtn("#2e8b57"), padding:"7px 14px", fontSize:12}}>✅ 최종 확정</button>
          )}
        </div>
      )}

      {compareSlotOptions.length > 0 && (
        <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", padding:16,
          display:"flex", gap:10, alignItems:"center", flexWrap:"wrap"}}>
          <span style={{fontSize:12.5, fontWeight:700, color:"#444"}}>O/X 비교 라운드</span>
          {[0,1,2].map(i=>(
            <select key={i} value={compareSlots[i]||""} onChange={e=>updateCompareSlot(i, e.target.value)}
              style={{padding:"6px 10px", borderRadius:8, border:"1px solid #ddd",
                fontSize:12, fontWeight:600, color:"#444", background:"#fff", cursor:"pointer"}}>
              <option value="">비워짐</option>
              {compareSlotOptions.map(r=>(
                <option key={r.name} value={r.name}>{r.name}</option>
              ))}
            </select>
          ))}
        </div>
      )}

      {viewData.length > 0 && (
        <>
          <div style={{display:"flex", gap:8, alignItems:"center", flexWrap:"wrap"}}>
            <span style={{fontSize:13, fontWeight:700, color:"#444"}}>📋 선정 매장 ({viewData.length})</span>
            {hqOptions.length > 0 && (
              <select value={hqFilter} onChange={e=>setHqFilter(e.target.value)}
                style={{marginLeft:"auto", padding:"7px 14px", borderRadius:8, border:"1px solid #ddd",
                  fontSize:12.5, fontWeight:600, color:"#444", background:"#fff", cursor:"pointer"}}>
                <option value="전체">전체 본부 ({viewData.length})</option>
                {hqOptions.map(hq=>(
                  <option key={hq} value={hq}>{hq} ({viewData.filter(r=>r.본부===hq).length})</option>
                ))}
              </select>
            )}
            {isDraft && hqFilter !== "전체" && (
              <>
                {isSubmitted && (
                  <span style={{fontSize:12, fontWeight:700, color:"#2e8b57", background:"#e8f8ee",
                    padding:"6px 12px", borderRadius:8}}>
                    ✅ 제출완료 ({new Date(submission.submittedAt).toLocaleString('ko-KR')})
                  </span>
                )}
                <button
                  onClick={()=> isSubmitted ? handleCancelSubmit() : setShowSubmitNote(true)}
                  style={{...settingsBtn(isSubmitted?"#aaa":"#e8420a"), padding:"7px 18px", fontSize:12.5}}>
                  {isSubmitted ? "제출 취소" : "제출하기"}
                </button>
              </>
            )}
          </div>

          <div style={{background:"#fff", borderRadius:14, boxShadow:"0 2px 16px rgba(0,0,0,0.07)", overflow:"hidden"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse", width:"100%", minWidth:900+compareRounds.length*80, fontSize:12}}>
                <thead>
                  <tr style={{background:"#f0f4f8"}}>
                    <th style={tH}>#</th>
                    <th style={tH}>본부</th>
                    <th style={tH}>매장코드</th>
                    <th style={{...tH, minWidth:140}}>매장명</th>
                    <th style={{...tH, minWidth:180}}>주소</th>
                    {compareRounds.map(r=>(<th key={r.name} style={tH}>{r.name}</th>))}
                    {showAccpCol && <th style={tH}>ACCP 매장</th>}
                    <th style={{...tH, minWidth:140}}>비고</th>
                    <th style={{...tH, minWidth:110}}>전면사진</th>
                    <th style={{...tH, minWidth:100}}>지도</th>
                    {isDraft && <th style={{...tH, width:50}}>삭제</th>}
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((row,i)=>{
                    const imgs = photos[row.매장코드] || [];
                    return (
                      <tr key={row.id} style={{background:i%2===0?"#fafafa":"#fff"}}>
                        <td style={{...tD, textAlign:"center", color:"#aaa"}}>{i+1}</td>
                        <td style={{...tD, textAlign:"center"}}>{row.본부}</td>
                        <td style={{...tD, fontSize:11, color:"#666"}}>{row.매장코드}</td>
                        <td style={{...tD, fontWeight:600}}>{row.매장명}</td>
                        <td style={{...tD, fontSize:11, color:"#666"}}>{row.주소}</td>
                        {compareRounds.map(r=>(
                          <td key={r.name} style={{...tD, textAlign:"center"}}><CompareCell round={r} row={row} /></td>
                        ))}
                        {showAccpCol && <td style={{...tD, textAlign:"center"}}><OXField row={row} field="ACCP매장" /></td>}
                        <td style={tD}>
                          {isDraft ? (
                            <input type="text" value={row.비고||""} placeholder="-"
                              onChange={e=>updateDraftRow(row.id,"비고",e.target.value)}
                              style={{...numInput, width:"100%", textAlign:"left", fontSize:11}} />
                          ) : (row.비고||"-")}
                        </td>
                        <td style={{...tD, textAlign:"center"}}>
                          <button
                            onClick={()=>setPhotoPopupCode(row.매장코드)}
                            style={{...settingsBtn(imgs.length>0?"#2e8b57":"#888"), padding:"5px 12px", fontSize:11.5}}>
                            📷 {imgs.length>0 ? `${imgs.length}장` : "등록"}
                          </button>
                        </td>
                        <td style={{...tD, textAlign:"center"}}>
                          <a
                            href={`https://map.naver.com/p/search/${encodeURIComponent(row.매장명||"")}`}
                            target="_blank" rel="noopener noreferrer"
                            style={{...settingsBtn("#03c75a"), padding:"5px 12px", fontSize:11.5,
                              textDecoration:"none", display:"inline-block"}}>
                            지도로 보기
                          </a>
                        </td>
                        {isDraft && (
                          <td style={{...tD, textAlign:"center"}}>
                            <button onClick={()=>deleteDraftRow(row.id)}
                              style={{width:24, height:24, borderRadius:"50%", border:"none", background:"#c00",
                                color:"#fff", fontSize:13, fontWeight:800, cursor:"pointer", lineHeight:1}}>×</button>
                          </td>
                        )}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {photoPopupRow && (
        <SianPopup
          itemId={photoPopupRow.매장코드}
          itemName={photoPopupRow.매장명}
          images={photos[photoPopupRow.매장코드] || []}
          isAdmin={isAdmin}
          onSave={(imgs)=>{ setPhotos(prev=>({...prev, [photoPopupRow.매장코드]: imgs})); setPhotoPopupCode(null); }}
          onClose={()=>setPhotoPopupCode(null)}
        />
      )}

      {showAddStore && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:380, maxWidth:440, alignItems:"stretch", gap:10}}>
            <div style={{fontSize:16, fontWeight:800, marginBottom:4}}>+ 매장 추가</div>
            {[["본부","본부"],["매장코드","매장코드"],["매장명","매장명"],["주소","주소"]].map(([field,label])=>(
              <div key={field} style={{display:"flex", flexDirection:"column", gap:4}}>
                <label style={{fontSize:11.5, color:"#888", fontWeight:600}}>{label}</label>
                <input type="text" value={addStoreForm[field]}
                  onChange={e=>setAddStoreForm(prev=>({...prev,[field]:e.target.value}))}
                  style={{...numInput, width:"100%", textAlign:"left"}} />
              </div>
            ))}
            <div style={{display:"flex", gap:10, marginTop:8}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>setShowAddStore(false)}>취소</button>
              <button style={{...styles2.popupBtn, flex:1}} onClick={addDraftRow}>추가</button>
            </div>
          </div>
        </div>
      )}

      {showSubmitNote && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:400, maxWidth:460, alignItems:"stretch", gap:10}}>
            <div style={{fontSize:16, fontWeight:800, marginBottom:4}}>제출하기</div>
            <div style={{fontSize:12.5, color:"#888"}}>이번에 수정한 내용을 간단히 남겨주세요 (히스토리로 기록됩니다).</div>
            <textarea
              value={submitNote} onChange={e=>setSubmitNote(e.target.value)}
              placeholder="예) 20번 매장 삭제하고 OO대리점 OO점 추가했습니다."
              rows={4}
              style={{...numInput, width:"100%", textAlign:"left", resize:"vertical", fontFamily:"inherit"}} />
            <div style={{display:"flex", gap:10, marginTop:8}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>{setShowSubmitNote(false); setSubmitNote("");}}>취소</button>
              <button style={{...styles2.popupBtn, flex:1, background:"linear-gradient(135deg,#e8420a,#ff6b35)"}}
                onClick={handleSubmitConfirm}>제출하기</button>
            </div>
          </div>
        </div>
      )}

      {showFinalize && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:380, maxWidth:440, alignItems:"stretch", gap:10}}>
            <div style={{fontSize:16, fontWeight:800, marginBottom:4}}>✅ 최종 확정</div>
            <div style={{fontSize:12.5, color:"#888"}}>어떤 이름으로 저장하시겠습니까? (예: 폴더블8)</div>
            <input type="text" value={finalizeName} placeholder="라운드 이름 입력"
              onChange={e=>setFinalizeName(e.target.value)}
              style={{...numInput, width:"100%", textAlign:"left"}} />
            <div style={{fontSize:11, color:"#e85d26"}}>확정 후에는 이 리스트를 수정/삭제할 수 없고, 히스토리 조회 및 다른 라운드의 O/X 대조 기준으로만 사용됩니다.</div>
            <div style={{display:"flex", gap:10, marginTop:8}}>
              <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", flex:1}}
                onClick={()=>{setShowFinalize(false); setFinalizeName("");}}>취소</button>
              <button style={{...styles2.popupBtn, flex:1, background:"linear-gradient(135deg,#2e8b57,#1d6fa4)"}}
                onClick={handleFinalize}>확정</button>
            </div>
          </div>
        </div>
      )}

      {showBaseModal && (
        <div style={styles2.popupOverlay}>
          <div style={{...styles2.popupBox, minWidth:380, maxWidth:440, alignItems:"stretch", gap:10}}>
            <div style={{fontSize:16, fontWeight:800, marginBottom:4}}>어떤 리스트를 기반으로 시작할까요?</div>
            <div style={{fontSize:12.5, color:"#888", marginBottom:4}}>선택한 라운드의 매장 목록을 복사해서 신규 선정 작업을 시작합니다.</div>
            {draft.length > 0 && (
              <div style={{fontSize:11.5, color:"#e85d26", background:"#fff8f0", padding:"8px 10px", borderRadius:8}}>
                ⚠️ 지금 작업 중인 신규 선정 리스트({draft.length}개 매장)가 선택한 리스트로 덮어써집니다.
              </div>
            )}
            <div style={{display:"flex", flexDirection:"column", gap:8}}>
              {[...rounds].reverse().filter(r=>r.data.length>0).map(r=>(
                <button key={r.name} onClick={()=>startDraftFrom(r.name)}
                  style={{...settingsBtn("#1d6fa4"), width:"100%", padding:"10px 14px", fontSize:13, textAlign:"left"}}>
                  {r.label} ({r.data.length}개 매장)
                </button>
              ))}
              <button onClick={()=>startDraftFrom(null)}
                style={{...settingsBtn("#888"), width:"100%", padding:"10px 14px", fontSize:13, textAlign:"left"}}>
                빈 리스트로 시작
              </button>
            </div>
            <button style={{...styles2.popupBtn, background:"#aaa", color:"#333", marginTop:8}}
              onClick={()=>setShowBaseModal(false)}>취소</button>
          </div>
        </div>
      )}
    </div>
  );
}

function settingsBtn(color, textColor="#fff") {
  return {
    background:color, color:textColor, border:"none",
    borderRadius:6, padding:"4px 12px", fontSize:12,
    fontWeight:600, cursor:"pointer",
  };
}

const styles2 = {
  popupOverlay: {
    position:"fixed", inset:0, background:"rgba(0,0,0,0.7)",
    backdropFilter:"blur(6px)",
    display:"flex", alignItems:"center", justifyContent:"center", zIndex:9999,
    animation:"fadeIn 0.2s ease both",
  },
  popupBox: {
    background:"var(--c-popup-bg)",
    border:`1px solid ${C.borderAccent}`,
    borderRadius:20, padding:"40px 48px",
    display:"flex", flexDirection:"column", alignItems:"center", gap:12,
    boxShadow:"var(--c-card-shadow)",
    minWidth:320,
    animation:"fadeUp 0.3s cubic-bezier(.22,1,.36,1) both",
  },
  popupIcon:  { fontSize:48 },
  popupTitle: { fontSize:20, fontWeight:800, color:C.text0, textAlign:"center" },
  popupSub:   { fontSize:14, color:C.text1, textAlign:"center" },
  popupBtn: {
    marginTop:8, padding:"11px 24px",
    background:"linear-gradient(135deg,#ff6b35,#e8420a)",
    color:"#fff", border:"none", borderRadius:50,
    fontSize:14, fontWeight:700, cursor:"pointer",
    boxShadow:"none",
    transition:"all 0.2s",
  },
};



// ─── 시안 이미지 팝업 ──────────────────────────────────────────────────────
function SianPopup({ itemId, itemName, images, isAdmin, onSave, onClose }) {
  const [localImgs, setLocalImgs] = useState(images||[]);
  const [uploading, setUploading] = useState(0);
  const fileRef = useRef(null);
  const [viewImg, setViewImg] = useState(null);
  const imgSrc = img => img.url || img.dataUrl;
  const handleUpload = async e => {
    const files = Array.from(e.target.files);
    e.target.value = "";
    setUploading(n => n + files.length);
    for (const file of files) {
      try {
        const uploaded = await dataClient.storage.uploadImage(file);
        if (uploaded) {
          setLocalImgs(prev => [...prev, { id: Date.now()+Math.random(), name: file.name, url: uploaded.url, path: uploaded.path }]);
          setUploading(n => n - 1);
          continue;
        }
      } catch (err) {
        console.error("[SianPopup] Storage 업로드 실패, base64로 폴백:", err);
      }
      // Playground(스토리지 미지원) 또는 업로드 실패 시 기존 방식으로 폴백
      const reader = new FileReader();
      reader.onload = ev => {
        setLocalImgs(prev => [...prev, { id: Date.now()+Math.random(), name: file.name, dataUrl: ev.target.result }]);
        setUploading(n => n - 1);
      };
      reader.readAsDataURL(file);
    }
  };
  const handleSave = async () => {
    const removed = (images||[]).filter(orig => !localImgs.some(l => l.id === orig.id));
    await Promise.all(removed.map(img => dataClient.storage.deleteImage(img.path)));
    onSave(localImgs);
  };
  const downloadImg = async (img) => {
    const src = imgSrc(img);
    let href = src, revoke = null;
    if (/^https?:\/\//.test(src)) {
      try {
        const blob = await (await fetch(src)).blob();
        href = URL.createObjectURL(blob);
        revoke = href;
      } catch (err) {
        console.error("[SianPopup] 다운로드 실패:", err);
      }
    }
    const a = document.createElement("a");
    a.href = href;
    a.download = img.name || "시안.png";
    document.body.appendChild(a);
    a.click();
    a.remove();
    if (revoke) URL.revokeObjectURL(revoke);
  };
  return (
    <div style={styles2.popupOverlay}>
      <div style={{...styles2.popupBox,minWidth:420,maxWidth:560,gap:14,maxHeight:"85vh",overflowY:"auto"}}>
        <div style={{fontSize:17,fontWeight:800}}>🖼 시안 — {itemName}</div>
        {isAdmin && (<>
          <button style={{...styles2.popupBtn,background:"#2e8b57",width:"100%"}}
            onClick={()=>fileRef.current?.click()}>+ 이미지 업로드</button>
          <input ref={fileRef} type="file" accept="image/*" multiple style={{display:"none"}} onChange={handleUpload}/>
        </>)}
        {uploading>0 && <div style={{fontSize:12,color:"#2e8b57"}}>⏳ 업로드 중... ({uploading})</div>}
        {localImgs.length===0
          ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:24}}>{isAdmin?"이미지를 업로드해주세요.":"등록된 시안이 없습니다."}</div>
          : <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,width:"100%"}}>
              {localImgs.map(img=>(
                <div key={img.id} style={{position:"relative",borderRadius:8,overflow:"hidden",border:"1px solid #eee",cursor:"pointer"}}
                  onClick={()=>setViewImg(img)}>
                  <img src={imgSrc(img)} alt={img.name} style={{width:"100%",height:120,objectFit:"cover"}}/>
                  <div style={{fontSize:11,color:"#666",padding:"4px 6px",background:"#f8f8f8",
                    overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{img.name}</div>
                  <button style={{position:"absolute",top:4,right:isAdmin?28:4,background:"rgba(0,0,0,0.5)",
                    color:"#fff",border:"none",borderRadius:"50%",width:20,height:20,fontSize:11,cursor:"pointer",
                    display:"flex",alignItems:"center",justifyContent:"center"}}
                    title="다운로드"
                    onClick={e=>{e.stopPropagation();downloadImg(img);}}>⬇</button>
                  {isAdmin && (
                    <button style={{position:"absolute",top:4,right:4,background:"rgba(0,0,0,0.5)",
                      color:"#fff",border:"none",borderRadius:"50%",width:20,height:20,fontSize:12,cursor:"pointer"}}
                      onClick={e=>{e.stopPropagation();setLocalImgs(p=>p.filter(i=>i.id!==img.id));}}>×</button>
                  )}
                </div>
              ))}
            </div>
        }
        <div style={{display:"flex",gap:10,width:"100%"}}>
          <button style={{...styles2.popupBtn,background:"#aaa",color:"#333",flex:1}} onClick={onClose}>취소</button>
          {isAdmin
            ? <button style={{...styles2.popupBtn,flex:1}} onClick={handleSave}>저장</button>
            : <button style={{...styles2.popupBtn,flex:1}} onClick={onClose}>닫기</button>}
        </div>
      </div>
      {viewImg && createPortal(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:10000,
          display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setViewImg(null)}>
          <img src={imgSrc(viewImg)} alt={viewImg.name}
            style={{maxWidth:"90vw",maxHeight:"90vh",width:"auto",height:"auto",borderRadius:8,objectFit:"contain",display:"block"}}/>
          <button style={{position:"absolute",top:24,right:76,background:"rgba(255,255,255,0.15)",
            color:"#fff",border:"none",borderRadius:"50%",width:40,height:40,fontSize:18,cursor:"pointer"}}
            title="다운로드"
            onClick={e=>{e.stopPropagation();downloadImg(viewImg);}}>⬇</button>
          <button style={{position:"absolute",top:24,right:24,background:"rgba(255,255,255,0.15)",
            color:"#fff",border:"none",borderRadius:"50%",width:40,height:40,fontSize:18,cursor:"pointer"}}
            title="닫기"
            onClick={()=>setViewImg(null)}>×</button>
        </div>,
        document.body
      )}
    </div>
  );
}

const styles = {
  // ── login / auth ──
  loginBg: {
    minHeight:"100vh",
    background: C.bg0,
    backgroundImage: `
      radial-gradient(ellipse 80% 60% at 50% -10%, rgba(255,107,53,0.15) 0%, transparent 70%),
      linear-gradient(180deg, ${C.bg0} 0%, ${C.bg1} 100%)
    `,
    display:"flex", alignItems:"center", justifyContent:"center",
    position:"relative", overflow:"hidden",
  },
  loginCard: {
    background: `linear-gradient(145deg, ${C.bg2}, ${C.bg1})`,
    border: `1px solid ${C.border}`,
    borderRadius:20,
    boxShadow:"var(--c-card-shadow)",
    padding:"48px 44px 44px",
    width:420, maxWidth:"92vw",
    display:"flex", flexDirection:"column", alignItems:"center",
  },
  loginLogo: {
    width:52, height:52, borderRadius:14,
    background:"linear-gradient(135deg,#ff6b35,#e8420a)",
    display:"flex", alignItems:"center", justifyContent:"center",
    fontSize:26, marginBottom:16,
    boxShadow:"0 8px 24px rgba(255,107,53,0.4)",
  },
  loginTitle: { fontSize:22, fontWeight:800, color:C.text0, marginBottom:4, letterSpacing:-0.5 },
  loginSub:   { fontSize:13, color:C.text1, marginBottom:28 },
  loginField: { width:"100%", marginBottom:16 },
  loginLabel: { display:"block", fontSize:12, fontWeight:600, color:C.text1, marginBottom:7, letterSpacing:0.5, textTransform:"uppercase" },
  loginInput: {
    width:"100%",
    padding:"11px 14px", borderRadius:10,
    background:C.surface,
    border:`1.5px solid var(--c-input-border)`,
    color:C.text0, fontSize:14,
    transition:"all 0.2s",
  },
  loginErr:   { color:"#fb7185", fontSize:12, marginBottom:8, alignSelf:"flex-start", display:"flex", alignItems:"center", gap:5 },
  loginBtn: {
    width:"100%", padding:"13px",
    background:"linear-gradient(135deg,#ff6b35,#e8420a)",
    color:"#fff", border:"none", borderRadius:12,
    fontSize:15, fontWeight:700, cursor:"pointer",
    marginTop:10, letterSpacing:0.5,
    boxShadow:"none",
    transition:"all 0.2s",
  },

  // ── layout ──
  app: { display:"flex", height:"100vh", background:C.bg0, overflow:"hidden" },
  sidebar: {
    width:220, flexShrink:0,
    background:`linear-gradient(180deg, ${C.bg1} 0%, ${C.bg0} 100%)`,
    borderRight:`1px solid ${C.border}`,
    display:"flex", flexDirection:"column",
    padding:"0",
  },
  sidebarLogo: {
    padding:"24px 20px 20px",
    borderBottom:`1px solid ${C.border}`,
    display:"flex", alignItems:"center", gap:10,
  },
  sidebarLogoIcon: {
    width:34, height:34, borderRadius:9,
    background:"linear-gradient(135deg,#ff6b35,#c73500)",
    display:"flex", alignItems:"center", justifyContent:"center",
    fontSize:18, flexShrink:0,
    boxShadow:"0 4px 14px rgba(255,107,53,0.35)",
  },
  sidebarLogoText: { color:C.text0, fontWeight:800, fontSize:16, letterSpacing:-0.3 },
  nav: { flex:1, display:"flex", flexDirection:"column", gap:2, padding:"14px 10px" },
  navBtn: {
    background:"transparent", border:"none", borderLeft:"2px solid transparent",
    color:C.text1,
    padding:"10px 14px", borderRadius:"0 10px 10px 0", cursor:"pointer", textAlign:"left",
    fontSize:13.5, fontWeight:500, transition:"all 0.15s",
    display:"flex", alignItems:"center", gap:10,
  },
  navBtnActive: {
    background:C.orangeDim, color:C.orange,
    fontWeight:700, borderLeft:`2px solid ${C.orange}`,
  },
  sidebarFooter: {
    padding:"16px 14px",
    borderTop:`1px solid ${C.border}`,
    display:"flex", flexDirection:"column", gap:10,
  },
  roleBadge: {
    color:"#fff", borderRadius:8, padding:"5px 10px",
    fontSize:11, fontWeight:700, textAlign:"center", letterSpacing:0.5,
  },
  userId: { color:C.text2, fontSize:11.5, textAlign:"center", fontFamily:C.mono },
  logoutBtn: {
    background:"rgba(255,255,255,0.05)",
    border:`1px solid ${C.border}`,
    color:C.text1, padding:"7px 12px", borderRadius:8,
    cursor:"pointer", fontSize:12, fontWeight:500,
    transition:"all 0.15s",
  },
  main: { flex:1, overflow:"auto" },

  // ── page ──
  page: { padding:"32px 36px", minHeight:"100%" },
  pageHeader: { display:"flex", alignItems:"baseline", gap:14, marginBottom:28 },
  pageTitle: { fontSize:22, fontWeight:800, color:C.text0, margin:0, letterSpacing:-0.5 },
  periodNote: { fontSize:12, color:C.text2 },

  // ── schedule ──
  scheduleLayout: { display:"flex", gap:24, alignItems:"flex-start" },
  calSection: {
    flex:"0 0 500px",
    background:`linear-gradient(145deg, ${C.bg2}, ${C.bg1})`,
    border:`1px solid ${C.border}`,
    borderRadius:16,
    boxShadow:"var(--c-section-shadow)",
    padding:22,
  },
  calToolbar: { display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:18, flexWrap:"wrap", gap:10 },
  monthNav: { display:"flex", alignItems:"center", gap:10 },
  monthLabel: { fontSize:17, fontWeight:800, color:C.text0, minWidth:126, textAlign:"center", letterSpacing:-0.3 },
  navArrow: {
    background:C.overlay, border:`1px solid ${C.border}`,
    borderRadius:8, width:30, height:30, fontSize:16, cursor:"pointer", color:C.text1,
    transition:"all 0.15s",
  },
  calActions: { display:"flex", gap:7, flexWrap:"wrap" },

  aiBtn: {
    background:"rgba(59,130,246,0.15)", color:"#60a5fa",
    border:"1px solid rgba(59,130,246,0.3)",
    borderRadius:8, padding:"6px 13px", fontSize:12, fontWeight:600, cursor:"pointer",
    transition:"all 0.15s",
  },
  tempBtn: {
    background:C.orangeDim, color:C.orange,
    border:`1px solid ${C.borderAccent}`,
    borderRadius:8, padding:"6px 13px", fontSize:12, fontWeight:600, cursor:"pointer",
    transition:"all 0.15s",
  },
  editBtn: {
    background:C.surface, color:C.text1,
    border:`1px solid ${C.border}`,
    borderRadius:8, padding:"6px 13px", fontSize:12, fontWeight:600, cursor:"pointer",
    transition:"all 0.15s",
  },
  skBtn: {
    background:"rgba(16,185,129,0.15)", color:"#34d399",
    border:"1px solid rgba(16,185,129,0.3)",
    borderRadius:8, padding:"6px 14px", fontSize:12, fontWeight:600, cursor:"pointer",
    transition:"all 0.15s",
  },

  calendar: { userSelect:"none" },
  calHead: { display:"grid", gridTemplateColumns:"repeat(7,1fr)", marginBottom:6 },
  calHeadCell: { textAlign:"center", padding:"5px 0", fontSize:11, fontWeight:700, color:C.text2, letterSpacing:0.5, textTransform:"uppercase" },
  calBody: { display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:"1px 0" },
  calCell: {
    position:"relative", height:50, display:"flex", flexDirection:"column",
    alignItems:"center", justifyContent:"center", borderRadius:8,
    transition:"background 0.1s",
  },
  circle: {
    position:"absolute", width:34, height:34, borderRadius:"50%",
    top:"50%", left:"50%", transform:"translate(-50%,-50%)", zIndex:1,
  },
  // "오늘" 표시는 원/링 형태로 만들면 확정(빨간 원)·선택 중(파란 원)·해제 표시와 계속 헷갈리므로
  // 원이 아닌 아래쪽 작은 밑줄로만 표시한다.
  todayUnderline: {
    position:"absolute", width:16, height:3, borderRadius:2,
    background:"#94a3b8", left:"50%", bottom:4,
    transform:"translateX(-50%)", zIndex:0,
  },
  dayNum: { fontSize:13, position:"relative", zIndex:2, fontWeight:500 },
  holDot: { fontSize:7, position:"absolute", bottom:3, right:3, opacity:0.7 },

  legend: { display:"flex", gap:14, marginTop:14, flexWrap:"wrap", alignItems:"center" },

  // ── table ──
  tableSection: {
    flex:1,
    background:`linear-gradient(145deg, ${C.bg2}, ${C.bg1})`,
    border:`1px solid ${C.border}`,
    borderRadius:16,
    boxShadow:"var(--c-section-shadow)",
    padding:22,
  },
  tableTitleRow: { marginBottom:14 },
  tableTitle: { fontSize:15, fontWeight:700, color:C.text0, letterSpacing:-0.2 },
  table: { width:"100%", borderCollapse:"collapse", fontSize:13.5 },
  th: {
    background:C.surface, padding:"10px 14px", textAlign:"left",
    fontWeight:700, color:C.text2, borderBottom:`1px solid ${C.border}`,
    fontSize:11.5, letterSpacing:0.5, textTransform:"uppercase",
  },
  td: { padding:"10px 14px", borderBottom:`1px solid ${C.border}`, color:C.text0 },
  noteInput: {
    width:"100%", border:"none",
    background:"transparent", fontSize:13.5,
    outline:"none", color:C.text0,
    borderBottom:`1px solid ${C.border}`,
    padding:"2px 0",
  },
  editHint: {
    marginTop:10, padding:"10px 14px",
    background:"rgba(234,179,8,0.08)",
    borderRadius:8, fontSize:12, color:"#fbbf24",
    border:"1px solid rgba(234,179,8,0.2)",
  },
};
