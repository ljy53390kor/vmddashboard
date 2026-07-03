/**
 * 공공데이터포털 - 한국천문연구원 특일정보 API
 * /getRestDeInfo : 공휴일(대체공휴일 포함) 조회
 *
 * - 캐싱: localStorage에 7일 보관 (매번 36번 요청 방지)
 * - 연도 범위: 현재 연도 기준 ±2년 자동 계산
 */

const API_KEY  = import.meta.env.VITE_HOLIDAY_API_KEY;
const BASE_URL = '/holiday-api/B090041/openapi/service/SpcdeInfoService/getRestDeInfo';
const CACHE_KEY = 'vmd-holidays-v1';
const CACHE_TTL = 7 * 24 * 60 * 60 * 1000; // 7일

/** 특정 연·월의 공휴일 날짜 배열 반환 ("YYYY-MM-DD") */
async function fetchMonth(year, month) {
  const mm  = String(month).padStart(2, '0');
  const url = `${BASE_URL}?serviceKey=${encodeURIComponent(API_KEY)}`
            + `&solYear=${year}&solMonth=${mm}&numOfRows=30&_type=json`;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const json = await res.json();
  const body = json?.response?.body;
  if (!body || body.totalCount === 0) return [];

  const item = body.items?.item;
  if (!item) return [];

  return (Array.isArray(item) ? item : [item]).map(({ locdate }) => {
    const d = String(locdate);
    return `${d.slice(0, 4)}-${d.slice(4, 6)}-${d.slice(6, 8)}`;
  });
}

/** 여러 연도의 전체 공휴일을 Set으로 반환 (캐시 우선) */
export async function loadHolidays(years) {
  // ── 캐시 확인 ──────────────────────────────────────────────────
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { dates, ts, cachedYears } = JSON.parse(raw);
      const yearsMatch = JSON.stringify([...years].sort()) ===
                         JSON.stringify([...cachedYears].sort());
      if (yearsMatch && Date.now() - ts < CACHE_TTL) {
        console.log('[holidays] 캐시 사용:', dates.length, '일');
        return new Set(dates);
      }
    }
  } catch { /* 캐시 실패 무시 */ }

  // ── API 호출 (연도별 12개월) ────────────────────────────────────
  const allDates = [];
  for (const year of years) {
    for (let m = 1; m <= 12; m++) {
      try {
        const dates = await fetchMonth(year, m);
        allDates.push(...dates);
      } catch (e) {
        console.warn(`[holidays] ${year}-${String(m).padStart(2,'0')} 로드 실패:`, e.message);
      }
    }
  }

  // ── 캐시 저장 ──────────────────────────────────────────────────
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({
      dates: allDates,
      ts: Date.now(),
      cachedYears: years,
    }));
  } catch { /* 스토리지 가득 참 등 무시 */ }

  console.log('[holidays] API에서 로드 완료:', allDates.length, '일');
  return new Set(allDates);
}

/** 캐시 강제 초기화 (업데이트 필요 시 수동 호출) */
export function clearHolidayCache() {
  localStorage.removeItem(CACHE_KEY);
}
