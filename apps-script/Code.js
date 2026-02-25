/**
 * 급식 + 시간표 통합 관리 시스템
 * 1단계: 평일 날짜 페이지 일괄 생성
 * 2단계: 급식 메뉴 업데이트 (NEIS API → 기존 페이지 patch)
 * 3단계: 시간표 이미지 삽입 (선택속성 1~7 → 이미지 블록)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('급식관리')
    .addItem('1. 날짜 페이지 생성', 'manualCreatePages')
    .addItem('2. 급식 메뉴 업데이트', 'manualUpdateMeals')
    .addItem('3. 시간표 이미지 삽입', 'manualUpdateTimetable')
    .addSeparator()
    .addItem('전체 실행 (1→2→3)', 'manualFullUpdate')
    .addToUi();
}

// ===== 설정 =====
// initializeConfig()는 Apps Script 에디터에서 최초 1회만 실행하세요.
// 실행 후 PropertiesService에 저장되므로 값은 코드에 남기지 않습니다.
function initializeConfig() {
  PropertiesService.getScriptProperties().setProperties({
    'NEIS_API_KEY': '여기에_NEIS_API_KEY',
    'NOTION_TOKEN': '여기에_NOTION_TOKEN',
    'NOTION_DB_ID': '여기에_NOTION_DB_ID',
    'ATPT_OFCDC_SC_CODE': '여기에_교육청코드',
    'SD_SCHUL_CODE': '여기에_학교코드',
    'SPREADSHEET_ID': '여기에_스프레드시트_ID',
    'TIMETABLE_PROP_NAME': '시간표'
  });
  Logger.log('설정 저장 완료');
}

function getConfig() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ===== 이미지 URL 매핑 (GitHub 등에 업로드 후 수정) =====
const TIMETABLE_IMAGES = {
  '1': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/1.png',
  '2': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/2.png',
  '3': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/3.png',
  '4': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/4.png',
  '5': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/5.png',
  '6': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/6.png',
  '7': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/7.png',
  '8': 'https://raw.githubusercontent.com/getbetterwithme/meal-to-notion/main/timetable/8.png'
};

// ===== 트리거 전용: 매일 실행, 20일~말일만 동작 =====
function scheduledNextMonthUpdate() {
  const now = new Date();
  const today = now.getDate();

  // 20일 이전이면 스킵
  if (today < 20) {
    Logger.log(`[자동실행] ${today}일 — 20일 이전이므로 스킵`);
    return;
  }

  const nextMonthDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const yearMonth = Utilities.formatDate(nextMonthDate, Session.getScriptTimeZone(), "yyyyMM");

  Logger.log(`[자동실행] ${today}일, 타겟 연월: ${yearMonth}`);

  createMonthPages(yearMonth);  // 이미 있으면 스킵됨
  updateMealData(yearMonth);    // 이미 메뉴 있으면 덮어쓰지만 동일 데이터
  Logger.log(`[자동실행] ${yearMonth} 완료`);
}

// ===== 수동 실행 (UI) =====
function manualCreatePages() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const count = createMonthPages(yearMonth);
  SpreadsheetApp.getUi().alert(`${yearMonth} 날짜 페이지 ${count}개 생성 완료`);
}

function manualUpdateMeals() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const count = updateMealData(yearMonth);
  SpreadsheetApp.getUi().alert(`${yearMonth} 급식 메뉴 ${count}건 업데이트 완료`);
}

function manualUpdateTimetable() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const count = updateTimetableImages(yearMonth);
  SpreadsheetApp.getUi().alert(`${yearMonth} 시간표 이미지 ${count}건 삽입 완료`);
}

function manualFullUpdate() {
  const yearMonth = promptYearMonth();
  if (!yearMonth) return;
  const ui = SpreadsheetApp.getUi();

  const pages = createMonthPages(yearMonth);
  const meals = updateMealData(yearMonth);
  const images = updateTimetableImages(yearMonth);

  ui.alert(`${yearMonth} 전체 완료\n- 페이지 생성: ${pages}건\n- 급식 업데이트: ${meals}건\n- 시간표 이미지: ${images}건`);
}

function promptYearMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('연월을 입력하세요 (예: 202603)');
  if (response.getSelectedButton() === ui.Button.CANCEL) return null;
  const yearMonth = response.getResponseText().trim();
  if (!/^\d{6}$/.test(yearMonth)) {
    ui.alert('YYYYMM 형식으로 입력해주세요.');
    return null;
  }
  return yearMonth;
}

// ===== 1단계: 평일 날짜 페이지 생성 =====
function createMonthPages(yearMonth) {
  const config = getConfig();
  const weekdays = getWeekdays(yearMonth);
  const existingMap = getNotionPagesMap(yearMonth, config);
  let created = 0;

  for (const dateStr of weekdays) {
    if (existingMap[dateStr]) continue; // 이미 존재

    const pageId = createNotionPage(dateStr, yearMonth, config);
    if (pageId) created++;
    Utilities.sleep(400);
  }

  Logger.log(`[1단계] ${yearMonth} 페이지 ${created}개 생성 (기존 ${Object.keys(existingMap).length}개)`);
  return created;
}

// ===== 2단계: 급식 메뉴 업데이트 =====
function updateMealData(yearMonth) {
  const config = getConfig();
  const meals = getMonthlyMeals(yearMonth, config);
  if (!meals || meals.length === 0) {
    Logger.log(`[2단계] ${yearMonth} 급식 데이터 없음`);
    return 0;
  }

  const pagesMap = getNotionPagesMap(yearMonth, config);
  let updated = 0;

  for (const meal of meals) {
    const pageId = pagesMap[meal.date];
    if (!pageId) {
      Logger.log(`[2단계] ${meal.date} 페이지 없음, 스킵`);
      continue;
    }

    const success = patchNotionPage(pageId, {
      "메뉴": { rich_text: [{ text: { content: meal.menu } }] }
    }, config);

    if (success) updated++;
    Utilities.sleep(400);
  }

  Logger.log(`[2단계] ${yearMonth} 급식 ${updated}건 업데이트`);
  return updated;
}

// ===== 3단계: 시간표 이미지 삽입 =====
function updateTimetableImages(yearMonth) {
  const config = getConfig();
  const timetableProp = config.TIMETABLE_PROP_NAME || '시간표';
  const pagesMap = getNotionPagesMap(yearMonth, config);
  let inserted = 0;

  for (const [dateStr, pageId] of Object.entries(pagesMap)) {
    // 페이지 상세 조회하여 시간표 속성값 확인
    const pageData = getNotionPage(pageId, config);
    if (!pageData) continue;

    const selectProp = pageData.properties[timetableProp]?.select;
    if (!selectProp || !selectProp.name) continue; // 시간표 미지정

    const imageUrl = TIMETABLE_IMAGES[selectProp.name];
    if (!imageUrl) continue; // 매핑 없음

    // 이미 이미지 블록이 있는지 확인
    const existingBlocks = getPageBlocks(pageId, config);
    const hasImage = existingBlocks.some(b => b.type === 'image');
    if (hasImage) continue; // 이미 이미지가 있으면 스킵

    const success = appendImageBlock(pageId, imageUrl, config);
    if (success) inserted++;
    Utilities.sleep(400);
  }

  Logger.log(`[3단계] ${yearMonth} 시간표 이미지 ${inserted}건 삽입`);
  return inserted;
}

// ===== 헬퍼: 평일 날짜 계산 =====
function getWeekdays(yearMonth) {
  const year = parseInt(yearMonth.substring(0, 4));
  const month = parseInt(yearMonth.substring(4, 6)) - 1;
  const dates = [];
  const lastDay = new Date(year, month + 1, 0).getDate();

  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month, d);
    const day = date.getDay();
    if (day >= 1 && day <= 5) { // 월~금
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      dates.push(dateStr);
    }
  }
  return dates;
}

// ===== 헬퍼: Notion DB에서 해당 월 페이지 맵 조회 =====
function getNotionPagesMap(yearMonth, config) {
  const url = `https://api.notion.com/v1/databases/${config.NOTION_DB_ID}/query`;
  const map = {};
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const body = {
      page_size: 100,
      filter: {
        property: "Month",
        select: { equals: String(yearMonth) }
      }
    };
    if (startCursor) body.start_cursor = startCursor;

    const options = {
      method: "post",
      headers: {
        "Authorization": `Bearer ${config.NOTION_TOKEN}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    };

    try {
      const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
      (res.results || []).forEach(p => {
        const dateStart = p.properties["날짜"]?.date?.start;
        if (dateStart) map[dateStart] = p.id;
      });
      hasMore = res.has_more;
      startCursor = res.next_cursor;
    } catch (e) {
      Logger.log(`getNotionPagesMap 오류: ${e}`);
      break;
    }
  }
  return map;
}

// ===== 헬퍼: Notion 페이지 생성 (빈 페이지) =====
function createNotionPage(dateStr, yearMonth, config) {
  const url = 'https://api.notion.com/v1/pages';
  const payload = {
    parent: { database_id: config.NOTION_DB_ID },
    properties: {
      "이름": { title: [{ text: { content: dateStr } }] },
      "날짜": { date: { start: dateStr } },
      "Month": { select: { name: String(yearMonth) } }
    }
  };
  const options = {
    method: "post",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).id;
    Logger.log(`createNotionPage 실패 (${dateStr}): ${res.getContentText()}`);
    return null;
  } catch (e) { return null; }
}

// ===== 헬퍼: Notion 페이지 속성 업데이트 =====
function patchNotionPage(pageId, properties, config) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  const options = {
    method: "patch",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify({ properties: properties }),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

// ===== 헬퍼: Notion 페이지 상세 조회 =====
function getNotionPage(pageId, config) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Notion-Version": "2022-06-28"
    },
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText());
    return null;
  } catch (e) { return null; }
}

// ===== 헬퍼: 페이지 블록 목록 조회 =====
function getPageBlocks(pageId, config) {
  const url = `https://api.notion.com/v1/blocks/${pageId}/children?page_size=100`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Notion-Version": "2022-06-28"
    },
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).results || [];
    return [];
  } catch (e) { return []; }
}

// ===== 헬퍼: 이미지 블록 추가 =====
function appendImageBlock(pageId, imageUrl, config) {
  const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
  const options = {
    method: "patch",
    headers: {
      "Authorization": `Bearer ${config.NOTION_TOKEN}`,
      "Content-Type": "application/json",
      "Notion-Version": "2022-06-28"
    },
    payload: JSON.stringify({
      children: [{
        object: "block",
        type: "image",
        image: {
          type: "external",
          external: { url: imageUrl }
        }
      }]
    }),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

// ===== NEIS 급식 데이터 조회 =====
function getMonthlyMeals(yearMonth, config) {
  const url = `https://open.neis.go.kr/hub/mealServiceDietInfo?KEY=${config.NEIS_API_KEY}&Type=json&ATPT_OFCDC_SC_CODE=${config.ATPT_OFCDC_SC_CODE}&SD_SCHUL_CODE=${config.SD_SCHUL_CODE}&MLSV_FROM_YMD=${yearMonth}01&MLSV_TO_YMD=${yearMonth}31`;
  try {
    const res = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    if (!res.mealServiceDietInfo) return null;

    return res.mealServiceDietInfo[1].row.map(r => {
      const menuList = r.DDISH_NM.split('<br/>');
      const cleanedItems = menuList.map(item => {
        return item
          .replace(/\([^)]+\)/g, "")
          .replace(/[@*#]/g, "")
          .replace(/\s+/g, " ")
          .trim();
      }).filter(item => item.length > 0);

      return {
        date: `${r.MLSV_YMD.substring(0,4)}-${r.MLSV_YMD.substring(4,6)}-${r.MLSV_YMD.substring(6,8)}`,
        menu: cleanedItems.join(", ")
      };
    });
  } catch (e) { return null; }
}
