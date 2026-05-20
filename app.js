console.log("app.js version: 20260520d");
const statusLabels = {
  reserved: "預約",
  scheduled: "已排定",
  changed: "改期",
  reschedule: "再次開會",
  withdrawn: "撤回",
  cancelled: "取消",
  finished: "完成",
};

const reportCategoryLabels = {
  mediator: "調解人",
  committee: "調解委員會",
  arbitration: "仲裁人(委員會)",
  transfer: "轉外縣市/中科/加工區",
  laborAssocCity: "台中市勞資關係協會",
  laborAssocCounty: "台中縣勞資關係協會",
  laborEmploymentAssoc: "台中市勞雇關係協會",
  occupationalInjuryAssoc: "職業災害法律權益服務協會",
  bureau: "本局召開",
};

const roomNameMap = {
  調解會議室: "晤談室(一)",
  第一會議室: "晤談室(二)",
  第二會議室: "晤談室(三)",
  晤談1: "晤談室(一)",
  晤談2: "晤談室(二)",
  晤談3: "晤談室(三)",
  晤談4: "晤談室(四)",
  惠中602: "其他",
  線上會議: "其他",
};

const fixedRooms = ["晤談室(一)", "晤談室(二)", "晤談室(三)", "晤談室(四)", "勞資爭議調解會議室"];

const timeSlots = ["09:00", "10:30", "13:30", "15:00"];
const weekdayNames = ["週一", "週二", "週三", "週四", "週五"];
const storageKey = "mediation-scheduler-cases";
const intakeSource = "data/intake-2026-05-15.json";
const applicationBatchSource = "data/application-batch.json";
const supabaseConfig = window.APP_SUPABASE_CONFIG || {};
const supabaseClient =
  window.supabase && supabaseConfig.url && supabaseConfig.anonKey
    ? window.supabase.createClient(supabaseConfig.url, supabaseConfig.anonKey)
    : null;

const sampleCases = [
  {
    id: crypto.randomUUID(),
    sourceCaseNo: "",
    caseNo: "115-H-00041",
    meetingDate: "2026-05-18",
    meetingTime: "09:00",
    room: "晤談室(一)",
    worker: "王○○",
    employer: "甲○公司",
    mediator: "林委員",
    recorder: "李紀錄",
    owner: "陳承辦",
    disputeType: "工資",
    reportCategory: "mediator",
    workerGender: "男",
    workerAge: "35",
    maleCount: "1",
    femaleCount: "0",
    specialCaseType: "無",
    mediationMethodCode: "H - 指派調解人",
    status: "scheduled",
    notes: "電話通知完成",
  },
  {
    id: crypto.randomUUID(),
    sourceCaseNo: "",
    caseNo: "115-H-00052",
    meetingDate: "2026-05-18",
    meetingTime: "10:30",
    room: "晤談室(二)",
    worker: "李○○",
    employer: "乙○企業",
    mediator: "黃委員",
    recorder: "王紀錄",
    owner: "張承辦",
    disputeType: "資遣費",
    reportCategory: "mediator",
    workerGender: "女",
    workerAge: "42",
    maleCount: "0",
    femaleCount: "1",
    specialCaseType: "無",
    mediationMethodCode: "H - 指派調解人",
    status: "changed",
    notes: "原 5/15 改期",
  },
  {
    id: crypto.randomUUID(),
    sourceCaseNo: "",
    caseNo: "115-H-00073",
    meetingDate: "2026-05-19",
    meetingTime: "13:30",
    room: "晤談室(三)",
    worker: "周○○",
    employer: "丙○商行",
    mediator: "吳委員",
    recorder: "協會指派",
    owner: "陳承辦",
    disputeType: "解僱",
    reportCategory: "laborEmploymentAssoc",
    workerGender: "男",
    workerAge: "48",
    maleCount: "1",
    femaleCount: "0",
    specialCaseType: "無",
    mediationMethodCode: "E - 台中市勞雇關係協會",
    status: "scheduled",
    notes: "",
  },
  {
    id: crypto.randomUUID(),
    sourceCaseNo: "",
    caseNo: "115-H-00088",
    meetingDate: "2026-05-20",
    meetingTime: "15:00",
    room: "晤談室(四)",
    worker: "劉○○",
    employer: "丁○公司",
    mediator: "林委員",
    recorder: "李紀錄",
    owner: "許承辦",
    disputeType: "職災",
    reportCategory: "occupationalInjuryAssoc",
    workerGender: "女",
    workerAge: "51",
    maleCount: "0",
    femaleCount: "1",
    specialCaseType: "職業災害",
    mediationMethodCode: "J - 臺中市職業災害法律服務協會",
    status: "withdrawn",
    notes: "申請人撤回",
  },
];

let cases = loadCases();
let intakeCases = [];

let applicationBatchCases = [];
let selectedWeekStart = startOfWeek(new Date());

const elements = {
  weekPicker: document.querySelector("#weekPicker"),
  keyword: document.querySelector("#keyword"),
  statusFilter: document.querySelector("#statusFilter"),
  calendar: document.querySelector("#calendar"),
  caseList: document.querySelector("#caseList"),
  intakeList: document.querySelector("#intakeList"),
  intakeSummary: document.querySelector("#intakeSummary"),
  weekRange: document.querySelector("#weekRange"),
  totalCount: document.querySelector("#totalCount"),
  roomCount: document.querySelector("#roomCount"),
  changedCount: document.querySelector("#changedCount"),
  mediatorCount: document.querySelector("#mediatorCount"),
  dialog: document.querySelector("#caseDialog"),
  form: document.querySelector("#caseForm"),
  dialogTitle: document.querySelector("#dialogTitle"),
  conflictWarning: document.querySelector("#conflictWarning"),
  deleteCase: document.querySelector("#deleteCase"),
};

document.querySelector("#newCase").addEventListener("click", () => openCaseDialog());
initImportExcel();
document.querySelector("#importIntakeBtn").addEventListener("click", () => document.querySelector("#intakeExcelInput").click());
document.querySelector("#intakeExcelInput").addEventListener("change", importIntakeExcel);
document.querySelector("#importAssocBtn").addEventListener("click", () => document.querySelector("#assocExcelInput").click());
document.querySelector("#assocExcelInput").addEventListener("change", importAssocExcel);
document.querySelector("#closeDialog").addEventListener("click", closeDialog);
document.querySelector("#cancelEdit").addEventListener("click", closeDialog);
document.querySelector("#exportCsv").addEventListener("click", exportCsv);
document.querySelector("#exportWeeklyReport").addEventListener("click", exportWeeklyReport);
elements.weekPicker.addEventListener("change", handleWeekChange);
document.querySelector("#prevWeek").addEventListener("click", () => {
  selectedWeekStart = addDays(selectedWeekStart, -7);
  elements.weekPicker.value = dateToWeekValue(selectedWeekStart);
  render();
});
document.querySelector("#nextWeek").addEventListener("click", () => {
  selectedWeekStart = addDays(selectedWeekStart, 7);
  elements.weekPicker.value = dateToWeekValue(selectedWeekStart);
  render();
});
document.querySelector("#todayWeek").addEventListener("click", () => {
  selectedWeekStart = startOfWeek(new Date());
  elements.weekPicker.value = dateToWeekValue(selectedWeekStart);
  render();
});
elements.keyword.addEventListener("input", render);
elements.statusFilter.addEventListener("change", render);
elements.form.addEventListener("submit", saveCase);
elements.deleteCase.addEventListener("click", deleteCurrentCase);

["meetingDate", "meetingTime", "room", "mediator"].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener("change", updateConflictWarning);
});
document.querySelector("#room").addEventListener("change", toggleOtherRoom);

initialize();

function initialize() {
  elements.weekPicker.value = dateToWeekValue(selectedWeekStart);
  loadRemoteCases();
  renderIntakeList();
  render();
}

async function loadRemoteCases() {
  if (!supabaseClient) return;
  const pageSize = 1000;
  let allData = [];
  let from = 0;

  while (true) {
    const { data, error } = await supabaseClient
      .from("mediation_schedules")
      .select("*")
      .order("meeting_date", { ascending: true })
      .order("meeting_time", { ascending: true })
      .range(from, from + pageSize - 1);

    if (error) {
      console.warn("Supabase load failed; using local data", error);
      return;
    }

    allData = allData.concat(data);
    if (data.length < pageSize) break;
    from += pageSize;
  }

  cases = allData.map(fromDbCase);
  persist();
  render();
}

function loadCases() {
  const saved = localStorage.getItem(storageKey);
  if (!saved) {
    localStorage.setItem(storageKey, JSON.stringify(sampleCases));
    return sampleCases;
  }
  try {
    const parsed = JSON.parse(saved);
    const normalized = parsed.map((item) => ({
      ...item,
      room: normalizeRoom(item.room).room,
      roomOther: normalizeRoom(item.room).roomOther || item.roomOther || "",
    }));
    localStorage.setItem(storageKey, JSON.stringify(normalized));
    return normalized;
  } catch {
    return sampleCases;
  }
}

async function importIntakeExcel(e) {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    fixSheetRef(ws);

    // header:1 取二維陣列
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    if (!rows.length) {
      alert("Excel 讀取失敗：無法讀到任何列。");
      e.target.value = "";
      return;
    }

    // 從第一列建立「欄位名 → 索引」對應
    const headerRow = rows[0];
    const headerMap = {};
    headerRow.forEach((h, i) => { if (h) headerMap[String(h).trim()] = i; });

    const col = name => headerMap[name] ?? -1;
    const get = (row, name) => col(name) >= 0 ? String(row[col(name)] ?? "").trim() : "";
    const getNum = (row, name) => col(name) >= 0 ? Number(row[col(name)]) || 0 : 0;

    // 格式偵測
    const isSystemExport = "錄案日期" in headerMap;

    if (!isSystemExport) {
      // 批次輸入表（無固定標題，用索引）
      const dataRows = rows.slice(4).filter(r => String(r[3] || "").trim());
      if (!dataRows.length) {
        alert("找不到資料列（批次輸入表格式）。");
        e.target.value = "";
        return;
      }
      intakeCases = dataRows.map((r, idx) => {
        const mediMode = String(r[18] || "").trim();
        const dispute  = String(r[32] || "").trim();
        return {
          sourceCaseNo: `batch-${String(r[1] || idx)}`,
          customCaseNo: String(r[1] || "").trim(),
          worker: String(r[3] || "").trim(),
          employer: String(r[10] || "").trim(),
          owner: String(r[2] || "").trim(),
          mediationMethodCode: mediMode,
          disputeType: guessDisputeType(dispute),
          claim: dispute,
          specialCaseType: "無",
          receivedDate: String(r[0] || "").trim(),
          acceptanceMethod: "",
          maleCount: 0, femaleCount: 0, workerGender: "", workerAge: "",
          reportCategory: reportCategoryFromMethod(mediMode),
        };
      }).filter(item => item.worker || item.employer);
    } else {
      // 系統匯出格式（用欄位名稱）
      const dataRows = rows.slice(1).filter(r =>
        String(r[col("自編案號")] || "").trim() || String(r[col("勞方")] || "").trim()
      );
      if (!dataRows.length) {
        alert("找不到資料列（系統匯出格式）。\n偵測到欄位：" + Object.keys(headerMap).slice(0,5).join("、"));
        e.target.value = "";
        return;
      }
      intakeCases = dataRows.map((r, idx) => {
        const maleCount   = getNum(r, "男性人數");
        const femaleCount = getNum(r, "女性人數");
        const age         = getNum(r, "年齡");
        const mediMode    = get(r, "案件類別");
        const dispute     = get(r, "主要爭議");
        const special     = get(r, "特殊案件類型");
        return {
          sourceCaseNo:     get(r, "自編案號") || `import-${idx}`,
          customCaseNo:     get(r, "自編案號"),
          worker:           get(r, "勞方"),
          employer:         get(r, "資方"),
          owner:            get(r, "承辦人"),
          mediationMethodCode: mediMode,
          disputeType:      guessDisputeType(dispute),
          claim:            dispute,
          specialCaseType:  special && special !== "無" ? special : "無",
          receivedDate:     get(r, "錄案日期"),
          acceptanceMethod: get(r, "調解地點"),
          maleCount,
          femaleCount,
          workerGender:     maleCount > 0 ? "男" : (femaleCount > 0 ? "女" : ""),
          workerAge:        age > 0 ? String(age) : "",
          reportCategory:   reportCategoryFromMethod(mediMode),
        };
      }).filter(item => item.worker || item.employer);
    }

    if (!intakeCases.length) {
      alert("匯入後沒有有效資料（勞方或資方欄位皆為空）。");
      e.target.value = "";
      return;
    }

    applicationBatchCases = [];
    renderIntakeList();
    alert(`匯入完成，共 ${intakeCases.length} 筆待排案件。`);
  } catch (err) {
    alert("Excel 讀取失敗：" + err.message);
  }
  e.target.value = "";
}

async function importAssocExcel(e) {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    fixSheetRef(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    // 第1列標題、第2列說明，資料從第3列開始
    if (rows.length < 3) {
      alert("找不到資料（資料需從第3列開始）。");
      e.target.value = "";
      return;
    }

    // 建立 headerMap
    const headerMap = {};
    rows[0].forEach((h, i) => { if (h) headerMap[String(h).trim()] = i; });
    const get = (row, name) => (name in headerMap) ? String(row[headerMap[name]] ?? "").trim() : "";

    // 欄位對應（依範例格式）
    // 開會日期|開會時間|地點|其他地點|勞方|資方|主要爭議|承辦人|調解人|紀錄|調解方式|案號|備註
    const dataRows = rows.slice(2).filter(r => get(r, "勞方") || get(r, "資方"));

    if (!dataRows.length) {
      alert("找不到有效資料列（勞方或資方須有值）。");
      e.target.value = "";
      return;
    }

    // 日期格式轉換（支援 2025-10-08 及 115/10/08）
    const parseAssocDate = (val) => {
      if (!val && val !== 0) return "";
      // Excel 日期序號（數字）→ 轉換為西元日期
      if (typeof val === "number" || (!isNaN(Number(val)) && String(val).length <= 5)) {
        const num = Number(val);
        if (num > 40000 && num < 60000) {
          const base = new Date(1899, 11, 30);
          base.setDate(base.getDate() + num);
          return base.toISOString().slice(0, 10);
        }
      }
      const s = String(val).trim();
      if (!s) return "";
      // 西元 2025-10-08 或 2025/10/08
      if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) return s.replace(/\//g, "-");
      // 民國 115/10/08 或 115-10-08
      const m = s.match(/^(\d{2,3})[-/](\d{1,2})[-/](\d{1,2})$/);
      if (m) {
        const y = parseInt(m[1]) + 1911;
        return `${y}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
      }
      return s;
    };

    // 時間格式轉換（0900 → 09:00）
    const parseAssocTime = (val) => {
      const s = String(val || "").trim().replace(":", "").padStart(4, "0");
      return /^\d{4}$/.test(s) ? `${s.slice(0,2)}:${s.slice(2)}` : "09:00";
    };

    // 地點對應
    const roomMapAssoc = {
      "晤談室(一)": "晤談室(一)", "晤談室(二)": "晤談室(二)",
      "晤談室(三)": "晤談室(三)", "晤談室(四)": "晤談室(四)",
      "勞資爭議調解會議室": "勞資爭議調解會議室",
    };

    const newCases = dataRows.map(r => {
      const rawRoom  = get(r, "地點");
      const room     = roomMapAssoc[rawRoom] || "其他";
      const roomOther = room === "其他" ? (get(r, "其他地點") || rawRoom) : "";
      const mediMode = get(r, "調解方式").toUpperCase();
      const dispute  = get(r, "主要爭議");
      const rawDateVal = headerMap["開會日期"] >= 0 ? r[headerMap["開會日期"]] : "";
      const dateVal  = parseAssocDate(rawDateVal);
      const today    = toDateInputValue(new Date());
      return {
        id:                 crypto.randomUUID(),
        sourceCaseNo:       "",
        caseNo:             get(r, "案號"),
        meetingDate:        dateVal,
        meetingTime:        parseAssocTime(get(r, "開會時間")),
        room,
        roomOther,
        worker:             get(r, "勞方"),
        employer:           get(r, "資方"),
        mediator:           get(r, "調解人"),
        recorder:           get(r, "紀錄"),
        owner:              get(r, "承辦人"),
        disputeType:        guessDisputeType(dispute) || dispute || "其他",
        reportCategory:     mediMode || "A",
        workerGender:       "",
        workerAge:          "",
        maleCount:          0,
        femaleCount:        0,
        specialCaseType:    "無",
        mediationMethodCode: mediMode,
        status:             dateVal && dateVal < today ? "finished" : "scheduled",
        notes:              get(r, "備註"),
      };
    });

    if (!newCases.length) {
      alert("沒有可匯入的資料。");
      e.target.value = "";
      return;
    }

    // 批量寫入 Supabase
    const toDb = (item) => ({
      id:                   item.id,
      source_case_no:       item.sourceCaseNo || null,
      case_no:              item.caseNo || null,
      meeting_date:         item.meetingDate,
      meeting_time:         item.meetingTime,
      room:                 item.room,
      room_other:           item.roomOther || null,
      worker:               item.worker || null,
      employer:             item.employer || null,
      mediator:             item.mediator || null,
      recorder:             item.recorder || null,
      owner:                item.owner || null,
      dispute_type:         item.disputeType || null,
      report_category:      item.reportCategory || null,
      worker_gender:        null,
      worker_age:           null,
      male_count:           0,
      female_count:         0,
      special_case_type:    item.specialCaseType || null,
      mediation_method_code: item.mediationMethodCode || null,
      status:               item.status,
      notes:                item.notes || null,
    });

    const batchSize = 100;
    let total = 0, errors = 0;
    for (let i = 0; i < newCases.length; i += batchSize) {
      const batch = newCases.slice(i, i + batchSize).map(toDb);
      const { error } = await supabaseClient
        .from("mediation_schedules")
        .upsert(batch, { onConflict: "id" });
      if (error) {
        console.error("批次失敗:", error.message, error.details);
        errors++;
      } else {
        total += batch.length;
      }
    }

    // 重新從 Supabase 讀取確保資料同步
    await loadRemoteCases();

    alert(`協會案件匯入完成：${total} 筆成功${errors ? "，" + errors + " 批失敗" : ""}。`);
  } catch (err) {
    alert("Excel 讀取失敗：" + err.message);
  }
  e.target.value = "";
}

async function loadIntakeCases() {
  const [intake, batch] = await Promise.all([loadJson(intakeSource), loadJson(applicationBatchSource)]);
  applicationBatchCases = batch.map(normalizeBatchItem);
  intakeCases = mergeIntakeSources(intake, applicationBatchCases);
  renderIntakeList();
}

async function loadJson(source) {
  try {
    const response = await fetch(source, { cache: "no-store" });
    if (!response.ok) throw new Error(`Cannot load ${source}`);
    return await response.json();
  } catch {
    return [];
  }
}

function persist() {
  localStorage.setItem(storageKey, JSON.stringify(cases));
}

function render() {
  const weekDates = Array.from({ length: 5 }, (_, index) => addDays(selectedWeekStart, index));
  const filtered = getFilteredCases();
  elements.weekRange.textContent = `${formatDate(weekDates[0])} 至 ${formatDate(weekDates[4])}`;
  renderStats(filtered);
  renderIntakeList();
  renderCalendar(weekDates, filtered);
  renderCaseList(filtered);
}

function renderStats(filtered) {
  elements.totalCount.textContent = filtered.length;
  elements.roomCount.textContent = new Set(filtered.map((item) => item.room)).size;
  elements.changedCount.textContent = filtered.filter((item) =>
    ["changed", "withdrawn", "cancelled"].includes(item.status)
  ).length;
  elements.mediatorCount.textContent = new Set(filtered.map((item) => item.mediator)).size;
}

function renderIntakeList() {
  const scheduledSourceNos = new Set(cases.map((item) => item.sourceCaseNo).filter(Boolean));
  const waiting = intakeCases.filter((item) => !scheduledSourceNos.has(item.sourceCaseNo));
  if (intakeCases.length === 0) {
    elements.intakeSummary.textContent = "尚未匯入收案資料，請點「匯入收案 Excel」";
  } else {
    elements.intakeSummary.textContent = `共 ${intakeCases.length} 筆，尚待排程 ${waiting.length} 筆`;
  }
  elements.intakeList.innerHTML = "";

  if (waiting.length === 0) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = intakeCases.length ? "本批收案都已建立排程" : "請點「匯入收案 Excel」載入當天收案資料";
    elements.intakeList.append(empty);
    return;
  }

  waiting.forEach((item) => {
    const card = document.createElement("article");
    card.className = "intake-card";
    card.innerHTML = `
      <strong>${escapeHtml(item.customCaseNo || item.sourceCaseNo)}｜${escapeHtml(item.worker)}</strong>
      <span>${escapeHtml(item.employer)}</span>
      <span>收件：${escapeHtml(item.receivedDate)}　方式：${escapeHtml(item.acceptanceMethod)}</span>
      <span>${escapeHtml(item.claim || item.disputeType || "未填請求事項")}</span>
      <span>${escapeHtml(item.workerGender || "性別未填")}　${escapeHtml(item.workerAge ? `${item.workerAge}歲` : "年齡未填")}　${escapeHtml(reportCategoryLabels[item.reportCategory] || "")}</span>
    `;
    const button = document.createElement("button");
    button.className = "primary-button";
    button.type = "button";
    button.textContent = "安排";
    button.addEventListener("click", () => openCaseDialog(null, item));
    const dismissBtn = document.createElement("button");
    dismissBtn.className = "danger-button";
    dismissBtn.type = "button";
    dismissBtn.style.minHeight = "32px";
    dismissBtn.style.fontSize = "13px";
    dismissBtn.textContent = "不排程";
    dismissBtn.addEventListener("click", () => {
      intakeCases = intakeCases.filter(c => c.sourceCaseNo !== item.sourceCaseNo);
      renderIntakeList();
    });
    const btnRow = document.createElement("div");
    btnRow.style.cssText = "display:flex;gap:8px;";
    btnRow.append(button, dismissBtn);
    card.append(btnRow);
    elements.intakeList.append(card);
  });
}

function renderCalendar(weekDates, filtered) {
  elements.calendar.innerHTML = "";
  elements.calendar.append(createCell("", "day-head"));
  weekDates.forEach((date, index) => {
    elements.calendar.append(createCell(`${weekdayNames[index]} ${formatMonthDay(date)}`, "day-head"));
  });

  const roomOrder = ["晤談室(一)", "晤談室(四)", "晤談室(二)", "晤談室(三)", "勞資爭議調解會議室"];
  const roomRank = (item) => {
    const r = displayRoom(item);
    const idx = roomOrder.indexOf(r);
    return idx >= 0 ? idx : roomOrder.length;
  };

  // 建立衝突 key set：同日期+時段+地點出現 2 筆以上視為衝突
  const slotRoomCount = {};
  filtered.forEach(item => {
    const key = `${item.meetingDate}|${item.meetingTime}|${displayRoom(item)}`;
    slotRoomCount[key] = (slotRoomCount[key] || 0) + 1;
  });
  const conflictKeys = new Set(Object.keys(slotRoomCount).filter(k => slotRoomCount[k] > 1));

  timeSlots.forEach((slot) => {
    elements.calendar.append(createCell(slot, "time-cell"));
    weekDates.forEach((date) => {
      const cell = createCell("", "calendar-cell");
      const dateKey = toDateInputValue(date);
      filtered
        .filter((item) => item.meetingDate === dateKey && item.meetingTime === slot)
        .sort((a, b) => roomRank(a) - roomRank(b) || a.caseNo.localeCompare(b.caseNo))
        .forEach((item) => {
          const key = `${item.meetingDate}|${item.meetingTime}|${displayRoom(item)}`;
          cell.append(createCaseCard(item, conflictKeys.has(key)));
        });
      elements.calendar.append(cell);
    });
  });
}

function renderCaseList(filtered) {
  elements.caseList.innerHTML = "";
  if (filtered.length === 0) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = "目前沒有符合條件的案件";
    elements.caseList.append(empty);
    return;
  }

  filtered
    .slice()
    .sort((a, b) => `${a.meetingDate} ${a.meetingTime}`.localeCompare(`${b.meetingDate} ${b.meetingTime}`))
    .forEach((item) => {
      const button = document.createElement("button");
      button.className = "list-item";
      button.type = "button";
      button.innerHTML = `
        <strong>${escapeHtml(item.caseNo)}｜${escapeHtml(statusLabels[item.status])}</strong>
        <span>${escapeHtml(item.meetingDate)} ${escapeHtml(item.meetingTime)}　${escapeHtml(item.worker)} / ${escapeHtml(item.employer)}</span>
        <span>${escapeHtml(displayRoom(item))}　${escapeHtml(getMediatorDisplay(item))}　${escapeHtml(item.owner)}</span>
      `;
      button.addEventListener("click", () => openCaseDialog(item.id));
      elements.caseList.append(button);
    });
}

function createCaseCard(item, isConflict = false) {
  const button = document.createElement("button");
  const isSpecial = item.specialCaseType && item.specialCaseType !== "無" && item.specialCaseType !== "";
  const isReserved = item.status === "reserved";
  button.className = `case-card ${item.status}${isSpecial ? " special" : ""}${isConflict ? " conflict" : ""}`;
  button.type = "button";
  const specialTag = isSpecial ? `<small style="color:#7c3aed;font-weight:700;">⚠ ${escapeHtml(item.specialCaseType)}</small>` : "";
  const conflictTag = isConflict ? `<small style="color:#dc2626;font-weight:700;">⚠ 衝期</small>` : "";
  const mainLine = isReserved
    ? `<strong>【預約】${escapeHtml(getMediatorDisplay(item))}</strong>`
    : `<strong>${escapeHtml(item.caseNo)} ${escapeHtml(statusLabels[item.status])}</strong>`;
  const workerLine = isReserved
    ? `<small style="color:#6b7280;">待補：案號、勞資雙方</small>`
    : `<small>${escapeHtml(item.worker)} / ${escapeHtml(item.employer)}</small>`;
  button.innerHTML = `
    ${mainLine}
    ${workerLine}
    <small>${escapeHtml(displayRoom(item))}｜${escapeHtml(getMediatorDisplay(item))}</small>
    ${specialTag}${conflictTag}
  `;
  button.addEventListener("click", () => openCaseDialog(item.id));
  return button;
}

function createCell(text, className) {
  const cell = document.createElement("div");
  cell.className = className;
  cell.textContent = text;
  return cell;
}

function getFilteredCases() {
  const weekEnd = addDays(selectedWeekStart, 4);
  const keyword = elements.keyword.value.trim().toLowerCase();
  const status = elements.statusFilter.value;

  return cases.filter((item) => {
    const date = parseDate(item.meetingDate);
    const inWeek = date >= selectedWeekStart && date <= weekEnd;
    const inStatus = status === "all" || item.status === status;
    const text = [
      item.sourceCaseNo,
      item.caseNo,
      item.worker,
      item.employer,
      item.mediator,
      item.recorder,
      item.owner,
      item.room,
      item.notes,
    ]
      .join(" ")
      .toLowerCase();
    const inKeyword = !keyword || text.includes(keyword);
    return inWeek && inStatus && inKeyword;
  });
}

function openCaseDialog(id, intakeItem) {
  elements.form.reset();
  elements.conflictWarning.hidden = true;
  const item = cases.find((entry) => entry.id === id);
  elements.dialogTitle.textContent = item ? "編輯排程" : "新增排程";
  elements.deleteCase.hidden = !item;
  document.querySelector("#caseId").value = item?.id ?? "";
  document.querySelector("#sourceCaseNo").value = item?.sourceCaseNo ?? intakeItem?.sourceCaseNo ?? "";
  document.querySelector("#caseNo").value = item?.caseNo ?? intakeItem?.customCaseNo ?? "";
  document.querySelector("#meetingDate").value = item?.meetingDate ?? toDateInputValue(selectedWeekStart);
  document.querySelector("#meetingTime").value = item?.meetingTime ?? "09:00";
  const roomValue = normalizeRoom(item?.room);
  document.querySelector("#room").value = roomValue.room ?? "晤談室(一)";
  document.querySelector("#roomOther").value = item?.roomOther ?? roomValue.roomOther ?? "";
  document.querySelector("#worker").value = item?.worker ?? intakeItem?.worker ?? "";
  document.querySelector("#employer").value = item?.employer ?? intakeItem?.employer ?? "";
  document.querySelector("#mediator").value = item?.mediator ?? "";
  // 調解委員會三欄
  if (item && isCommittee(item)) {
    const obj = parseMediatorField(item.mediator);
    document.getElementById("mediatorChair").value = obj.chair || "";
    document.getElementById("mediatorLabor").value = obj.labor || "";
    document.getElementById("mediatorMgmt").value = obj.mgmt || "";
  } else {
    document.getElementById("mediatorChair").value = "";
    document.getElementById("mediatorLabor").value = "";
    document.getElementById("mediatorMgmt").value = "";
  }
  onReportCategoryChange();
  document.querySelector("#recorder").value = item?.recorder ?? "";
  document.querySelector("#owner").value = item?.owner ?? "";
  document.querySelector("#owner").value = item?.owner ?? intakeItem?.owner ?? "";
  document.querySelector("#disputeType").value = item?.disputeType ?? intakeItem?.disputeType ?? guessDisputeType(intakeItem?.claim) ?? "工資";
  document.querySelector("#reportCategory").value = item?.reportCategory ?? intakeItem?.mediationMethodCode?.split(" - ")[0] ?? intakeItem?.reportCategory ?? "";
  document.querySelector("#workerGender").value = item?.workerGender ?? intakeItem?.workerGender ?? "";
  document.querySelector("#workerAge").value = item?.workerAge ?? intakeItem?.workerAge ?? "";
  document.querySelector("#maleCount").value = item?.maleCount ?? intakeItem?.maleCount ?? inferGenderCount(item?.workerGender ?? intakeItem?.workerGender, "男");
  document.querySelector("#femaleCount").value = item?.femaleCount ?? intakeItem?.femaleCount ?? inferGenderCount(item?.workerGender ?? intakeItem?.workerGender, "女");
  setSpecialCaseCheckboxes(item?.specialCaseType ?? intakeItem?.specialCaseType ?? "無");
  document.querySelector("#mediationMethodCode").value = item?.mediationMethodCode ?? intakeItem?.mediationMethodCode ?? document.querySelector("#reportCategory").value ?? "";
  document.querySelector("#notes").value = item?.notes ?? buildIntakeNotes(intakeItem) ?? "";

  // 自動判斷狀態：編輯時保留原狀態；新增時依日期自動判斷
  const today = toDateInputValue(new Date());
  const meetingDateVal = item?.meetingDate ?? toDateInputValue(selectedWeekStart);
  const autoStatus = item ? item.status : (meetingDateVal < today ? "finished" : "scheduled");
  document.querySelector("#caseStatus").value = autoStatus;
  onStatusChange();
  // 重設子表單
  const panel = document.getElementById("secondMeetingPanel");
  panel.hidden = true;
  document.getElementById("meetingDate2").value = "";
  document.getElementById("meetingTime2").value = "";
  document.getElementById("room2").value = "晤談室(一)";
  document.getElementById("roomOther2").value = "";
  document.getElementById("mediator2").value = "";
  document.getElementById("roomOther2Label").hidden = true;
  document.getElementById("conflictWarning2").hidden = true;

  toggleOtherRoom();
  updateConflictWarning();
  elements.dialog.showModal();
}

// ── Excel 匯入 ──
function initImportExcel() {
  const input = document.getElementById("importExcel");
  if (!input) return;
  input.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById("importFileName").textContent = file.name;
    try {
      const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];

      // 修正 xlsx.js !ref 範圍 bug
      fixSheetRef(ws);

      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (rows.length < 2) { alert("找不到資料列。"); return; }

      // 建立標題對應
      const headerMap = {};
      rows[0].forEach((h, i) => { if (h) headerMap[String(h).trim()] = i; });
      const get = (row, name) => (name in headerMap) ? String(row[headerMap[name]] ?? "").trim() : "";
      const getNum = (row, name) => (name in headerMap) ? Number(row[headerMap[name]]) || 0 : 0;

      const isSystemExport = "錄案日期" in headerMap;
      const dataRows = isSystemExport
        ? rows.slice(1).filter(r => get(r, "自編案號") || get(r, "勞方"))
        : rows.slice(4).filter(r => String(r[3] || "").trim());

      if (!dataRows.length) { alert("找不到資料列。"); return; }

      // 填入第1筆到表單
      const r = dataRows[0];
      if (isSystemExport) {
        const caseNo = get(r, "自編案號");
        const worker = get(r, "勞方");
        const employer = get(r, "資方");
        const owner = get(r, "承辦人");
        const mediMode = get(r, "案件類別");
        const dispute = get(r, "主要爭議");
        const special = get(r, "特殊案件類型");
        const meetingDate = get(r, "會議日期");
        const maleCount = getNum(r, "男性人數");
        const femaleCount = getNum(r, "女性人數");

        if (caseNo) document.getElementById("caseNo").value = caseNo;
        if (worker) document.getElementById("worker").value = worker;
        if (employer) document.getElementById("employer").value = employer;
        if (owner) document.getElementById("owner").value = owner;
        if (dispute) {
          const sel = document.getElementById("disputeType");
          if ([...sel.options].some(o => o.text === dispute || o.value === dispute)) sel.value = dispute;
        }
        setSpecialCaseCheckboxes(special && special !== "無" ? "是" : "無");
        if (mediMode) {
          const code = mediMode.split(" - ")[0].trim().toUpperCase();
          const sel = document.getElementById("reportCategory");
          if ([...sel.options].some(o => o.value === code)) sel.value = code;
          document.getElementById("mediationMethodCode").value = mediMode;
        }
        if (meetingDate && meetingDate.includes("/")) {
          const parts = meetingDate.split("/");
          if (parts.length === 3) {
            const year = parseInt(parts[0]) + 1911;
            document.getElementById("meetingDate").value = year + "-" + parts[1].padStart(2,"0") + "-" + parts[2].padStart(2,"0");
          }
        }
        // 性別 / 人數
        document.getElementById("workerGender").value = maleCount > 0 ? "男" : (femaleCount > 0 ? "女" : "");
        document.getElementById("maleCount").value = maleCount;
        document.getElementById("femaleCount").value = femaleCount;
      } else {
        // 批次輸入表
        if (r[3]) document.getElementById("worker").value = String(r[3]).trim();
        if (r[10]) document.getElementById("employer").value = String(r[10]).trim();
        if (r[2]) document.getElementById("owner").value = String(r[2]).trim();
        const mediMode = String(r[18] || "").trim();
        if (mediMode) {
          const code = mediMode.split(" - ")[0].trim().toUpperCase();
          const sel = document.getElementById("reportCategory");
          if ([...sel.options].some(o => o.value === code)) sel.value = code;
          document.getElementById("mediationMethodCode").value = mediMode;
        }
      }

      if (dataRows.length > 1) {
        alert("共有 " + dataRows.length + " 筆資料，已填入第 1 筆。");
      }
    } catch (err) {
      alert("Excel 讀取失敗：" + err.message);
    }
  });
}

function syncSpecialCaseType() {
  const yes = document.getElementById("special_yes");
  document.getElementById("specialCaseType").value = (yes && yes.checked) ? "是" : "無";
}

function setSpecialCaseCheckboxes(value) {
  const isSpecial = value && value !== "無" && value !== "否" && value !== "";
  const yes = document.getElementById("special_yes");
  const no = document.getElementById("special_no");
  if (yes) yes.checked = isSpecial;
  if (no) no.checked = !isSpecial;
  syncSpecialCaseType();
}

function onStatusChange() {
  const status = document.querySelector("#caseStatus").value;
  const isReserved = status === "reserved";
  const panel = document.getElementById("secondMeetingPanel");
  const title = document.getElementById("secondMeetingTitle");
  const show = status === "changed" || status === "reschedule";
  panel.hidden = !show;
  if (show) {
    title.textContent = status === "changed" ? "改期後的新時間" : "再次開會時間";
  }

  // 預約模式：勞資雙方、承辦人、調解人不必填
  document.getElementById("worker").required   = !isReserved;
  document.getElementById("employer").required = !isReserved;
  document.getElementById("owner").required    = !isReserved;
  // 調解人欄位必填跟著調解方式走，預約時全部不必填
  if (isReserved) {
    document.getElementById("mediator").required      = false;
    document.getElementById("mediatorChair").required = false;
    document.getElementById("mediatorLabor").required = false;
    document.getElementById("mediatorMgmt").required  = false;
    // 預約時顯示單一調解人欄即可
    document.getElementById("mediatorSingleLabel").style.display = "";
    document.getElementById("mediatorChairLabel").style.display  = "none";
    document.getElementById("mediatorLaborLabel").style.display  = "none";
    document.getElementById("mediatorMgmtLabel").style.display   = "none";
  } else {
    onReportCategoryChange();
  }

  updateConflictWarning2();
}

function updateConflictWarning2() {
  const panel = document.getElementById("secondMeetingPanel");
  if (panel.hidden) return;
  const id = document.querySelector("#caseId").value;
  const meetingDate2 = document.getElementById("meetingDate2").value;
  const meetingTime2 = document.getElementById("meetingTime2").value;
  const room2val = document.getElementById("room2").value;
  const roomOther2 = document.getElementById("roomOther2").value.trim();
  const roomForConflict2 = room2val === "其他" ? roomOther2 || room2val : room2val;
  const mediator = document.querySelector("#mediator").value.trim();
  const warn = document.getElementById("conflictWarning2");
  if (!meetingDate2 || !meetingTime2) { warn.hidden = true; return; }
  const conflicts = cases.filter((item) => {
    if (item.id === id) return false;
    if (item.meetingDate !== meetingDate2 || item.meetingTime !== meetingTime2) return false;
    return displayRoom(item) === roomForConflict2 || (mediator && item.mediator === mediator);
  });
  warn.hidden = !conflicts.length;
  if (conflicts.length) warn.textContent = `提醒：新時間已有 ${conflicts.length} 筆可能衝突的排程，請確認。`;
}


function closeDialog() {
  elements.dialog.close();
}

async function saveCase(event) {
  event.preventDefault();
  const id = document.querySelector("#caseId").value || crypto.randomUUID();
  const next = {
    id,
    sourceCaseNo: document.querySelector("#sourceCaseNo").value.trim(),
    caseNo: document.querySelector("#caseNo").value.trim(),
    meetingDate: document.querySelector("#meetingDate").value,
    meetingTime: document.querySelector("#meetingTime").value,
    room: document.querySelector("#room").value,
    roomOther: document.querySelector("#room").value === "其他" ? document.querySelector("#roomOther").value.trim() : "",
    worker: document.querySelector("#worker").value.trim(),
    employer: document.querySelector("#employer").value.trim(),
    mediator: (() => {
      if (document.querySelector("#reportCategory").value === "F") {
        const chair = document.getElementById("mediatorChair").value.trim();
        const labor = document.getElementById("mediatorLabor").value.trim();
        const mgmt  = document.getElementById("mediatorMgmt").value.trim();
        return JSON.stringify({ chair, labor, mgmt });
      }
      return document.querySelector("#mediator").value.trim();
    })(),
    recorder: document.querySelector("#recorder").value.trim(),
    owner: document.querySelector("#owner").value.trim(),
    disputeType: document.querySelector("#disputeType").value,
    reportCategory: document.querySelector("#reportCategory").value,
    workerGender: document.querySelector("#workerGender").value,
    workerAge: document.querySelector("#workerAge").value,
    maleCount: document.querySelector("#maleCount").value,
    femaleCount: document.querySelector("#femaleCount").value,
    specialCaseType: document.querySelector("#specialCaseType").value.trim(),
    mediationMethodCode: document.querySelector("#mediationMethodCode").value.trim(),
    status: document.querySelector("#caseStatus").value,
    notes: document.querySelector("#notes").value.trim(),
  };

  const index = cases.findIndex((item) => item.id === id);
  if (index >= 0) {
    cases[index] = next;
  } else {
    cases.push(next);
  }
  await saveRemoteCase(next);

  // 若有第二次開會資料，自動新增一筆
  const status = next.status;
  if (status === "changed" || status === "reschedule") {
    const date2 = document.getElementById("meetingDate2").value;
    const time2 = document.getElementById("meetingTime2").value;
    if (date2 && time2) {
      const room2val = document.getElementById("room2").value;
      const roomOther2 = document.getElementById("roomOther2").value.trim();
      const mediator2 = document.getElementById("mediator2").value.trim();
      const next2 = {
        ...next,
        id: crypto.randomUUID(),
        meetingDate: date2,
        meetingTime: time2,
        room: room2val,
        roomOther: room2val === "其他" ? roomOther2 : "",
        mediator: mediator2 || next.mediator,
        status: "scheduled",
        notes: (next.notes ? next.notes + "　" : "") + (status === "changed" ? "【改期】" : "【再次開會】"),
      };
      cases.push(next2);
      await saveRemoteCase(next2);
    }
  }

  selectedWeekStart = startOfWeek(parseDate(next.meetingDate));
  elements.weekPicker.value = dateToWeekValue(selectedWeekStart);
  persist();
  closeDialog();
  render();
}

async function deleteCurrentCase() {
  const id = document.querySelector("#caseId").value;
  cases = cases.filter((item) => item.id !== id);
  await deleteRemoteCase(id);
  persist();
  closeDialog();
  render();
}

function updateConflictWarning() {
  const id = document.querySelector("#caseId").value;
  const meetingDate = document.querySelector("#meetingDate").value;
  const meetingTime = document.querySelector("#meetingTime").value;
  const room = document.querySelector("#room").value;
  const roomOther = document.querySelector("#roomOther").value.trim();
  const roomForConflict = room === "其他" ? roomOther || room : room;
  const mediator = document.querySelector("#mediator").value.trim();

  const conflicts = cases.filter((item) => {
    if (item.id === id) return false;
    if (item.meetingDate !== meetingDate || item.meetingTime !== meetingTime) return false;
    return displayRoom(item) === roomForConflict || (mediator && item.mediator === mediator);
  });

  if (!conflicts.length) {
    elements.conflictWarning.hidden = true;
    return;
  }

  elements.conflictWarning.hidden = false;
  elements.conflictWarning.textContent = `提醒：此時段已有 ${conflicts.length} 筆可能衝突的排程，請確認會議室或調解人。`;
}

function exportCsv() {
  const rows = [
    ["來源案件編號", "案號", "日期", "時間", "會議室", "勞方", "資方", "調解人", "紀錄", "承辦人", "爭議類型", "週報分類", "勞方性別", "勞方年齡", "男性人數", "女性人數", "特殊案件類型", "調解方式代碼", "狀態", "備註"],
    ...getFilteredCases().map((item) => [
      item.sourceCaseNo,
      item.caseNo,
      item.meetingDate,
      item.meetingTime,
      displayRoom(item),
      item.worker,
      item.employer,
      getMediatorDisplay(item),
      item.recorder,
      item.owner,
      item.disputeType,
      reportCategoryLabels[item.reportCategory] ?? item.reportCategory,
      item.workerAge,
      item.maleCount,
      item.femaleCount,
      item.specialCaseType,
      item.mediationMethodCode,
      statusLabels[item.status],
      item.notes,
    ]),
  ];
  const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `調解會議排程_${toDateInputValue(selectedWeekStart)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function exportWeeklyReport() {
  const rows = getFilteredCases()
    .slice()
    .sort((a, b) => `${a.meetingDate} ${a.meetingTime} ${a.room}`.localeCompare(`${b.meetingDate} ${b.meetingTime} ${b.room}`));
  const weekEnd = addDays(selectedWeekStart, 4);
  const reportTitle = `${toMinguoDate(selectedWeekStart)}-${toMinguoDate(weekEnd)}`;
  const dispatcher = "";
  const counts = buildWeeklyCounts(rows);
  const detailRows = rows.map((item) => [
    formatDate(parseDate(item.meetingDate)),
    getWeekdayName(parseDate(item.meetingDate)),
    displayRoom(item),
    item.meetingTime.replace(":", ""),
    ["withdrawn", "cancelled"].includes(item.status) ? statusLabels[item.status] : item.worker,
    ["withdrawn", "cancelled"].includes(item.status) ? "" : item.employer,
    ["withdrawn", "cancelled"].includes(item.status) ? "" : item.disputeType,
    item.owner,
    getMediatorDisplay(item),
    item.recorder,
    item.notes,
  ]);

  const detailTable = tableHtml([
    ["開會日期", "星期", "地點", "開會時間", "勞方", "資方", "主要爭議", "承辦人", "調解人", "紀錄", "備註"],
    ...detailRows,
  ]);
  const countTable = tableHtml([
    ["期間(週)", "調解人", "調解委員會", "仲裁人(委員會)", "轉外縣市/中科/加工區", "台中市勞資關係協會", "台中縣勞資關係協會", "台中市勞雇關係協會", "職災法律權益服務協會", "本局召開案件數", "男", "女", "中高齡"],
    [`${formatMonthDay(selectedWeekStart)}~${formatMonthDay(weekEnd)}(預計召開)`, counts.mediator, counts.committee, counts.arbitration, counts.transfer, counts.laborAssocCity, counts.laborAssocCounty, counts.laborEmploymentAssoc, counts.occupationalInjuryAssoc, counts.bureau, counts.male, counts.female, counts.middleAged],
  ]);

  const html = `
    <html>
      <head>
        <meta charset="utf-8" />
        <style>
          body { font-family: "Microsoft JhengHei", Arial, sans-serif; }
          h1 { text-align: center; font-size: 20px; }
          .meta { display: flex; justify-content: space-between; margin: 8px 0 12px; }
          table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
          th, td { border: 1px solid #333; padding: 6px; font-size: 12px; vertical-align: middle; }
          th { background: #e8eef5; font-weight: 700; text-align: center; }
          .count-title { font-weight: 700; margin-top: 20px; }
        </style>
      </head>
      <body>
        <h1>勞資爭議調解(仲裁)會議週報表</h1>
        <div class="meta">
          <span>日期：${escapeHtml(reportTitle)}</span>
          <span>案件數：共${rows.length}件</span>
          <span>本週派案人員：${escapeHtml(dispatcher)}</span>
        </div>
        ${detailTable}
        <div class="count-title">件數統計</div>
        ${countTable}
      </body>
    </html>
  `;

  const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `勞資爭議調解會議週報表_${toDateInputValue(selectedWeekStart)}.xls`;
  link.click();
  URL.revokeObjectURL(url);
}

function handleWeekChange(event) {
  selectedWeekStart = weekValueToDate(event.target.value);
  render();
}

function startOfWeek(date) {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + diff);
  return copy;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function parseDate(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function toDateInputValue(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

function formatDate(date) {
  return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

function formatMonthDay(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

function dateToWeekValue(date) {
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((target - yearStart) / 86400000 + 1) / 7);
  return `${target.getUTCFullYear()}-W${String(week).padStart(2, "0")}`;
}

function weekValueToDate(value) {
  const [yearText, weekText] = value.split("-W");
  const year = Number(yearText);
  const week = Number(weekText);
  const simple = new Date(year, 0, 1 + (week - 1) * 7);
  return startOfWeek(simple);
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeRoom(room) {
  if (!room) return { room: "晤談室(一)", roomOther: "" };
  const mapped = roomNameMap[room] ?? room;
  if (fixedRooms.includes(mapped) || mapped === "其他") {
    return { room: mapped, roomOther: mapped === "其他" && !roomNameMap[room] ? room : "" };
  }
  return { room: "其他", roomOther: room };
}

function displayRoom(item) {
  return item.room === "其他" ? item.roomOther || "其他" : item.room;
}

function isCommittee(item) {
  return (item?.reportCategory || "") === "F";
}

// 從 item 取得顯示用調解人字串
function getMediatorDisplay(item) {
  if (isCommittee(item)) {
    try {
      const obj = JSON.parse(item.mediator || "{}");
      return [obj.chair, obj.labor, obj.mgmt].filter(Boolean).join("／");
    } catch { return item.mediator || ""; }
  }
  return item.mediator || "";
}

// 切換調解委員會欄位
function onReportCategoryChange() {
  const isF = document.querySelector("#reportCategory").value === "F";
  document.getElementById("mediatorSingleLabel").style.display = isF ? "none" : "";
  document.getElementById("mediatorChairLabel").style.display  = isF ? "" : "none";
  document.getElementById("mediatorLaborLabel").style.display  = isF ? "" : "none";
  document.getElementById("mediatorMgmtLabel").style.display   = isF ? "" : "none";
  document.getElementById("mediator").required      = !isF;
  document.getElementById("mediatorChair").required = isF;
  document.getElementById("mediatorLabor").required = isF;
  document.getElementById("mediatorMgmt").required  = isF;
}

// 從 mediator 欄位字串解析三委員
function parseMediatorField(val) {
  try { return JSON.parse(val || "{}"); } catch { return {}; }
}

function toggleOtherRoom() {
  const isOther = document.querySelector("#room").value === "其他";
  document.querySelector("#roomOtherLabel").hidden = !isOther;
  document.querySelector("#roomOther").required = isOther;
}

function tableHtml(rows) {
  return `<table>${rows
    .map((row, rowIndex) => `<tr>${row.map((cell) => `${rowIndex === 0 ? "<th>" : "<td>"}${escapeHtml(cell ?? "")}${rowIndex === 0 ? "</th>" : "</td>"}`).join("")}</tr>`)
    .join("")}</table>`;
}

function buildWeeklyCounts(rows) {
  const counts = Object.fromEntries(Object.keys(reportCategoryLabels).map((key) => [key, 0]));
  counts.male = 0;
  counts.female = 0;
  counts.middleAged = 0;
  rows.forEach((item) => {
    const key = item.reportCategory || "mediator";
    counts[key] = (counts[key] || 0) + 1;
    counts.male += toNumber(item.maleCount) || inferGenderCount(item.workerGender, "男");
    counts.female += toNumber(item.femaleCount) || inferGenderCount(item.workerGender, "女");
    if (toNumber(item.workerAge) >= 45) counts.middleAged += 1;
  });
  return counts;
}

function toMinguoDate(date) {
  return `${date.getFullYear() - 1911}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}`;
}

function getWeekdayName(date) {
  const day = date.getDay();
  return ["週日", "週一", "週二", "週三", "週四", "週五", "週六"][day];
}

async function saveRemoteCase(item) {
  if (!supabaseClient) return;
  const { error } = await supabaseClient
    .from("mediation_schedules")
    .upsert(toDbCase(item), { onConflict: "id" });

  if (error) {
    alert(`Supabase 儲存失敗：${error.message}`);
  }
}

async function deleteRemoteCase(id) {
  if (!supabaseClient) return;
  const { error } = await supabaseClient.from("mediation_schedules").delete().eq("id", id);
  if (error) {
    alert(`Supabase 刪除失敗：${error.message}`);
  }
}

function toDbCase(item) {
  return {
    id: item.id,
    source_case_no: item.sourceCaseNo || null,
    case_no: item.caseNo,
    meeting_date: item.meetingDate,
    meeting_time: item.meetingTime,
    room: item.room,
    room_other: item.roomOther || null,
    worker: item.worker,
    employer: item.employer,
    mediator: item.mediator || null,
    recorder: item.recorder || null,
    owner: item.owner || null,
    dispute_type: item.disputeType || null,
    report_category: item.reportCategory || null,
    worker_gender: item.workerGender || null,
    worker_age: item.workerAge ? Number(item.workerAge) : null,
    male_count: item.maleCount ? Number(item.maleCount) : 0,
    female_count: item.femaleCount ? Number(item.femaleCount) : 0,
    special_case_type: item.specialCaseType || null,
    mediation_method_code: item.mediationMethodCode || null,
    status: item.status || "scheduled",
    notes: item.notes || null,
  };
}

function fromDbCase(row) {
  return {
    id: row.id,
    sourceCaseNo: row.source_case_no || "",
    caseNo: row.case_no || "",
    meetingDate: row.meeting_date || "",
    meetingTime: row.meeting_time || "",
    room: row.room || "晤談室(一)",
    roomOther: row.room_other || "",
    worker: row.worker || "",
    employer: row.employer || "",
    mediator: row.mediator || "",
    recorder: row.recorder || "",
    owner: row.owner || "",
    disputeType: row.dispute_type || "其他",
    reportCategory: row.report_category || "mediator",
    workerGender: row.worker_gender || "",
    workerAge: row.worker_age ?? "",
    maleCount: row.male_count ?? 0,
    femaleCount: row.female_count ?? 0,
    specialCaseType: row.special_case_type || "",
    mediationMethodCode: row.mediation_method_code || "",
    status: row.status || "scheduled",
    notes: row.notes || "",
  };
}

function guessDisputeType(claim = "") {
  if (claim.includes("資遣")) return "資遣費";
  if (claim.includes("職災")) return "職災";
  if (claim.includes("解僱")) return "解僱";
  if (claim.includes("退休")) return "退休金";
  if (claim.includes("工資") || claim.includes("加班")) return "工資";
  return "其他";
}

function normalizeBatchItem(item) {
  return {
    ...item,
    sourceCaseNo: `batch-${item.batchNo || item.worker || crypto.randomUUID()}`,
    customCaseNo: item.batchNo ? `批次${item.batchNo}` : "",
    claim: [item.claim1, item.claim2, item.claim3].filter(Boolean).join("、"),
    acceptanceMethod: item.acceptanceMethod || "",
    receivedDate: item.receivedDate || "",
    reportCategory: item.reportCategory || reportCategoryFromMethod(item.mediationMethodCode),
  };
}

function mergeIntakeSources(intake, batch) {
  const batchByPeople = new Map(
    batch.map((item) => [`${item.worker}||${item.employer}`, item])
  );
  const merged = intake.map((item) => {
    const match = batchByPeople.get(`${item.worker}||${item.employer}`);
    if (!match) return item;
    return {
      ...item,
      ...match,
      sourceCaseNo: item.sourceCaseNo || match.sourceCaseNo,
      customCaseNo: item.customCaseNo || match.customCaseNo,
      claim: item.claim || match.claim,
      receivedDate: item.receivedDate || match.receivedDate,
    };
  });
  const intakePeople = new Set(intake.map((item) => `${item.worker}||${item.employer}`));
  const batchOnly = batch.filter((item) => !intakePeople.has(`${item.worker}||${item.employer}`));
  return [...merged, ...batchOnly];
}

function reportCategoryFromMethod(method = "") {
  const text = String(method);
  if (text.startsWith("F")) return "committee";
  if (text.startsWith("I")) return "arbitration";
  if (text.startsWith("G")) return "transfer";
  if (text.startsWith("A") || text.startsWith("AH")) return "laborAssocCity";
  if (text.startsWith("B") || text.startsWith("BH")) return "laborAssocCounty";
  if (text.startsWith("E") || text.startsWith("EH")) return "laborEmploymentAssoc";
  if (text.startsWith("J") || text.startsWith("JH")) return "occupationalInjuryAssoc";
  return "mediator";
}

function inferGenderCount(gender, target) {
  return gender === target ? 1 : 0;
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function buildIntakeNotes(item) {
  if (!item) return "";
  const parts = [];
  if (item.claim) parts.push(item.claim);
  if (item.mediationMethodCode) parts.push(item.mediationMethodCode);
  if (item.mediationLocation) parts.push(`原填地點：${item.mediationLocation}`);
  if (item.specialCaseType && item.specialCaseType !== "無") parts.push(`特殊案件：${item.specialCaseType}`);
  return parts.join("；");
}

// xlsx.js 0.18.5 bug 修正：某些 Excel 的 ws['!ref'] 只包含標題列
// 掃描 ws 所有 cell key，找出實際最大列號並修正範圍
function fixSheetRef(ws) {
  const refMatch = (ws['!ref'] || 'A1').match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!refMatch) return;
  let maxRow = parseInt(refMatch[4]);
  for (const key of Object.keys(ws)) {
    if (key[0] === '!' || key[0] === undefined) continue;
    const m = key.match(/\d+/);
    if (m) { const rn = parseInt(m[0]); if (rn > maxRow) maxRow = rn; }
  }
  if (maxRow > parseInt(refMatch[4])) {
    ws['!ref'] = refMatch[1] + refMatch[2] + ':' + refMatch[3] + maxRow;
  }
}
