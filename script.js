const fields = {
  schoolName: document.querySelector("#schoolName"),
  activityName: document.querySelector("#activityName"),
  applicantUnit: document.querySelector("#applicantUnit"),
  applicantName: document.querySelector("#applicantName"),
  activityDate: document.querySelector("#activityDate"),
  activityEndDate: document.querySelector("#activityEndDate"),
  location: document.querySelector("#location"),
  startTime: document.querySelector("#startTime"),
  endTime: document.querySelector("#endTime"),
  reason: document.querySelector("#reason"),
  manualStudentInput: document.querySelector("#manualStudentInput"),
  studentInput: document.querySelector("#studentInput"),
};

const previewTargets = {
  schoolEyebrow: document.querySelector('[data-preview="schoolEyebrow"]'),
  permitTitle: document.querySelector('[data-preview="permitTitle"]'),
  activityName: document.querySelector('[data-preview="activityName"]'),
  applicantUnit: document.querySelector('[data-preview="applicantUnit"]'),
  applicantName: document.querySelector('[data-preview="applicantName"]'),
  activityDate: document.querySelector('[data-preview="activityDate"]'),
  timeRange: document.querySelector('[data-preview="timeRange"]'),
  location: document.querySelector('[data-preview="location"]'),
  reason: document.querySelector('[data-preview="reason"]'),
};

const groupedStudents = document.querySelector("#groupedStudents");
const mentorSignatures = document.querySelector("#mentorSignatures");
const warnings = document.querySelector("#warnings");
const studentCount = document.querySelector("#studentCount");
const classCount = document.querySelector("#classCount");
const rosterStatus = document.querySelector("#rosterStatus");
const printPages = document.querySelector("#printPages");
const permitPreviewEl = document.querySelector("#permitPreview");
const signInPreviewEl = document.querySelector("#signInPreview");
const keepApplicantCheck = document.querySelector("#keepApplicantCheck");

const storageKey = "nanning-public-leave-form-v1";
const customRosterStorageKey = "nanning-custom-roster-v1";

// 南寧高中節次／時段定義
const PERIOD_LIST = [
  { id: "morning", label: "早自習", isClass: false },
  { id: "p1", label: "第1節", isClass: true, num: 1 },
  { id: "p2", label: "第2節", isClass: true, num: 2 },
  { id: "p3", label: "第3節", isClass: true, num: 3 },
  { id: "p4", label: "第4節", isClass: true, num: 4 },
  { id: "lunch", label: "午休", isClass: false },
  { id: "p5", label: "第5節", isClass: true, num: 5 },
  { id: "p6", label: "第6節", isClass: true, num: 6 },
  { id: "p7", label: "第7節", isClass: true, num: 7 },
  { id: "clean", label: "打掃", isClass: false },
  { id: "p8", label: "第8節", isClass: true, num: 8 },
];

const PERIOD_PRESETS = {
  all: PERIOD_LIST.map((p) => p.id),
  morning: ["morning", "p1", "p2", "p3", "p4"],
  afternoon: ["lunch", "p5", "p6", "p7", "clean", "p8"],
  classes: ["p1", "p2", "p3", "p4", "p5", "p6", "p7"],
  clear: [],
};

const WEEKDAY_NAMES = ["日", "一", "二", "三", "四", "五", "六"];

// 時段與預覽狀態管理
let scheduleMode = "uniform"; // "uniform" | "daily"
let uniformPeriods = []; // Array of period IDs
let dailySchedules = []; // Array of { date: "YYYY-MM-DD", periods: string[], startTime: string, endTime: string }
let activePreviewMode = "permit"; // "permit" | "signin"

function getTodayIso() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function getWeekdayName(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-").map(Number);
  if (!y || !m || !d) return "";
  const dt = new Date(y, m - 1, d);
  return WEEKDAY_NAMES[dt.getDay()] || "";
}

function formatDate(value) {
  if (!value) return "未填寫";
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return value;
  return `${Number(year) - 1911} 年 ${Number(month)} 月 ${Number(day)} 日`;
}

function formatDateWithWeekday(value) {
  if (!value) return "未填寫";
  const base = formatDate(value);
  const wk = getWeekdayName(value);
  return wk ? `${base}（${wk}）` : base;
}

function formatShortDateWithWeekday(value) {
  if (!value) return "未填日期";
  const [, month, day] = value.split("-");
  if (!month || !day) return value;
  const wk = getWeekdayName(value);
  return `${Number(month)}/${Number(day)}${wk ? `(${wk})` : ""}`;
}

function formatRocSlashDate(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return value;
  return `${Number(year) - 1911}/${Number(month)}/${Number(day)}`;
}

function getDateRangeArray(startStr, endStr) {
  if (!startStr) return [];
  const finalEndStr = !endStr || endStr < startStr ? startStr : endStr;
  const [sy, sm, sd] = startStr.split("-").map(Number);
  const [ey, em, ed] = finalEndStr.split("-").map(Number);
  if (!sy || !sm || !sd || !ey || !em || !ed) return [startStr];

  const startDate = new Date(sy, sm - 1, sd);
  const endDate = new Date(ey, em - 1, ed);
  const result = [];
  const maxDays = 60;

  const cursor = new Date(startDate);
  while (cursor <= endDate && result.length < maxDays) {
    const yyyy = cursor.getFullYear();
    const mm = String(cursor.getMonth() + 1).padStart(2, "0");
    const dd = String(cursor.getDate()).padStart(2, "0");
    result.push(`${yyyy}-${mm}-${dd}`);
    cursor.setDate(cursor.getDate() + 1);
  }
  return result.length ? result : [startStr];
}

function isConsecutiveDates(dateStrings) {
  if (dateStrings.length <= 1) return true;
  const range = getDateRangeArray(dateStrings[0], dateStrings[dateStrings.length - 1]);
  if (range.length !== dateStrings.length) return false;
  return range.every((d, i) => d === dateStrings[i]);
}

function formatPeriods(selectedIds) {
  if (!Array.isArray(selectedIds) || selectedIds.length === 0) return "";
  const selectedSet = new Set(selectedIds);
  const ordered = PERIOD_LIST.filter((p) => selectedSet.has(p.id));
  if (ordered.length === 0) return "";

  if (ordered.length === PERIOD_LIST.length) {
    return "全天（早自習～第8節）";
  }

  const hasMorning = selectedSet.has("morning");
  const hasLunch = selectedSet.has("lunch");
  const hasClean = selectedSet.has("clean");

  const classNums = ordered.filter((p) => p.isClass).map((p) => p.num);
  const classRanges = [];
  let i = 0;
  while (i < classNums.length) {
    let start = classNums[i];
    let end = start;
    while (i + 1 < classNums.length && classNums[i + 1] === end + 1) {
      end = classNums[i + 1];
      i++;
    }
    if (start === end) {
      classRanges.push(`第${start}節`);
    } else {
      classRanges.push(`第${start}節～第${end}節`);
    }
    i++;
  }

  const parts = [];
  if (hasMorning) parts.push("早自習");

  const onlyAfternoonClasses = classNums.length > 0 && classNums[0] >= 5;
  if (hasLunch && onlyAfternoonClasses) {
    parts.push("午休");
  }

  parts.push(...classRanges);

  if (hasLunch && !onlyAfternoonClasses) {
    parts.push("午休");
  }
  if (hasClean) {
    parts.push("打掃");
  }

  return parts.join("、");
}

function formatScheduleSlot(periods, startTime, endTime) {
  const periodText = formatPeriods(periods);
  const start = (startTime || "").trim();
  const end = (endTime || "").trim();
  const timeText = start && end ? `${start} 至 ${end}` : start || end || "";

  if (periodText && timeText) {
    return `${periodText}（${timeText}）`;
  }
  return periodText || timeText || "未填寫";
}

function getResolvedSchedule() {
  const startStr = fields.activityDate.value;
  const endStr = fields.activityEndDate.value || startStr;

  if (scheduleMode === "daily" && dailySchedules.length > 0) {
    const validDays = dailySchedules
      .filter((item) => item && item.date)
      .slice()
      .sort((a, b) => a.date.localeCompare(b.date));

    if (validDays.length === 0) {
      return {
        dateDisplay: "未填寫",
        timeDisplay: "未填寫",
        days: [],
      };
    }

    const days = validDays.map((item) => ({
      date: item.date,
      formattedDate: formatDate(item.date),
      formattedDateWithWeekday: formatDateWithWeekday(item.date),
      shortDate: formatShortDateWithWeekday(item.date),
      timeText: formatScheduleSlot(item.periods, item.startTime, item.endTime),
    }));

    if (days.length === 1) {
      return {
        dateDisplay: days[0].formattedDate,
        timeDisplay: days[0].timeText,
        days,
      };
    }

    const dateList = days.map((d) => d.date);
    const dateDisplay = isConsecutiveDates(dateList)
      ? `${days[0].formattedDate} 至 ${days[days.length - 1].formattedDate}（共 ${days.length} 日）`
      : `${dateList.map((d) => formatRocSlashDate(d)).join("、")}（共 ${days.length} 日）`;

    const allSameTime = days.every((d) => d.timeText === days[0].timeText);
    const timeDisplay = allSameTime
      ? days[0].timeText === "未填寫"
        ? "未填寫"
        : `每日：${days[0].timeText}`
      : days.map((d) => `${d.shortDate}：${d.timeText}`).join("\n");

    return {
      dateDisplay,
      timeDisplay,
      days,
    };
  }

  const dates = getDateRangeArray(startStr, endStr);
  const uniformTimeText = formatScheduleSlot(uniformPeriods, fields.startTime.value, fields.endTime.value);

  if (dates.length === 0) {
    return {
      dateDisplay: "未填寫",
      timeDisplay: uniformTimeText,
      days: [],
    };
  }

  const days = dates.map((d) => ({
    date: d,
    formattedDate: formatDate(d),
    formattedDateWithWeekday: formatDateWithWeekday(d),
    shortDate: formatShortDateWithWeekday(d),
    timeText: uniformTimeText,
  }));

  if (days.length === 1) {
    return {
      dateDisplay: days[0].formattedDate,
      timeDisplay: uniformTimeText,
      days,
    };
  }

  return {
    dateDisplay: `${days[0].formattedDate} 至 ${days[days.length - 1].formattedDate}（共 ${days.length} 日）`,
    timeDisplay: uniformTimeText === "未填寫" ? "未填寫" : `每日：${uniformTimeText}`,
    days,
  };
}

function renderUniformPeriodButtons() {
  const container = document.querySelector("#uniformPeriodGroup");
  if (!container) return;

  const selectedSet = new Set(uniformPeriods);
  container.innerHTML = PERIOD_LIST.map((p) => {
    const isActive = selectedSet.has(p.id);
    const specialClass = !p.isClass ? "special-slot" : "";
    return `<button type="button" class="period-chip ${specialClass} ${isActive ? "active" : ""}" data-period-id="${p.id}" aria-pressed="${isActive}">${escapeHtml(p.label)}</button>`;
  }).join("");
}

function syncDailySchedulesFromRange(forceReset = false) {
  const startStr = fields.activityDate.value || getTodayIso();
  const endStr = fields.activityEndDate.value || startStr;
  const dates = getDateRangeArray(startStr, endStr);

  const existingByDate = new Map(dailySchedules.map((item) => [item.date, item]));

  dailySchedules = dates.map((dateStr) => {
    if (!forceReset && existingByDate.has(dateStr)) {
      return existingByDate.get(dateStr);
    }
    return {
      date: dateStr,
      periods: [...uniformPeriods],
      startTime: fields.startTime.value || "",
      endTime: fields.endTime.value || "",
    };
  });
}

// 當逐日卡片增刪或改日期時，自動同步上方起迄日期欄位
function syncDateInputsFromDailySchedules() {
  const validDates = dailySchedules
    .map((item) => item && item.date)
    .filter(Boolean)
    .sort();
  if (validDates.length > 0) {
    if (fields.activityDate) fields.activityDate.value = validDates[0];
    if (fields.activityEndDate) fields.activityEndDate.value = validDates[validDates.length - 1];
  }
}

function renderDailyScheduleList() {
  const container = document.querySelector("#dailyScheduleList");
  if (!container) return;

  if (dailySchedules.length === 0) {
    syncDailySchedulesFromRange();
  }

  container.innerHTML = dailySchedules
    .map((item, idx) => {
      const selectedSet = new Set(item.periods || []);
      const wk = getWeekdayName(item.date);
      const chipsHtml = PERIOD_LIST.map((p) => {
        const isActive = selectedSet.has(p.id);
        const specialClass = !p.isClass ? "special-slot" : "";
        return `<button type="button" class="period-chip ${specialClass} ${isActive ? "active" : ""}" data-daily-index="${idx}" data-period-id="${p.id}" aria-pressed="${isActive}">${escapeHtml(p.label)}</button>`;
      }).join("");

      return `
        <div class="daily-card" data-daily-card="${idx}">
          <div class="daily-card-header">
            <div class="daily-date-wrap">
              <span class="period-label">第 ${idx + 1} 日：</span>
              <input type="date" class="daily-date-input" data-daily-index="${idx}" value="${escapeHtml(item.date || "")}">
              ${wk ? `<span class="daily-weekday-badge">星期${wk}</span>` : ""}
            </div>
            <div class="period-quick-actions" data-daily-quick="${idx}">
              <button type="button" class="quick-btn" data-daily-index="${idx}" data-preset="all">全天</button>
              <button type="button" class="quick-btn" data-daily-index="${idx}" data-preset="morning">上午</button>
              <button type="button" class="quick-btn" data-daily-index="${idx}" data-preset="afternoon">下午</button>
              <button type="button" class="quick-btn" data-daily-index="${idx}" data-preset="classes">第1~7節</button>
              <button type="button" class="quick-btn clear-preset" data-daily-index="${idx}" data-preset="clear">清除</button>
              ${
                dailySchedules.length > 1
                  ? `<button type="button" class="daily-remove-btn" data-remove-daily="${idx}" title="移除此日">移除此日</button>`
                  : ""
              }
            </div>
          </div>
          <div class="period-chips">${chipsHtml}</div>
          <div class="custom-time-row" style="padding-top: 6px;">
            <span class="sub-label">當日補充或自訂起迄時間（選填）：</span>
            <div class="time-inputs-inline">
              <input type="time" class="daily-start-time" data-daily-index="${idx}" value="${escapeHtml(item.startTime || "")}" aria-label="當日開始時間">
              <span class="time-sep">至</span>
              <input type="time" class="daily-end-time" data-daily-index="${idx}" value="${escapeHtml(item.endTime || "")}" aria-label="當日結束時間">
            </div>
          </div>
        </div>
      `;
    })
    .join("");
}

function setScheduleMode(mode) {
  scheduleMode = mode === "daily" ? "daily" : "uniform";
  const uniformTab = document.querySelector("#uniformScheduleTab");
  const dailyTab = document.querySelector("#dailyScheduleTab");
  const uniformPanel = document.querySelector("#uniformSchedulePanel");
  const dailyPanel = document.querySelector("#dailySchedulePanel");

  if (uniformTab && dailyTab && uniformPanel && dailyPanel) {
    if (scheduleMode === "daily") {
      uniformTab.classList.remove("active");
      dailyTab.classList.add("active");
      uniformPanel.style.display = "none";
      dailyPanel.style.display = "grid";
      if (dailySchedules.length === 0) {
        syncDailySchedulesFromRange();
      }
      renderDailyScheduleList();
    } else {
      uniformTab.classList.add("active");
      dailyTab.classList.remove("active");
      uniformPanel.style.display = "grid";
      dailyPanel.style.display = "none";
      renderUniformPeriodButtons();
    }
  }
  update();
}

function setPreviewMode(mode) {
  activePreviewMode = mode === "signin" ? "signin" : "permit";
  const permitTab = document.querySelector("#previewPermitTab");
  const signInTab = document.querySelector("#previewSignInTab");

  if (permitTab && signInTab) {
    if (activePreviewMode === "signin") {
      permitTab.classList.remove("active");
      signInTab.classList.add("active");
    } else {
      permitTab.classList.add("active");
      signInTab.classList.remove("active");
    }
  }

  if (permitPreviewEl && signInPreviewEl) {
    permitPreviewEl.style.display = activePreviewMode === "permit" ? "block" : "none";
    signInPreviewEl.style.display = activePreviewMode === "signin" ? "grid" : "none";
  }

  update();
}

function getLoadedRoster() {
  const raw = localStorage.getItem(customRosterStorageKey);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return { data: parsed, isCustom: true };
      }
    } catch (e) {
      console.error("Failed to parse custom roster", e);
    }
  }
  return { data: window.STUDENT_ROSTER || [], isCustom: false };
}

function normalizeClassCodeText(raw) {
  const trimmed = String(raw || "").trim().replace(/班$/, "");
  return trimmed;
}

function findMatchingKey(keys, candidates) {
  const normalizedCandidates = candidates.map((c) => c.toLowerCase());
  return keys.find((k) => normalizedCandidates.includes(k.trim().toLowerCase()));
}

function handleExcelUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (!jsonData.length) {
        alert("Excel 檔案為空，請確認檔案內容！");
        return;
      }

      const firstRow = jsonData[0];
      const keys = Object.keys(firstRow);

      const headerMap = {
        studentIdKey: findMatchingKey(keys, ["學號", "學生學號", "studentid", "student_id", "id"]),
        classCodeKey: findMatchingKey(keys, ["班級", "班級代號", "班級代碼", "班級名稱", "年班", "class", "classcode", "class_code"]),
        seatNumberKey: findMatchingKey(keys, ["座號", "班級座號", "學生座號", "seat", "seatnumber", "seat_number"]),
        nameKey: findMatchingKey(keys, ["姓名", "學生姓名", "name", "studentname"]),
      };

      if (!headerMap.studentIdKey || !headerMap.classCodeKey || !headerMap.seatNumberKey || !headerMap.nameKey) {
        alert("Excel 欄位格式錯誤！必須包含以下欄位：\n「學號」、「班級」、「座號」、「姓名」\n\n請下載名冊範例 Excel 對照格式。");
        return;
      }

      const rosterData = jsonData
        .map((row) => {
          return {
            studentId: String(row[headerMap.studentIdKey] || "").trim(),
            classCode: normalizeClassCodeText(row[headerMap.classCodeKey]),
            seatNumber: Number(row[headerMap.seatNumberKey]) || 0,
            name: String(row[headerMap.nameKey] || "").trim(),
          };
        })
        .filter((item) => item.name && item.classCode);

      if (rosterData.length === 0) {
        alert("沒有成功解析出有效的學生資料，請檢查 Excel 內容！");
        return;
      }

      localStorage.setItem(customRosterStorageKey, JSON.stringify(rosterData));
      alert(`名冊上傳成功！已載入 ${rosterData.length} 筆學生資料。`);
      update();
    } catch (error) {
      console.error(error);
      alert("解析 Excel 檔案失敗，請確保檔案格式正確！\n錯誤原因：" + error.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function normalizeSeatCode(raw) {
  const value = String(raw || "").trim().toUpperCase();
  const compact = value.replace(/\s+/g, "").replace(/班/g, "");
  const dashed = compact.match(/^(\d{3})-(\d{1,2})$/);
  const plain = compact.match(/^(\d{3})(\d{1,2})$/);

  if (!dashed && !plain) return null;

  const classCode = dashed ? dashed[1] : plain[1];
  const seatNumber = Number(dashed ? dashed[2] : plain[2]);
  if (!seatNumber || seatNumber > 99) return null;

  return {
    classCode,
    seatNumber,
    displayCode: `${classCode}-${String(seatNumber).padStart(2, "0")}`,
    sortCode: Number(`${classCode}${String(seatNumber).padStart(2, "0")}`),
  };
}

function makeStudentRecord(studentId, seatCodeRaw, name) {
  const seat = normalizeSeatCode(seatCodeRaw);
  if (!seat || !name) return null;

  return {
    studentId: String(studentId || "").trim(),
    name: String(name || "").trim(),
    ...seat,
  };
}

function makeStudentRecordFromParts(studentId, classCode, seatNumber, name) {
  const classText = normalizeClassCodeText(classCode);
  const seatText = String(seatNumber || "").trim().padStart(2, "0");
  return makeStudentRecord(studentId, `${classText}-${seatText}`, name);
}

function normalizeName(value) {
  return String(value || "").replace(/\s+/g, "").trim();
}

function maskName(name) {
  if (!name) return "";
  const len = name.length;
  if (len <= 1) return name;
  if (len === 2) {
    return name[0] + "○";
  }
  if (len === 3) {
    return name[0] + "○" + name[2];
  }
  return name[0] + "○".repeat(len - 2) + name[len - 1];
}

// 南寧高中班級顯示格式：例如 401 -> 401班
function readableClass(classCode) {
  const cleaned = String(classCode || "").trim();
  if (!cleaned || cleaned === "未分班") return "未分班";
  if (cleaned.endsWith("班")) return cleaned;
  return `${cleaned}班`;
}

function splitLine(line) {
  return line
    .replace(/[，、;\t；]/g, ",")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
}

function expandSeatRange(classCode, startSeat, endSeat, roster, lineNumber) {
  const min = Math.min(startSeat, endSeat);
  const max = Math.max(startSeat, endSeat);
  const result = [];
  const classKey = String(classCode).trim();

  for (let s = min; s <= max; s++) {
    const seatText = String(s).padStart(2, "0");
    const displayCode = `${classKey}-${seatText}`;
    const compactCode = `${classKey}${seatText}`;
    const rosterStudent = roster.bySeatCode.get(displayCode) || roster.bySeatCode.get(compactCode);

    if (rosterStudent) {
      result.push({
        ...rosterStudent,
        sourceLine: lineNumber,
      });
    } else {
      result.push({
        studentId: "",
        name: "未填姓名",
        classCode: classKey,
        seatNumber: s,
        displayCode,
        sortCode: Number(compactCode),
        sourceLine: lineNumber,
      });
    }
  }
  return result;
}

function parseSingleManualChunk(tokens, lineNumber) {
  if (tokens.length >= 3) {
    let studentId = "";
    let seatCodeRaw = "";
    let name = "";

    let seatIndex = tokens.findIndex((t) => normalizeSeatCode(t) !== null);
    if (seatIndex !== -1) {
      seatCodeRaw = tokens[seatIndex];
      const remaining = tokens.filter((_, idx) => idx !== seatIndex);
      let idIndex = remaining.findIndex((t) => /^\d{5,}$/.test(t));
      if (idIndex !== -1) {
        studentId = remaining[idIndex];
        name = remaining.filter((_, idx) => idx !== idIndex).join(" ");
      } else {
        studentId = "";
        name = remaining.join(" ");
      }
    } else {
      studentId = tokens[0];
      seatCodeRaw = tokens[1];
      name = tokens.slice(2).join(" ");
    }

    const seat = normalizeSeatCode(seatCodeRaw);
    if (!seat) {
      return {
        warning: `第 ${lineNumber} 行手動格式錯誤（座號格式不對，應如 40101 或 401-01）：${tokens.join(" ")}`,
      };
    }

    return {
      studentId,
      name: name || "未填姓名",
      ...seat,
      sourceLine: lineNumber,
    };
  } else if (tokens.length === 2) {
    const seatCodeRaw = tokens[0];
    const name = tokens[1];
    const seat = normalizeSeatCode(seatCodeRaw);
    if (!seat) {
      const seat2 = normalizeSeatCode(name);
      if (seat2 && /^\d{5,}$/.test(seatCodeRaw)) {
        return {
          studentId: seatCodeRaw,
          name: "未填姓名",
          ...seat2,
          sourceLine: lineNumber,
        };
      }
      if (/^\d{5,}$/.test(name)) {
        return {
          studentId: name,
          name: seatCodeRaw,
          classCode: "未分班",
          seatNumber: 0,
          displayCode: "未分班",
          sortCode: 999999,
          sourceLine: lineNumber,
        };
      }
      return {
        warning: `第 ${lineNumber} 行手動格式錯誤（應包含班級座號與姓名，如「40101 蘇大輔」）：${tokens.join(" ")}`,
      };
    }
    return {
      studentId: "",
      name,
      ...seat,
      sourceLine: lineNumber,
    };
  } else if (tokens.length === 1) {
    const seat = normalizeSeatCode(tokens[0]);
    if (seat) {
      return {
        studentId: "",
        name: "未填姓名",
        ...seat,
        sourceLine: lineNumber,
      };
    }
    return {
      warning: `第 ${lineNumber} 行格式不足（手動輸入每人需至少包含班級座號與姓名，如「40101 蘇大輔」）：${tokens[0]}`,
    };
  } else {
    return {
      warning: `第 ${lineNumber} 行格式不足（手動輸入需包含班級座號與姓名，如「40101 蘇大輔」）`,
    };
  }
}

function parseManualInputLine(line, lineNumber) {
  const trimmed = line.trim();
  if (!trimmed) return null;

  // 1. 若含有頓號或分號，先拆分為多個子段
  if (/[、;；]/.test(trimmed)) {
    const subSegments = trimmed.split(/[、;；]+/).map((s) => s.trim()).filter(Boolean);
    const results = [];
    subSegments.forEach((seg) => {
      const res = parseManualInputLine(seg, lineNumber);
      if (res) {
        if (Array.isArray(res)) results.push(...res);
        else results.push(res);
      }
    });
    return results.length ? results : null;
  }

  // 2. 將逗號與空白拆分為原始 tokens
  const rawTokens = trimmed
    .replace(/[，,]/g, " ")
    .split(/\s+/)
    .filter(Boolean);

  if (rawTokens.length === 0) return null;

  // 找出所有座號所在索引
  const seatIndices = [];
  rawTokens.forEach((t, idx) => {
    if (normalizeSeatCode(t) !== null) {
      seatIndices.push(idx);
    }
  });

  // 若只有 0 或 1 個座號，直接由單筆解析器處理
  if (seatIndices.length <= 1) {
    return parseSingleManualChunk(rawTokens, lineNumber);
  }

  // 若同一行有多個座號，依邊界切分成多個學生 chunks
  const startIndices = [];
  seatIndices.forEach((sIdx, i) => {
    if (i === 0) {
      startIndices.push(0);
    } else {
      const prevTok = rawTokens[sIdx - 1];
      const prevPrevTok = rawTokens[sIdx - 2];
      // 如果座號前一個詞是學號（且不是前一個學生的姓名）
      if (prevTok && /^\d{5,}$/.test(prevTok) && sIdx - 1 > (startIndices[i - 1] || 0)) {
        startIndices.push(sIdx - 1);
      } else {
        startIndices.push(sIdx);
      }
    }
  });

  const parsedStudents = [];
  for (let k = 0; k < startIndices.length; k++) {
    const start = startIndices[k];
    const end = k + 1 < startIndices.length ? startIndices[k + 1] : rawTokens.length;
    const chunk = rawTokens.slice(start, end);
    if (chunk.length > 0) {
      const parsed = parseSingleManualChunk(chunk, lineNumber);
      if (parsed) {
        parsedStudents.push(parsed);
      }
    }
  }

  return parsedStudents.length ? parsedStudents : null;
}

function parseRoster() {
  const byStudentId = new Map();
  const bySeatCode = new Map();
  const byName = new Map();
  const byClass = new Map();

  function addRecord(record) {
    if (!record) return;
    if (record.studentId) byStudentId.set(record.studentId.toUpperCase(), record);
    bySeatCode.set(record.displayCode.replace("-", ""), record);
    bySeatCode.set(record.displayCode, record);

    const nameKey = normalizeName(record.name);
    if (nameKey) {
      if (!byName.has(nameKey)) byName.set(nameKey, []);
      byName.get(nameKey).push(record);
    }

    if (record.classCode) {
      if (!byClass.has(record.classCode)) byClass.set(record.classCode, []);
      byClass.get(record.classCode).push(record);
    }
  }

  const { data: baseRoster } = getLoadedRoster();

  baseRoster.forEach((student) => {
    addRecord(makeStudentRecordFromParts(student.studentId, student.classCode, student.seatNumber, student.name));
  });

  if (fields.manualStudentInput && fields.manualStudentInput.value) {
    const rosterLines = fields.manualStudentInput.value.split(/\r?\n/);
    rosterLines.forEach((line, index) => {
      const parsed = parseManualInputLine(line, index + 1);
      if (parsed) {
        const toAdd = Array.isArray(parsed) ? parsed : [parsed];
        toAdd.forEach((p) => {
          if (p && !p.warning) addRecord(p);
        });
      }
    });
  }

  return { byStudentId, bySeatCode, byName, byClass };
}

function parseStudentLine(line, roster, lineNumber) {
  const trimmed = line.trim();
  if (!trimmed) return null;

  // 1. 同班多座號速記，如：401: 1 3 5 12 或 401：1~5 8 或 401# 1 3 5
  const classPrefixMatch = trimmed.match(/^(\d{3})班?[:：#]\s*(.+)$/);
  if (classPrefixMatch) {
    const classCode = classPrefixMatch[1];
    const rest = classPrefixMatch[2].trim();
    const subTokens = rest.split(/[,\s，、;；]+/).filter(Boolean);
    const students = [];
    const warnings = [];

    for (let j = 0; j < subTokens.length; j++) {
      const st = subTokens[j];
      const rangeMatch = st.match(/^(\d{1,2})\s*[~～至到]\s*(\d{1,2})$/);
      if (rangeMatch) {
        const s1 = Number(rangeMatch[1]);
        const s2 = Number(rangeMatch[2]);
        students.push(...expandSeatRange(classCode, s1, s2, roster, lineNumber));
        continue;
      }

      const singleSeatMatch = st.match(/^(\d{1,2})$/);
      if (singleSeatMatch) {
        const seatNum = Number(singleSeatMatch[1]);
        const nextTok = subTokens[j + 1];
        let customName = "";
        if (nextTok && !nextTok.match(/^\d+/) && !nextTok.match(/^[~～至到-]/)) {
          customName = nextTok;
          j++;
        }

        const seatCodeStr = `${classCode}-${String(seatNum).padStart(2, "0")}`;
        const compactCode = `${classCode}${String(seatNum).padStart(2, "0")}`;
        const rosterStudent = roster.bySeatCode.get(seatCodeStr) || roster.bySeatCode.get(compactCode);

        if (rosterStudent) {
          students.push({
            ...rosterStudent,
            name: customName || rosterStudent.name,
            sourceLine: lineNumber,
          });
        } else {
          students.push({
            studentId: "",
            name: customName || "未填姓名",
            classCode,
            seatNumber: seatNum,
            displayCode: seatCodeStr,
            sortCode: Number(compactCode),
            sourceLine: lineNumber,
          });
        }
        continue;
      }

      warnings.push(`第 ${lineNumber} 行同班座號無法辨識：${st}`);
    }

    if (warnings.length > 0 && students.length === 0) {
      return { warning: warnings.join("；") };
    }
    if (warnings.length > 0) {
      students.forEach((s) => (s._extraWarnings = warnings));
    }
    return students;
  }

  // 2. 整班連號區間速記，如：40101~40110 或 401-01~401-10 或 40101~10
  const fullRangeMatch = trimmed.match(/^(\d{3})-?(\d{1,2})\s*[~～至到]\s*(\d{3})?-?(\d{1,2})$/);
  if (fullRangeMatch) {
    const classCode = fullRangeMatch[1];
    const s1 = Number(fullRangeMatch[2]);
    const endClass = fullRangeMatch[3] || classCode;
    const s2 = Number(fullRangeMatch[4]);
    if (classCode === endClass) {
      return expandSeatRange(classCode, s1, s2, roster, lineNumber);
    }
  }

  // 3. 通用多詞解析（支援換行、空格、頓號、逗號等隔開多筆）
  const normalizedLine = trimmed.replace(/[，、;\t；,]/g, " ");
  const tokens = normalizedLine.split(/\s+/).filter(Boolean);

  if (tokens.length === 0) return null;

  const resultStudents = [];
  const tokenWarnings = [];
  let i = 0;

  while (i < tokens.length) {
    const token = tokens[i];

    // 判斷 3A: 獨立整班帶入代號，例如「401」或「401班」
    const classMatch = token.match(/^(\d{3})班?$/);
    if (classMatch) {
      const code = classMatch[1];
      const nextTok = tokens[i + 1];
      // 若後面緊接著 1~2 碼數字（如 401 1 2 5），代表為同班座號速記
      if (nextTok && nextTok.match(/^\d{1,2}$/)) {
        i++;
        while (i < tokens.length) {
          const seatTok = tokens[i];
          const rangeM = seatTok.match(/^(\d{1,2})\s*[~～至到]\s*(\d{1,2})$/);
          if (rangeM) {
            resultStudents.push(...expandSeatRange(code, Number(rangeM[1]), Number(rangeM[2]), roster, lineNumber));
            i++;
            continue;
          }
          const singleM = seatTok.match(/^(\d{1,2})$/);
          if (singleM) {
            const seatNum = Number(singleM[1]);
            const seatCodeStr = `${code}-${String(seatNum).padStart(2, "0")}`;
            const compactCode = `${code}${String(seatNum).padStart(2, "0")}`;
            const rosterStudent = roster.bySeatCode.get(seatCodeStr) || roster.bySeatCode.get(compactCode);
            resultStudents.push(
              rosterStudent
                ? { ...rosterStudent, sourceLine: lineNumber }
                : {
                    studentId: "",
                    name: "未填姓名",
                    classCode: code,
                    seatNumber: seatNum,
                    displayCode: seatCodeStr,
                    sortCode: Number(compactCode),
                    sourceLine: lineNumber,
                  }
            );
            i++;
            continue;
          }
          break;
        }
        continue;
      }

      // 整班帶入
      const classStudents = roster.byClass.get(code);
      if (classStudents && classStudents.length > 0) {
        resultStudents.push(
          ...classStudents.map((student) => ({
            ...student,
            sourceLine: lineNumber,
          }))
        );
        i++;
        continue;
      }
      tokenWarnings.push(`第 ${lineNumber} 行查無此班級代號：${token}`);
      i++;
      continue;
    }

    // 判斷 3B: 單詞內含連號區間，例如「40101~40105」
    const inlineRange = token.match(/^(\d{3})-?(\d{1,2})\s*[~～至到]\s*(\d{3})?-?(\d{1,2})$/);
    if (inlineRange) {
      const cCode = inlineRange[1];
      const s1 = Number(inlineRange[2]);
      const s2 = Number(inlineRange[4]);
      resultStudents.push(...expandSeatRange(cCode, s1, s2, roster, lineNumber));
      i++;
      continue;
    }

    // 判斷 3C: 座號格式（如 40101 或 401-01）
    const seat = normalizeSeatCode(token);
    if (seat) {
      const rosterStudent = roster.bySeatCode.get(seat.displayCode) || roster.bySeatCode.get(seat.displayCode.replace("-", ""));
      const nextTok = tokens[i + 1];

      // 檢查下一個詞是否為姓名
      if (nextTok && !normalizeSeatCode(nextTok) && !nextTok.match(/^\d{3,}$/) && !nextTok.match(/^[~～至到]/)) {
        resultStudents.push({
          ...seat,
          studentId: rosterStudent?.studentId || "",
          name: nextTok || rosterStudent?.name || "未填姓名",
          sourceLine: lineNumber,
        });
        i += 2;
        continue;
      }

      resultStudents.push({
        ...seat,
        studentId: rosterStudent?.studentId || "",
        name: rosterStudent?.name || "未填姓名",
        sourceLine: lineNumber,
      });
      i++;
      continue;
    }

    // 判斷 3D: 學號（如 999001）
    const rosterByStudentId = roster.byStudentId.get(token.toUpperCase());
    if (rosterByStudentId) {
      const nextTok = tokens[i + 1];
      if (nextTok && normalizeName(nextTok) === normalizeName(rosterByStudentId.name)) {
        i++;
      }
      resultStudents.push({
        ...rosterByStudentId,
        sourceLine: lineNumber,
      });
      i++;
      continue;
    }

    // 判斷 3E: 姓名查詢
    const nameMatches = roster.byName.get(normalizeName(token)) || [];
    if (nameMatches.length === 1) {
      resultStudents.push({
        ...nameMatches[0],
        sourceLine: lineNumber,
      });
      i++;
      continue;
    }
    if (nameMatches.length > 1) {
      const candidates = nameMatches
        .slice(0, 5)
        .map((student) => `${student.displayCode}/${student.studentId}`)
        .join("、");
      tokenWarnings.push(`第 ${lineNumber} 行姓名重複，請改用學號或班級座號：${token}（${candidates}）`);
      i++;
      continue;
    }

    // 無法辨識
    tokenWarnings.push(`第 ${lineNumber} 行無法辨識：${token}`);
    i++;
  }

  if (resultStudents.length === 0 && tokenWarnings.length > 0) {
    return { warning: tokenWarnings.join("；") };
  }
  if (tokenWarnings.length > 0) {
    resultStudents.forEach((s) => (s._extraWarnings = tokenWarnings));
  }
  return resultStudents;
}

function getActiveInputMode() {
  const manualInputTab = document.querySelector("#manualInputTab");
  return manualInputTab && manualInputTab.classList.contains("active") ? "manual" : "roster";
}

function getStudents() {
  const roster = parseRoster();
  const result = [];
  const seen = new Set();
  const warningMessages = [];

  const mode = getActiveInputMode();

  if (mode === "roster") {
    fields.studentInput.value.split(/\r?\n/).forEach((line, index) => {
      const parsed = parseStudentLine(line, roster, index + 1);
      if (!parsed) return;

      if (parsed.warning) {
        warningMessages.push(parsed.warning);
        return;
      }

      const studentsToAdd = Array.isArray(parsed) ? parsed : [parsed];
      if (studentsToAdd[0]?._extraWarnings) {
        studentsToAdd[0]._extraWarnings.forEach((w) => {
          if (!warningMessages.includes(w)) warningMessages.push(w);
        });
      }

      let duplicateCount = 0;

      studentsToAdd.forEach((student) => {
        const uniqueKey = student.studentId || student.displayCode;
        if (seen.has(uniqueKey)) {
          duplicateCount++;
          return;
        }
        seen.add(uniqueKey);
        result.push(student);
      });

      if (duplicateCount > 0 && !Array.isArray(parsed)) {
        warningMessages.push(`第 ${index + 1} 行重複，已略過：${parsed.studentId || parsed.displayCode}`);
      }
    });
  } else {
    if (fields.manualStudentInput) {
      fields.manualStudentInput.value.split(/\r?\n/).forEach((line, index) => {
        const parsed = parseManualInputLine(line, index + 1);
        if (!parsed) return;

        const toAdd = Array.isArray(parsed) ? parsed : [parsed];
        toAdd.forEach((p) => {
          if (p.warning) {
            warningMessages.push(p.warning);
            return;
          }

          const uniqueKey = p.studentId || p.displayCode;
          if (seen.has(uniqueKey)) {
            warningMessages.push(`第 ${index + 1} 行重複，已略過：${uniqueKey}`);
            return;
          }

          seen.add(uniqueKey);
          result.push(p);
        });
      });
    }
  }

  result.sort((a, b) => a.sortCode - b.sortCode || a.name.localeCompare(b.name, "zh-Hant"));

  return { students: result, warningMessages };
}

function groupByClass(students) {
  return students.reduce((groups, student) => {
    if (!groups.has(student.classCode)) groups.set(student.classCode, []);
    groups.get(student.classCode).push(student);
    return groups;
  }, new Map());
}

function renderActivityInfo() {
  const schedule = getResolvedSchedule();
  const school = fields.schoolName.value.trim() || "南寧高級中學";

  if (previewTargets.schoolEyebrow) {
    previewTargets.schoolEyebrow.textContent = school;
  }
  if (previewTargets.permitTitle) {
    previewTargets.permitTitle.textContent = `${school}學生公假單請示單`;
  }
  previewTargets.activityName.textContent = fields.activityName.value.trim() || "未填寫";
  previewTargets.applicantUnit.textContent = fields.applicantUnit.value.trim() || "未填寫";
  previewTargets.applicantName.textContent = fields.applicantName.value.trim() || "未填寫";
  previewTargets.activityDate.textContent = schedule.dateDisplay;
  previewTargets.timeRange.textContent = schedule.timeDisplay;
  previewTargets.location.textContent = fields.location.value.trim() || "未填寫";
  previewTargets.reason.textContent = fields.reason.value.trim() || "未填寫";
}

function renderStudents() {
  const { students, warningMessages } = getStudents();
  const groups = groupByClass(students);

  studentCount.textContent = `${students.length} 人`;
  classCount.textContent = `${groups.size} 班`;
  const loaded = getLoadedRoster();
  const baseCount = loaded.data.length;
  rosterStatus.textContent = `${loaded.isCustom ? "自訂" : "預設"}(${baseCount}筆)`;

  const clearRosterBtn = document.querySelector("#clearRosterBtn");
  if (clearRosterBtn) {
    clearRosterBtn.style.display = loaded.isCustom ? "inline-flex" : "none";
  }
  warnings.innerHTML = warningMessages.map((message) => `<div>${escapeHtml(message)}</div>`).join("");

  const statsDashboard = document.querySelector("#statsDashboard");
  const statsTotalCount = document.querySelector("#statsTotalCount");
  const statsClassCount = document.querySelector("#statsClassCount");
  const statsBreakdown = document.querySelector("#statsBreakdown");
  const previewStatsSummary = document.querySelector("#previewStatsSummary");

  if (!students.length) {
    groupedStudents.className = "empty-state";
    groupedStudents.textContent = "尚未輸入學生名單";
    mentorSignatures.innerHTML = '<div class="signature-line"><span>班級</span><b>導師簽名</b></div>';

    if (statsDashboard) {
      statsDashboard.style.display = "none";
    }
    if (previewStatsSummary) {
      previewStatsSummary.style.display = "none";
      previewStatsSummary.innerHTML = "";
    }

    renderSignInPreview(students, groups);
    if (activePreviewMode === "signin") {
      renderSignInPages(students, groups);
    } else {
      renderPrintPages(students, groups);
    }
    return;
  }

  groupedStudents.className = "";
  groupedStudents.innerHTML = [...groups.entries()]
    .map(([classCode, classStudents]) => renderClassBlock(classCode, classStudents))
    .join("");

  mentorSignatures.innerHTML = [...groups.keys()]
    .map(
      (classCode) =>
        `<div class="signature-line"><span>${escapeHtml(readableClass(classCode))}</span><b>導師簽名</b></div>`
    )
    .join("");

  renderSignInPreview(students, groups);
  if (activePreviewMode === "signin") {
    renderSignInPages(students, groups);
  } else {
    renderPrintPages(students, groups);
  }

  if (statsDashboard && statsTotalCount && statsClassCount && statsBreakdown) {
    statsDashboard.style.display = "block";
    statsTotalCount.textContent = `${students.length} 人`;
    statsClassCount.textContent = `${groups.size} 班`;

    const sortedClasses = [...groups.keys()].sort();
    statsBreakdown.innerHTML = sortedClasses
      .map((classCode) => {
        const count = groups.get(classCode).length;
        return `<span>${escapeHtml(readableClass(classCode))}：<b>${count}</b>人</span>`;
      })
      .join("");
  }

  if (previewStatsSummary) {
    previewStatsSummary.style.display = "block";
    const sortedClasses = [...groups.keys()].sort();
    const breakdownText = sortedClasses.map((code) => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
    previewStatsSummary.innerHTML = `公假統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）`;
  }
}

function renderClassBlock(classCode, classStudents) {
  const rows = classStudents
    .map(
      (student) => `
        <tr>
          <td class="class-name">${escapeHtml(readableClass(student.classCode))}</td>
          <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
          <td>${escapeHtml(maskName(student.name))}</td>
        </tr>
      `
    )
    .join("");

  return `
    <div class="class-block">
      <h4>${escapeHtml(readableClass(classCode))}</h4>
      <table class="student-table">
        <thead>
          <tr>
            <th class="class-name">班級</th>
            <th class="seat">座號</th>
            <th>姓名</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function getActivityDetails() {
  const schedule = getResolvedSchedule();
  return {
    schoolName: fields.schoolName.value.trim() || "南寧高級中學",
    activityName: fields.activityName.value.trim() || "未填寫",
    applicantUnit: fields.applicantUnit.value.trim() || "未填寫",
    applicantName: fields.applicantName.value.trim() || "未填寫",
    activityDate: schedule.dateDisplay,
    timeRange: schedule.timeDisplay,
    days: schedule.days,
    location: fields.location.value.trim() || "未填寫",
    reason: fields.reason.value.trim() || "未填寫",
  };
}

// 動態計算 A4 列印分頁容量，避免跨多班級導師簽名或多日時段明細爆版
function getPrintPaginationMetrics(details, classCountNum) {
  const timeLines = (details.timeRange || "").split("\n").length;
  const reasonText = details.reason || "";
  const reasonLines = Math.max(
    1,
    reasonText.split("\n").reduce((acc, line) => acc + Math.max(1, Math.ceil(line.length / 32)), 0)
  );

  const extraHeaderLines = Math.max(0, timeLines - 1) + Math.max(0, reasonLines - 1);
  const extraRowsPenalty = Math.ceil(extraHeaderLines * 0.75);

  const firstPageCapacity = Math.max(8, 18 - extraRowsPenalty);
  const middlePageCapacity = 27;

  // 導師簽名區為雙欄，每列高度約等於 1.35 列學生表格，外加核章欄基本保留高度
  const mentorGridRows = Math.ceil(Math.max(1, classCountNum) / 2);
  const extraMentorRows = Math.max(0, mentorGridRows - 2);
  const lastPageReserve = Math.min(22, 8 + Math.ceil(extraMentorRows * 1.35));

  return {
    firstPageCapacity,
    middlePageCapacity,
    lastPageReserve,
  };
}

function paginateClassBlocks(groups, firstPageCapacity, middlePageCapacity, lastPageReserve) {
  function getRemainingRows(blocks) {
    return blocks.reduce((sum, b) => sum + 1 + Math.ceil(b.students.length / 2), 0);
  }

  const sortedClasses = [...groups.keys()].sort();
  const blocks = sortedClasses.map((code) => ({
    classCode: code,
    students: [...groups.get(code)],
    isContinuation: false,
  }));

  const pages = [];
  let currentPage = [];
  let currentCapacity = firstPageCapacity;
  let remainingCapacity = currentCapacity;

  while (blocks.length > 0) {
    const totalRemainingRows = getRemainingRows(blocks);

    // 判斷本頁若作為最後一頁，扣除簽名保留區後是否放得下所有剩餘學生
    const maxRowsIfLastPage = remainingCapacity - lastPageReserve;
    let capacityLimit = remainingCapacity;

    if (totalRemainingRows > maxRowsIfLastPage && totalRemainingRows <= remainingCapacity + 1) {
      // 剩餘學生加上簽名區會超過本頁，故保留至少 2 列學生推到下一頁與簽名區一起印出
      capacityLimit = Math.max(2, Math.min(remainingCapacity, totalRemainingRows - 2));
    }

    if (capacityLimit <= 1) {
      pages.push(currentPage);
      currentPage = [];
      currentCapacity = middlePageCapacity;
      remainingCapacity = currentCapacity;
      continue;
    }

    const block = blocks[0];
    const rowsNeeded = 1 + Math.ceil(block.students.length / 2);

    if (rowsNeeded <= capacityLimit) {
      currentPage.push({
        classCode: block.classCode,
        students: block.students,
        isContinuation: block.isContinuation,
        hasContinuation: false,
      });
      remainingCapacity -= rowsNeeded;
      blocks.shift();
    } else {
      const maxStudentRows = capacityLimit - 1;
      const maxStudents = maxStudentRows * 2;

      if (maxStudents > 0) {
        const fitStudents = block.students.slice(0, maxStudents);
        currentPage.push({
          classCode: block.classCode,
          students: fitStudents,
          isContinuation: block.isContinuation,
          hasContinuation: true,
        });
        block.students = block.students.slice(maxStudents);
        block.isContinuation = true;
      }

      pages.push(currentPage);
      currentPage = [];
      currentCapacity = middlePageCapacity;
      remainingCapacity = currentCapacity;
    }
  }

  if (currentPage.length > 0) {
    pages.push(currentPage);
  }

  return pages;
}

function renderPrintClassBlock(classCode, students, isContinuation, hasContinuation) {
  const half = Math.ceil(students.length / 2);
  const leftStudents = students.slice(0, half);
  const rightStudents = students.slice(half);

  function renderColumnBody(colStudents) {
    return colStudents
      .map(
        (student) => `
        <tr>
          <td class="class-name">${escapeHtml(readableClass(student.classCode))}</td>
          <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
          <td>${escapeHtml(maskName(student.name))}</td>
        </tr>
      `
      )
      .join("");
  }

  const leftBody = renderColumnBody(leftStudents);
  const rightBody = renderColumnBody(rightStudents);

  return `
    <div class="print-class-section" style="margin-bottom: 12px; break-inside: avoid;">
      <h4 style="font-size: 0.92rem; margin-top: 0; margin-bottom: 6px; border-bottom: 1px solid #333; padding-bottom: 2px; text-align: left; font-weight: 700;">
        ${escapeHtml(readableClass(classCode))}${isContinuation ? "（續）" : ""}${hasContinuation ? "（接下頁）" : ""}
      </h4>
      <div class="print-student-grid" style="display: flex; gap: 20px;">
        <div class="print-student-col" style="flex: 1; min-width: 0;">
          <table class="student-table print-student-table">
            <thead>
              <tr>
                <th class="class-name">班級</th>
                <th class="seat">座號</th>
                <th>姓名</th>
              </tr>
            </thead>
            <tbody>${leftBody}</tbody>
          </table>
        </div>
        <div class="print-student-col" style="flex: 1; min-width: 0;">
          ${
            rightBody.length
              ? `
            <table class="student-table print-student-table">
              <thead>
                <tr>
                  <th class="class-name">班級</th>
                  <th class="seat">座號</th>
                  <th>姓名</th>
                </tr>
              </thead>
              <tbody>${rightBody}</tbody>
            </table>
          `
              : ""
          }
        </div>
      </div>
      ${
        hasContinuation
          ? `
        <div class="class-continuation-note" style="text-align: right; font-size: 0.8rem; font-weight: 700; margin-top: 6px; color: #333;">
          （以下接下頁）
        </div>
      `
          : ""
      }
    </div>
  `;
}

function renderPrintPages(students, groups) {
  const details = getActivityDetails();
  const { firstPageCapacity, middlePageCapacity, lastPageReserve } = getPrintPaginationMetrics(details, groups.size);
  const pages = paginateClassBlocks(groups, firstPageCapacity, middlePageCapacity, lastPageReserve);
  const totalPages = Math.max(1, pages.length);
  const school = details.schoolName;

  const sortedClasses = [...groups.keys()].sort();
  const breakdownText = sortedClasses.map((code) => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
  const statsSummaryHtml = students.length
    ? `<div class="print-stats-summary" style="font-size: 0.85rem; margin-bottom: 8px; color: #333; font-weight: 700;">公假統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）</div>`
    : "";

  const effectivePages = pages.length ? pages : [[]];

  printPages.innerHTML = effectivePages
    .map((pageBlocks, pageIndex) => {
      const isFirst = pageIndex === 0;
      const isLast = pageIndex === totalPages - 1;

      const blocksHtml = pageBlocks
        .map((pb) => renderPrintClassBlock(pb.classCode, pb.students, pb.isContinuation, pb.hasContinuation))
        .join("");

      return `
        <article class="print-page">
          <header class="print-header">
            <h2>${escapeHtml(school)}學生公假單請示單</h2>
          </header>
          ${isFirst ? renderPrintInfoTable() : renderContinuationLabel()}
          <section class="print-students">
            <h3>公假學生名單${!isFirst ? "（續）" : ""}</h3>
            ${isFirst ? statsSummaryHtml : ""}
            ${students.length ? blocksHtml : '<div class="empty-state">尚未輸入學生名單</div>'}
          </section>
          ${isLast ? renderPrintSignatures([...groups.keys()]) : ""}
          <footer class="print-footer">第 ${pageIndex + 1} 頁，共 ${totalPages} 頁</footer>
        </article>
      `;
    })
    .join("");
}

function renderPrintInfoTable() {
  const details = getActivityDetails();
  return `
    <section class="info-table print-info-table">
      <div><strong>活動名稱</strong><span>${escapeHtml(details.activityName)}</span></div>
      <div><strong>申請單位</strong><span>${escapeHtml(details.applicantUnit)}</span></div>
      <div><strong>申請人</strong><span>${escapeHtml(details.applicantName)}</span></div>
      <div><strong>活動地點</strong><span>${escapeHtml(details.location)}</span></div>
      <div><strong>活動日期</strong><span>${escapeHtml(details.activityDate)}</span></div>
      <div><strong>活動時間</strong><span>${escapeHtml(details.timeRange)}</span></div>
      <div class="full"><strong>活動事由</strong><span>${escapeHtml(details.reason)}</span></div>
    </section>
  `;
}

function renderContinuationLabel(customDateLabel) {
  const details = getActivityDetails();
  const dateLabel = customDateLabel || details.activityDate;
  return `
    <section class="continuation-label">
      <span>${escapeHtml(details.activityName)}</span>
      <span>${escapeHtml(dateLabel)}</span>
    </section>
  `;
}

function renderPrintSignatures(classCodes) {
  const signatures = classCodes.length
    ? classCodes
        .map(
          (classCode) =>
            `<div class="signature-line"><span>${escapeHtml(readableClass(classCode))}</span><b>導師簽名</b></div>`
        )
        .join("")
    : '<div class="signature-line"><span>班級</span><b>導師簽名</b></div>';

  return `
    <section class="mentor-section print-mentor-section">
      <h3>導師簽名</h3>
      <div class="mentor-grid">${signatures}</div>
    </section>
    <section class="approval-section print-approval-section">
      <div>申請單位</div>
      <div>生輔組</div>
      <div>學務處</div>
      <div>校長</div>
    </section>
  `;
}

function renderSignInInfoTable(dayInfo, dayIndex, totalDays) {
  const details = getActivityDetails();
  const dateText = dayInfo
    ? `${dayInfo.formattedDateWithWeekday || dayInfo.formattedDate}${totalDays > 1 ? `（第 ${dayIndex + 1}/${totalDays} 日）` : ""}`
    : details.activityDate;
  const timeText = dayInfo ? dayInfo.timeText : details.timeRange;

  return `
    <section class="info-table print-info-table signin-info-table">
      <div><strong>活動名稱</strong><span>${escapeHtml(details.activityName)}</span></div>
      <div><strong>申請單位</strong><span>${escapeHtml(details.applicantUnit)}</span></div>
      <div><strong>活動日期</strong><span>${escapeHtml(dateText)}</span></div>
      <div><strong>公假時段</strong><span>${escapeHtml(timeText)}</span></div>
      <div><strong>活動地點</strong><span>${escapeHtml(details.location)}</span></div>
      <div><strong>帶隊教師</strong><span>${escapeHtml(details.applicantName)}</span></div>
      <div class="full"><strong>帶隊教師簽章</strong><span style="border: 0; background: transparent;"></span></div>
    </section>
  `;
}

function paginateSignInClassBlocks(groups, firstPageCapacity, middlePageCapacity) {
  const sortedClasses = [...groups.keys()].sort();
  const blocks = sortedClasses.map((code) => ({
    classCode: code,
    students: [...groups.get(code)],
    isContinuation: false,
  }));

  const pages = [];
  let currentPage = [];
  let currentCapacity = firstPageCapacity;
  let remainingCapacity = currentCapacity;

  while (blocks.length > 0) {
    if (remainingCapacity <= 1) {
      pages.push(currentPage);
      currentPage = [];
      currentCapacity = middlePageCapacity;
      remainingCapacity = currentCapacity;
      continue;
    }

    const block = blocks[0];
    const rowsNeeded = 1 + block.students.length;

    if (rowsNeeded <= remainingCapacity) {
      currentPage.push({
        classCode: block.classCode,
        students: block.students,
        isContinuation: block.isContinuation,
        hasContinuation: false,
      });
      remainingCapacity -= rowsNeeded;
      blocks.shift();
    } else {
      const maxStudents = remainingCapacity - 1;

      if (maxStudents > 0) {
        const fitStudents = block.students.slice(0, maxStudents);
        currentPage.push({
          classCode: block.classCode,
          students: fitStudents,
          isContinuation: block.isContinuation,
          hasContinuation: true,
        });
        block.students = block.students.slice(maxStudents);
        block.isContinuation = true;
      }

      pages.push(currentPage);
      currentPage = [];
      currentCapacity = middlePageCapacity;
      remainingCapacity = currentCapacity;
    }
  }

  if (currentPage.length > 0) {
    pages.push(currentPage);
  }

  return pages;
}

function renderSignInClassBlock(classCode, students, isContinuation, hasContinuation) {
  const body = students
    .map(
      (student) => `
      <tr>
        <td class="class-name">${escapeHtml(readableClass(student.classCode))}</td>
        <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
        <td class="student-id">${escapeHtml(student.studentId || "-")}</td>
        <td class="student-name" style="font-weight: 700;">${escapeHtml(student.name)}</td>
        <td class="signin-box"></td>
        <td class="signin-box"></td>
      </tr>
    `
    )
    .join("");

  return `
    <div class="print-class-section signin-class-section" style="margin-bottom: 16px; break-inside: avoid;">
      <h4 style="font-size: 0.92rem; margin-top: 0; margin-bottom: 6px; border-bottom: 1px solid #333; padding-bottom: 2px; text-align: left; font-weight: 700;">
        ${escapeHtml(readableClass(classCode))}${isContinuation ? "（續）" : ""}${hasContinuation ? "（接下頁）" : ""}
      </h4>
      <table class="student-table print-student-table signin-student-table">
        <thead>
          <tr>
            <th class="class-name">班級</th>
            <th class="seat">座號</th>
            <th style="width: 90px;">學號</th>
            <th style="width: 90px;">姓名</th>
            <th style="width: 120px; text-align: center;">簽到</th>
            <th style="width: 120px; text-align: center;">簽退</th>
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
      ${
        hasContinuation
          ? `
        <div class="class-continuation-note" style="text-align: right; font-size: 0.8rem; font-weight: 700; margin-top: 6px; color: #333;">
          （以下接下頁）
        </div>
      `
          : ""
      }
    </div>
  `;
}

// 螢幕即時預覽簽到表（支援多日卡片檢視）
function renderSignInPreview(students, groups) {
  if (!signInPreviewEl) return;
  const details = getActivityDetails();
  const school = details.schoolName;
  const scheduleDays =
    details.days && details.days.length > 0
      ? details.days
      : [
          {
            date: "",
            formattedDate: details.activityDate,
            formattedDateWithWeekday: details.activityDate,
            shortDate: details.activityDate,
            timeText: details.timeRange,
          },
        ];

  const sortedClasses = [...groups.keys()].sort();
  const breakdownText = sortedClasses.map((code) => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
  const statsSummaryHtml = students.length
    ? `<div class="preview-stats-summary">簽到統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）</div>`
    : "";

  const blocksHtml = sortedClasses
    .map((classCode) => renderSignInClassBlock(classCode, groups.get(classCode), false, false))
    .join("");

  const totalDays = scheduleDays.length;
  signInPreviewEl.innerHTML = scheduleDays
    .map((dayInfo, dayIndex) => {
      const dayBadge = totalDays > 1 ? `（${dayInfo.shortDate}）` : "";
      return `
        <article class="permit-paper">
          <header class="permit-header">
            <h2>${escapeHtml(school)}學生公假活動簽到表</h2>
          </header>
          ${renderSignInInfoTable(dayInfo, dayIndex, totalDays)}
          <section class="student-preview">
            <h3>簽到學生名單${escapeHtml(dayBadge)}</h3>
            ${statsSummaryHtml}
            ${students.length ? blocksHtml : '<div class="empty-state">尚未輸入學生名單</div>'}
          </section>
        </article>
      `;
    })
    .join("");
}

// 方案一：跨日活動時，依每一天獨立產生一份簽到表
function renderSignInPages(students, groups) {
  const details = getActivityDetails();
  const school = details.schoolName;
  const scheduleDays =
    details.days && details.days.length > 0
      ? details.days
      : [
          {
            date: "",
            formattedDate: details.activityDate,
            formattedDateWithWeekday: details.activityDate,
            shortDate: details.activityDate,
            timeText: details.timeRange,
          },
        ];

  const sortedClasses = [...groups.keys()].sort();
  const breakdownText = sortedClasses.map((code) => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
  const statsSummaryHtml = students.length
    ? `<div class="print-stats-summary" style="font-size: 0.85rem; margin-bottom: 8px; color: #333; font-weight: 700;">簽到統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）</div>`
    : "";

  const totalDays = scheduleDays.length;
  const allDayArticles = [];

  scheduleDays.forEach((dayInfo, dayIndex) => {
    const pages = paginateSignInClassBlocks(groups, 11, 18);
    const effectivePages = pages.length ? pages : [[]];
    const totalPagesForDay = effectivePages.length;
    const dayHeaderSuffix = totalDays > 1 ? `（${dayInfo.shortDate}）` : "";

    effectivePages.forEach((pageBlocks, pageIndex) => {
      const isFirst = pageIndex === 0;
      const blocksHtml = pageBlocks
        .map((pb) => renderSignInClassBlock(pb.classCode, pb.students, pb.isContinuation, pb.hasContinuation))
        .join("");

      const footerPrefix = totalDays > 1 ? `${escapeHtml(dayInfo.formattedDateWithWeekday)} — ` : "";

      allDayArticles.push(`
        <article class="print-page">
          <header class="print-header">
            <h2>${escapeHtml(school)}學生公假活動簽到表</h2>
          </header>
          ${
            isFirst
              ? renderSignInInfoTable(dayInfo, dayIndex, totalDays)
              : renderContinuationLabel(`${dayInfo.formattedDateWithWeekday}（${dayInfo.timeText}）`)
          }
          <section class="print-students">
            <h3>簽到學生名單${escapeHtml(dayHeaderSuffix)}${!isFirst ? "（續）" : ""}</h3>
            ${isFirst ? statsSummaryHtml : ""}
            ${students.length ? blocksHtml : '<div class="empty-state">尚未輸入學生名單</div>'}
          </section>
          <footer class="print-footer">${footerPrefix}第 ${pageIndex + 1} 頁，共 ${totalPagesForDay} 頁</footer>
        </article>
      `);
    });
  });

  printPages.innerHTML = allDayArticles.join("");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function saveState() {
  const data = Object.fromEntries(Object.entries(fields).map(([key, field]) => [key, field ? field.value : ""]));
  data.activeTabMode = getActiveInputMode();
  data.scheduleMode = scheduleMode;
  data.uniformPeriods = uniformPeriods;
  data.dailySchedules = dailySchedules;
  data.keepApplicant = keepApplicantCheck ? Boolean(keepApplicantCheck.checked) : true;
  localStorage.setItem(storageKey, JSON.stringify(data));
}

function loadState() {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return;

  try {
    const data = JSON.parse(raw);
    Object.entries(fields).forEach(([key, field]) => {
      if (field && typeof data[key] === "string") field.value = data[key];
    });

    if (Array.isArray(data.uniformPeriods)) {
      uniformPeriods = data.uniformPeriods;
    }
    if (Array.isArray(data.dailySchedules)) {
      dailySchedules = data.dailySchedules;
    }
    if (data.scheduleMode === "daily" || data.scheduleMode === "uniform") {
      scheduleMode = data.scheduleMode;
    }
    if (keepApplicantCheck && typeof data.keepApplicant === "boolean") {
      keepApplicantCheck.checked = data.keepApplicant;
    }

    if (data.activeTabMode) {
      switchTabMode(data.activeTabMode);
    }
  } catch (e) {
    console.error("Failed to load state", e);
    localStorage.removeItem(storageKey);
  }
}

function update() {
  renderActivityInfo();
  renderStudents();
  saveState();
}

function loadSample() {
  fields.schoolName.value = "南寧高級中學";
  fields.activityName.value = "校園交通安全宣導勤務";
  fields.applicantUnit.value = "學務處生活輔導組";
  fields.applicantName.value = "賴祥宇";
  fields.activityDate.value = "2026-06-23";
  fields.activityEndDate.value = "2026-06-23";
  fields.startTime.value = "08:10";
  fields.endTime.value = "10:00";
  fields.location.value = "校門口及活動中心";
  fields.reason.value = "協助交通安全宣導、場地引導及活動秩序維護。";

  uniformPeriods = ["p1", "p2"];
  dailySchedules = [];
  setScheduleMode("uniform");

  if (fields.manualStudentInput) fields.manualStudentInput.value = "";
  fields.studentInput.value = ["40101 蘇大輔", "401-02 賴小祥", "40201 蘇小祥", "402-02 賴大寶"].join("\n");
  switchTabMode("roster");
  update();
}

function clearAll() {
  const shouldKeepApplicant = keepApplicantCheck && keepApplicantCheck.checked;
  const savedSchool = fields.schoolName ? fields.schoolName.value : "南寧高級中學";
  const savedUnit = shouldKeepApplicant && fields.applicantUnit ? fields.applicantUnit.value : "";
  const savedName = shouldKeepApplicant && fields.applicantName ? fields.applicantName.value : "";

  Object.values(fields).forEach((field) => {
    if (field) field.value = "";
  });
  uniformPeriods = [];
  dailySchedules = [];
  localStorage.removeItem(storageKey);

  fields.schoolName.value = savedSchool || "南寧高級中學";
  if (fields.applicantUnit) fields.applicantUnit.value = savedUnit;
  if (fields.applicantName) fields.applicantName.value = savedName;

  const todayStr = getTodayIso();
  fields.activityDate.value = todayStr;
  fields.activityEndDate.value = todayStr;

  setScheduleMode("uniform");
  update();
}

Object.entries(fields).forEach(([key, field]) => {
  if (!field) return;
  field.addEventListener("input", () => {
    if (key === "activityDate") {
      if (!fields.activityEndDate.value || fields.activityEndDate.value < fields.activityDate.value) {
        fields.activityEndDate.value = fields.activityDate.value;
      }
      if (scheduleMode === "daily") {
        syncDailySchedulesFromRange();
        renderDailyScheduleList();
      }
    } else if (key === "activityEndDate") {
      if (fields.activityEndDate.value && fields.activityDate.value && fields.activityEndDate.value < fields.activityDate.value) {
        fields.activityDate.value = fields.activityEndDate.value;
      }
      if (scheduleMode === "daily") {
        syncDailySchedulesFromRange();
        renderDailyScheduleList();
      }
    }
    update();
  });
});

if (keepApplicantCheck) {
  keepApplicantCheck.addEventListener("change", saveState);
}

const previewPermitTab = document.querySelector("#previewPermitTab");
const previewSignInTab = document.querySelector("#previewSignInTab");
if (previewPermitTab) {
  previewPermitTab.addEventListener("click", () => setPreviewMode("permit"));
}
if (previewSignInTab) {
  previewSignInTab.addEventListener("click", () => setPreviewMode("signin"));
}

const uniformScheduleTab = document.querySelector("#uniformScheduleTab");
const dailyScheduleTab = document.querySelector("#dailyScheduleTab");
if (uniformScheduleTab) {
  uniformScheduleTab.addEventListener("click", () => setScheduleMode("uniform"));
}
if (dailyScheduleTab) {
  dailyScheduleTab.addEventListener("click", () => setScheduleMode("daily"));
}

const uniformPeriodGroup = document.querySelector("#uniformPeriodGroup");
if (uniformPeriodGroup) {
  uniformPeriodGroup.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-period-id]");
    if (!btn) return;
    const periodId = btn.getAttribute("data-period-id");
    const idx = uniformPeriods.indexOf(periodId);
    if (idx === -1) {
      uniformPeriods.push(periodId);
    } else {
      uniformPeriods.splice(idx, 1);
    }
    renderUniformPeriodButtons();
    update();
  });
}

const uniformQuickActions = document.querySelector('.period-quick-actions[data-target="uniform"]');
if (uniformQuickActions) {
  uniformQuickActions.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-preset]");
    if (!btn) return;
    const preset = btn.getAttribute("data-preset");
    if (PERIOD_PRESETS[preset]) {
      uniformPeriods = [...PERIOD_PRESETS[preset]];
      renderUniformPeriodButtons();
      update();
    }
  });
}

const clearTimeBtn = document.querySelector("#clearTimeBtn");
if (clearTimeBtn) {
  clearTimeBtn.addEventListener("click", () => {
    fields.startTime.value = "";
    fields.endTime.value = "";
    update();
  });
}

const dailyScheduleListEl = document.querySelector("#dailyScheduleList");
if (dailyScheduleListEl) {
  dailyScheduleListEl.addEventListener("click", (e) => {
    const chipBtn = e.target.closest("button.period-chip[data-daily-index]");
    if (chipBtn) {
      const dIdx = Number(chipBtn.getAttribute("data-daily-index"));
      const periodId = chipBtn.getAttribute("data-period-id");
      const dayItem = dailySchedules[dIdx];
      if (!dayItem) return;
      const pIdx = dayItem.periods.indexOf(periodId);
      if (pIdx === -1) {
        dayItem.periods.push(periodId);
      } else {
        dayItem.periods.splice(pIdx, 1);
      }
      renderDailyScheduleList();
      update();
      return;
    }

    const presetBtn = e.target.closest("button.quick-btn[data-daily-index]");
    if (presetBtn) {
      const dIdx = Number(presetBtn.getAttribute("data-daily-index"));
      const preset = presetBtn.getAttribute("data-preset");
      const dayItem = dailySchedules[dIdx];
      if (dayItem && PERIOD_PRESETS[preset]) {
        dayItem.periods = [...PERIOD_PRESETS[preset]];
        renderDailyScheduleList();
        update();
      }
      return;
    }

    const removeBtn = e.target.closest("button[data-remove-daily]");
    if (removeBtn) {
      const dIdx = Number(removeBtn.getAttribute("data-remove-daily"));
      if (dailySchedules.length > 1) {
        dailySchedules.splice(dIdx, 1);
        syncDateInputsFromDailySchedules();
        renderDailyScheduleList();
        update();
      }
    }
  });

  dailyScheduleListEl.addEventListener("input", (e) => {
    const target = e.target;
    const dIdx = Number(target.getAttribute("data-daily-index"));
    const dayItem = dailySchedules[dIdx];
    if (!dayItem) return;

    if (target.classList.contains("daily-date-input")) {
      dayItem.date = target.value;
      syncDateInputsFromDailySchedules();
      renderDailyScheduleList();
      update();
    } else if (target.classList.contains("daily-start-time")) {
      dayItem.startTime = target.value;
      update();
    } else if (target.classList.contains("daily-end-time")) {
      dayItem.endTime = target.value;
      update();
    }
  });
}

const syncDailyRangeBtn = document.querySelector("#syncDailyRangeBtn");
if (syncDailyRangeBtn) {
  syncDailyRangeBtn.addEventListener("click", () => {
    syncDailySchedulesFromRange(true);
    renderDailyScheduleList();
    update();
  });
}

const addDailyRowBtn = document.querySelector("#addDailyRowBtn");
if (addDailyRowBtn) {
  addDailyRowBtn.addEventListener("click", () => {
    let nextDateStr = getTodayIso();
    if (dailySchedules.length > 0) {
      const lastDateStr = dailySchedules[dailySchedules.length - 1].date;
      if (lastDateStr) {
        const [y, m, d] = lastDateStr.split("-").map(Number);
        const dt = new Date(y, m - 1, d);
        dt.setDate(dt.getDate() + 1);
        nextDateStr = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
      }
    }
    dailySchedules.push({
      date: nextDateStr,
      periods: [...uniformPeriods],
      startTime: fields.startTime.value || "",
      endTime: fields.endTime.value || "",
    });
    syncDateInputsFromDailySchedules();
    renderDailyScheduleList();
    update();
  });
}

document.querySelector("#loadSampleBtn").addEventListener("click", loadSample);
document.querySelector("#clearBtn").addEventListener("click", clearAll);
document.querySelector("#printBtn").addEventListener("click", () => {
  const { students } = getStudents();
  const groups = groupByClass(students);
  renderPrintPages(students, groups);
  window.print();
});

document.querySelector("#printSignInBtn").addEventListener("click", () => {
  const { students } = getStudents();
  const groups = groupByClass(students);
  renderSignInPages(students, groups);
  window.print();
});

const excelUpload = document.querySelector("#excelUpload");
const clearRosterBtn = document.querySelector("#clearRosterBtn");

if (excelUpload) {
  excelUpload.addEventListener("change", handleExcelUpload);
}

function handleClearRoster() {
  if (confirm("確定要清除自訂的名冊資料，還原為預設範例名冊嗎？")) {
    localStorage.removeItem(customRosterStorageKey);
    if (excelUpload) excelUpload.value = "";
    update();
  }
}

if (clearRosterBtn) {
  clearRosterBtn.addEventListener("click", handleClearRoster);
}

// Toggle Guide Panel
const toggleGuideBtn = document.querySelector("#toggleGuideBtn");
const guideContent = document.querySelector("#guideContent");
const guidePanel = document.querySelector(".guide-panel");
const guideStorageKey = "nanning-guide-collapsed-v1";

function setGuideState(isCollapsed) {
  if (isCollapsed) {
    guideContent.classList.add("collapsed");
    guidePanel.classList.add("collapsed-panel");
    toggleGuideBtn.textContent = "展開指引";
    toggleGuideBtn.setAttribute("aria-expanded", "false");
  } else {
    guideContent.classList.remove("collapsed");
    guidePanel.classList.remove("collapsed-panel");
    toggleGuideBtn.textContent = "收合指引";
    toggleGuideBtn.setAttribute("aria-expanded", "true");
  }
}

if (toggleGuideBtn && guideContent && guidePanel) {
  const isCollapsed = localStorage.getItem(guideStorageKey) === "true";
  setGuideState(isCollapsed);

  toggleGuideBtn.addEventListener("click", () => {
    const currentlyCollapsed = guideContent.classList.contains("collapsed");
    const newState = !currentlyCollapsed;
    setGuideState(newState);
    localStorage.setItem(guideStorageKey, String(newState));
  });
}

// Switch Tab Mode Function
function switchTabMode(mode) {
  const useRosterTab = document.querySelector("#useRosterTab");
  const manualInputTab = document.querySelector("#manualInputTab");
  const useRosterMode = document.querySelector("#useRosterMode");
  const manualInputMode = document.querySelector("#manualInputMode");

  if (!useRosterTab || !manualInputTab || !useRosterMode || !manualInputMode) return;

  if (mode === "manual") {
    useRosterTab.classList.remove("active");
    manualInputTab.classList.add("active");
    useRosterMode.style.display = "none";
    manualInputMode.style.display = "block";
  } else {
    useRosterTab.classList.add("active");
    manualInputTab.classList.remove("active");
    useRosterMode.style.display = "block";
    manualInputMode.style.display = "none";
  }
  update();
}

const useRosterTab = document.querySelector("#useRosterTab");
const manualInputTab = document.querySelector("#manualInputTab");

if (useRosterTab) {
  useRosterTab.addEventListener("click", () => switchTabMode("roster"));
}
if (manualInputTab) {
  manualInputTab.addEventListener("click", () => switchTabMode("manual"));
}

// 初始化載入
loadState();
const defaultToday = getTodayIso();
if (fields.activityDate && !fields.activityDate.value) {
  fields.activityDate.value = defaultToday;
}
if (fields.activityEndDate && (!fields.activityEndDate.value || fields.activityEndDate.value < fields.activityDate.value)) {
  fields.activityEndDate.value = fields.activityDate.value;
}
setScheduleMode(scheduleMode);
