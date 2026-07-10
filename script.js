const fields = {
  schoolName: document.querySelector("#schoolName"),
  activityName: document.querySelector("#activityName"),
  applicantUnit: document.querySelector("#applicantUnit"),
  applicantName: document.querySelector("#applicantName"),
  activityDate: document.querySelector("#activityDate"),
  location: document.querySelector("#location"),
  startTime: document.querySelector("#startTime"),
  endTime: document.querySelector("#endTime"),
  reason: document.querySelector("#reason"),
  manualStudentInput: document.querySelector("#manualStudentInput"),
  studentInput: document.querySelector("#studentInput"),
};

const previewTargets = {
  schoolEyebrow: document.querySelector('[data-preview="schoolEyebrow"]'),
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

const storageKey = "nanning-public-leave-form-v1";
const customRosterStorageKey = "nanning-custom-roster-v1";

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

function handleExcelUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
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
        studentIdKey: keys.find(k => k.trim() === "學號"),
        classCodeKey: keys.find(k => k.trim() === "班級"),
        seatNumberKey: keys.find(k => k.trim() === "座號"),
        nameKey: keys.find(k => k.trim() === "姓名"),
      };

      if (!headerMap.studentIdKey || !headerMap.classCodeKey || !headerMap.seatNumberKey || !headerMap.nameKey) {
        alert("Excel 欄位格式錯誤！必須包含以下欄位：\n「學號」、「班級」、「座號」、「姓名」\n\n請下載名冊範例 Excel 對照格式。");
        return;
      }

      const rosterData = jsonData.map((row) => {
        return {
          studentId: String(row[headerMap.studentIdKey] || "").trim(),
          classCode: String(row[headerMap.classCodeKey] || "").trim(),
          seatNumber: Number(row[headerMap.seatNumberKey]) || 0,
          name: String(row[headerMap.nameKey] || "").trim()
        };
      }).filter(item => item.name && item.classCode);

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
  const compact = value.replace(/\s+/g, "");
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
  const classText = String(classCode || "").trim();
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

function readableClass(classCode) {
  if (!classCode || classCode.length !== 3) return classCode || "未分班";
  const grade = Number(classCode[0]);
  const room = Number(classCode.slice(1));
  return `${grade}年${room}班`;
}

function splitLine(line) {
  return line
    .replace(/[，、\t]/g, ",")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
}

function parseManualInputLine(line, lineNumber) {
  const trimmed = line.trim();
  if (!trimmed) return null;

  const parts = splitLine(trimmed);
  const tokens = parts.length > 1 ? parts : trimmed.split(/\s+/).filter(Boolean);

  if (tokens.length >= 3) {
    let studentId = "";
    let seatCodeRaw = "";
    let name = "";

    let seatIndex = tokens.findIndex(t => normalizeSeatCode(t) !== null);
    if (seatIndex !== -1) {
      seatCodeRaw = tokens[seatIndex];
      const remaining = tokens.filter((_, idx) => idx !== seatIndex);
      let idIndex = remaining.findIndex(t => /^\d{5,}$/.test(t));
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
        warning: `第 ${lineNumber} 行手動格式錯誤（座號格式不對，應如 40101 或 401-01）：${trimmed}`
      };
    }

    return {
      studentId,
      name: name || "未填姓名",
      ...seat,
      sourceLine: lineNumber
    };
  } else if (tokens.length === 2) {
    const seatCodeRaw = tokens[0];
    const name = tokens[1];
    const seat = normalizeSeatCode(seatCodeRaw);
    if (!seat) {
      if (/^\d{5,}$/.test(name)) {
        return {
          studentId: name,
          name: seatCodeRaw,
          classCode: "未分班",
          seatNumber: 0,
          displayCode: "未分班",
          sortCode: 999999,
          sourceLine: lineNumber
        };
      }
      return {
        warning: `第 ${lineNumber} 行手動格式錯誤（應包含班級座號與姓名，如「40101 蘇大輔」）：${trimmed}`
      };
    }
    return {
      studentId: "",
      name,
      ...seat,
      sourceLine: lineNumber
    };
  } else {
    return {
      warning: `第 ${lineNumber} 行格式不足（手動輸入每行需至少包含班級座號與姓名，如「40101 蘇大輔」）：${trimmed}`
    };
  }
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
      if (parsed && !parsed.warning) {
        addRecord(parsed);
      }
    });
  }

  return { byStudentId, bySeatCode, byName, byClass };
}

function parseStudentLine(line, roster, lineNumber) {
  const trimmed = line.trim();
  if (!trimmed) return null;

  if (/^\d{3}$/.test(trimmed)) {
    const classStudents = roster.byClass.get(trimmed);
    if (classStudents && classStudents.length > 0) {
      return classStudents.map(student => ({
        ...student,
        sourceLine: lineNumber
      }));
    }
    return {
      warning: `第 ${lineNumber} 行查無此班級代號：${trimmed}`,
    };
  }

  const parts = splitLine(trimmed);
  const tokens = parts.length > 1 ? parts : trimmed.split(/\s+/).filter(Boolean);
  const firstToken = tokens[0] || "";
  const seat = normalizeSeatCode(firstToken);

  if (seat) {
    const rosterStudent = roster.bySeatCode.get(seat.displayCode) || roster.bySeatCode.get(seat.displayCode.replace("-", ""));
    if (rosterStudent && tokens.length === 1) {
      return {
        ...rosterStudent,
        sourceLine: lineNumber,
      };
    }

    return {
      ...seat,
      studentId: rosterStudent?.studentId || "",
      name: tokens.slice(1).join(" ").trim() || rosterStudent?.name || "未填姓名",
      sourceLine: lineNumber,
    };
  }

  const rosterStudent = roster.byStudentId.get(firstToken.toUpperCase());
  if (rosterStudent) {
    return {
      ...rosterStudent,
      sourceLine: lineNumber,
    };
  }

  const nameMatches = roster.byName.get(normalizeName(trimmed)) || [];
  if (nameMatches.length === 1) {
    return {
      ...nameMatches[0],
      sourceLine: lineNumber,
    };
  }

  if (nameMatches.length > 1) {
    const candidates = nameMatches
      .slice(0, 5)
      .map((student) => `${student.displayCode}/${student.studentId}`)
      .join("、");
    return {
      warning: `第 ${lineNumber} 行姓名重複，請改用學號或班級座號：${trimmed}（${candidates}）`,
    };
  }

  return {
    warning: `第 ${lineNumber} 行無法辨識：${trimmed}`,
  };
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

        if (parsed.warning) {
          warningMessages.push(parsed.warning);
          return;
        }

        const uniqueKey = parsed.studentId || parsed.displayCode;
        if (seen.has(uniqueKey)) {
          warningMessages.push(`第 ${index + 1} 行重複，已略過：${uniqueKey}`);
          return;
        }

        seen.add(uniqueKey);
        result.push(parsed);
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
  const date = fields.activityDate.value ? formatDate(fields.activityDate.value) : "未填寫";
  const start = fields.startTime.value || "";
  const end = fields.endTime.value || "";
  const timeRange = start && end ? `${start} 至 ${end}` : start || end || "未填寫";

  const school = fields.schoolName.value.trim() || "南寧高級中學";
  if (previewTargets.schoolEyebrow) {
    previewTargets.schoolEyebrow.textContent = school;
  }
  previewTargets.activityName.textContent = fields.activityName.value.trim() || "未填寫";
  previewTargets.applicantUnit.textContent = fields.applicantUnit.value.trim() || "未填寫";
  previewTargets.applicantName.textContent = fields.applicantName.value.trim() || "未填寫";
  previewTargets.activityDate.textContent = date;
  previewTargets.timeRange.textContent = timeRange;
  previewTargets.location.textContent = fields.location.value.trim() || "未填寫";
  previewTargets.reason.textContent = fields.reason.value.trim() || "未填寫";
}

function formatDate(value) {
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return value;
  return `${Number(year) - 1911} 年 ${Number(month)} 月 ${Number(day)} 日`;
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

  if (!students.length) {
    groupedStudents.className = "empty-state";
    groupedStudents.textContent = "尚未輸入學生名單";
    mentorSignatures.innerHTML = '<div class="signature-line"><span>班級</span><b>導師簽名</b></div>';
    
    const previewStatsSummary = document.querySelector("#previewStatsSummary");
    if (previewStatsSummary) {
      previewStatsSummary.style.display = "none";
      previewStatsSummary.innerHTML = "";
    }

    renderPrintPages(students, groups);
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
  renderPrintPages(students, groups);

  // Update Statistics Dashboard
  const statsDashboard = document.querySelector("#statsDashboard");
  const statsTotalCount = document.querySelector("#statsTotalCount");
  const statsClassCount = document.querySelector("#statsClassCount");
  const statsBreakdown = document.querySelector("#statsBreakdown");

  if (statsDashboard && statsTotalCount && statsClassCount && statsBreakdown) {
    if (students.length > 0) {
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
    } else {
      statsDashboard.style.display = "none";
    }
  }

  // Update Preview Stats Summary (On-screen Preview)
  const previewStatsSummary = document.querySelector("#previewStatsSummary");
  if (previewStatsSummary) {
    if (students.length > 0) {
      previewStatsSummary.style.display = "block";
      const sortedClasses = [...groups.keys()].sort();
      const breakdownText = sortedClasses.map(code => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
      previewStatsSummary.innerHTML = `公假統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）`;
    } else {
      previewStatsSummary.style.display = "none";
      previewStatsSummary.innerHTML = "";
    }
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
  const date = fields.activityDate.value ? formatDate(fields.activityDate.value) : "未填寫";
  const start = fields.startTime.value || "";
  const end = fields.endTime.value || "";
  return {
    schoolName: fields.schoolName.value.trim() || "南寧高級中學",
    activityName: fields.activityName.value.trim() || "未填寫",
    applicantUnit: fields.applicantUnit.value.trim() || "未填寫",
    applicantName: fields.applicantName.value.trim() || "未填寫",
    activityDate: date,
    timeRange: start && end ? `${start} 至 ${end}` : start || end || "未填寫",
    location: fields.location.value.trim() || "未填寫",
    reason: fields.reason.value.trim() || "未填寫",
  };
}

function flattenStudentRows(groups) {
  const rows = [];
  [...groups.entries()].forEach(([classCode, classStudents]) => {
    rows.push({ type: "class", classCode });
    classStudents.forEach((student) => rows.push({ type: "student", student }));
  });
  return rows;
}
function paginateClassBlocks(groups, firstPageCapacity, middlePageCapacity, lastPageReserve) {
  function getRemainingRows(blocks) {
    return blocks.reduce((sum, b) => sum + 1 + Math.ceil(b.students.length / 2), 0);
  }

  const sortedClasses = [...groups.keys()].sort();
  const blocks = sortedClasses.map(code => ({
    classCode: code,
    students: [...groups.get(code)],
    isContinuation: false
  }));

  const pages = [];
  let currentPage = [];
  let currentCapacity = firstPageCapacity;
  let remainingCapacity = currentCapacity;

  while (blocks.length > 0) {
    const totalRemainingRows = getRemainingRows(blocks);

    let capacityLimit = remainingCapacity;
    if (totalRemainingRows > remainingCapacity && totalRemainingRows - remainingCapacity <= lastPageReserve) {
      capacityLimit = Math.max(2, remainingCapacity - lastPageReserve);
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
        hasContinuation: false
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
          hasContinuation: true
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
      .map((student) => `
        <tr>
          <td class="class-name">${escapeHtml(readableClass(student.classCode))}</td>
          <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
          <td>${escapeHtml(maskName(student.name))}</td>
        </tr>
      `)
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
          ${rightBody.length ? `
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
          ` : ""}
        </div>
      </div>
      ${hasContinuation ? `
        <div class="class-continuation-note" style="text-align: right; font-size: 0.8rem; font-weight: 700; margin-top: 6px; color: #333;">
          （以下接下頁）
        </div>
      ` : ""}
    </div>
  `;
}

function renderPrintPages(students, groups) {
  const pages = paginateClassBlocks(groups, 18, 27, 8);
  const totalPages = pages.length;
  const details = getActivityDetails();
  const school = details.schoolName;

  const sortedClasses = [...groups.keys()].sort();
  const breakdownText = sortedClasses.map(code => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
  const statsSummaryHtml = students.length 
    ? `<div class="print-stats-summary" style="font-size: 0.85rem; margin-bottom: 8px; color: #333; font-weight: 700;">公假統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）</div>` 
    : '';

  printPages.innerHTML = pages
    .map((pageBlocks, pageIndex) => {
      const isFirst = pageIndex === 0;
      const isLast = pageIndex === totalPages - 1;
      
      const blocksHtml = pageBlocks
        .map(pb => renderPrintClassBlock(pb.classCode, pb.students, pb.isContinuation, pb.hasContinuation))
        .join("");

      return `
        <article class="print-page">
          <header class="print-header">
            <h2>${escapeHtml(school)}學生公假單請示單</h2>
          </header>
          ${isFirst ? renderPrintInfoTable() : renderContinuationLabel()}
          <section class="print-students">
            <h3>公假學生名單${!isFirst ? "（續）" : ""}</h3>
            ${isFirst ? statsSummaryHtml : ''}
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
      <div><strong>活動日期</strong><span>${escapeHtml(details.activityDate)}</span></div>
      <div><strong>活動時間</strong><span>${escapeHtml(details.timeRange)}</span></div>
      <div><strong>活動地點</strong><span>${escapeHtml(details.location)}</span></div>
      <div class="full"><strong>活動事由</strong><span>${escapeHtml(details.reason)}</span></div>
    </section>
  `;
}

function renderContinuationLabel() {
  const details = getActivityDetails();
  return `
    <section class="continuation-label">
      <span>${escapeHtml(details.activityName)}</span>
      <span>${escapeHtml(details.activityDate)}</span>
    </section>
  `;
}

function renderPrintStudentColumn(rows) {
  let currentClassCode = "";
  const body = rows
    .map((row) => {
      if (row.type === "class") {
        currentClassCode = row.classCode;
        return `<tr class="class-row"><td colspan="3" style="background: #f1f5f9; font-weight: 700; text-align: left;">${escapeHtml(readableClass(row.classCode))}</td></tr>`;
      }

      const student = row.student;
      const classCode = student.classCode || currentClassCode;
      return `
        <tr>
          <td class="class-name">${escapeHtml(readableClass(classCode))}</td>
          <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
          <td>${escapeHtml(maskName(student.name))}</td>
        </tr>
      `;
    })
    .join("");

  return `
    <table class="student-table print-student-table">
      <thead>
        <tr>
          <th class="class-name">班級</th>
          <th class="seat">座號</th>
          <th>姓名</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  `;
}

function renderPrintStudentRows(rows) {
  if (!rows.length) return "";
  const half = Math.ceil(rows.length / 2);
  const leftRows = rows.slice(0, half);
  const rightRows = rows.slice(half);

  return `
    <div class="print-student-grid">
      <div class="print-student-col">
        ${renderPrintStudentColumn(leftRows)}
      </div>
      <div class="print-student-col">
        ${rightRows.length ? renderPrintStudentColumn(rightRows) : ""}
      </div>
    </div>
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

function renderSignInInfoTable() {
  const details = getActivityDetails();
  return `
    <section class="info-table print-info-table signin-info-table">
      <div><strong>活動名稱</strong><span>${escapeHtml(details.activityName)}</span></div>
      <div><strong>申請單位</strong><span>${escapeHtml(details.applicantUnit)}</span></div>
      <div><strong>帶隊教師</strong><span>${escapeHtml(details.applicantName)}</span></div>
      <div><strong>活動日期</strong><span>${escapeHtml(details.activityDate)}</span></div>
      <div><strong>活動地點</strong><span>${escapeHtml(details.location)}</span></div>
      <div><strong>帶隊教師簽章</strong><span style="border: 0; background: transparent;"></span></div>
    </section>
  `;
}

function paginateSignInClassBlocks(groups, firstPageCapacity, middlePageCapacity) {
  const sortedClasses = [...groups.keys()].sort();
  const blocks = sortedClasses.map(code => ({
    classCode: code,
    students: [...groups.get(code)],
    isContinuation: false
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
        hasContinuation: false
      });
      remainingCapacity -= rowsNeeded;
      blocks.shift();
    } else {
      const maxStudents = remainingCapacity - 1; // 1 for header

      if (maxStudents > 0) {
        const fitStudents = block.students.slice(0, maxStudents);
        currentPage.push({
          classCode: block.classCode,
          students: fitStudents,
          isContinuation: block.isContinuation,
          hasContinuation: true
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
    .map((student) => `
      <tr>
        <td class="class-name">${escapeHtml(readableClass(student.classCode))}</td>
        <td class="seat">${String(student.seatNumber).padStart(2, "0")}</td>
        <td class="student-id">${escapeHtml(student.studentId || "-")}</td>
        <td class="student-name" style="font-weight: 700;">${escapeHtml(student.name)}</td>
        <td class="signin-box"></td>
        <td class="signin-box"></td>
      </tr>
    `)
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
      ${hasContinuation ? `
        <div class="class-continuation-note" style="text-align: right; font-size: 0.8rem; font-weight: 700; margin-top: 6px; color: #333;">
          （以下接下頁）
        </div>
      ` : ""}
    </div>
  `;
}

function renderSignInPages(students, groups) {
  const pages = paginateSignInClassBlocks(groups, 12, 18);
  const totalPages = pages.length;
  const details = getActivityDetails();
  const school = details.schoolName;

  const sortedClasses = [...groups.keys()].sort();
  const breakdownText = sortedClasses.map(code => `${readableClass(code)}：${groups.get(code).length}人`).join("、");
  const statsSummaryHtml = students.length 
    ? `<div class="print-stats-summary" style="font-size: 0.85rem; margin-bottom: 8px; color: #333; font-weight: 700;">簽到統計：共 ${groups.size} 班，計 ${students.length} 人（${breakdownText}）</div>` 
    : '';

  printPages.innerHTML = pages
    .map((pageBlocks, pageIndex) => {
      const isFirst = pageIndex === 0;
      
      const blocksHtml = pageBlocks
        .map(pb => renderSignInClassBlock(pb.classCode, pb.students, pb.isContinuation, pb.hasContinuation))
        .join("");

      return `
        <article class="print-page">
          <header class="print-header">
            <h2>${escapeHtml(school)}學生公假活動簽到表</h2>
          </header>
          ${isFirst ? renderSignInInfoTable() : renderContinuationLabel()}
          <section class="print-students">
            <h3>簽到學生名單${!isFirst ? "（續）" : ""}</h3>
            ${isFirst ? statsSummaryHtml : ''}
            ${students.length ? blocksHtml : '<div class="empty-state">尚未輸入學生名單</div>'}
          </section>
          <footer class="print-footer">第 ${pageIndex + 1} 頁，共 ${totalPages} 頁</footer>
        </article>
      `;
    })
    .join("");
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
  fields.startTime.value = "08:10";
  fields.endTime.value = "10:00";
  fields.location.value = "校門口及活動中心";
  fields.reason.value = "協助交通安全宣導、場地引導及活動秩序維護。";
  if (fields.manualStudentInput) fields.manualStudentInput.value = "";
  fields.studentInput.value = [
    "40101 蘇大輔",
    "401-02 賴小祥",
    "40201 蘇小祥",
    "402-02 賴大寶"
  ].join("\n");
  update();
}

function clearAll() {
  Object.values(fields).forEach((field) => {
    field.value = "";
  });
  localStorage.removeItem(storageKey);

  // Reset school name to Nanning
  fields.schoolName.value = "南寧高級中學";

  // Reset date to today
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  fields.activityDate.value = `${yyyy}-${mm}-${dd}`;

  update();
}

Object.values(fields).forEach((field) => {
  field.addEventListener("input", update);
});

document.querySelector("#loadSampleBtn").addEventListener("click", loadSample);
document.querySelector("#clearBtn").addEventListener("click", clearAll);
document.querySelector("#printBtn").addEventListener("click", () => {
  renderStudents();
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
    if (excelUpload) excelUpload.value = ""; // clear file selection
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
  // Load initial state
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

loadState();
if (fields.activityDate && !fields.activityDate.value) {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  fields.activityDate.value = `${yyyy}-${mm}-${dd}`;
}
update();
