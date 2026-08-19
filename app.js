(() => {
  const COLS_PER_PAGE = 28;
  const ROWS_PER_PAGE = 28;
  const PAGE_START = 80;
  const MIN_GRADE_ROWS = 22;

  const MONTHS = [
    "вересень",
    "жовтень",
    "листопад",
    "грудень",
    "січень",
    "лютий",
    "березень",
    "квітень",
    "травень",
    "червень",
  ];
  const SPECIAL_MARK = ["темат", "семестр", "залік", "екзамен", "іспит", "скориг"];
  const EXAM_WORDS = ["залік", "екзамен", "іспит"];
  const SHORT_MARK_RE =
    /^(?:н(?:[\\/]?а)?|н\/?б|нб|н\/д|н\/\d+|\d{1,2}(?:[.,]\d+)?)$/i;

  const fileInput = document.getElementById("file-input");
  const csvInput = document.getElementById("csv-input");
  const dropzone = document.getElementById("dropzone");
  const reportEl = document.getElementById("report");
  const previewEl = document.getElementById("preview");
  const printBtn = document.getElementById("print-btn");
  const pdfBtn = document.getElementById("pdf-btn");

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function fold(s) {
    return String(s ?? "")
      .toLowerCase()
      .replace(/[’'`ʼ]/g, "'")
      .replace(/\./g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function rowFold(row) {
    return (row || []).map(fold).join(" ");
  }

  function cellAfterLabel(row, needles) {
    for (let i = 0; i < row.length; i++) {
      const t = fold(row[i]);
      if (needles.some((n) => t.includes(fold(n)))) {
        for (let j = i + 1; j < row.length; j++) {
          if (String(row[j]).trim()) return String(row[j]).trim();
        }
      }
    }
    return "";
  }

  function rowHasAny(row, needles) {
    const t = rowFold(row);
    return needles.some((n) => t.includes(fold(n)));
  }

  function parseDelimited(text, delimiter) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const rows = [];
    let row = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') {
            cur += '"';
            i += 1;
          } else {
            inQuotes = false;
          }
        } else {
          cur += c;
        }
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === delimiter) {
        row.push(cur);
        cur = "";
      } else if (c === "\r") {
        continue;
      } else if (c === "\n") {
        row.push(cur);
        rows.push(row);
        row = [];
        cur = "";
      } else {
        cur += c;
      }
    }
    if (cur.length || row.length) {
      row.push(cur);
      rows.push(row);
    }
    return rows;
  }

  function parseCsvText(text) {
    const firstLine = text.split(/\r?\n/, 1)[0] || "";
    const delimiter = (firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length
      ? "\t"
      : ",";
    return parseDelimited(text, delimiter);
  }

  function excelSerialToDate(n) {
    const serial = Math.round(Number(n));
    const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const yy = String(d.getUTCFullYear()).slice(2);
    return `${dd}.${mm}.${yy}`;
  }

  function looksLikeExcelSerial(v) {
    const n = Number(String(v).trim());
    return Number.isFinite(n) && n >= 30000 && n <= 70000;
  }

  function formatJsDate(d) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = String(d.getFullYear()).slice(2);
    return `${dd}.${mm}.${yy}`;
  }

  function formatDateValue(val) {
    if (val instanceof Date && !isNaN(val.getTime())) return formatJsDate(val);
    const s = String(val ?? "").trim();
    if (!s) return "";
    if (looksLikeExcelSerial(s)) return excelSerialToDate(s);
    let m = s.match(/^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?/);
    if (m) {
      const d = String(Number(m[1])).padStart(2, "0");
      const mo = String(Number(m[2])).padStart(2, "0");
      const rest = s.slice(m[0].length).trim();
      let core = `${d}.${mo}`;
      if (m[3]) {
        const y = m[3].length === 4 ? m[3].slice(2) : m[3];
        core = `${d}.${mo}.${y}`;
      }
      return rest ? `${core} ${rest}` : core;
    }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      const mo = String(Number(m[1])).padStart(2, "0");
      const d = String(Number(m[2])).padStart(2, "0");
      const y = m[3].length === 4 ? m[3].slice(2) : m[3];
      return `${d}.${mo}.${y}`;
    }
    return s;
  }

  function stripTrailingZero(s) {
    const t = String(s ?? "").trim();
    if (/^\d+\.0+$/.test(t)) return String(parseInt(t, 10));
    return t;
  }

  function asLessonNum(val) {
    const s = String(val ?? "").trim();
    if (/^\d+$/.test(s)) return s;
    if (/^\d+\.0+$/.test(s)) return String(parseInt(s, 10));
    if (/^\d+\.$/.test(s)) return s.replace(/\.$/, "");
    return null;
  }

  function shortenDate(s) {
    return formatDateValue(s);
  }

  function shortenName(fullName) {
    const parts = String(fullName).trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return parts.slice(0, 2).join(" ");
    return String(fullName).trim();
  }

  function isLongStatus(val) {
    const s = stripTrailingZero(val);
    if (!s) return false;
    return !SHORT_MARK_RE.test(s);
  }

  function firstLongStatus(marks) {
    for (let i = 0; i < marks.length; i++) {
      if (isLongStatus(marks[i])) return { col: i, text: String(marks[i]).trim() };
    }
    return { col: null, text: null };
  }

  function isSpecialTitle(title) {
    const t = fold(title);
    return SPECIAL_MARK.some((k) => t.includes(k));
  }

  function isLessonDateTitle(title) {
    const s = String(title || "").trim();
    if (!s || isSpecialTitle(s)) return false;
    return /^\d{1,2}\.\d{1,2}(\.\d{2,4})?$/.test(s);
  }

  function isMetaSheet(name) {
    const t = fold(name);
    return (
      t.startsWith("код") ||
      t.startsWith("звітн") ||
      t.includes("звітність") ||
      t.includes("рейтинг")
    );
  }

  function buildSpreads(gradeColumns, lessons) {
    const regular = lessons.filter((l) => l.type === "lesson");
    const tail = lessons.filter((l) => l.type !== "lesson");
    const chunks = [];
    if (!gradeColumns.length) {
      chunks.push([]);
    } else {
      for (let i = 0; i < gradeColumns.length; i += COLS_PER_PAGE) {
        chunks.push(gradeColumns.slice(i, i + COLS_PER_PAGE));
      }
    }
    const spreads = [];
    let lessonIdx = 0;
    for (let i = 0; i < chunks.length; i++) {
      const cols = chunks[i];
      const dateCount = cols.filter((c) => isLessonDateTitle(c.title)).length;
      const isLast = i === chunks.length - 1;
      let pageLessons;
      if (isLast) {
        pageLessons = regular.slice(lessonIdx).concat(tail);
      } else {
        pageLessons = regular.slice(lessonIdx, lessonIdx + dateCount);
        lessonIdx += dateCount;
      }
      spreads.push({
        cols,
        colStart: i * COLS_PER_PAGE,
        pageLessons,
        isLast,
        dateCount,
      });
    }
    return spreads;
  }

  function isMonthJunk(row) {
    const cells = (row || []).map((c) => String(c).trim());
    const first = cells[0] || "";
    const second = cells[1] || "";
    const third = cells[2] || "";
    const rest = cells.slice(3).filter(Boolean).join(" ");
    const blob = fold(cells.join(" "));
    if (MONTHS.includes(second.toLowerCase())) return true;
    if (blob.includes("разом за")) return true;
    if (MONTHS.some((m) => blob === m || blob.startsWith(m + " "))) return true;
    if (!first && !second && !rest && /^\d+([.,]\d+)?$/.test(third)) return true;
    return false;
  }

  function statusFontPt(text, span) {
    const markRowMm = 140;
    const widthMm = markRowMm * (span / COLS_PER_PAGE) - 1.5;
    const widthPt = Math.max(8, (widthMm * 72) / 25.4);
    const heightPt = 16;
    const n = Math.max(1, String(text).length);
    let size = 8.5;
    while (size > 4.5) {
      const charsPerLine = Math.max(1, widthPt / (size * 0.48));
      const lines = Math.ceil(n / charsPerLine);
      if (lines * size * 1.08 <= heightPt) return Math.round(size * 100) / 100;
      size -= 0.25;
    }
    return 4.5;
  }

  function statusTdHtml(text, span) {
    const fontPt = statusFontPt(text, span);
    return (
      `<td class="col-mark status-span" colspan="${span}">` +
      `<div class="status-fit" style="font-size:${fontPt}pt">${escapeHtml(text)}</div></td>`
    );
  }

  function findHeaderIndex(rows, needles) {
    for (let i = 0; i < rows.length; i++) {
      if (rowHasAny(rows[i], needles)) return i;
    }
    return -1;
  }

  /**
   * Closed-fail parse. Never throws on a weird sheet — returns {ok:false, reason}.
   */
  function parseJournalSheet(rows, sheetMeta) {
    const label = sheetMeta.name
      ? `Аркуш ${sheetMeta.index + 1} «${sheetMeta.name}»`
      : `Аркуш ${sheetMeta.index + 1}`;

    if (isMetaSheet(sheetMeta.name || "")) {
      return {
        ok: false,
        label,
        reason: "службовий аркуш (коди / звітність), не журнал дисципліни",
      };
    }

    if (!rows || !rows.length) {
      return { ok: false, label, reason: "порожній аркуш" };
    }

    let discipline = "";
    let teacher = "";
    for (const row of rows.slice(0, 8)) {
      if (!discipline) {
        discipline = cellAfterLabel(row, [
          "назва дисципліни",
          "назва предмета",
          "дисципліна",
        ]);
      }
      if (!teacher) {
        teacher = cellAfterLabel(row, ["викладач"]);
      }
    }

    const gradesHeaderIdx = findHeaderIndex(rows, [
      "піб студента",
      "прізвище та ініціали студента",
      "прізвище студента",
      "піб",
    ]);
    const topicsHeaderIdx = findHeaderIndex(rows, [
      "тема заняття",
      "короткий зміст",
      "зміст заняття",
    ]);

    if (gradesHeaderIdx === -1) {
      return {
        ok: false,
        label,
        reason: "не знайдено таблицю студентів (очікується колонка на кшталт «ПІБ»)",
      };
    }
    if (topicsHeaderIdx === -1) {
      return {
        ok: false,
        label,
        reason: "не знайдено таблицю тем занять (очікується «Тема заняття» або «Короткий зміст»)",
      };
    }
    if (topicsHeaderIdx <= gradesHeaderIdx) {
      return {
        ok: false,
        label,
        reason: "структура зʼїхала: таблиця тем стоїть не після таблиці оцінок",
      };
    }

    const headerRow = rows[gradesHeaderIdx] || [];
    const gradeColumns = [];
    for (let colIdx = 2; colIdx < headerRow.length; colIdx++) {
      const val = formatDateValue(headerRow[colIdx]);
      if (val) gradeColumns.push({ colIdx, title: val });
    }
    if (!gradeColumns.length) {
      return {
        ok: false,
        label,
        reason: "у шапці оцінок немає колонок з датами",
      };
    }

    const students = [];
    for (let r = gradesHeaderIdx + 1; r < topicsHeaderIdx; r++) {
      const row = rows[r] || [];
      if (!row.some((c) => String(c).trim())) continue;
      const num =
        asLessonNum(row[0]) || String(row[0] || "").trim().replace(/\.+$/, "");
      const name = shortenName(String(row[1] || "").replace(/\/\d+$/, ""));
      if (!name && !num) continue;
      const marks = gradeColumns.map((c) =>
        c.colIdx < row.length ? stripTrailingZero(row[c.colIdx]) : ""
      );
      const status = firstLongStatus(marks);
      students.push({
        num,
        name,
        marks,
        statusCol: status.col,
        statusText: status.text,
      });
    }
    if (!students.length) {
      return {
        ok: false,
        label,
        reason: "таблицю студентів знайдено, але рядків студентів немає",
      };
    }

    const lessons = [];
    for (let r = topicsHeaderIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const rStr = row.map((c) => String(c).trim()).join(" ").trim();
      if (!rStr) continue;
      if (isMonthJunk(row)) continue;

      const first = String(row[0] || "").trim();
      const second = String(row[1] || "").trim();
      const third = stripTrailingZero(row[2]);
      const fourth = String(row[3] || "").trim();
      const hw = row.slice(4).map((c) => String(c).trim()).find(Boolean) || "";
      const lessonNum = asLessonNum(first);

      if (lessonNum) {
        lessons.push({
          type: "lesson",
          num: lessonNum,
          date: formatDateValue(second),
          hours: third,
          topic: fourth,
          hw,
        });
      } else if (
        EXAM_WORDS.some((k) => (fourth + second).toLowerCase().includes(k))
      ) {
        lessons.push({
          type: "exam",
          num: first,
          date: formatDateValue(second),
          hours: third,
          topic: fourth || second,
          hw,
        });
      } else {
        const topic = fourth || row.map((c) => String(c).trim()).filter(Boolean).join(" ");
        lessons.push({
          type: "note",
          num: first,
          date: second ? formatDateValue(second) : "",
          hours: third,
          topic,
          hw,
        });
      }
    }

    return {
      ok: true,
      label,
      data: {
        discipline,
        teacher,
        gradeColumns,
        students,
        lessons,
        sheetName: sheetMeta.name || "",
      },
    };
  }

  function renderJournal(data) {
    const {
      discipline,
      teacher,
      gradeColumns,
      students,
      lessons,
    } = data;
    const spreads = buildSpreads(gradeColumns, lessons);

    let html = "";
    let pageNum = PAGE_START;

    for (const spread of spreads) {
      const sCols = spread.cols;
      const colPtr = spread.colStart;
      const cEnd = colPtr + sCols.length;
      const emptyC = COLS_PER_PAGE - sCols.length;

      html += `<div class="page" data-kind="grades">
  <div class="page-content">
    <div style="margin-bottom:6px;">
      <div class="header-line"><span>Назва предмета: <span class="header-underline">${escapeHtml(discipline)}</span></span></div>
    </div>
    <table class="journal-table">
      <thead><tr>
        <th class="col-num">№<br>з/п</th>
        <th class="col-name">Прізвище та ініціали студента</th>`;

      for (const col of sCols) {
        const cls = "vertical-th" + (isSpecialTitle(col.title) ? " special-mark" : "");
        html += `<th class="${cls}"><div class="vertical-text">${escapeHtml(col.title)}</div></th>`;
      }
      for (let i = 0; i < emptyC; i++) {
        html += `<th class="vertical-th"><div class="vertical-text">&nbsp;</div></th>`;
      }
      html += `</tr></thead><tbody>`;

      for (const st of students) {
        html += `<tr class="grades-row"><td class="col-num">${escapeHtml(st.num)}</td><td class="col-name">${escapeHtml(st.name)}</td>`;
        let spanned = false;
        if (st.statusCol !== null && colPtr > st.statusCol) {
          html += statusTdHtml(st.statusText, COLS_PER_PAGE);
          spanned = true;
        } else {
          for (let cI = colPtr; cI < cEnd; cI++) {
            const val = cI < st.marks.length ? st.marks[cI] : "";
            if (isLongStatus(val)) {
              html += statusTdHtml(val, cEnd - cI + emptyC);
              spanned = true;
              break;
            }
            const spec = isSpecialTitle(gradeColumns[cI].title);
            html += `<td class="col-mark${spec ? " special-mark" : ""}">${escapeHtml(val)}</td>`;
          }
        }
        if (!spanned) {
          for (let i = 0; i < emptyC; i++) html += `<td class="col-mark">&nbsp;</td>`;
        }
        html += `</tr>`;
      }

      for (let extra = students.length + 1; extra <= Math.max(MIN_GRADE_ROWS, students.length); extra++) {
        html += `<tr class="filler-row"><td class="col-num">${extra}</td><td class="col-name">&nbsp;</td>${"<td></td>".repeat(COLS_PER_PAGE)}</tr>`;
      }

      html += `</tbody></table></div><div class="page-footer">${pageNum}</div></div>`;
      pageNum += 1;

      const sLessons = spread.pageLessons;
      const emptyL = spread.isLast
        ? Math.max(0, ROWS_PER_PAGE - sLessons.length)
        : 0;

      html += `<div class="page" data-kind="lessons">
  <div class="page-content">
    <div style="margin-bottom:10px; font-size:10pt;"><span>Прізвище та ініціали викладача: <span class="header-underline">${escapeHtml(teacher)}</span></span></div>
    <table class="journal-table">
      <thead><tr>
        <th class="col-lesson-num">№<br>зан.</th>
        <th class="col-lesson-date">Дата<br>пров.</th>
        <th class="col-lesson-hours">К-сть<br>год</th>
        <th class="col-lesson-topic">Короткий зміст заняття</th>
        <th class="col-lesson-hw">Що задано: назва підручника, параграф, стор. тощо</th>
        <th class="col-lesson-sign">Підпис<br>викл.</th>
      </tr></thead><tbody>`;

      for (const item of sLessons) {
        html += `<tr><td>${escapeHtml(item.num || "")}</td><td>${escapeHtml(item.date || "")}</td><td>${escapeHtml(item.hours || "")}</td><td class="col-lesson-topic">${escapeHtml(item.topic)}</td><td class="col-lesson-hw">${escapeHtml(item.hw || "")}</td><td></td></tr>`;
      }
      for (let i = 0; i < emptyL; i++) {
        html += `<tr class="filler-row"><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>`;
      }

      html += `</tbody></table>`;
      html += `</div><div class="page-footer">${pageNum}</div></div>`;
      pageNum += 1;
    }

    const pack = document.createElement("section");
    pack.className = "journal-pack";
    pack.innerHTML = html;
    return pack;
  }

  function fitStatusCells(root) {
    root.querySelectorAll(".status-fit").forEach((el) => {
      let size = parseFloat(el.style.fontSize) || 8.5;
      const min = 4.5;
      let guard = 0;
      while (guard++ < 50 && size > min && el.scrollHeight > el.clientHeight + 0.5) {
        size -= 0.25;
        el.style.fontSize = size + "pt";
      }
    });
  }

  function afterLayout() {
    fitStatusCells(previewEl);
  }

  function guessGroupName(fileName, okResults) {
    const base = (fileName || "").replace(/\.[^.]+$/, "");
    const fromFile = base.match(/([A-Za-zА-Яа-яІіЇїЄєҐґ]{1,5}[-–]?\d{1,3})/);
    if (fromFile) return fromFile[1].replace("–", "-").toUpperCase();
    for (const r of okResults) {
      const m = (r.data.discipline || "").match(/\(([^)]+)\)/);
      if (m) return m[1].trim();
    }
    return "";
  }

  function renderCover(groupName, subjects) {
    const pack = document.createElement("section");
    pack.className = "journal-pack";
    const title = groupName ? `Журнал групи ${groupName}` : "Журнал групи";
    pack.innerHTML = `<div class="page cover-page">
      <div class="cover-inner">
        <p class="cover-kicker">Навчальний журнал</p>
        <h2>${escapeHtml(title)}</h2>
        <ol class="cover-list">${subjects
          .map(
            (s) =>
              `<li><strong>${escapeHtml(s.discipline || "без назви")}</strong>` +
              (s.teacher ? ` <span>— ${escapeHtml(s.teacher)}</span>` : "") +
              `</li>`
          )
          .join("")}</ol>
      </div>
      <div class="page-footer"></div>
    </div>`;
    return pack;
  }

  function renderReport(results) {
    reportEl.hidden = false;
    reportEl.innerHTML = "";
    for (const r of results) {
      const item = document.createElement("div");
      item.className = "report-item " + (r.ok ? "ok" : "skip");
      if (r.ok) {
        const d = r.data;
        item.innerHTML =
          `<strong>${escapeHtml(r.label)}</strong> — у журнал. ` +
          `<code>${escapeHtml(d.discipline || "без назви")}</code>` +
          (d.teacher ? ` · ${escapeHtml(d.teacher)}` : "") +
          ` · ${d.students.length} студ. · ${d.gradeColumns.length} дат`;
      } else {
        item.textContent = `${r.label} не розпізнано, пропущено: ${r.reason}.`;
      }
      reportEl.appendChild(item);
    }
  }

  let lastPdfName = "Журнал групи.pdf";

  function processSheets(sheets, opts) {
    const fileName = (opts && opts.fileName) || "";
    previewEl.innerHTML = "";
    const results = sheets.map((sheet, index) =>
      parseJournalSheet(sheet.rows, { index, name: sheet.name })
    );
    renderReport(results);
    const ok = results.filter((r) => r.ok);
    const group = guessGroupName(fileName, ok);
    lastPdfName = group ? `Журнал групи ${group}.pdf` : "Журнал групи.pdf";
    const titleEl = document.getElementById("journal-title");
    if (titleEl) {
      titleEl.textContent = group ? `Журнал групи ${group}` : "Журнал на папір";
    }
    if (ok.length) {
      previewEl.appendChild(
        renderCover(
          group,
          ok.map((r) => ({
            discipline: r.data.discipline,
            teacher: r.data.teacher,
          }))
        )
      );
    }
    for (const r of ok) {
      previewEl.appendChild(renderJournal(r.data));
    }
    const ready = ok.length > 0;
    printBtn.disabled = !ready;
    if (pdfBtn) pdfBtn.disabled = !ready;
    requestAnimationFrame(() => requestAnimationFrame(afterLayout));
  }

  function cellToRaw(v) {
    if (v instanceof Date && !isNaN(v.getTime())) return formatJsDate(v);
    if (typeof v === "number") return String(v);
    return v == null ? "" : String(v);
  }

  function trimSheetRows(rows) {
    const mapped = rows.map((row) => (row || []).map(cellToRaw));
    let end = mapped.length;
    while (end > 0 && !mapped[end - 1].some((c) => String(c).trim())) end -= 1;
    return mapped.slice(0, end);
  }

  function rowsFromWorkbook(buffer) {
    if (typeof XLSX === "undefined") {
      throw new Error("Бібліотека Excel не завантажилась. Спробуйте CSV.");
    }
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    return wb.SheetNames.map((name) => {
      const sheet = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
        blankrows: false,
      });
      return { name, rows: trimSheetRows(rows) };
    });
  }

  async function handleFile(file) {
    const name = file.name || "file";
    const lower = name.toLowerCase();
    if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
      const buf = await file.arrayBuffer();
      processSheets(rowsFromWorkbook(buf), { fileName: name });
      return;
    }
    const text = await file.text();
    processSheets([{ name, rows: parseCsvText(text) }], { fileName: name });
  }

  async function downloadPdf() {
    if (typeof html2canvas === "undefined" || !window.jspdf) {
      window.print();
      return;
    }
    const pages = previewEl.querySelectorAll(".page");
    if (!pages.length) return;
    pdfBtn.disabled = true;
    const status = document.getElementById("pdf-status");
    if (status) status.textContent = "Готуємо PDF…";
    document.body.classList.add("pdf-capture");
    try {
      const JsPDF = window.jspdf.jsPDF;
      const pdf = new JsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
      for (let i = 0; i < pages.length; i++) {
        if (status) status.textContent = `Готуємо PDF… ${i + 1}/${pages.length}`;
        const canvas = await html2canvas(pages[i], {
          scale: 2,
          backgroundColor: "#ffffff",
          useCORS: true,
        });
        const img = canvas.toDataURL("image/jpeg", 0.92);
        if (i > 0) pdf.addPage();
        pdf.addImage(img, "JPEG", 0, 0, 210, 297);
      }
      pdf.save(lastPdfName);
      if (status) status.textContent = "";
    } catch (err) {
      if (status) status.textContent = "Не вдалось зібрати PDF, відкриваю друк.";
      window.print();
    } finally {
      document.body.classList.remove("pdf-capture");
      pdfBtn.disabled = false;
    }
  }

  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (file) handleFile(file).catch((err) => {
      renderReport([{ ok: false, label: "Файл", reason: err.message }]);
    });
  });

  ["dragenter", "dragover"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });
  dropzone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) handleFile(file).catch((err) => {
      renderReport([{ ok: false, label: "Файл", reason: err.message }]);
    });
  });

  printBtn.addEventListener("click", () => {
    fitStatusCells(previewEl);
    window.print();
  });
  if (pdfBtn) pdfBtn.addEventListener("click", () => downloadPdf());
  window.addEventListener("beforeprint", () => fitStatusCells(previewEl));

  if (csvInput) {
    csvInput.addEventListener("change", () => {
      const file = csvInput.files && csvInput.files[0];
      if (file) handleFile(file).catch((err) => {
        renderReport([{ ok: false, label: "Файл", reason: err.message }]);
      });
    });
  }

  async function tryAutoload() {
    const q = new URLSearchParams(location.search).get("file");
    const candidates = q ? [q] : ["Км-42.xlsx", "input.csv"];
    for (const name of candidates) {
      try {
        const res = await fetch(encodeURI(name));
        if (!res.ok) continue;
        const lower = name.toLowerCase();
        if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
          processSheets(rowsFromWorkbook(await res.arrayBuffer()), { fileName: name });
        } else {
          processSheets([{ name, rows: parseCsvText(await res.text()) }], { fileName: name });
        }
        return;
      } catch (e) {
        /* file:// or missing sample */
      }
    }
  }
  tryAutoload();
})();
