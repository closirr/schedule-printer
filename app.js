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
  const SHORT_MARK_RE = /^(?:н(?:[\\/]?а)?|н\/?б|нб|\d{1,2})$/i;

  const fileInput = document.getElementById("file-input");
  const dropzone = document.getElementById("dropzone");
  const reportEl = document.getElementById("report");
  const previewEl = document.getElementById("preview");
  const printBtn = document.getElementById("print-btn");

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

  function shortenDate(s) {
    const m = String(s).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (!m) return String(s).trim();
    const d = String(Number(m[1])).padStart(2, "0");
    const mo = String(Number(m[2])).padStart(2, "0");
    return `${d}.${mo}.${m[3].slice(2)}`;
  }

  function shortenName(fullName) {
    const parts = String(fullName).trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return parts.slice(0, 2).join(" ");
    return String(fullName).trim();
  }

  function isLongStatus(val) {
    const s = String(val || "").trim();
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
    return /^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(s);
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
      const val = String(headerRow[colIdx] || "").trim();
      if (val) gradeColumns.push({ colIdx, title: shortenDate(val) });
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
      const num = String(row[0] || "").trim().replace(/\.+$/, "");
      const name = shortenName(row[1] || "");
      if (!name && !num) continue;
      const marks = gradeColumns.map((c) =>
        c.colIdx < row.length ? String(row[c.colIdx] || "").trim() : ""
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
      const third = String(row[2] || "").trim();
      const fourth = String(row[3] || "").trim();
      const hw = row.slice(4).map((c) => String(c).trim()).find(Boolean) || "";

      if (/^\d+$/.test(first)) {
        lessons.push({
          type: "lesson",
          num: first,
          date: shortenDate(second),
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
          date: shortenDate(second),
          hours: third,
          topic: fourth || second,
          hw,
        });
      } else {
        const topic = fourth || row.map((c) => String(c).trim()).filter(Boolean).join(" ");
        lessons.push({
          type: "note",
          num: first,
          date: second ? shortenDate(second) : "",
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

  function renderReport(results) {
    reportEl.hidden = false;
    reportEl.innerHTML = "";
    for (const r of results) {
      const item = document.createElement("div");
      item.className = "report-item " + (r.ok ? "ok" : "skip");
      if (r.ok) {
        const d = r.data;
        item.innerHTML =
          `<strong>${escapeHtml(r.label)}</strong> — друкуємо. ` +
          `<code>${escapeHtml(d.discipline || "без назви")}</code>` +
          (d.teacher ? ` · ${escapeHtml(d.teacher)}` : "") +
          ` · ${d.students.length} студ. · ${d.gradeColumns.length} дат`;
      } else {
        item.textContent = `${r.label} не розпізнано, пропущено: ${r.reason}.`;
      }
      reportEl.appendChild(item);
    }
  }

  function processSheets(sheets) {
    previewEl.innerHTML = "";
    const results = sheets.map((sheet, index) =>
      parseJournalSheet(sheet.rows, { index, name: sheet.name })
    );
    renderReport(results);
    const ok = results.filter((r) => r.ok);
    for (const r of ok) {
      previewEl.appendChild(renderJournal(r.data));
    }
    printBtn.disabled = ok.length === 0;
    requestAnimationFrame(() => requestAnimationFrame(afterLayout));
  }

  function rowsFromWorkbook(buffer) {
    if (typeof XLSX === "undefined") {
      throw new Error("Бібліотека Excel не завантажилась. Спробуйте CSV.");
    }
    const wb = XLSX.read(buffer, { type: "array", cellDates: false, raw: false });
    return wb.SheetNames.map((name) => {
      const sheet = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: true,
      });
      return { name, rows };
    });
  }

  async function handleFile(file) {
    const name = (file.name || "").toLowerCase();
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const buf = await file.arrayBuffer();
      processSheets(rowsFromWorkbook(buf));
      return;
    }
    const text = await file.text();
    processSheets([{ name: file.name, rows: parseCsvText(text) }]);
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
  window.addEventListener("beforeprint", () => fitStatusCells(previewEl));

  const sampleName = new URLSearchParams(location.search).get("file") || "input.csv";
  fetch(sampleName)
    .then((res) => (res.ok ? res.text() : Promise.reject()))
    .then((text) => processSheets([{ name: sampleName, rows: parseCsvText(text) }]))
    .catch(() => {});
})();
