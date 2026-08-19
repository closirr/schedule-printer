(() => {
  const COLS_PER_PAGE = 28;
  const ROWS_PER_PAGE = 28;
  const PAGE_START = 1;
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
    /^(?:н(?:[\\/]?а)?|н\/?б|нб|н\/д|н\/з|н[\\/]\d+(?:[.,]\d+)?|\d{1,2}(?:[.,]\d+)?|відм\.?|доб\.?|задов\.?)$/i;

  const fileInput = document.getElementById("file-input");
  const csvInput = document.getElementById("csv-input");
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
    // Excel serials via SheetJS often land at 23:59:56 of the previous
    // civil day. Shift toward noon UTC so the calendar day matches Excel.
    const shifted = new Date(d.getTime() + 12 * 60 * 60 * 1000);
    const dd = String(shifted.getUTCDate()).padStart(2, "0");
    const mm = String(shifted.getUTCMonth() + 1).padStart(2, "0");
    const yy = String(shifted.getUTCFullYear()).slice(2);
    return `${dd}.${mm}.${yy}`;
  }

  function looksLikeFormattedDate(s) {
    const t = String(s || "").trim();
    // Require a year so hour values like "2.6" are not treated as dates.
    return (
      /^\d{1,2}\.\d{1,2}\.\d{2,4}/.test(t) ||
      /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(t)
    );
  }

  function isDateObj(v) {
    return (
      v instanceof Date ||
      (v != null &&
        typeof v === "object" &&
        typeof v.getTime === "function" &&
        !isNaN(v.getTime()))
    );
  }

  function formatDateValue(val) {
    if (val instanceof Date && !isNaN(val.getTime())) return formatJsDate(val);
    const s = String(val ?? "").trim();
    if (!s) return "";
    if (looksLikeExcelSerial(s)) return excelSerialToDate(s);
    const compact = s.replace(/\s+/g, " ").trim();
    let m = compact.match(/^(\d{1,2})\s*\.\s*(\d{1,2})(?:\s*\.\s*(\d{2,4}))?/);
    if (m) {
      const d = String(Number(m[1])).padStart(2, "0");
      const mo = String(Number(m[2])).padStart(2, "0");
      let core = `${d}.${mo}`;
      if (m[3]) {
        const y = m[3].length === 4 ? m[3].slice(2) : m[3];
        core = `${d}.${mo}.${y}`;
      }
      return core;
    }
    m = compact.match(/^(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{2,4})/);
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

  function normalizeMarkText(val) {
    return String(val ?? "")
      .replace(/[\t\r\n\u00A0\u1680\u2000-\u200D\u2028\u2029\u202F\u205F\u3000\uFEFF]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function cleanMark(val) {
    let s = stripTrailingZero(normalizeMarkText(val));
    if (!s || /^[-–—−]+$/.test(s)) return "";
    const nSlash = s.match(/^[nн]\s*[\\/]\s*(\d+(?:[.,]\d+)?)$/i);
    if (nSlash) return "н/" + stripTrailingZero(nSlash[1]);
    if (/^[nн]\s*[\\/]?\s*а$/i.test(s) || /^n\/?a$/i.test(s)) return "н/а";
    if (/^[nн]\s*[\\/]?\s*д$/i.test(s) || /^n\/?d$/i.test(s)) return "н/д";
    if (/^[nн]\s*[\\/]?\s*з$/i.test(s) || /^n\/?z$/i.test(s)) return "н/з";
    if (/^[nн]\s*[\\/]\s*б$/i.test(s) || /^n\/?b$/i.test(s)) return "н/б";
    const tokens = s.split(" ").filter(Boolean);
    if (tokens.length > 1) {
      const rest = tokens.filter((t) => !/^[nн]$/i.test(t));
      if (rest.length && rest.length < tokens.length) {
        return cleanMark(rest[rest.length - 1]);
      }
    }
    if (/^[nн]$/i.test(s)) return "н";
    const verbal = abbrevVerbalGrade(s);
    if (verbal) return verbal;
    return s;
  }

  function abbrevVerbalGrade(val) {
    const t = fold(val).replace(/[.]/g, "");
    if (!t) return "";
    if (t === "відмінно" || t === "відмін" || t === "відм") return "відм.";
    if (t === "добре" || t === "добр" || t === "доб") return "доб.";
    if (t === "задовільно" || t === "задов" || t === "зад") return "задов.";
    if (t === "незадовільно" || t === "незад" || t === "нз" || t === "н/з") return "н/з";
    return "";
  }

  function isNaNdMark(val) {
    const t = String(val || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
    return /^(н[\\/]?а|н[\\/]?д|н[\\/]?з)$/.test(t);
  }

  function isCompactMark(val) {
    return isNaNdMark(val) || !!abbrevVerbalGrade(val);
  }

  function isHundredthsMark(val) {
    return /^\d{1,2}[.,]\d{2,}$/.test(String(val || "").trim());
  }

  function combineStatusParts(parts) {
    const cleaned = parts
      .map((p) => String(p).replace(/\s+/g, " ").trim())
      .filter(Boolean);
    const kept = [];
    for (const s of cleaned) {
      if (kept.some((k) => k.includes(s))) continue;
      const i = kept.findIndex((k) => s.includes(k));
      if (i >= 0) kept[i] = s;
      else kept.push(s);
    }
    return kept.join(" ").replace(/\s+/g, " ").trim();
  }

  function mergeSplitStatus(marks) {
    const out = marks.slice();
    for (let i = 0; i < out.length; i++) {
      if (!isLongStatus(out[i])) continue;
      const parts = [out[i]];
      let j = i + 1;
      while (j < out.length) {
        const v = out[j];
        if (!String(v || "").trim()) {
          j += 1;
          continue;
        }
        if (isLongStatus(v)) {
          parts.push(v);
          out[j] = "";
          j += 1;
          continue;
        }
        break;
      }
      out[i] = combineStatusParts(parts);
      i = j - 1;
    }
    return out;
  }

  function isHoursValue(val) {
    return /^\d+([.,]\d+)?$/.test(String(val ?? "").trim());
  }

  function looksLikeExamTopic(text) {
    const t = fold(text).replace(/[.]+$/g, "");
    if (!t) return false;
    if (/план|фактичн|годин|програм/.test(t)) return false;
    if (EXAM_WORDS.some((k) => t === k || t.startsWith(k + " "))) return true;
    return (
      t.length <= 40 &&
      EXAM_WORDS.some((k) => new RegExp("(?:^|\\s)" + k + "(?:\\s|$)").test(t))
    );
  }

  function isTeacherSummary(text) {
    const t = fold(text);
    return (
      t.includes("за планом") ||
      t.includes("фактично") ||
      t.includes("програму виконано") ||
      t.startsWith("програму ")
    );
  }

  function looksLikeConsultation(text) {
    return fold(text).includes("консульт");
  }

  function looksLikeHwText(text) {
    const t = fold(text);
    if (!t) return false;
    return (
      t.includes("google classroom") ||
      t.includes("classroom") ||
      t.includes("завдання в") ||
      t.includes("домашн") ||
      t.includes("підручник") ||
      t.includes("параграф") ||
      t.startsWith("повтор") ||
      t === "підготовка" ||
      t.startsWith("підготовка ")
    );
  }

  function looksLikeLooseDate(val) {
    const t = String(val ?? "").replace(/\s+/g, " ").trim();
    if (!t) return false;
    if (looksLikeFormattedDate(t) || looksLikeExcelSerial(t)) return true;
    const m = t.match(/^(\d{1,2})\s*[./]\s*(\d{1,2})(?:\s*[./]\s*(\d{2,4}))?/);
    if (!m) return false;
    if (m[3]) return true;
    const d = Number(m[1]);
    const mo = Number(m[2]);
    return d >= 1 && d <= 31 && mo >= 1 && mo <= 12;
  }

  function asLessonNum(val) {
    const s = String(val ?? "").trim();
    const m = s.match(/^(\d+)\.+$/);
    if (m) return m[1];
    if (/^\d+$/.test(s)) return s;
    if (/^\d+\.0+$/.test(s)) return String(parseInt(s, 10));
    return null;
  }

  function parseTopicRow(row) {
    const cells = (row || []).map((c) =>
      String(c ?? "")
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .replace(/[\t\r]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
    );
    if (!cells.some(Boolean)) return null;
    if (isMonthJunk(cells)) return null;

    const first = cells[0] || "";
    const second = cells[1] || "";
    const third = cells[2] || "";
    const lessonNum = asLessonNum(first);
    const date = looksLikeLooseDate(second) ? formatDateValue(second) : "";

    let hours = "";
    const pending = [];
    if (isHoursValue(third)) hours = stripTrailingZero(third);
    else if (third && !looksLikeLooseDate(third)) pending.push(third);

    for (let i = 3; i < cells.length; i++) {
      if (cells[i]) pending.push(cells[i]);
    }

    let topic = "";
    const hwParts = [];
    for (const t of pending) {
      if (looksLikeLooseDate(t) && date) continue;
      if (!topic) {
        if (looksLikeHwText(t) && pending.some((x) => x !== t && !looksLikeHwText(x))) {
          hwParts.push(t);
        } else {
          topic = t;
        }
        continue;
      }
      const frag = t.replace(/^[.,;:\-\s]+/, "").trim();
      if (frag && fold(topic).includes(fold(frag))) continue;
      if (/^[.,;]/.test(t)) {
        topic = (topic.replace(/[.,;\s]+$/, "") + " " + t).replace(/\s+/g, " ").trim();
        continue;
      }
      hwParts.push(t);
    }
    let hw = hwParts.join(" ").replace(/\s+/g, " ").trim();
    if (!topic && hw) {
      if (
        looksLikeExamTopic(hw) ||
        looksLikeConsultation(hw) ||
        !looksLikeHwText(hw)
      ) {
        topic = hw;
        hw = "";
      }
    }
    if (looksLikeExamTopic(hw) && !topic) {
      topic = hw;
      hw = "";
    }
    if (fold(hw) && fold(topic) && fold(hw) === fold(topic)) hw = "";

    if (!date && !topic && !hw) return null;

    const examish =
      looksLikeExamTopic(topic) ||
      looksLikeExamTopic(first) ||
      looksLikeExamTopic(second);
    const summary = isTeacherSummary(topic) || isTeacherSummary(cells.filter(Boolean).join(" "));
    if (summary && !examish) {
      return {
        type: "note",
        num: first,
        date,
        hours,
        topic,
        hw,
      };
    }
    if (lessonNum) {
      return {
        type: examish ? "exam" : "lesson",
        num: lessonNum,
        date,
        hours,
        topic,
        hw,
      };
    }
    if (examish) {
      return {
        type: "exam",
        num: first,
        date,
        hours,
        topic: topic || (looksLikeExamTopic(second) ? second : ""),
        hw,
      };
    }
    if (!summary && (date || hours) && topic) {
      return {
        type: "lesson",
        num: first,
        date,
        hours,
        topic,
        hw,
      };
    }
    return {
      type: "note",
      num: first,
      date,
      hours,
      topic,
      hw,
    };
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
    const s = cleanMark(val);
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
    while (size > 3.5) {
      const charsPerLine = Math.max(1, widthPt / (size * 0.48));
      const lines = Math.ceil(n / charsPerLine);
      if (lines * size * 1.05 <= heightPt) return Math.round(size * 100) / 100;
      size -= 0.25;
    }
    return 3.5;
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

  function looksLikeTopicsHeader(row) {
    if (
      rowHasAny(row, [
        "тема заняття",
        "короткий зміст",
        "зміст заняття",
        "тема занять",
      ])
    ) {
      return true;
    }
    const cells = (row || []).map((c) => fold(c));
    const hasDate = cells.some((c) => c === "дата" || c.startsWith("дата "));
    const hasHours = cells.some(
      (c) => c.includes("годин") || c === "год" || c === "к-сть год"
    );
    const hasNum = cells.some(
      (c) => c === "№" || c === "n" || c.includes("зан")
    );
    const hasHw = cells.some(
      (c) =>
        c.includes("домашн") ||
        c.includes("задано") ||
        c === "дз" ||
        c.includes("підручник")
    );
    if (hasDate && hasHours) return true;
    if (hasDate && hasNum && hasHw) return true;
    return false;
  }

  function findTopicsHeaderIndex(rows, afterIdx) {
    for (let i = afterIdx + 1; i < rows.length; i++) {
      if (looksLikeTopicsHeader(rows[i])) return i;
    }
    return -1;
  }

  function parseDateParts(s) {
    const m = String(s || "")
      .trim()
      .match(/^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?/);
    if (!m) return null;
    return { d: Number(m[1]), mo: Number(m[2]), y: m[3] || "" };
  }

  function datesLookSame(a, b) {
    const pa = parseDateParts(a);
    const pb = parseDateParts(b);
    if (!pa || !pb) return false;
    if (pa.d !== pb.d || pa.mo !== pb.mo) return false;
    if (pa.y && pb.y) {
      const ya = pa.y.length === 4 ? pa.y.slice(2) : pa.y;
      const yb = pb.y.length === 4 ? pb.y.slice(2) : pb.y;
      return ya === yb;
    }
    return true;
  }

  function collectSheetWarnings(gradeColumns, lessons) {
    const warnings = [];
    const dateCols = gradeColumns.filter((c) => isLessonDateTitle(c.title));
    const regular = lessons.filter((l) => l.type === "lesson");
    if (dateCols.length !== regular.length) {
      const dates = dateCols.length;
      const topics = regular.length;
      const gap = Math.abs(dates - topics);
      const pct = gap / Math.max(dates, topics, 1);
      const severe = pct > 0.2;
      warnings.push({
        detail: false,
        severe,
        text:
          `Кількість дат у сітці (${dates}) не збігається з кількістю тем (${topics}).` +
          (severe
            ? ` Розбіжність ${Math.round(pct * 100)}% — перевірте аркуш.`
            : " Зайві теми без дат лишаємо як є."),
      });
    }
    const n = Math.min(dateCols.length, regular.length);
    let shown = 0;
    for (let i = 0; i < n; i++) {
      if (datesLookSame(dateCols[i].title, regular[i].date)) continue;
      shown += 1;
      if (shown <= 8) {
        warnings.push({
          detail: true,
          text: `Дата в сітці ${dateCols[i].title} ≠ у темах ${regular[i].date || "немає"} (заняття ${regular[i].num}).`,
        });
      }
    }
    if (shown > 8) {
      warnings.push({
        detail: true,
        text: `Ще ${shown - 8} розбіжностей дат сітка/теми.`,
      });
    }
    return warnings;
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
        kind: "skip",
        label,
        reason: "службовий аркуш (коди / звітність), не журнал дисципліни",
      };
    }

    if (!rows || !rows.length) {
      return { ok: false, kind: "error", label, reason: "порожній аркуш" };
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
    const topicsHeaderIdx =
      gradesHeaderIdx === -1
        ? -1
        : findTopicsHeaderIndex(rows, gradesHeaderIdx);

    if (gradesHeaderIdx === -1) {
      return {
        ok: false,
        kind: discipline ? "error" : "skip",
        label,
        reason: "не знайдено таблицю студентів (очікується колонка на кшталт «ПІБ»)",
      };
    }
    if (topicsHeaderIdx === -1) {
      return {
        ok: false,
        kind: "error",
        label,
        reason: "не знайдено таблицю тем занять (очікується «Тема заняття» / «Дата»+«Години»)",
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
        kind: "error",
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
      const marks = mergeSplitStatus(
        gradeColumns.map((c) =>
          c.colIdx < row.length ? cleanMark(row[c.colIdx]) : ""
        )
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
        kind: "error",
        label,
        reason: "таблицю студентів знайдено, але рядків студентів немає",
      };
    }

    const lessons = [];
    let lastLessonNum = 0;
    for (let r = topicsHeaderIdx + 1; r < rows.length; r++) {
      const parsed = parseTopicRow(rows[r] || []);
      if (!parsed) continue;
      if (parsed.type === "lesson") {
        if (!asLessonNum(parsed.num)) {
          lastLessonNum += 1;
          parsed.num = String(lastLessonNum);
        } else {
          lastLessonNum = Number(parsed.num) || lastLessonNum;
        }
      }
      lessons.push(parsed);
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
        warnings: collectSheetWarnings(gradeColumns, lessons),
        sheetName: sheetMeta.name || "",
      },
    };
  }

  function renderJournal(data, startPage) {
    const {
      discipline,
      teacher,
      gradeColumns,
      students,
      lessons,
    } = data;
    const spreads = buildSpreads(gradeColumns, lessons);

    let html = "";
    let pageNum = startPage || PAGE_START;

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
            let markCls = "";
            if (isHundredthsMark(val)) markCls = " mark-decimal";
            else if (isCompactMark(val)) markCls = " mark-compact";
            html += `<td class="col-mark${spec ? " special-mark" : ""}${markCls}">${escapeHtml(val)}</td>`;
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
    return { pack, nextPage: pageNum };
  }

  function fitStatusCells(root) {
    root.querySelectorAll(".status-fit").forEach((el) => {
      let size = parseFloat(el.style.fontSize) || 8.5;
      const min = 3.5;
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

  function journalPageCount(data) {
    const spreads = buildSpreads(data.gradeColumns, data.lessons);
    return Math.max(1, spreads.length) * 2;
  }

  function formatPageRange(start, end) {
    if (start === end) return String(start);
    return `${start}–${end}`;
  }

  function renderCover(entries, pageNum) {
    const pack = document.createElement("section");
    pack.className = "journal-pack";
    const rows = (entries || [])
      .map((s, i) => {
        const name = s.discipline || s.sheetName || "без назви";
        return (
          `<div class="toc-row">` +
          `<span class="toc-index">${i + 1}.</span>` +
          `<span class="toc-body">` +
          `<span class="toc-name">${escapeHtml(name)}</span>` +
          `<span class="toc-lead"></span>` +
          `<span class="toc-pages">${escapeHtml(formatPageRange(s.start, s.end))}</span>` +
          `</span></div>`
        );
      })
      .join("");
    pack.innerHTML = `<div class="page cover-page">
      <div class="cover-inner">
        <h2 class="toc-heading">Зміст</h2>
        <div class="toc-list">${rows}</div>
      </div>
      <div class="page-footer">${pageNum || PAGE_START}</div>
    </div>`;
    return pack;
  }

  function isDetailMode() {
    const el = document.getElementById("detail-checks");
    return !!(el && el.checked);
  }

  function visibleWarnings(list) {
    const all = list || [];
    const detail = isDetailMode();
    return all.filter((w) => w && (detail || !w.detail));
  }

  function renderReport(results) {
    lastResults = results || lastResults;
    if (!lastResults || !lastResults.length) return;
    reportEl.hidden = false;
    reportEl.innerHTML = "";
    for (const r of lastResults) {
      const item = document.createElement("div");
      if (!r.ok) {
        item.className =
          "report-item " + (r.kind === "skip" ? "skip" : "error");
        item.textContent = `${r.label} не розпізнано, пропущено: ${r.reason}.`;
      } else {
        const d = r.data;
        const warnings = visibleWarnings(d.warnings);
        const severe = warnings.some((w) => w && w.severe);
        item.className = "report-item " + (severe ? "mismatch" : "ok");
        let html =
          `<strong>${escapeHtml(r.label)}</strong>` +
          (d.teacher ? ` ${escapeHtml(d.teacher)}` : "");
        if (warnings.length) {
          html +=
            "<ul class=\"warn-list\">" +
            warnings
              .map((w) => `<li>${escapeHtml(w.text || w)}</li>`)
              .join("") +
            "</ul>";
        }
        item.innerHTML = html;
      }
      reportEl.appendChild(item);
    }
  }

  let lastPdfName = "Журнал групи.pdf";
  let lastResults = [];

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
    const pickedEl = document.getElementById("picked-name");
    if (pickedEl) {
      pickedEl.textContent = group ? `Журнал групи ${group}` : "Журнал групи";
    }
    const stagePrint = document.getElementById("stage-print");
    if (stagePrint) {
      stagePrint.hidden = ok.length === 0;
      if (ok.length) stagePrint.scrollIntoView({ behavior: "smooth", block: "start" });
    }
    if (ok.length) {
      let pageNum = PAGE_START + 1;
      const tocEntries = ok.map((r) => {
        const start = pageNum;
        const end = start + journalPageCount(r.data) - 1;
        pageNum = end + 1;
        return {
          discipline: r.data.discipline,
          sheetName: r.data.sheetName,
          start,
          end,
        };
      });
      previewEl.appendChild(renderCover(tocEntries, PAGE_START));
      pageNum = PAGE_START + 1;
      for (const r of ok) {
        const rendered = renderJournal(r.data, pageNum);
        previewEl.appendChild(rendered.pack);
        pageNum = rendered.nextPage;
      }
    }
    if (printBtn) printBtn.disabled = ok.length === 0;
    requestAnimationFrame(() => requestAnimationFrame(afterLayout));
  }

  function excelCellText(cell) {
    if (!cell) return "";
    let w = cell.w != null ? String(cell.w).trim() : "";
    if (!w) {
      try {
        w = String(XLSX.utils.format_cell(cell) || "").trim();
      } catch (err) {
        w = "";
      }
    }
    const v = cell.v;
    const asDate = cell.t === "d" || isDateObj(v);
    if (asDate) {
      if (w && looksLikeFormattedDate(w)) return formatDateValue(w);
      if (isDateObj(v)) return formatJsDate(v);
    }
    if (cell.t === "n" && w && looksLikeFormattedDate(w)) {
      return formatDateValue(w);
    }
    if (typeof v === "number") return String(v);
    if (v == null || v === "") return "";
    return String(v);
  }

  function rowsFromSheet(sheet) {
    if (!sheet || !sheet["!ref"]) return [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      let has = false;
      for (let c = 0; c <= range.e.c; c++) {
        const v = excelCellText(sheet[XLSX.utils.encode_cell({ r, c })]);
        row.push(v);
        if (String(v).trim()) has = true;
      }
      if (has) rows.push(row);
    }
    return rows;
  }

  function rowsFromWorkbook(buffer) {
    if (typeof XLSX === "undefined") {
      throw new Error("Бібліотека Excel не завантажилась. Спробуйте CSV.");
    }
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    return wb.SheetNames.map((name) => ({
      name,
      rows: rowsFromSheet(wb.Sheets[name]),
    }));
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

  function openPrintView() {
    fitStatusCells(previewEl);
    if (!previewEl.querySelector(".page")) return;
    const w = window.open("", "_blank");
    if (!w) {
      window.print();
      return;
    }
    const title = lastPdfName.replace(/\.pdf$/i, "") || "Журнал групи";
    const styles = Array.from(
      document.querySelectorAll("style, link[rel='stylesheet']")
    )
      .map((el) => el.outerHTML)
      .join("\n");
    w.document.open();
    w.document.write(`<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<title>${escapeHtml(title)}</title>
${styles}
<style>
  html, body { background: #fff !important; padding: 0 !important; margin: 0 !important; }
  .app-chrome, .print-hint { display: none !important; }
  .page { margin: 0 auto 0 auto; box-shadow: none; }
</style>
</head>
<body>${previewEl.innerHTML}</body>
</html>`);
    w.document.close();
    w.focus();
    setTimeout(() => {
      try {
        w.print();
      } catch (e) {
        window.print();
      }
    }, 300);
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

  printBtn.addEventListener("click", () => openPrintView());
  const detailChecks = document.getElementById("detail-checks");
  if (detailChecks) {
    detailChecks.addEventListener("change", () => renderReport(lastResults));
  }
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
    if (!q) return;
    try {
      const res = await fetch(encodeURI(q));
      if (!res.ok) return;
      const lower = q.toLowerCase();
      if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        processSheets(rowsFromWorkbook(await res.arrayBuffer()), { fileName: q });
      } else {
        processSheets([{ name: q, rows: parseCsvText(await res.text()) }], { fileName: q });
      }
    } catch (e) {
      /* ignore */
    }
  }
  tryAutoload();
})();
