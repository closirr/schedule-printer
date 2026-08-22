(() => {
  const COLS_PER_PAGE = 28;
  const ROWS_PER_PAGE = 28;
  const PAGE_START = 1;
  const MIN_GRADE_ROWS = 30;

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
    /^(?:н(?:[\\/]?а)?|н\/?б|нб|н\/д|н\/з|н[\\/]\d+(?:[.,]\d+)?|\d{1,2}(?:[.,]\d+)?|\d{1,2}[\\/]\d{1,2}|відм\.?|доб\.?|задов\.?|п)$/i;

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
    if (!s || /^[-–—−]+$/.test(s) || /^[!?.,…]+$/.test(s)) return "";
    if (/^[nн]\s*[\\/]\s*$/i.test(s)) return "н";
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
    const asNum = s.replace(",", ".");
    if (/^\d{1,2}([.]\d+)?$/.test(asNum)) {
      const n = Number(asNum);
      if (Number.isFinite(n)) return String(Math.round(n));
    }
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
    return t === "н/а" || t === "н/д" || t === "н/з";
  }

  function isSlashPairMark(val) {
    return /^\d{1,2}[\\/]\d{1,2}$/.test(String(val || "").trim());
  }

  function isCompactMark(val) {
    const t = String(val || "").trim();
    if (/^[nн]$/i.test(t)) return false;
    return isNaNdMark(t) || isSlashPairMark(t) || !!abbrevVerbalGrade(t);
  }

  function isTransferTick(val) {
    return /^[пp]$/i.test(String(val || "").trim());
  }

  function looksLikeTransfer(text) {
    return fold(text).includes("перезарахув");
  }

  function normalizeStatusText(s) {
    const raw = String(s || "").replace(/\s+/g, " ").trim();
    if (!raw) return "";
    if (looksLikeTransfer(raw)) return "перезараховано";
    return raw.replace(/\s+[пp]\s*$/i, "").trim();
  }

  function combineStatusParts(parts) {
    const cleaned = parts
      .map((p) => String(p).replace(/\s+/g, " ").trim())
      .filter((p) => p && !isTransferTick(p));
    if (cleaned.some(looksLikeTransfer)) return "перезараховано";
    const kept = [];
    for (const s of cleaned) {
      if (kept.some((k) => k.includes(s))) continue;
      const i = kept.findIndex((k) => s.includes(k));
      if (i >= 0) kept[i] = s;
      else kept.push(s);
    }
    return normalizeStatusText(kept.join(" ").replace(/\s+/g, " ").trim());
  }

  function mergeSplitStatus(marks) {
    const out = marks.slice();
    for (let i = 0; i < out.length; i++) {
      if (!isLongStatus(out[i])) continue;
      const parts = [out[i]];
      let j = i + 1;
      while (j < out.length) {
        const v = out[j];
        if (!String(v || "").trim() || isTransferTick(v)) {
          if (isTransferTick(v)) out[j] = "";
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
      t.includes("класрум") ||
      t.includes("завдання в") ||
      t.includes("домашн") ||
      t.includes("підручник") ||
      t.includes("параграф") ||
      t.startsWith("повтор") ||
      t.startsWith("повт ") ||
      t === "підготовка"
    );
  }

  function findHomeworkCol(headerRow) {
    const cells = headerRow || [];
    for (let i = 3; i < cells.length; i++) {
      const t = fold(cells[i]);
      if (
        t.includes("домашн") ||
        t.includes("задано") ||
        t.includes("підручник") ||
        t === "дз"
      ) {
        return i;
      }
    }
    return -1;
  }

  function joinUniqueTexts(parts) {
    const out = [];
    for (const raw of parts) {
      const t = String(raw || "").replace(/\s+/g, " ").trim();
      if (!t) continue;
      const frag = t.replace(/^[.,;:\-\s]+/, "").trim();
      if (frag && out.some((x) => fold(x).includes(fold(frag)))) continue;
      out.push(t);
    }
    return out.join(" ").replace(/\s+/g, " ").trim();
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

  function parseTopicRow(row, hwCol) {
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
    if (isHoursValue(third)) hours = stripTrailingZero(third);

    const pieces = [];
    if (third && !hours && !looksLikeLooseDate(third)) {
      pieces.push({ i: 2, t: third });
    }
    for (let i = 3; i < cells.length; i++) {
      if (cells[i]) pieces.push({ i, t: cells[i] });
    }

    let topic = "";
    let hw = "";
    const knownHw = Number.isInteger(hwCol) && hwCol >= 3;
    if (knownHw) {
      topic = joinUniqueTexts(
        pieces.filter((p) => p.i < hwCol).map((p) => p.t)
      );
      hw = joinUniqueTexts(pieces.filter((p) => p.i >= hwCol).map((p) => p.t));
    } else {
      const GAP = 4;
      const TOPIC_ZONE = 8;
      if (pieces.length && pieces[0].i >= TOPIC_ZONE) {
        hw = joinUniqueTexts(pieces.map((p) => p.t));
      } else if (pieces.length) {
        const topicBits = [pieces[0].t];
        const hwBits = [];
        let last = pieces[0].i;
        for (let k = 1; k < pieces.length; k++) {
          if (pieces[k].i - last >= GAP) {
            hwBits.push(...pieces.slice(k).map((p) => p.t));
            break;
          }
          if (looksLikeHwText(pieces[k].t) && !looksLikeHwText(topicBits.join(" "))) {
            hwBits.push(pieces[k].t);
          } else {
            topicBits.push(pieces[k].t);
          }
          last = pieces[k].i;
        }
        topic = joinUniqueTexts(topicBits);
        hw = joinUniqueTexts(hwBits);
      }
    }

    if (!topic && hw) {
      if (looksLikeExamTopic(hw) || looksLikeConsultation(hw)) {
        topic = hw;
        hw = "";
      }
    }
    if (fold(hw) && fold(topic) && fold(hw) === fold(topic)) hw = "";

    const examish =
      looksLikeExamTopic(topic) ||
      looksLikeExamTopic(first) ||
      looksLikeExamTopic(second);
    if (!examish && !topic && hw) return null;
    if (!examish && !hours && !topic) return null;
    if (!date && !topic && !hw) return null;
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
    if (looksLikeFormattedDate(s) || isSlashPairMark(s)) return false;
    if (/^\d{1,2}\.\d{1,2}(\.\d{2,4})?$/.test(s)) return false;
    return !SHORT_MARK_RE.test(s);
  }

  function firstLongStatus(marks) {
    for (let i = 0; i < marks.length; i++) {
      if (isLongStatus(marks[i])) {
        return { col: i, text: normalizeStatusText(marks[i]) };
      }
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

  function cleanHeaderTitle(s) {
    return String(s || "")
      .replace(/[\r\n\t]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function isPlausibleGradeHeader(title) {
    const raw = cleanHeaderTitle(title);
    if (!raw) return false;
    if (isLessonDateTitle(raw) || isSpecialTitle(raw)) return true;
    if (/^\d{1,3}$/.test(raw)) return false;
    const t = fold(raw);
    return /тест|зан|зошит|діагност|контрол|бігунок|додатков|вступн|підсумков|академ|бальна|тематич|семестр|залік|екзамен|іспит|скориг|д\/з|^кр$|^рзм$|^ткр$/.test(
      t
    );
  }

  function isMetaSheet(name) {
    const t = fold(name);
    return (
      t === "код" ||
      t === "коди" ||
      t.startsWith("коди ") ||
      t.startsWith("код ") ||
      t.startsWith("звітн") ||
      t.includes("звітність") ||
      t.includes("рейтинг") ||
      t.includes("карта студент") ||
      t.includes("титул") ||
      t === "зміст" ||
      t.startsWith("зміст ") ||
      t.includes("список груп") ||
      t.includes("зведен") ||
      t.includes("відомість")
    );
  }

  function afterColon(s) {
    const m = String(s || "").match(/^[^:]{1,48}:\s*(.+)$/);
    return m ? m[1].trim() : "";
  }

  function looksLikeGroupCode(s) {
    const t = String(s || "").replace(/\s+/g, "");
    return /^[A-Za-zА-Яа-яІіЇїЄєҐґ]{1,8}[-–]?\d{1,3}[A-Za-zА-Яа-яІіЇїЄєҐґ]{0,3}$/.test(
      t
    );
  }

  function parseTitlePage(rows) {
    const cells = [];
    for (const row of rows || []) {
      for (const c of row || []) {
        const t = String(c || "")
          .replace(/\s+/g, " ")
          .trim();
        if (t) cells.push(t);
      }
    }
    const collegeParts = [];
    let journalTitle = "";
    let subtitle = "";
    let department = "";
    let group = "";
    let period = "";
    let course = "";
    let specialty = "";
    for (const t of cells) {
      const f = fold(t);
      if (f === "журнал" || f.startsWith("журнал ")) {
        journalTitle = t.toUpperCase() === t ? t : "ЖУРНАЛ";
        continue;
      }
      if (f.includes("обліку роботи")) {
        subtitle = t.replace(/\s+/g, " ").trim();
        continue;
      }
      if (f.includes("відділен") || f.includes("віділен")) {
        department = afterColon(t) || t;
        continue;
      }
      if (/^курс\b/.test(f) || f.startsWith("курс:")) {
        course = afterColon(t) || t.replace(/^курс:?\s*/i, "").trim();
        continue;
      }
      if (f.includes("спеціальн")) {
        specialty = afterColon(t);
        continue;
      }
      if (f.includes("семестр") || /\bн\.?\s*р\.?\b/i.test(t)) {
        period = t.replace(/^на\s+/i, "").trim();
        continue;
      }
      if (f.startsWith("груп")) {
        group = afterColon(t) || group;
        continue;
      }
      if (!journalTitle && t.length > 18) {
        collegeParts.push(t);
        continue;
      }
      if (!group && looksLikeGroupCode(t)) group = t;
    }
    return {
      college: collegeParts.join(" ").replace(/\s+/g, " ").trim(),
      collegeLines: collegeParts,
      journalTitle: journalTitle || "ЖУРНАЛ",
      subtitle,
      department,
      group,
      period,
      course,
      specialty,
    };
  }

  function titlePageUseful(title) {
    return !!(title && (title.college || title.group || title.period));
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

  function collectSheetWarnings(gradeColumns, lessons, students) {
    const warnings = [];
    const dateCols = gradeColumns.filter((c) => isLessonDateTitle(c.title));
    const regular = lessons.filter((l) => l.type === "lesson");
    const weirdHeaders = (gradeColumns || []).filter(
      (c) => !isLessonDateTitle(c.title) && !isSpecialTitle(c.title)
    );
    const weirdShare = weirdHeaders.length / Math.max(gradeColumns.length, 1);
    if (weirdHeaders.length >= 8 && weirdShare > 0.15) {
      const sample = weirdHeaders[0].title;
      warnings.push({
        detail: false,
        severe: true,
        text:
          `У шапці оцінок замість дат стоять сторонні назви (${weirdHeaders.length} кол., наприклад «${sample}»). ` +
          `Аркуш заповнено нестандартно — перевірте Excel.`,
      });
    }

    let dateMarks = 0;
    let dateStatus = 0;
    let filled = 0;
    for (const st of students || []) {
      if (st.statusText && looksLikeFormattedDate(st.statusText)) dateStatus += 1;
      for (const m of st.marks || []) {
        if (!String(m || "").trim()) continue;
        filled += 1;
        if (looksLikeFormattedDate(m) || /^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(String(m).trim())) {
          dateMarks += 1;
        }
      }
    }
    if (dateStatus >= 2 || dateMarks >= 4 || (filled && dateMarks / filled > 0.08)) {
      warnings.push({
        detail: false,
        severe: true,
        text: "У клітинках оцінок знайдено дати. Ймовірно Excel сприйняв дроби на кшталт 10/5 як дату, і рядок журналу роз'їхався. Перевірте аркуш.",
      });
    }
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
        reason: "службовий, пропущено",
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
      const raw = cleanHeaderTitle(headerRow[colIdx]);
      if (!raw) continue;
      const asDate = looksLikeLooseDate(raw) ? formatDateValue(raw) : "";
      const title =
        asDate && isLessonDateTitle(asDate) ? asDate : raw;
      if (!isPlausibleGradeHeader(title)) continue;
      gradeColumns.push({ colIdx, title });
    }
    if (!gradeColumns.length) {
      let named = 0;
      for (let r = gradesHeaderIdx + 1; r < topicsHeaderIdx; r++) {
        const name = shortenName(String((rows[r] || [])[1] || "").replace(/\/\d+$/, ""));
        if (name) named += 1;
      }
      if (!named) {
        return {
          ok: false,
          kind: "skip",
          label,
          reason: "порожній шаблон, пропущено",
        };
      }
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
      if (!name) continue;
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
        kind: "skip",
        label,
        reason: "порожній шаблон, пропущено",
      };
    }

    const hwCol = findHomeworkCol(rows[topicsHeaderIdx] || []);
    const lessons = [];
    let lastLessonNum = 0;
    for (let r = topicsHeaderIdx + 1; r < rows.length; r++) {
      const parsed = parseTopicRow(rows[r] || [], hwCol);
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
        warnings: collectSheetWarnings(gradeColumns, lessons, students),
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
              html += statusTdHtml(normalizeStatusText(val), cEnd - cI + emptyC);
              spanned = true;
              break;
            }
            const spec = isSpecialTitle(gradeColumns[cI].title);
            const compact = isCompactMark(val) ? " mark-compact" : "";
            html += `<td class="col-mark${spec ? " special-mark" : ""}${compact}">${escapeHtml(val)}</td>`;
          }
        }
        if (!spanned) {
          for (let i = 0; i < emptyC; i++) html += `<td class="col-mark">&nbsp;</td>`;
        }
        html += `</tr>`;
      }

      for (let extra = students.length + 1; extra <= Math.max(MIN_GRADE_ROWS, students.length); extra++) {
        html += `<tr class="filler-row grades-row"><td class="col-num">${extra}</td><td class="col-name">&nbsp;</td>${'<td class="col-mark">&nbsp;</td>'.repeat(COLS_PER_PAGE)}</tr>`;
      }

      html += `</tbody></table></div><div class="page-footer">${pageNum}</div></div>`;
      pageNum += 1;

      const sLessons = spread.pageLessons;
      const emptyL = Math.max(0, ROWS_PER_PAGE - sLessons.length);

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
        const topic = item.topic
          ? `<div class="topic-fit">${escapeHtml(item.topic)}</div>`
          : "";
        const hw = item.hw
          ? `<div class="hw-fit">${escapeHtml(item.hw)}</div>`
          : "";
        html += `<tr class="lesson-row"><td>${escapeHtml(item.num || "")}</td><td>${escapeHtml(item.date || "")}</td><td>${escapeHtml(item.hours || "")}</td><td class="col-lesson-topic">${topic}</td><td class="col-lesson-hw">${hw}</td><td></td></tr>`;
      }
      for (let i = 0; i < emptyL; i++) {
        html += `<tr class="filler-row lesson-row"><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>`;
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

  function fitClippedCells(root) {
    root.querySelectorAll(".hw-fit, .topic-fit").forEach((el) => {
      const start = el.classList.contains("topic-fit") ? 8.5 : 8;
      let size = parseFloat(el.style.fontSize) || start;
      const min = 5.5;
      let guard = 0;
      while (guard++ < 40 && size > min && el.scrollHeight > el.clientHeight + 0.5) {
        size -= 0.25;
        el.style.fontSize = size + "pt";
      }
      if (el.scrollHeight <= el.clientHeight + 0.5) return;
      let text = el.textContent || "";
      guard = 0;
      while (guard++ < 80 && text.length > 3 && el.scrollHeight > el.clientHeight + 0.5) {
        const cut = text.replace(/\s+\S+\s*$/, "").replace(/\s+$/, "");
        text = cut !== text && cut.length ? cut : text.slice(0, -1);
        el.textContent = text + "…";
      }
    });
  }

  function fitPrintCells(root) {
    fitStatusCells(root);
    fitClippedCells(root);
  }

  function afterLayout() {
    fitPrintCells(previewEl);
  }

  function guessGroupName(fileName, okResults, title) {
    if (title && title.group) return String(title.group).replace("–", "-");
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

  function titleFieldRow(label, value) {
    if (!value) return "";
    return (
      `<div class="title-field">` +
      `<span class="title-field-label">${escapeHtml(label)}:</span>` +
      `<span class="title-field-rule"></span>` +
      `<span class="title-field-value">${escapeHtml(value)}</span>` +
      `</div>`
    );
  }

  function renderTitlePage(title) {
    const pack = document.createElement("section");
    pack.className = "journal-pack";
    const collegeHtml = (title.collegeLines && title.collegeLines.length
      ? title.collegeLines
      : title.college
        ? [title.college]
        : []
    )
      .map((line) => `<div>${escapeHtml(line)}</div>`)
      .join("");
    pack.innerHTML = `<div class="page title-page">
      <div class="page-content title-content">
        <div class="title-frame">
          <div class="title-inner">
            <div class="title-college">${collegeHtml}</div>
            <div class="title-word">${escapeHtml(title.journalTitle || "ЖУРНАЛ")}</div>
            <div class="title-sub">${escapeHtml(title.subtitle || "обліку роботи академічної групи та викладачів")}</div>
            <div class="title-fields">
              ${titleFieldRow("Відділення", title.department)}
              ${titleFieldRow("Група", title.group)}
              ${titleFieldRow("Курс", title.course)}
              ${titleFieldRow("Спеціальність", title.specialty)}
            </div>
            <div class="title-period">${title.period ? escapeHtml("на " + title.period.replace(/^на\s+/i, "")) : ""}</div>
          </div>
        </div>
      </div>
    </div>`;
    return pack;
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
    if (!reportEl || !lastResults || !lastResults.length) return;
    reportEl.hidden = false;
    reportEl.innerHTML = "";
    for (const r of lastResults) {
      const item = document.createElement("div");
      if (!r.ok) {
        item.className =
          "report-item " + (r.kind === "skip" ? "skip" : "error");
        item.textContent =
          r.kind === "skip"
            ? `${r.label} — ${r.reason || "службовий, пропущено"}.`
            : `${r.label} не розпізнано, пропущено: ${r.reason}.`;
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
    if (!previewEl) return;
    const fileName = (opts && opts.fileName) || "";
    previewEl.innerHTML = "";
    const titleSheet = (sheets || []).find((s) =>
      fold(s.name || "").includes("титул")
    );
    const title = titleSheet ? parseTitlePage(titleSheet.rows) : null;
    const hasTitle = titlePageUseful(title);
    const results = sheets.map((sheet, index) =>
      parseJournalSheet(sheet.rows, { index, name: sheet.name })
    );
    renderReport(results);
    const ok = results.filter((r) => r.ok);
    const group = guessGroupName(fileName, ok, title);
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
      let pageNum = PAGE_START;
      if (hasTitle) {
        previewEl.appendChild(renderTitlePage(title));
        pageNum += 1;
      }
      const tocPage = pageNum;
      pageNum += 1;
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
      previewEl.appendChild(renderCover(tocEntries, tocPage));
      pageNum = tocPage + 1;
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
    const fmt = String(cell.z || "");
    const fmtHasYear = /y/i.test(fmt);
    const wNorm = w.replace(/\s/g, "");
    if (asDate) {
      if (!fmtHasYear && /^\d{1,2}[./]\d{1,2}$/.test(wNorm)) return wNorm;
      if (w && looksLikeFormattedDate(w)) return formatDateValue(w);
      if (isDateObj(v)) return formatJsDate(v);
    }
    if (cell.t === "n" && w && looksLikeFormattedDate(w)) {
      if (!fmtHasYear && /^\d{1,2}[./]\d{1,2}$/.test(wNorm)) return wNorm;
      return formatDateValue(w);
    }
    if (typeof v === "number") {
      const nw = w.replace(/\s/g, "").replace(",", ".");
      if (nw && /^-?\d+([.]\d+)?$/.test(nw)) return nw;
      return String(v);
    }
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

  function xmlEscape(s) {
    return String(s ?? "")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function xmlUnescape(s) {
    return String(s ?? "")
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&");
  }

  function createSharedStrings(xml) {
    if (!xml) {
      return { exists: false, add() { return null; }, serialize() { return ""; } };
    }
    let sst = xml;
    let siCount = (sst.match(/<si\b/g) || []).length;
    return {
      exists: true,
      add(text) {
        const t = String(text ?? "");
        const body =
          /^\s|\s$/.test(t)
            ? `<t xml:space="preserve">${xmlEscape(t)}</t>`
            : `<t>${xmlEscape(t)}</t>`;
        sst = sst.replace("</sst>", `<si>${body}</si></sst>`);
        const idx = siCount;
        siCount += 1;
        sst = sst
          .replace(/\bcount="\d+"/, `count="${siCount}"`)
          .replace(/\buniqueCount="\d+"/, `uniqueCount="${siCount}"`);
        return idx;
      },
      serialize() {
        return sst;
      },
    };
  }

  function makePatchedCell(addr, text, attrs, sst) {
    const style = ((attrs || "").match(/\bs="[^"]*"/) || [""])[0];
    const styleBit = style ? " " + style : "";
    if (sst && sst.exists) {
      const idx = sst.add(text);
      return `<c r="${addr}"${styleBit} t="s"><v>${idx}</v></c>`;
    }
    return `<c r="${addr}"${styleBit} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
  }

  function setSheetCellXml(xml, addr, text, sst) {
    const re = new RegExp(
      `<c([^>]*\\br="${addr}"(?![A-Z0-9])[^>]*)(\/>|>[\\s\\S]*?</c>)`
    );
    if (re.test(xml)) {
      return xml.replace(re, (full, attrs) => makePatchedCell(addr, text, attrs, sst));
    }
    const next = makePatchedCell(addr, text, "", sst);
    const rowNum = String(addr).replace(/^[A-Z]+/i, "");
    const rowRe = new RegExp(`(<row[^>]*\\br="${rowNum}"(?![0-9])[^>]*>)`);
    if (rowRe.test(xml)) return xml.replace(rowRe, `$1${next}`);
    return xml.replace(
      "</sheetData>",
      `<row r="${rowNum}">${next}</row></sheetData>`
    );
  }

  function zipNormalize(p) {
    const stack = [];
    for (const seg of String(p || "").replace(/\\/g, "/").split("/")) {
      if (!seg || seg === ".") continue;
      if (seg === "..") stack.pop();
      else stack.push(seg);
    }
    return stack.join("/");
  }

  function zipDir(p) {
    return String(p || "").replace(/\/[^/]+$/, "");
  }

  function resolveRelTarget(partPath, target) {
    const t = String(target || "").replace(/^\/+/, "");
    if (/^[a-z][a-z0-9+.-]*:/i.test(t)) return t;
    return zipNormalize(zipDir(partPath) + "/" + t);
  }

  function relTargetFrom(partPath, absPath) {
    const from = zipDir(partPath).split("/").filter(Boolean);
    const to = String(absPath).split("/").filter(Boolean);
    let i = 0;
    while (i < from.length && i < to.length && from[i] === to[i]) i += 1;
    return from.slice(i).map(() => "..").concat(to.slice(i)).join("/");
  }

  function serializeRels(rels) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      (rels || [])
        .map((r) => {
          const mode = r.mode ? ` TargetMode="${xmlEscape(r.mode)}"` : "";
          return `<Relationship Id="${xmlEscape(r.id)}" Type="${xmlEscape(r.type)}" Target="${xmlEscape(r.target)}"${mode}/>`;
        })
        .join("") +
      `</Relationships>`
    );
  }

  function clonePartPath(srcAbs, stamp) {
    const m = String(srcAbs).match(/^(.*)\/([^/]+)\.([^.]+)$/);
    if (!m) return srcAbs + "_g" + stamp;
    return `${m[1]}/${m[2]}_g${stamp}.${m[3]}`;
  }

  async function remapSheetRels(zip, out, relsXml, srcPartPath, destPartPath, stamp) {
    if (!relsXml) return "";
    const next = [];
    for (const r of parseXmlRels(relsXml)) {
      if (r.mode === "External" || /^[a-z][a-z0-9+.-]*:/i.test(r.target || "")) {
        next.push(r);
        continue;
      }
      const srcAbs = resolveRelTarget(srcPartPath, r.target);
      const destAbs = clonePartPath(srcAbs, stamp);
      const srcFile = zip.file(srcAbs);
      if (srcFile && !out.file(destAbs)) {
        out.file(destAbs, await srcFile.async("uint8array"));
      }
      const srcRels = zipDir(srcAbs) + "/_rels/" + srcAbs.split("/").pop() + ".rels";
      const destRels = zipDir(destAbs) + "/_rels/" + destAbs.split("/").pop() + ".rels";
      const rf = zip.file(srcRels);
      if (rf && !out.file(destRels)) {
        out.file(destRels, await rf.async("uint8array"));
      }
      next.push({ ...r, target: relTargetFrom(destPartPath, destAbs) });
    }
    return serializeRels(next);
  }

  function attrOf(tag, name) {
    const m = String(tag || "").match(
      new RegExp(`\\b${name}="([^"]*)"`)
    );
    return m ? xmlUnescape(m[1]) : "";
  }

  function parseXmlSheets(wbXml) {
    const out = [];
    const re = /<sheet\b[^>]*\/>/g;
    let m;
    while ((m = re.exec(wbXml))) {
      out.push({
        name: attrOf(m[0], "name"),
        rId: attrOf(m[0], "r:id"),
      });
    }
    return out;
  }

  function parseXmlRels(relsXml) {
    const out = [];
    const re = /<Relationship\b[^>]*\/>/g;
    let m;
    while ((m = re.exec(relsXml))) {
      out.push({
        id: attrOf(m[0], "Id"),
        type: attrOf(m[0], "Type"),
        target: attrOf(m[0], "Target"),
        mode: attrOf(m[0], "TargetMode"),
      });
    }
    return out;
  }

  function findLabelValueAddr(sheet, needles) {
    if (!sheet || !sheet["!ref"]) return "";
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const maxR = Math.min(range.e.r, 14);
    const maxC = Math.min(range.e.c, 40);
    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const t = fold(excelCellText(sheet[addr]));
        if (!needles.some((n) => t.includes(fold(n)))) continue;
        return XLSX.utils.encode_cell({ r, c: c + 1 });
      }
    }
    return "";
  }

  function scanTitleAddrs(sheet) {
    const addrs = {};
    if (!sheet || !sheet["!ref"]) return addrs;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const raw = excelCellText(sheet[addr]);
        if (!String(raw).trim()) continue;
        const f = fold(raw);
        if (f.includes("відділен") || f.includes("віділен")) addrs.department = addr;
        else if (/^курс\b/.test(f) || f.startsWith("курс:")) addrs.course = addr;
        else if (f.includes("спеціальн")) addrs.specialty = addr;
        else if (f.includes("семестр") || /\bн\.?\s*р\.?\b/i.test(raw)) addrs.period = addr;
        else if (looksLikeGroupCode(raw)) addrs.group = addr;
      }
    }
    return addrs;
  }

  function rewriteLabeledCell(oldText, value, fallbackLabel) {
    const v = String(value || "").trim();
    if (!v) return String(oldText || "");
    const old = String(oldText || "");
    if (/^[^:]{1,48}:/.test(old)) return old.replace(/^([^:]{1,48}:)\s*.*$/, "$1 " + v);
    if (fallbackLabel) return fallbackLabel + ": " + v;
    return v;
  }

  function uniqueSheetName(name, used) {
    let n = String(name || "")
      .replace(/[:\\/?*\[\]]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!n) n = "Аркуш";
    n = n.slice(0, 31);
    let out = n;
    let i = 2;
    while (used.has(fold(out))) {
      const suffix = ` (${i})`;
      out = n.slice(0, Math.max(1, 31 - suffix.length)) + suffix;
      i += 1;
    }
    used.add(fold(out));
    return out;
  }

  function journalLayoutOf(rows) {
    const gradesHeaderIdx = findHeaderIndex(rows, [
      "піб студента",
      "прізвище та ініціали студента",
      "прізвище студента",
      "піб",
    ]);
    if (gradesHeaderIdx === -1) return null;
    const topicsHeaderIdx = findTopicsHeaderIndex(rows, gradesHeaderIdx);
    if (topicsHeaderIdx === -1) return null;
    let named = 0;
    for (let r = gradesHeaderIdx + 1; r < topicsHeaderIdx; r++) {
      const name = shortenName(String((rows[r] || [])[1] || "").replace(/\/\d+$/, ""));
      if (name) named += 1;
    }
    let dates = 0;
    const header = rows[gradesHeaderIdx] || [];
    for (let c = 2; c < header.length; c++) {
      const raw = cleanHeaderTitle(header[c]);
      if (raw && isPlausibleGradeHeader(raw)) dates += 1;
    }
    return { gradesHeaderIdx, topicsHeaderIdx, named, dates };
  }

  function pickPrototypeName(sheets) {
    let best = "";
    let bestScore = Infinity;
    for (const s of sheets || []) {
      if (isMetaSheet(s.name || "")) continue;
      const lay = journalLayoutOf(s.rows || []);
      if (!lay) continue;
      const score = lay.named * 1000 + lay.dates * 10;
      if (score < bestScore) {
        bestScore = score;
        best = s.name;
      }
    }
    return best;
  }

  function colLettersToNum(letters) {
    let n = 0;
    for (const ch of String(letters || "")) {
      const c = ch.toUpperCase().charCodeAt(0);
      if (c < 65 || c > 90) continue;
      n = n * 26 + (c - 64);
    }
    return n;
  }

  function colNumToLetters(n) {
    let s = "";
    let x = Number(n) || 0;
    while (x > 0) {
      const r = (x - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      x = Math.floor((x - 1) / 26);
    }
    return s || "A";
  }

  function xmlSheetBounds(xml) {
    let maxR = 0;
    let maxC = 0;
    let m;
    const rowRe = /<row r="(\d+)"/g;
    while ((m = rowRe.exec(xml))) maxR = Math.max(maxR, Number(m[1]) || 0);
    const refRe = /\b(?:r|ref)="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/g;
    while ((m = refRe.exec(xml))) {
      maxC = Math.max(maxC, colLettersToNum(m[1]));
      maxR = Math.max(maxR, Number(m[2]) || 0);
      if (m[3]) {
        maxC = Math.max(maxC, colLettersToNum(m[3]));
        maxR = Math.max(maxR, Number(m[4]) || 0);
      }
    }
    const colRe = /<col\b[^>]*\bmax="(\d+)"/g;
    while ((m = colRe.exec(xml))) maxC = Math.max(maxC, Number(m[1]) || 0);
    return { maxR, maxC };
  }

  function trimSheetXml(xml, bounds) {
    const maxR = bounds && bounds.maxR;
    const maxC = bounds && bounds.maxC;
    if (!xml || !maxR) return xml;
    let out = xml.replace(
      /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g,
      (full, r) => (Number(r) > maxR ? "" : full)
    );
    if (maxC) {
      out = out.replace(
        /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*\/>/g,
        (full, c1, r1, c2, r2) => {
          if (Number(r1) > maxR || Number(r2) > maxR) return "";
          if (colLettersToNum(c1) > maxC || colLettersToNum(c2) > maxC) return "";
          return full;
        }
      );
    } else {
      out = out.replace(
        /<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*\/>/g,
        (full, c1, r1, c2, r2) =>
          Number(r1) > maxR || Number(r2) > maxR ? "" : full
      );
    }
    const merges = out.match(/<mergeCell\b[^>]*\/>/g) || [];
    if (/<mergeCells\b/.test(out)) {
      if (!merges.length) {
        out = out.replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/, "");
      } else {
        out = out.replace(
          /<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/,
          `<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>`
        );
      }
    }
    const last = colNumToLetters(maxC || 1) + maxR;
    if (/<dimension\b/.test(out)) {
      out = out.replace(/<dimension\b[^>]*\/>/, `<dimension ref="A1:${last}"/>`);
    } else {
      out = out.replace(
        /<sheetViews\b/,
        `<dimension ref="A1:${last}"/><sheetViews`
      );
    }
    return out;
  }

  function isHistoryEtalonName(name) {
    const f = fold(name);
    return f.includes("всесвітн") && f.includes("істор");
  }

  function parseSubjectLine(line) {
    const t = String(line || "").replace(/\s+/g, " ").trim();
    if (!t) return null;
    const m = t.split(/\s*[—–\-|]\s+|\s+[—–]\s+/);
    if (m.length >= 2) {
      const subject = m[0].trim();
      const teacher = m.slice(1).join(" ").trim();
      if (subject) return { subject, teacher };
    }
    return { subject: t, teacher: "" };
  }

  const makeState = {
    buffer: null,
    fileName: "",
    prototypeName: "",
    titleName: "",
    titleAddrs: {},
    titleRaw: {},
    source: "",
  };

  const DEFAULT_TEMPLATE = "template.xlsx";
  let pasteTimer = 0;
  let lastPasteKey = "";

  function parsePasteSubjects(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map(parseSubjectLine)
      .filter(Boolean);
  }

  function applyPasteList() {
    const area = document.getElementById("subj-paste");
    if (!area) return;
    const parsed = parsePasteSubjects(area.value);
    if (!parsed.length) return;
    const key = parsed
      .map((p) => p.subject + "\t" + p.teacher)
      .join("\n");
    if (key === lastPasteKey) return;
    lastPasteKey = key;
    const list = document.getElementById("subj-list");
    if (!list) return;
    list.innerHTML = "";
    parsed.forEach((p) => addSubjectRow(p.subject, p.teacher));
  }

  function schedulePasteSync() {
    if (pasteTimer) clearTimeout(pasteTimer);
    pasteTimer = setTimeout(() => {
      pasteTimer = 0;
      applyPasteList();
    }, 180);
  }

  function setMakeStatus(text, isErr) {
    const el = document.getElementById("make-status");
    if (!el) return;
    el.textContent = text || "";
    el.className = isErr ? "err" : "";
  }

  function addSubjectRow(subject, teacher) {
    const list = document.getElementById("subj-list");
    if (!list) return;
    const row = document.createElement("div");
    row.className = "subj-row";
    const a = document.createElement("input");
    a.type = "text";
    a.placeholder = "Предмет";
    a.value = subject || "";
    a.setAttribute("data-k", "subject");
    const b = document.createElement("input");
    b.type = "text";
    b.placeholder = "Викладач";
    b.value = teacher || "";
    b.setAttribute("data-k", "teacher");
    const rm = document.createElement("button");
    rm.type = "button";
    rm.className = "subj-remove";
    rm.textContent = "Прибрати";
    rm.addEventListener("click", () => row.remove());
    row.appendChild(a);
    row.appendChild(b);
    row.appendChild(rm);
    list.appendChild(row);
  }

  function collectSubjectRows() {
    const list = document.getElementById("subj-list");
    if (!list) return [];
    const out = [];
    list.querySelectorAll(".subj-row").forEach((row) => {
      const subject = (row.querySelector('[data-k="subject"]') || {}).value || "";
      const teacher = (row.querySelector('[data-k="teacher"]') || {}).value || "";
      if (String(subject).trim()) {
        out.push({
          subject: String(subject).trim(),
          teacher: String(teacher).trim(),
        });
      }
    });
    return out;
  }

  function fillMakeFormFromTemplate(sheets, wb, opts) {
    const titleSheet = (sheets || []).find((s) =>
      fold(s.name || "").includes("титул")
    );
    const title = titleSheet ? parseTitlePage(titleSheet.rows) : {};
    const setVal = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.value = v || "";
    };
    setVal("make-group", title.group);
    setVal("make-course", title.course);
    setVal("make-dept", title.department);
    setVal("make-spec", title.specialty);
    setVal("make-period", title.period);
    makeState.titleName = titleSheet ? titleSheet.name : "";
    makeState.titleAddrs = {};
    makeState.titleRaw = {};
    if (titleSheet && wb && wb.Sheets[titleSheet.name]) {
      makeState.titleAddrs = scanTitleAddrs(wb.Sheets[titleSheet.name]);
      const ts = wb.Sheets[titleSheet.name];
      Object.keys(makeState.titleAddrs).forEach((k) => {
        makeState.titleRaw[k] = excelCellText(ts[makeState.titleAddrs[k]]);
      });
    }
    const list = document.getElementById("subj-list");
    if (list) list.innerHTML = "";
    if (opts && opts.skipSubjects) {
      addSubjectRow("", "");
      return;
    }
    let n = 0;
    for (const s of sheets || []) {
      if (isMetaSheet(s.name || "")) continue;
      if (!journalLayoutOf(s.rows || [])) continue;
      let discipline = "";
      let teacher = "";
      for (const row of (s.rows || []).slice(0, 8)) {
        if (!discipline) {
          discipline = cellAfterLabel(row, [
            "назва дисципліни",
            "назва предмета",
            "дисципліна",
          ]);
        }
        if (!teacher) teacher = cellAfterLabel(row, ["викладач"]);
      }
      addSubjectRow(discipline || s.name, teacher);
      n += 1;
    }
    if (!n) addSubjectRow("", "");
  }

  async function handleMakeTemplate(file, opts) {
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheets = wb.SheetNames.map((name) => ({
      name,
      rows: rowsFromSheet(wb.Sheets[name]),
    }));
    const proto = pickPrototypeName(sheets);
    if (!proto) {
      setMakeStatus(
        "У файлі немає аркуша-предмета з колонкою ПІБ і таблицею тем.",
        true
      );
      return;
    }
    makeState.buffer = buf;
    makeState.fileName = file.name || "шаблон.xlsx";
    makeState.prototypeName = proto;
    const source = (opts && opts.source) || "custom";
    fillMakeFormFromTemplate(sheets, wb, {
      skipSubjects: source === "standard",
    });
    const form = document.getElementById("make-form");
    if (form) form.hidden = false;
    makeState.source = source;
    const card = document.getElementById("tpl-standard");
    if (card) card.classList.toggle("is-on", source === "standard");
    const picked = document.getElementById("make-picked");
    if (picked) {
      picked.textContent =
        source === "standard"
          ? "Використовується зараз"
          : "Свій файл: " + (file.name || "шаблон.xlsx");
    }
    lastPasteKey = "";
    applyPasteList();
    setMakeStatus("");
  }

  function bindFileZone(zone, input, onFile) {
    if (!zone || !input) return;
    input.addEventListener("change", () => {
      const file = input.files && input.files[0];
      if (file) onFile(file);
    });
    ["dragenter", "dragover"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        zone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        zone.classList.remove("dragover");
      });
    });
    zone.addEventListener("drop", (e) => {
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) onFile(file);
    });
  }

  async function buildJournalXlsx(subjects, meta) {
    if (typeof JSZip === "undefined") {
      throw new Error("Бібліотека JSZip не завантажилась.");
    }
    if (!makeState.buffer) throw new Error("Спочатку оберіть шаблон.");
    const zip = await JSZip.loadAsync(makeState.buffer);
    const wbXml = await zip.file("xl/workbook.xml").async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const typesFile = zip.file("[Content_Types].xml");
    if (!typesFile) throw new Error("Файл шаблону пошкоджений (немає Content_Types).");
    const typesXml = await typesFile.async("string");
    const listed = parseXmlSheets(wbXml);
    const rels = parseXmlRels(relsXml);
    const relById = {};
    rels.forEach((r) => {
      relById[r.id] = r;
    });
    let proto = listed.find((s) => s.name === makeState.prototypeName) || null;
    let protoScore = Infinity;
    for (const s of listed) {
      if (isMetaSheet(s.name || "") || !relById[s.rId]) continue;
      const t = String(relById[s.rId].target || "").replace(/^\/+/, "");
      const p = t.startsWith("xl/") ? t : "xl/" + t;
      const file = zip.file(p);
      if (!file) continue;
      const xml = await file.async("string");
      const merges = Number((xml.match(/mergeCells count="(\d+)"/) || [])[1] || 0);
      if (merges < 100) continue;
      const score = (xml.match(/<v>/g) || []).length * 1000 + xml.length;
      if (score < protoScore) {
        protoScore = score;
        proto = s;
      }
    }
    if (!proto || !relById[proto.rId]) {
      throw new Error("Не знайдено аркуш-зразок у шаблоні.");
    }
    const protoName = proto.name;
    makeState.prototypeName = protoName;
    const protoTarget = String(relById[proto.rId].target || "").replace(/^\/+/, "");
    const protoPath = protoTarget.startsWith("xl/")
      ? protoTarget
      : "xl/" + protoTarget;
    let protoXml = await zip.file(protoPath).async("string");
    const protoRelsPath = protoPath
      .replace("worksheets/", "worksheets/_rels/")
      .replace(/\.xml$/, ".xml.rels");
    const protoRelsFile = zip.file(protoRelsPath);
    const protoRels = protoRelsFile ? await protoRelsFile.async("string") : "";

    const etalon = listed.find((s) => isHistoryEtalonName(s.name || "")) || proto;
    if (etalon && relById[etalon.rId]) {
      const etTarget = String(relById[etalon.rId].target || "").replace(/^\/+/, "");
      const etPath = etTarget.startsWith("xl/") ? etTarget : "xl/" + etTarget;
      const etXml =
        etalon.name === protoName ? protoXml : await zip.file(etPath).async("string");
      protoXml = trimSheetXml(protoXml, xmlSheetBounds(etXml));
    }

    const xlsxWb = XLSX.read(makeState.buffer, { type: "array", cellDates: true });
    const protoSheet = xlsxWb.Sheets[protoName];
    const discAddr =
      findLabelValueAddr(protoSheet, [
        "назва дисципліни",
        "назва предмета",
        "дисципліна",
      ]) || "C1";
    const teachAddr =
      findLabelValueAddr(protoSheet, ["викладач"]) || "C2";

    const keep = listed.filter((s) => isMetaSheet(s.name || ""));

    const sstFile = zip.file("xl/sharedStrings.xml");
    const sst = createSharedStrings(
      sstFile ? await sstFile.async("string") : ""
    );

    const usedNames = new Set();
    const outSheets = [];
    const copyKept = async (src) => {
      const rel = relById[src.rId];
      const target = String((rel && rel.target) || "").replace(/^\/+/, "");
      const path = target.startsWith("xl/") ? target : "xl/" + target;
      const xml = await zip.file(path).async("string");
      const relsP = path
        .replace("worksheets/", "worksheets/_rels/")
        .replace(/\.xml$/, ".xml.rels");
      const rf = zip.file(relsP);
      outSheets.push({
        name: uniqueSheetName(src.name, usedNames),
        xml,
        rels: rf ? await rf.async("string") : "",
        srcPart: path,
        cloneRels: false,
        patchTitle: fold(src.name || "").includes("титул"),
      });
    };
    for (const s of keep) await copyKept(s);
    for (const sub of subjects) {
      let xml = protoXml;
      xml = setSheetCellXml(xml, discAddr, sub.subject, sst);
      xml = setSheetCellXml(xml, teachAddr, sub.teacher || "", sst);
      outSheets.push({
        name: uniqueSheetName(sub.subject, usedNames),
        xml,
        rels: protoRels,
        srcPart: protoPath,
        cloneRels: true,
        patchTitle: false,
      });
    }

    const titleIdx = outSheets.findIndex((s) => s.patchTitle);
    if (titleIdx >= 0) {
      let xml = outSheets[titleIdx].xml;
      const addrs = makeState.titleAddrs || {};
      const raw = makeState.titleRaw || {};
      if (addrs.group && meta.group) {
        xml = setSheetCellXml(xml, addrs.group, meta.group, sst);
      }
      if (addrs.department && meta.department) {
        xml = setSheetCellXml(
          xml,
          addrs.department,
          rewriteLabeledCell(raw.department, meta.department, "Відділення"),
          sst
        );
      }
      if (addrs.course && meta.course) {
        xml = setSheetCellXml(
          xml,
          addrs.course,
          rewriteLabeledCell(raw.course, meta.course, "Курс"),
          sst
        );
      }
      if (addrs.specialty && meta.specialty) {
        xml = setSheetCellXml(
          xml,
          addrs.specialty,
          rewriteLabeledCell(raw.specialty, meta.specialty, "Спеціальність"),
          sst
        );
      }
      if (addrs.period && meta.period) {
        const p = String(meta.period).trim();
        const periodText = /^на\s+/i.test(p) ? p : "на " + p;
        xml = setSheetCellXml(xml, addrs.period, periodText, sst);
      }
      outSheets[titleIdx].xml = xml;
    }

    const out = new JSZip();
    const skip = new Set([
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "[Content_Types].xml",
      "xl/calcChain.xml",
    ]);
    const files = Object.keys(zip.files);
    for (const path of files) {
      const f = zip.files[path];
      if (f.dir) continue;
      if (skip.has(path)) continue;
      if (path.startsWith("xl/worksheets/")) continue;
      if (/calcChain/i.test(path)) continue;
      out.file(path, await f.async("uint8array"));
    }
    if (sst.exists) out.file("xl/sharedStrings.xml", sst.serialize());

    const sheetRels = [];
    const otherRels = rels.filter(
      (r) => !/\/relationships\/worksheet$/.test(r.type) && !/calcChain/i.test(r.target || "")
    );
    for (let i = 0; i < outSheets.length; i++) {
      const s = outSheets[i];
      const n = i + 1;
      const sheetPath = `xl/worksheets/sheet${n}.xml`;
      out.file(sheetPath, s.xml);
      let relsXml = s.rels || "";
      if (relsXml && s.cloneRels) {
        relsXml = await remapSheetRels(
          zip,
          out,
          relsXml,
          s.srcPart,
          sheetPath,
          n
        );
      }
      if (relsXml) out.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, relsXml);
      sheetRels.push({
        id: "rId" + n,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: `worksheets/sheet${n}.xml`,
      });
    }
    const wbRels = sheetRels.concat(
      otherRels.map((r, i) => ({
        id: "rId" + (sheetRels.length + i + 1),
        type: r.type,
        target: r.target,
        mode: r.mode,
      }))
    );
    out.file("xl/_rels/workbook.xml.rels", serializeRels(wbRels));

    const sheetsXml = outSheets
      .map(
        (s, i) =>
          `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
      )
      .join("");
    let newWb = wbXml
      .replace(/<sheets\b[^>]*>[\s\S]*?<\/sheets>/, `<sheets>${sheetsXml}</sheets>`)
      .replace(/<definedNames\b[^>]*\/>/g, "")
      .replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/g, "")
      .replace(/<calcPr\b[^>]*\/>/g, "")
      .replace(/<calcPr\b[^>]*>[\s\S]*?<\/calcPr>/g, "");
    out.file("xl/workbook.xml", newWb);

    let newTypes = typesXml.replace(/<Override\b[^>]*\/>/g, (tag) => {
      const part = (tag.match(/PartName="([^"]+)"/) || [])[1] || "";
      return /^\/xl\/worksheets\/sheet\d+\.xml$/i.test(part) ? "" : tag;
    });
    const extraDraw = Object.keys(out.files)
      .filter((p) => /xl\/drawings\/[^/]*_g\d+\.xml$/i.test(p))
      .map(
        (p) =>
          `<Override PartName="/${p.replace(/^\/+/, "")}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
      )
      .join("");
    const overrides =
      extraDraw +
      outSheets
        .map(
          (_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
        )
        .join("");
    if (newTypes.includes("<Override")) {
      newTypes = newTypes.replace("<Override", overrides + "<Override");
    } else {
      newTypes = newTypes.replace("</Types>", overrides + "</Types>");
    }
    out.file("[Content_Types].xml", newTypes);

    const appFile = zip.file("docProps/app.xml");
    if (appFile) {
      let appXml = await appFile.async("string");
      const n = outSheets.length;
      const titles = outSheets
        .map((s) => `<vt:lpstr>${xmlEscape(s.name)}</vt:lpstr>`)
        .join("");
      appXml = appXml.replace(
        /<vt:i4>\d+<\/vt:i4>/,
        `<vt:i4>${n}</vt:i4>`
      );
      appXml = appXml.replace(
        /<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/,
        `<TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts>`
      );
      out.file("docProps/app.xml", appXml);
    }

    return out.generateAsync({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      platform: "DOS",
    });
  }

  async function generateMadeJournal() {
    const subjects = collectSubjectRows();
    if (!subjects.length) {
      setMakeStatus("Додайте хоча б один предмет.", true);
      return;
    }
    const group = (document.getElementById("make-group") || {}).value || "";
    const meta = {
      group: String(group).trim(),
      course: String((document.getElementById("make-course") || {}).value || "").trim(),
      department: String((document.getElementById("make-dept") || {}).value || "").trim(),
      specialty: String((document.getElementById("make-spec") || {}).value || "").trim(),
      period: String((document.getElementById("make-period") || {}).value || "").trim(),
    };
    setMakeStatus("Збираю Excel…");
    try {
      const blob = await buildJournalXlsx(subjects, meta);
      const a = document.createElement("a");
      const name = meta.group
        ? `Журнал групи ${meta.group}.xlsx`
        : "Журнал групи.xlsx";
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      setMakeStatus(
        `Готово: ${subjects.length} предметів, зразок «${makeState.prototypeName}».`
      );
    } catch (err) {
      setMakeStatus(err && err.message ? err.message : String(err), true);
    }
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
    fitPrintCells(previewEl);
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

  if (fileInput) {
    fileInput.addEventListener("change", () => {
      const file = fileInput.files && fileInput.files[0];
      if (file) handleFile(file).catch((err) => {
        renderReport([{ ok: false, label: "Файл", reason: err.message }]);
      });
    });
  }

  if (dropzone) {
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
  }

  if (printBtn) printBtn.addEventListener("click", () => openPrintView());
  const detailChecks = document.getElementById("detail-checks");
  if (detailChecks) {
    detailChecks.addEventListener("change", () => renderReport(lastResults));
  }
  window.addEventListener("beforeprint", () => {
    if (previewEl) fitPrintCells(previewEl);
  });

  if (csvInput) {
    csvInput.addEventListener("change", () => {
      const file = csvInput.files && csvInput.files[0];
      if (file) handleFile(file).catch((err) => {
        renderReport([{ ok: false, label: "Файл", reason: err.message }]);
      });
    });
  }

  const makeInput = document.getElementById("make-input");
  if (makeInput) {
    makeInput.addEventListener("change", () => {
      const file = makeInput.files && makeInput.files[0];
      if (!file) return;
      handleMakeTemplate(file, { source: "custom" }).catch((err) =>
        setMakeStatus(err && err.message ? err.message : String(err), true)
      );
    });
  }
  const subjAdd = document.getElementById("subj-add");
  if (subjAdd) subjAdd.addEventListener("click", () => addSubjectRow("", ""));
  const subjPaste = document.getElementById("subj-paste");
  if (subjPaste) {
    subjPaste.addEventListener("input", schedulePasteSync);
    subjPaste.addEventListener("paste", () => {
      setTimeout(applyPasteList, 0);
    });
  }
  const makeBtn = document.getElementById("make-btn");
  if (makeBtn) makeBtn.addEventListener("click", () => generateMadeJournal());

  const tplThumb = document.getElementById("tpl-thumb");
  const tplZoom = document.getElementById("tpl-zoom");
  if (tplThumb && tplZoom) {
    tplThumb.addEventListener("click", (e) => {
      e.stopPropagation();
      tplZoom.hidden = false;
    });
    tplZoom.addEventListener("click", () => {
      tplZoom.hidden = true;
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !tplZoom.hidden) tplZoom.hidden = true;
    });
  }

  async function loadDefaultTemplate() {
    if (!document.getElementById("stage-make")) return;
    if (makeState.source === "standard" && makeState.buffer) return;
    try {
      const res = await fetch(DEFAULT_TEMPLATE);
      if (!res.ok) throw new Error("Не вдалося підвантажити стандартний шаблон.");
      const blob = await res.blob();
      const file = new File(
        [blob],
        "Стандартний шаблон.xlsx",
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
      );
      await handleMakeTemplate(file, { source: "standard" });
    } catch (err) {
      const picked = document.getElementById("make-picked");
      if (picked) picked.textContent = "Стандартний шаблон не підвантажився.";
      setMakeStatus(err && err.message ? err.message : String(err), true);
    }
  }
  const tplCard = document.getElementById("tpl-standard");
  if (tplCard) {
    tplCard.addEventListener("click", (e) => {
      if (e.target.closest("#tpl-thumb")) return;
      loadDefaultTemplate();
    });
    tplCard.style.cursor = "pointer";
  }
  loadDefaultTemplate();

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
