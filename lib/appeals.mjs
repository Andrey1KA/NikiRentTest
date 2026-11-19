export const HOUSEHOLD_CATEGORY = "🔧быт.вопрос";

const HEADER_RULES = [
  ["date", ["дата поступления"]],
  ["apartment", ["квартира"]],
  ["summary", ["суть вопроса", "суть проблемы", "коротко суть"]],
  ["category", ["категория"]],
  ["owner", ["ответствен"]],
  ["status", ["статус"]],
];

export function parseWorkbook(XLSX, buffer) {
  let matrix;
  try {
    matrix = readMatrix(XLSX, buffer);
  } catch {
    throw new Error("Файл не похож на таблицу Excel или CSV");
  }
  return parseMatrix(matrix);
}

export function parseMatrix(matrix) {
  if (!Array.isArray(matrix) || matrix.length === 0) {
    throw new Error("В файле нет строк");
  }

  const headerIndex = findHeaderIndex(matrix);
  if (headerIndex < 0) {
    throw new Error("В файле не найдена строка заголовков с датой и категорией");
  }

  const columns = mapColumns(matrix[headerIndex]);
  if (columns.date == null || columns.category == null) {
    throw new Error("В файле не найдена строка заголовков с датой и категорией");
  }

  const appeals = [];
  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const date = parseDate(row[columns.date]);
    if (!date) continue;

    appeals.push({
      date,
      apartment: cellText(row[columns.apartment]),
      summary: cellText(row[columns.summary]),
      category: cellText(row[columns.category]) || "Без категории",
      owner: cellText(row[columns.owner]),
      status: cellText(row[columns.status]),
      index: appeals.length,
    });
  }

  return appeals;
}

export function monthKey(date) {
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${date.getFullYear()}-${month}`;
}

export function listMonths(appeals) {
  return [...new Set(appeals.map((appeal) => monthKey(appeal.date)))].sort().reverse();
}

export function defaultMonthKey(appeals, today) {
  const months = listMonths(appeals);
  const current = monthKey(today);
  if (months.includes(current)) return current;
  return months[0] || "";
}

export function appealsForMonth(appeals, key, today) {
  if (!key) return [];
  return appeals.filter((appeal) => isInSlice(appeal.date, key, today));
}

export function categoryStats(appeals) {
  const counts = new Map();
  const order = [];

  for (const appeal of appeals) {
    if (!counts.has(appeal.category)) {
      counts.set(appeal.category, 0);
      order.push(appeal.category);
    }
    counts.set(appeal.category, counts.get(appeal.category) + 1);
  }

  return order
    .map((category, index) => ({ category, count: counts.get(category), index }))
    .sort((left, right) => right.count - left.count || left.index - right.index)
    .map(({ category, count }) => ({ category, count }));
}

export function unresolvedHousehold(appeals) {
  return appeals
    .filter(
      (appeal) =>
        appeal.category === HOUSEHOLD_CATEGORY &&
        appeal.status.toLowerCase() !== "решен",
    )
    .slice()
    .sort((left, right) => left.date - right.date || left.index - right.index);
}

export function isCurrentMonth(key, today) {
  return key === monthKey(today);
}

function isInSlice(date, key, today) {
  const [yearText, monthText] = key.split("-");
  const year = Number(yearText);
  const monthIndex = Number(monthText) - 1;
  const start = new Date(year, monthIndex, 1);
  const end = isCurrentMonth(key, today)
    ? new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999)
    : new Date(year, monthIndex + 1, 0, 23, 59, 59, 999);
  return date >= start && date <= end;
}

function readMatrix(XLSX, buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const isZip = bytes.length > 1 && bytes[0] === 0x50 && bytes[1] === 0x4b;

  if (!isZip && looksLikeText(bytes)) {
    let text = decodeCsv(bytes);
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const firstLine = text.split(/\r?\n/, 1)[0] || "";
    const commas = (firstLine.match(/,/g) || []).length;
    const semicolons = (firstLine.match(/;/g) || []).length;
    const workbook = XLSX.read(text, {
      type: "string",
      FS: semicolons > commas ? ";" : ",",
      raw: true,
    });
    return firstSheet(XLSX, workbook);
  }

  const workbook = XLSX.read(bytes, { type: "array" });
  return firstSheet(XLSX, workbook);
}

function firstSheet(XLSX, workbook) {
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("В книге нет листов");
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    raw: true,
    defval: "",
    blankrows: true,
  });
}

function looksLikeText(bytes) {
  const sample = bytes.subarray(0, Math.min(bytes.length, 240));
  for (const byte of sample) {
    if (byte === 0) return false;
  }
  return sample.length > 0;
}

function decodeCsv(bytes) {
  const utf8 = new TextDecoder("utf-8").decode(bytes);
  if (/категор/i.test(utf8) || /дата поступления/i.test(utf8)) return utf8;
  return new TextDecoder("windows-1251").decode(bytes);
}

function findHeaderIndex(matrix) {
  const limit = Math.min(matrix.length, 15);
  for (let index = 0; index < limit; index += 1) {
    const cells = (matrix[index] || []).map(normalize);
    const hasDate = cells.some((cell) => cell.includes("дата поступления"));
    const hasCategory = cells.some((cell) => cell.includes("категория"));
    if (hasDate && hasCategory) return index;
  }
  return -1;
}

function mapColumns(headerRow) {
  const columns = {};
  headerRow.forEach((cell, index) => {
    const text = normalize(cell);
    if (!text) return;
    for (const [field, needles] of HEADER_RULES) {
      if (columns[field] != null) continue;
      if (needles.some((needle) => text.includes(needle))) columns[field] = index;
    }
  });
  return columns;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToDate(value);
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (/^\d+(\.\d+)?$/.test(trimmed)) return excelSerialToDate(Number(trimmed));
    const dayFirst = trimmed.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (dayFirst) return dateFromParts(Number(dayFirst[3]), Number(dayFirst[2]), Number(dayFirst[1]));
    const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return dateFromParts(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }
  return null;
}

function excelSerialToDate(serial) {
  const whole = Math.floor(serial);
  if (whole < 1 || whole > 120000) return null;
  const utc = new Date(Date.UTC(1899, 11, 30));
  utc.setUTCDate(utc.getUTCDate() + whole);
  return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
}

function dateFromParts(year, month, day) {
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}

function cellText(value) {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function normalize(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}
