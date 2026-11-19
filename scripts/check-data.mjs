import { readFileSync } from "node:fs";
import * as XLSXNamespace from "xlsx";
import {
  appealsForMonth,
  categoryStats,
  defaultMonthKey,
  parseMatrix,
  parseWorkbook,
  unresolvedHousehold,
} from "../lib/appeals.mjs";

const XLSX = XLSXNamespace.default ?? XLSXNamespace;
const today = new Date(2026, 10, 19);
let failed = 0;

function assert(condition, message) {
  if (condition) return;
  failed += 1;
  console.error(`FAIL ${message}`);
}

const buffer = readFileSync(new URL("../public/data/appeals.xlsx", import.meta.url));
const appeals = parseWorkbook(XLSX, buffer);

assert(appeals.length === 28, `ожидалось 28 обращений, получено ${appeals.length}`);
assert(defaultMonthKey(appeals, today) === "2026-11", "текущий месяц должен быть ноябрём 2026");

const october = categoryStats(appealsForMonth(appeals, "2026-10", today));
const november = categoryStats(appealsForMonth(appeals, "2026-11", today));

assert(october.reduce((sum, item) => sum + item.count, 0) === 15, "в октябре должно быть 15 обращений");
assert(november.reduce((sum, item) => sum + item.count, 0) === 13, "в ноябре должно быть 13 обращений");

assert(
  october.map((item) => `${item.category}:${item.count}`).join("|") ===
    "🏠Заселение:5|🔧быт.вопрос:4|🔁продление/оплата:2|📄Документы:2|новая продажа/заявка:1|🌐 Wi-fi, связь:1",
  `неверные категории октября: ${JSON.stringify(october)}`,
);

assert(
  november.map((item) => `${item.category}:${item.count}`).join("|") ===
    "🏠Заселение:4|🔁продление/оплата:3|📄Документы:2|🌐 Wi-fi, связь:2|🔧быт.вопрос:2",
  `неверные категории ноября: ${JSON.stringify(november)}`,
);

const openIssues = unresolvedHousehold(appeals);
assert(openIssues.length === 3, `нерешённых бытовых вопросов должно быть 3, получено ${openIssues.length}`);
const localIso = (date) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
assert(
  openIssues.map((item) => localIso(item.date)).join(",") === "2026-10-01,2026-10-14,2026-11-11",
  `неверный порядок нерешённых: ${openIssues.map((item) => localIso(item.date)).join(", ")}`,
);
assert(openIssues[0].apartment === "БЗах 142-121", "первая квартира");
assert(openIssues[2].status === "отложен", "статус третьего вопроса");
assert(openIssues[2].apartment === "Триумф", "пробел в названии квартиры должен срезаться");

const futureMatrix = [
  ["Дата поступления вопроса", "квартира/дом", "Коротко суть вопроса", "Категория", "Ответственный кому передали вопрос", "Статус"],
  ["Октябрь"],
  ["01.11.2026", "Дом", "Уже было", "🏠Заселение", "Отдел продаж ", "Решен"],
  ["20.11.2026", "Дом", "Ещё не наступило", "🏠Заселение", "БА", ""],
];
const futureAppeals = parseMatrix(futureMatrix);
const futureSlice = appealsForMonth(futureAppeals, "2026-11", today);
assert(futureSlice.length === 1, "будущая дата текущего месяца не входит в срез");
assert(futureSlice[0].status === "Решен", "статус сохраняется как в файле");
assert(unresolvedHousehold(futureAppeals).length === 0, "решённое заселение не попадает в быт");

const semicolon = [
  "Дата поступления вопроса;квартира/дом;Коротко суть вопроса;Категория;Ответственный кому передали вопрос;Статус",
  "01.10.2026;Лагуна;Нет света;🔧быт.вопрос;Роз;",
  "02.10.2026;Жук;Ключи;🔧быт.вопрос;Самостоятельно;решен",
  "03.10.2026;Дом;Закрыто;🔧быт.вопрос;Роз;Решен",
].join("\n");
const fromCsv = parseWorkbook(XLSX, new TextEncoder().encode(semicolon));
assert(fromCsv.length === 3, "csv с точкой с запятой должен читаться");
assert(
  fromCsv[0].date.getFullYear() === 2026 && fromCsv[0].date.getMonth() === 9 && fromCsv[0].date.getDate() === 1,
  `дата 01.10.2026 прочитана как ${fromCsv[0].date.toDateString()}`,
);
assert(unresolvedHousehold(fromCsv).length === 1, "пустой статус нерешён, а «Решен» считается закрытым");

try {
  parseWorkbook(XLSX, new TextEncoder().encode("просто текст"));
  assert(false, "битый файл должен выбрасывать ошибку");
} catch (error) {
  assert(error.message.includes("заголовков") || error.message.includes("таблицу"), error.message);
}

if (failed) {
  console.error(`Проверок не пройдено: ${failed}`);
  process.exit(1);
}

console.log("Данные сходятся: октябрь 15, ноябрь 13, нерешённый быт 3");
