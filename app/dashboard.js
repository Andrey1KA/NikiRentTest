"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { BarController, BarElement, CategoryScale, Chart, LinearScale, Tooltip } from "chart.js";
import * as XLSX from "xlsx";
import {
  appealsForMonth,
  categoryStats,
  defaultMonthKey,
  isCurrentMonth,
  listMonths,
  parseWorkbook,
  unresolvedHousehold,
} from "../lib/appeals.mjs";

Chart.register(BarController, BarElement, CategoryScale, LinearScale, Tooltip);
Chart.defaults.font.family = '"Segoe UI", "Segoe UI Emoji", sans-serif';

const PALETTE = ["#0e6b62", "#c45c26", "#1f4e79", "#8a6a2f", "#5c6b73", "#7a3e5c", "#3d6b4f", "#a33b3b"];

const MONTHS = [
  ["Январь", "января"],
  ["Февраль", "февраля"],
  ["Март", "марта"],
  ["Апрель", "апреля"],
  ["Май", "мая"],
  ["Июнь", "июня"],
  ["Июль", "июля"],
  ["Август", "августа"],
  ["Сентябрь", "сентября"],
  ["Октябрь", "октября"],
  ["Ноябрь", "ноября"],
  ["Декабрь", "декабря"],
];

export default function Dashboard() {
  const [appeals, setAppeals] = useState([]);
  const [month, setMonth] = useState("");
  const [today, setToday] = useState(null);
  const [ready, setReady] = useState(false);
  const [loadError, setLoadError] = useState("");
  const chartRef = useRef(null);
  const chartInstance = useRef(null);

  useEffect(() => {
    const current = new Date();
    setToday(current);
    let cancelled = false;

    async function loadDefault() {
      try {
        const response = await fetch("/data/appeals.xlsx");
        if (!response.ok) throw new Error("Не удалось открыть встроенную таблицу");
        const buffer = await response.arrayBuffer();
        const nextAppeals = parseWorkbook(XLSX, buffer);
        if (cancelled) return;
        setAppeals(nextAppeals);
        setMonth(defaultMonthKey(nextAppeals, current));
      } catch (error) {
        if (cancelled) return;
        setLoadError(error.message || "Не удалось открыть таблицу обращений");
      } finally {
        if (!cancelled) setReady(true);
      }
    }

    loadDefault();
    return () => {
      cancelled = true;
    };
  }, []);

  const slice = useMemo(
    () => (today ? appealsForMonth(appeals, month, today) : []),
    [appeals, month, today],
  );
  const stats = useMemo(() => categoryStats(slice), [slice]);
  const openIssues = useMemo(() => unresolvedHousehold(appeals), [appeals]);
  const months = useMemo(() => listMonths(appeals), [appeals]);

  useEffect(() => {
    return () => {
      chartInstance.current?.destroy();
      chartInstance.current = null;
    };
  }, []);

  useEffect(() => {
    const canvas = chartRef.current;
    if (!canvas || stats.length === 0) {
      chartInstance.current?.destroy();
      chartInstance.current = null;
      return;
    }

    const labels = stats.map((item) => item.category);
    const values = stats.map((item) => item.count);
    const colors = stats.map((_, index) => PALETTE[index % PALETTE.length]);

    if (chartInstance.current) {
      const chart = chartInstance.current;
      chart.data.labels = labels;
      chart.data.datasets[0].data = values;
      chart.data.datasets[0].backgroundColor = colors;
      chart.update("none");
      return;
    }

    chartInstance.current = new Chart(canvas, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            data: values,
            backgroundColor: colors,
            borderRadius: 6,
            maxBarThickness: 22,
          },
        ],
      },
      options: {
        animation: false,
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(item) {
                return ` ${item.parsed.x}`;
              },
            },
          },
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: { precision: 0, color: "#6d645b" },
            grid: { color: "#efe7dc" },
            border: { display: false },
          },
          y: {
            ticks: { color: "#1e1a16", font: { size: 12 } },
            grid: { display: false },
            border: { display: false },
          },
        },
      },
    });
  }, [stats]);

  const total = slice.length;

  return (
    <div className="wrap">
      <header className="top">
        <div>
          <p className="eyebrow">NikiRent</p>
          <h1>Обращения гостей</h1>
          <p className="lede">{today ? `Срез на ${today.getDate()} ${MONTHS[today.getMonth()][1]} ${today.getFullYear()}` : "Загрузка таблицы…"}</p>
          {loadError ? <p className="message error">{loadError}</p> : null}
        </div>
      </header>

      <section className="panel" aria-labelledby="stats-title">
        <div className="panel-head">
          <div>
            <h2 id="stats-title">Статистика</h2>
            <p className="hint">{sliceCaption(month, today)}</p>
          </div>
          <label className="month-label">
            Месяц
            <select
              id="month"
              value={month}
              onChange={(event) => setMonth(event.target.value)}
              disabled={!months.length}
            >
              {months.length === 0 ? <option value="">Нет данных</option> : null}
              {months.map((key) => (
                <option key={key} value={key}>
                  {monthTitle(key)}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="stat-layout">
          <div className="total-card">
            <p className="total">{ready ? total : "—"}</p>
            <p className="hint">{appealsWord(total)}</p>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Категория</th>
                  <th className="num">Количество</th>
                  <th className="num">Доля</th>
                </tr>
              </thead>
              <tbody>
                {stats.map((item, index) => (
                  <tr key={item.category}>
                    <td>
                      <span className="swatch" style={{ background: PALETTE[index % PALETTE.length] }} />
                      {item.category}
                    </td>
                    <td className="num">{item.count}</td>
                    <td className="num">{formatPercent(item.count, total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {ready && stats.length === 0 ? <p className="empty">Нет обращений за этот период.</p> : null}
          </div>
          <div className="chart-wrap">
            <canvas ref={chartRef} hidden={stats.length === 0} aria-label="Распределение обращений по категориям" />
            {ready && stats.length === 0 ? <p className="empty">Нет обращений за этот период.</p> : null}
          </div>
        </div>
      </section>

      <section className="panel" aria-labelledby="open-title">
        <div className="panel-head">
          <div>
            <h2 id="open-title">Нерешённые бытовые вопросы</h2>
            <p className="hint">Категория «🔧быт.вопрос», статус не «решен». Все месяцы, сначала самые старые.</p>
          </div>
          <p className="count-pill">
            {openIssues.length} {questionsWord(openIssues.length)}
          </p>
        </div>
        <div className="table-wrap">
          <table className="issues">
            <colgroup>
              <col className="col-date" />
              <col className="col-apartment" />
              <col />
              <col className="col-owner" />
              <col className="col-status" />
            </colgroup>
            <thead>
              <tr>
                <th>Дата поступления</th>
                <th>Квартира</th>
                <th>Суть проблемы</th>
                <th>Ответственный</th>
                <th>Статус</th>
              </tr>
            </thead>
            <tbody>
              {openIssues.map((issue) => (
                <tr key={`${issue.index}-${issue.summary}`}>
                  <td>{formatDay(issue.date)}</td>
                  <td>{issue.apartment || "—"}</td>
                  <td>{issue.summary || "—"}</td>
                  <td>{issue.owner || "—"}</td>
                  <td>
                    <span className={issue.status ? (issue.status.toLowerCase() === "отложен" ? "pill warn" : "pill") : "pill muted"}>
                      {issue.status || "не указан"}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {ready && openIssues.length === 0 ? <p className="empty">Нерешённых бытовых вопросов нет.</p> : null}
        </div>
      </section>
    </div>
  );
}

function sliceCaption(month, today) {
  if (!month || !today) return "Период появится после загрузки файла";
  const [yearText, monthText] = month.split("-");
  const year = Number(yearText);
  const monthIndex = Number(monthText) - 1;
  if (isCurrentMonth(month, today)) {
    return `1–${today.getDate()} ${MONTHS[monthIndex][1]} ${year}`;
  }
  return "весь месяц";
}

function monthTitle(key) {
  const [yearText, monthText] = key.split("-");
  return `${MONTHS[Number(monthText) - 1][0]} ${yearText}`;
}

function formatDay(date) {
  return new Intl.DateTimeFormat("ru-RU").format(date);
}

function formatPercent(count, total) {
  if (!total) return "0,0 %";
  return `${new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  }).format((count / total) * 100)} %`;
}

function appealsWord(count) {
  return plural(count, "обращение", "обращения", "обращений");
}

function questionsWord(count) {
  return plural(count, "вопрос", "вопроса", "вопросов");
}

function plural(count, one, few, many) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 === 1 && mod100 !== 11) return one;
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return few;
  return many;
}
