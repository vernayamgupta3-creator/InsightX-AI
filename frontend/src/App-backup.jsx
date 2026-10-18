import { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";

const COLORS = ["#6366f1", "#14b8a6", "#f59e0b", "#ec4899", "#8b5cf6", "#0ea5e9"];
const numberFormat = (value) =>
  Number.isFinite(Number(value))
    ? new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 }).format(Number(value))
    : String(value ?? "—");

const isNumeric = (value) =>
  value !== null && value !== undefined && String(value).trim() !== "" &&
  Number.isFinite(Number(String(value).replace(/,/g, "")));

const toNumber = (value) => Number(String(value ?? "").replace(/,/g, "").trim());

function normalizeRows(rows) {
  return rows
    .filter((row) => row && typeof row === "object")
    .map((row) => {
      const result = {};
      Object.entries(row).forEach(([key, value]) => {
        const cleanKey = String(key ?? "").trim() || "Column";
        result[cleanKey] = value === undefined ? "" : value;
      });
      return result;
    });
}

function makeSample() {
  return [
    { Product: "Laptop", Sales: 50000, Profit: 8000, Region: "South" },
    { Product: "Phone", Sales: 30000, Profit: 5000, Region: "North" },
    { Product: "Tablet", Sales: 20000, Profit: 3000, Region: "East" },
    { Product: "Monitor", Sales: 15000, Profit: 2500, Region: "West" },
    { Product: "Keyboard", Sales: 5000, Profit: 1000, Region: "South" },
    { Product: "Mouse", Sales: 4000, Profit: 800, Region: "East" },
    { Product: "Headphones", Sales: 9000, Profit: 1600, Region: "North" },
  ];
}

function downloadBlob(filename, content, type = "text/plain;charset=utf-8") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function StatCard({ label, value, detail, icon }) {
  return (
    <div className="ix-stat">
      <div className="ix-stat-top"><span>{label}</span><span className="ix-stat-icon">{icon}</span></div>
      <div className="ix-stat-value">{value}</div>
      <div className="ix-stat-detail">{detail}</div>
    </div>
  );
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [activePage, setActivePage] = useState("Overview");
  const [notice, setNotice] = useState("");
  const [xColumn, setXColumn] = useState("");
  const [yColumn, setYColumn] = useState("");
  const [chartType, setChartType] = useState("bar");
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [search, setSearch] = useState("");
  const [showAllRows, setShowAllRows] = useState(false);
  const [cleanedRows, setCleanedRows] = useState(null);
  const fileInput = useRef(null);

  const currentRows = cleanedRows ?? rows;
  const columns = useMemo(() => currentRows.length ? Object.keys(currentRows[0]) : [], [currentRows]);
  const numericColumns = useMemo(
    () => columns.filter((column) => currentRows.some((row) => isNumeric(row[column]) && String(row[column]).trim() !== "")),
    [columns, currentRows]
  );
  const categoricalColumns = useMemo(
    () => columns.filter((column) => !numericColumns.includes(column)),
    [columns, numericColumns]
  );

  const analysis = useMemo(() => {
    const missingByColumn = {};
    columns.forEach((column) => {
      missingByColumn[column] = currentRows.filter((row) =>
        row[column] === null || row[column] === undefined || String(row[column]).trim() === ""
      ).length;
    });
    const missingCells = Object.values(missingByColumn).reduce((sum, value) => sum + value, 0);
    const seen = new Set();
    let duplicateRows = 0;
    currentRows.forEach((row) => {
      const key = JSON.stringify(columns.map((column) => row[column] ?? ""));
      if (seen.has(key)) duplicateRows += 1;
      else seen.add(key);
    });
    const numericStats = {};
    numericColumns.forEach((column) => {
      const values = currentRows.map((row) => row[column]).filter((value) => isNumeric(value) && String(value).trim() !== "").map(toNumber);
      const total = values.reduce((sum, value) => sum + value, 0);
      numericStats[column] = {
        count: values.length,
        total,
        average: values.length ? total / values.length : 0,
        min: values.length ? Math.min(...values) : 0,
        max: values.length ? Math.max(...values) : 0,
      };
    });
    const categoryStats = {};
    categoricalColumns.forEach((column) => {
      const counts = {};
      currentRows.forEach((row) => {
        const value = String(row[column] ?? "").trim() || "(Blank)";
        counts[value] = (counts[value] || 0) + 1;
      });
      categoryStats[column] = Object.entries(counts)
        .map(([name, count]) => ({ name, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 8);
    });
    return { missingByColumn, missingCells, duplicateRows, numericStats, categoryStats };
  }, [currentRows, columns, numericColumns, categoricalColumns]);

  const selectedX = xColumn || categoricalColumns[0] || columns[0] || "";
  const selectedY = yColumn || numericColumns[0] || "";
  const chartData = useMemo(() => {
    if (!selectedX || !selectedY) return [];
    const groups = {};
    currentRows.forEach((row) => {
      const label = String(row[selectedX] ?? "").trim() || "(Blank)";
      if (isNumeric(row[selectedY]) && String(row[selectedY]).trim() !== "") {
        groups[label] = (groups[label] || 0) + toNumber(row[selectedY]);
      }
    });
    return Object.entries(groups)
      .map(([name, value]) => ({ name: name.length > 24 ? `${name.slice(0, 21)}…` : name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 20);
  }, [currentRows, selectedX, selectedY]);

  const filteredRows = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return currentRows;
    return currentRows.filter((row) => columns.some((column) => String(row[column] ?? "").toLowerCase().includes(q)));
  }, [currentRows, columns, search]);

  const insights = useMemo(() => {
    if (!currentRows.length) return [];
    const items = [];
    items.push({ title: "Dataset loaded", text: `Your dataset contains ${numberFormat(currentRows.length)} rows and ${numberFormat(columns.length)} columns.` });
    if (numericColumns.length) {
      const best = numericColumns
        .map((column) => ({ column, ...analysis.numericStats[column] }))
        .sort((a, b) => b.total - a.total)[0];
      items.push({ title: "Largest numeric total", text: `${best.column} has the largest total (${numberFormat(best.total)}).` });
    } else {
      items.push({ title: "Numeric data not detected", text: "No numeric columns were detected. Check whether number values contain extra text or symbols." });
    }
    if (analysis.missingCells) {
      const mostMissing = Object.entries(analysis.missingByColumn).sort((a, b) => b[1] - a[1])[0];
      items.push({ title: "Missing values found", text: `${numberFormat(analysis.missingCells)} cells are blank. "${mostMissing[0]}" has the most blanks (${mostMissing[1]}).` });
    } else {
      items.push({ title: "No missing cells", text: "Every cell in the current dataset has a value." });
    }
    if (analysis.duplicateRows) {
      items.push({ title: "Possible duplicate rows", text: `${numberFormat(analysis.duplicateRows)} duplicate row(s) were detected. Review them before removing anything.` });
    } else {
      items.push({ title: "No duplicate rows", text: "No fully identical duplicate rows were detected." });
    }
    if (categoricalColumns.length) {
      const column = categoricalColumns[0];
      const top = analysis.categoryStats[column]?.[0];
      if (top) items.push({ title: `Most common ${column}`, text: `"${top.name}" appears ${top.count} time(s) in this column.` });
    }
    return items;
  }, [currentRows, columns.length, numericColumns, categoricalColumns, analysis]);

  function loadRows(data, name = "Sample dataset") {
    const normalized = normalizeRows(data);
    if (!normalized.length) {
      setNotice("No data rows found. Please upload a CSV or Excel file with a header row.");
      return;
    }
    const allKeys = [...new Set(normalized.flatMap((row) => Object.keys(row)))];
    const fixed = normalized.map((row) => Object.fromEntries(allKeys.map((key) => [key, row[key] ?? ""])));
    setRows(fixed);
    setCleanedRows(null);
    setFileName(name);
    setXColumn("");
    setYColumn("");
    setNotice(`Loaded ${fixed.length} rows successfully.`);
    setActivePage("Overview");
  }

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const extension = file.name.split(".").pop().toLowerCase();
      if (extension === "csv" || extension === "txt") {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: "greedy",
          dynamicTyping: false,
          complete: (result) => {
            if (result.errors?.length && !result.data?.length) {
              setNotice("Could not read this CSV. Please check the file format.");
              return;
            }
            loadRows(result.data, file.name);
          },
          error: () => setNotice("Could not read this file. Please try another CSV."),
        });
      } else if (["xlsx", "xls"].includes(extension)) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        loadRows(XLSX.utils.sheet_to_json(firstSheet, { defval: "" }), file.name);
      } else {
        setNotice("Unsupported file type. Upload a CSV, XLSX, or XLS file.");
      }
    } catch {
      setNotice("There was a problem reading the file. Please try a valid CSV or Excel file.");
    }
    event.target.value = "";
  }

  function cleanData() {
    const withoutDuplicates = [];
    const seen = new Set();
    currentRows.forEach((row) => {
      const key = JSON.stringify(columns.map((column) => row[column] ?? ""));
      if (!seen.has(key)) {
        seen.add(key);
        withoutDuplicates.push({ ...row });
      }
    });
    setCleanedRows(withoutDuplicates);
    setNotice(`Created a cleaned working copy with ${withoutDuplicates.length} rows. Your original upload remains unchanged.`);
  }

  function exportCsv() {
    if (!currentRows.length) return;
    downloadBlob("insightx-cleaned-data.csv", Papa.unparse(currentRows), "text/csv;charset=utf-8");
  }

  function exportReport() {
    const report = {
      application: "InsightX AI",
      generatedAt: new Date().toISOString(),
      sourceFile: fileName || "Sample dataset",
      rows: currentRows.length,
      columns,
      numericSummary: analysis.numericStats,
      missingValuesByColumn: analysis.missingByColumn,
      totalMissingCells: analysis.missingCells,
      duplicateRows: analysis.duplicateRows,
      automaticInsights: insights,
    };
    downloadBlob("insightx-analysis-report.json", JSON.stringify(report, null, 2), "application/json");
  }

  function askQuestion(event) {
    event.preventDefault();
    const q = question.trim().toLowerCase();
    if (!currentRows.length) {
      setAnswer("Upload a dataset first, or load the sample dataset.");
      return;
    }
    if (!q) {
      setAnswer("Type a question about your dataset first.");
      return;
    }
    const matchingNumeric = numericColumns.find((column) => q.includes(column.toLowerCase()));
    const matchingCategory = columns.find((column) => q.includes(column.toLowerCase()));
    if (q.includes("missing") || q.includes("blank") || q.includes("null")) {
      setAnswer(`There are ${analysis.missingCells} missing cell(s) in total. ${Object.entries(analysis.missingByColumn).map(([key, value]) => `${key}: ${value}`).join("; ")}.`);
    } else if (q.includes("duplicate")) {
      setAnswer(`I found ${analysis.duplicateRows} fully duplicate row(s). Use Data Cleaning to create a copy with duplicate rows removed.`);
    } else if (q.includes("how many rows") || q.includes("row count") || q.includes("number of rows")) {
      setAnswer(`Your dataset has ${currentRows.length} rows and ${columns.length} columns.`);
    } else if (q.includes("how many columns") || q.includes("column count")) {
      setAnswer(`Your dataset has ${columns.length} columns: ${columns.join(", ")}.`);
    } else if (q.includes("average") || q.includes("mean")) {
      const column = matchingNumeric || numericColumns[0];
      setAnswer(column ? `The average of ${column} is ${numberFormat(analysis.numericStats[column].average)}.` : "I couldn't find a numeric column to calculate an average.");
    } else if (q.includes("total") || q.includes("sum")) {
      const column = matchingNumeric || numericColumns[0];
      setAnswer(column ? `The total of ${column} is ${numberFormat(analysis.numericStats[column].total)}.` : "I couldn't find a numeric column to calculate a total.");
    } else if (q.includes("highest") || q.includes("maximum") || q.includes("max")) {
      const column = matchingNumeric || numericColumns[0];
      if (!column) setAnswer("I couldn't find a numeric column.");
      else {
        const top = [...currentRows].filter((row) => isNumeric(row[column])).sort((a, b) => toNumber(b[column]) - toNumber(a[column]))[0];
        setAnswer(top ? `The highest ${column} is ${numberFormat(top[column])}${matchingCategory ? ` (row value for ${matchingCategory}: ${top[matchingCategory]})` : ""}.` : `No numeric values were found in ${column}.`);
      }
    } else if (q.includes("lowest") || q.includes("minimum") || q.includes("min")) {
      const column = matchingNumeric || numericColumns[0];
      if (!column) setAnswer("I couldn't find a numeric column.");
      else {
        const bottom = [...currentRows].filter((row) => isNumeric(row[column])).sort((a, b) => toNumber(a[column]) - toNumber(b[column]))[0];
        setAnswer(bottom ? `The lowest ${column} is ${numberFormat(bottom[column])}.` : `No numeric values were found in ${column}.`);
      }
    } else if (q.includes("summary") || q.includes("describe") || q.includes("insight")) {
      setAnswer(insights.map((item) => `${item.title}: ${item.text}`).join("\n"));
    } else if (matchingNumeric) {
      const stat = analysis.numericStats[matchingNumeric];
      setAnswer(`${matchingNumeric}: total ${numberFormat(stat.total)}, average ${numberFormat(stat.average)}, minimum ${numberFormat(stat.min)}, maximum ${numberFormat(stat.max)} across ${stat.count} numeric values.`);
    } else {
      setAnswer("I can answer questions about row count, columns, missing values, duplicates, totals, averages, highest/lowest values, and summaries. Try: “What is the total Sales?” or “Are there missing values?”");
    }
  }

  const pages = [
    { name: "Overview", icon: "▦" },
    { name: "Data Workspace", icon: "▤" },
    { name: "AI Analyst", icon: "✳" },
    { name: "Visualization Studio", icon: "◒" },
    { name: "Data Cleaning", icon: "⌁" },
    { name: "Reports", icon: "⇩" },
  ];

  return (
    <div className="ix-app">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
        *{box-sizing:border-box}body{margin:0;background:#f6f7fb;font-family:'DM Sans',sans-serif;color:#182033}
        button,input,select{font:inherit}button{cursor:pointer}
        .ix-app{min-height:100vh;background:#f6f7fb}.ix-layout{display:flex;min-height:100vh}
        .ix-sidebar{width:245px;flex-shrink:0;background:#111827;color:#d1d5e3;padding:24px 16px;display:flex;flex-direction:column}
        .ix-brand{display:flex;align-items:center;gap:11px;padding:0 10px 30px;color:#fff;font-family:'Space Grotesk',sans-serif;font-size:20px;font-weight:700}
        .ix-logo{width:38px;height:38px;border-radius:12px;background:linear-gradient(135deg,#818cf8,#4f46e5);display:grid;place-items:center;color:white;font-size:22px}
        .ix-brand small{display:block;color:#8e98ad;font:11px 'DM Sans',sans-serif;margin-top:3px;letter-spacing:.5px}
        .ix-nav-label{font-size:10px;font-weight:700;letter-spacing:1.5px;color:#6f7b91;padding:0 12px 10px}
        .ix-nav{display:grid;gap:5px}.ix-nav button{border:0;background:transparent;color:#aeb7c8;text-align:left;padding:12px 13px;border-radius:10px;display:flex;gap:12px;align-items:center;font-size:13px}
        .ix-nav button:hover{background:#1e293b;color:#fff}.ix-nav button.active{background:#3730a3;color:#fff}
        .ix-nav-icon{width:20px;text-align:center;font-size:17px}.ix-side-bottom{margin-top:auto;padding:18px 10px 0;font-size:11px;line-height:1.7;color:#748096}
        .ix-side-pill{display:inline-block;border:1px solid #334155;border-radius:20px;padding:4px 9px;color:#a5b4fc;margin-bottom:10px}
        .ix-main{min-width:0;flex:1}.ix-topbar{height:76px;background:#fff;border-bottom:1px solid #e8eaf1;display:flex;align-items:center;justify-content:space-between;padding:0 34px}
        .ix-breadcrumb{font-size:12px;color:#8a92a5}.ix-breadcrumb strong{color:#293149}.ix-top-right{display:flex;align-items:center;gap:12px;font-size:12px;color:#667085}
        .ix-status{display:flex;align-items:center;gap:7px;background:#ecfdf5;color:#047857;border-radius:30px;padding:7px 11px;font-weight:700;font-size:11px}.ix-dot{width:7px;height:7px;border-radius:50%;background:#10b981}
        .ix-avatar{width:34px;height:34px;border-radius:50%;background:#e0e7ff;color:#4338ca;display:grid;place-items:center;font-weight:700}
        .ix-content{padding:30px 34px 50px;max-width:1500px;margin:auto}.ix-heading-row{display:flex;align-items:flex-start;justify-content:space-between;gap:20px;margin-bottom:24px}
        .ix-eyebrow{color:#6366f1;font-size:11px;font-weight:700;letter-spacing:1.6px;text-transform:uppercase;margin-bottom:8px}
        h1,h2,h3,p{margin-top:0}.ix-title{font:700 29px 'Space Grotesk',sans-serif;letter-spacing:-.8px;margin:0 0 7px;color:#151c31}
        .ix-subtitle{font-size:13px;color:#7c8498;margin:0;line-height:1.6}.ix-actions{display:flex;gap:9px;flex-wrap:wrap}
        .ix-btn{border:1px solid #e0e4ee;background:#fff;color:#344054;padding:10px 14px;border-radius:9px;font-size:12px;font-weight:700;display:inline-flex;align-items:center;gap:7px}
        .ix-btn:hover{border-color:#a5b4fc;color:#4338ca}.ix-btn-primary{background:#4f46e5;color:white;border-color:#4f46e5}.ix-btn-primary:hover{background:#4338ca;color:white}
        .ix-banner{background:linear-gradient(115deg,#4338ca,#6366f1 65%,#818cf8);border-radius:16px;color:white;padding:22px 25px;display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:22px;box-shadow:0 10px 25px #4f46e51c}
        .ix-banner h2{font:600 19px 'Space Grotesk',sans-serif;margin:0 0 6px}.ix-banner p{font-size:12px;color:#e0e7ff;margin:0;line-height:1.6}
        .ix-banner .ix-btn{background:#fff;color:#4338ca;border-color:white;white-space:nowrap}
        .ix-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px;margin-bottom:20px}
        .ix-stat,.ix-card{background:#fff;border:1px solid #e8eaf2;border-radius:14px;box-shadow:0 3px 12px #10182805}
        .ix-stat{padding:18px 19px}.ix-stat-top{display:flex;justify-content:space-between;align-items:center;color:#7d8599;font-size:11px;font-weight:600}
        .ix-stat-icon{display:grid;place-items:center;width:32px;height:32px;border-radius:10px;background:#eef2ff;color:#4f46e5;font-size:16px}
        .ix-stat-value{font:700 27px 'Space Grotesk',sans-serif;margin:13px 0 5px;letter-spacing:-.7px}.ix-stat-detail{font-size:10px;color:#9aa1b2}
        .ix-grid-2{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(280px,1fr);gap:18px;margin-bottom:19px}
        .ix-card{padding:21px;min-width:0}.ix-card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:18px}
        .ix-card-title{font:600 15px 'Space Grotesk',sans-serif;margin:0 0 5px;color:#202840}.ix-card-sub{font-size:11px;color:#9299aa;margin:0}
        .ix-tag{background:#f1f5f9;color:#64748b;border-radius:6px;padding:5px 8px;font-size:10px;font-weight:700;white-space:nowrap}
        .ix-insight-list{display:grid;gap:10px}.ix-insight{display:flex;gap:11px;padding:12px;background:#fafbff;border:1px solid #eef0f7;border-radius:10px}
        .ix-insight-mark{width:30px;height:30px;border-radius:9px;background:#eef2ff;color:#4f46e5;display:grid;place-items:center;flex-shrink:0}
        .ix-insight h3{font-size:12px;margin:0 0 4px}.ix-insight p{font-size:11px;line-height:1.5;color:#737c91;margin:0}
        .ix-empty{padding:32px 20px;text-align:center;border:1px dashed #d9deea;background:#fbfcff;border-radius:12px;color:#8a92a5;font-size:12px;line-height:1.7}
        .ix-empty-icon{font-size:28px;margin-bottom:8px}.ix-upload{border:1px dashed #a5b4fc;background:#f8faff;border-radius:12px;padding:22px;text-align:center}
        .ix-upload-icon{font-size:28px;color:#6366f1;margin-bottom:8px}.ix-upload strong{display:block;font-size:13px;margin-bottom:6px}.ix-upload p{font-size:11px;color:#8a92a5;margin:0 0 13px}
        .ix-file-input{display:none}.ix-notice{margin:0 0 18px;padding:11px 14px;border-radius:9px;background:#eff6ff;color:#1d4ed8;font-size:12px;white-space:pre-wrap;line-height:1.5}
        .ix-table-wrap{overflow:auto;border:1px solid #edf0f5;border-radius:10px}.ix-table{border-collapse:collapse;width:100%;font-size:11px;white-space:nowrap}
        .ix-table th{background:#f8f9fc;color:#697386;text-align:left;font-weight:700;padding:12px 14px;border-bottom:1px solid #e8eaf1}
        .ix-table td{padding:11px 14px;border-bottom:1px solid #f0f1f5;color:#414b60}.ix-table tr:last-child td{border-bottom:0}.ix-table tr:hover td{background:#fafbff}
        .ix-controls{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:15px}.ix-input,.ix-select{border:1px solid #e0e4ee;background:#fff;border-radius:9px;padding:10px 12px;font-size:12px;color:#344054;outline:none}
        .ix-input:focus,.ix-select:focus{border-color:#818cf8;box-shadow:0 0 0 3px #6366f11a}.ix-input{min-width:180px;flex:1}
        .ix-chart-area{height:300px;width:100%;margin-top:8px}.ix-chart-selects{display:flex;gap:9px;flex-wrap:wrap}
        .ix-chart-selects label{display:grid;gap:5px;font-size:10px;color:#7b8498;font-weight:700}.ix-chart-selects .ix-select{min-width:140px}
        .ix-question-form{display:flex;gap:9px;margin-bottom:15px}.ix-answer{white-space:pre-wrap;line-height:1.75;font-size:12px;color:#48536a;background:#f8faff;border:1px solid #e8ebf6;border-radius:11px;padding:16px}
        .ix-prompt-chips{display:flex;gap:7px;flex-wrap:wrap;margin-top:12px}.ix-prompt-chips button{border:1px solid #e2e6f0;background:#fff;color:#6d7488;padding:7px 10px;border-radius:20px;font-size:10px}
        .ix-prompt-chips button:hover{border-color:#a5b4fc;color:#4338ca}.ix-kpi-list{display:grid;gap:9px}.ix-kpi{display:flex;justify-content:space-between;gap:12px;border-bottom:1px solid #f0f1f5;padding:10px 0;font-size:12px}.ix-kpi:last-child{border:0}.ix-kpi span{color:#778096}.ix-kpi strong{color:#273149}
        .ix-footer{text-align:center;font-size:10px;color:#a0a6b5;padding:22px 0 0}.ix-footer strong{color:#727b90}
        @media(max-width:1000px){.ix-sidebar{width:205px}.ix-content{padding:25px 22px}.ix-stats{grid-template-columns:repeat(2,minmax(0,1fr))}.ix-grid-2{grid-template-columns:1fr}}
        @media(max-width:650px){.ix-layout{display:block}.ix-sidebar{width:100%;padding:14px}.ix-brand{padding:0 7px 14px}.ix-nav{grid-template-columns:repeat(3,minmax(0,1fr))}.ix-nav-label,.ix-side-bottom{display:none}.ix-nav button{justify-content:center;flex-direction:column;gap:4px;padding:9px 4px;font-size:10px}.ix-topbar{height:58px;padding:0 16px}.ix-content{padding:20px 13px}.ix-heading-row{display:block}.ix-actions{margin-top:15px}.ix-title{font-size:25px}.ix-stats{gap:9px}.ix-stat{padding:14px}.ix-stat-value{font-size:23px}.ix-banner{align-items:flex-start;flex-direction:column}.ix-question-form{flex-direction:column}.ix-card{padding:15px}}
      `}</style>
      <div className="ix-layout">
        <aside className="ix-sidebar">
          <div className="ix-brand"><div className="ix-logo">✳</div><div>InsightX AI<small>DISCOVER WHAT DATA HIDES</small></div></div>
          <div className="ix-nav-label">WORKSPACE</div>
          <nav className="ix-nav">
            {pages.map((page) => (
              <button key={page.name} className={activePage === page.name ? "active" : ""} onClick={() => setActivePage(page.name)}>
                <span className="ix-nav-icon">{page.icon}</span>{page.name}
              </button>
            ))}
          </nav>
          <div className="ix-side-bottom"><span className="ix-side-pill">● Local processing</span><br />Your files are analysed in your browser and are not uploaded to an InsightX server.<br /><br />InsightX AI · Portfolio edition</div>
        </aside>
        <main className="ix-main">
          <header className="ix-topbar">
            <div className="ix-breadcrumb">Workspace <span style={{ padding: "0 8px" }}>›</span> <strong>{activePage}</strong></div>
            <div className="ix-top-right"><span className="ix-status"><span className="ix-dot" /> SYSTEM READY</span><div className="ix-avatar">IX</div></div>
          </header>
          <div className="ix-content">
            {notice && <div className="ix-notice" role="status">{notice}<button onClick={() => setNotice("")} style={{ float: "right", border: 0, background: "transparent", color: "inherit" }}>✕</button></div>}
            <div className="ix-heading-row">
              <div><div className="ix-eyebrow">DATA INTELLIGENCE PLATFORM</div><h1 className="ix-title">{activePage === "Overview" ? "Discover what your data hides." : activePage}</h1><p className="ix-subtitle">Turn raw data into clear answers, useful insights, and confident decisions.</p></div>
              <div className="ix-actions">
                <button className="ix-btn" onClick={() => loadRows(makeSample(), "sample_sales.csv")}>＋ Load sample data</button>
                <button className="ix-btn ix-btn-primary" onClick={() => fileInput.current?.click()}>↑ Upload dataset</button>
                <input ref={fileInput} className="ix-file-input" type="file" accept=".csv,.txt,.xlsx,.xls" onChange={handleFile} />
              </div>
            </div>

            {!currentRows.length && (
              <div className="ix-banner">
                <div><h2>Your next data discovery starts here.</h2><p>Upload a CSV or Excel file to reveal patterns, check data quality, build charts, and ask questions about your data.</p></div>
                <button className="ix-btn" onClick={() => loadRows(makeSample(), "sample_sales.csv")}>Explore sample dataset →</button>
              </div>
            )}

            <div className="ix-stats">
              <StatCard label="TOTAL ROWS" value={numberFormat(currentRows.length)} detail={fileName || "Upload a dataset to begin"} icon="▤" />
              <StatCard label="DATA COLUMNS" value={numberFormat(columns.length)} detail={`${numericColumns.length} numeric · ${categoricalColumns.length} categorical`} icon="▦" />
              <StatCard label="MISSING CELLS" value={numberFormat(analysis.missingCells)} detail={analysis.missingCells ? "Review data quality" : "No blanks detected"} icon="⌁" />
              <StatCard label="DUPLICATE ROWS" value={numberFormat(analysis.duplicateRows)} detail="Exact matching rows" icon="⧉" />
            </div>

            {activePage === "Overview" && <>
              <div className="ix-grid-2">
                <section className="ix-card">
                  <div className="ix-card-head"><div><h2 className="ix-card-title">Automatic Data Insights</h2><p className="ix-card-sub">Quick observations generated from your current dataset</p></div><span className="ix-tag">AUTO ANALYSIS</span></div>
                  {insights.length ? <div className="ix-insight-list">{insights.map((item, index) => <div className="ix-insight" key={`${item.title}-${index}`}><div className="ix-insight-mark">{["✦", "↗", "◉", "✓", "⌁"][index % 5]}</div><div><h3>{item.title}</h3><p>{item.text}</p></div></div>)}</div> : <div className="ix-empty"><div className="ix-empty-icon">✦</div>Upload a dataset or load sample data to generate insights.</div>}
                </section>
                <section className="ix-card">
                  <div className="ix-card-head"><div><h2 className="ix-card-title">Quick Visualization</h2><p className="ix-card-sub">Compare a numeric metric by category</p></div><span className="ix-tag">LIVE</span></div>
                  {chartData.length ? <div className="ix-chart-area"><ResponsiveContainer width="100%" height="100%"><BarChart data={chartData} margin={{ top: 8, right: 8, left: -15, bottom: 25 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#edf0f6" /><XAxis dataKey="name" tick={{ fontSize: 10, fill: "#8991a4" }} angle={-20} textAnchor="end" interval={0} /><YAxis tick={{ fontSize: 10, fill: "#8991a4" }} /><Tooltip formatter={(value) => numberFormat(value)} /><Bar dataKey="value" name={selectedY} fill="#6366f1" radius={[5, 5, 0, 0]} /></BarChart></ResponsiveContainer></div> : <div className="ix-empty"><div className="ix-empty-icon">▥</div>Load data with at least one category column and one numeric column to view a chart.</div>}
                  {currentRows.length > 0 && <div className="ix-kpi-list" style={{ marginTop: 12 }}>{numericColumns.slice(0, 3).map((column) => <div className="ix-kpi" key={column}><span>{column} total</span><strong>{numberFormat(analysis.numericStats[column].total)}</strong></div>)}</div>}
                </section>
              </div>
              <section className="ix-card">
                <div className="ix-card-head"><div><h2 className="ix-card-title">Dataset Preview</h2><p className="ix-card-sub">{currentRows.length ? `Showing ${Math.min(5, currentRows.length)} of ${currentRows.length} rows` : "Your uploaded data will appear here"}</p></div><button className="ix-btn" onClick={() => setActivePage("Data Workspace")}>Open workspace →</button></div>
                {currentRows.length ? <Table rows={currentRows.slice(0, 5)} columns={columns} /> : <div className="ix-empty">No dataset loaded yet. Use <strong>Upload dataset</strong> or <strong>Load sample data</strong> to begin.</div>}
              </section>
            </>}

            {activePage === "Data Workspace" && <section className="ix-card">
              <div className="ix-card-head"><div><h2 className="ix-card-title">Data Workspace</h2><p className="ix-card-sub">Search, inspect, and explore your uploaded data.</p></div><span className="ix-tag">{currentRows.length} ROWS</span></div>
              {!currentRows.length ? <div className="ix-upload"><div className="ix-upload-icon">⇧</div><strong>Upload a dataset to start exploring</strong><p>CSV, XLSX, and XLS files are supported.</p><button className="ix-btn ix-btn-primary" onClick={() => fileInput.current?.click()}>Choose file</button></div> : <>
                <div className="ix-controls"><input className="ix-input" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search any value in the dataset…" /><button className="ix-btn" onClick={exportCsv}>↓ Export CSV</button><button className="ix-btn" onClick={() => setShowAllRows((value) => !value)}>{showAllRows ? "Show first 100" : "Show all rows"}</button></div>
                <div style={{ fontSize: 11, color: "#858da0", marginBottom: 10 }}>Showing {Math.min(showAllRows ? filteredRows.length : 100, filteredRows.length)} of {filteredRows.length} matching rows</div>
                <Table rows={filteredRows.slice(0, showAllRows ? 5000 : 100)} columns={columns} />
              </>}
            </section>}

            {activePage === "AI Analyst" && <div className="ix-grid-2">
              <section className="ix-card">
                <div className="ix-card-head"><div><h2 className="ix-card-title">Ask your data</h2><p className="ix-card-sub">Ask a question in plain English.</p></div><span className="ix-tag">LOCAL ANALYST</span></div>
                <form className="ix-question-form" onSubmit={askQuestion}><input className="ix-input" value={question} onChange={(e) => setQuestion(e.target.value)} placeholder="e.g. What is the total Sales?" /><button className="ix-btn ix-btn-primary" type="submit">Ask question ↗</button></form>
                {answer ? <div className="ix-answer">{answer}</div> : <div className="ix-empty"><div className="ix-empty-icon">✳</div>Ask about totals, averages, highest values, missing cells, duplicates, or dataset summaries.</div>}
                <div className="ix-prompt-chips">{["Give me a summary", "Are there missing values?", "How many rows?", "What is the average?", "Show duplicates"].map((text) => <button key={text} onClick={() => setQuestion(text)}>{text}</button>)}</div>
              </section>
              <section className="ix-card"><div className="ix-card-head"><div><h2 className="ix-card-title">Available columns</h2><p className="ix-card-sub">The analyst uses these columns for calculations.</p></div></div>{columns.length ? <div className="ix-kpi-list">{columns.map((column) => <div className="ix-kpi" key={column}><span>{column}</span><strong>{numericColumns.includes(column) ? "Numeric" : "Category / text"}</strong></div>)}</div> : <div className="ix-empty">Load a dataset to inspect its columns.</div>}<p style={{ fontSize: 10, lineHeight: 1.6, color: "#9299aa", marginTop: 15 }}>Note: this version uses built-in analysis rules, not a connected large language model. It works without an API key or paid AI service.</p></section>
            </div>}

            {activePage === "Visualization Studio" && <section className="ix-card">
              <div className="ix-card-head"><div><h2 className="ix-card-title">Visualization Studio</h2><p className="ix-card-sub">Build a chart from your current dataset.</p></div><button className="ix-btn" onClick={() => downloadBlob("insightx-chart-data.csv", Papa.unparse(chartData), "text/csv;charset=utf-8")}>↓ Export chart data</button></div>
              {!currentRows.length ? <div className="ix-empty">Load a dataset first, then choose the columns for your chart.</div> : <>
                <div className="ix-controls">
                  <label className="ix-chart-selects"><span>Chart type</span><select className="ix-select" value={chartType} onChange={(e) => setChartType(e.target.value)}><option value="bar">Bar chart</option><option value="line">Line chart</option><option value="pie">Pie chart</option></select></label>
                  <label className="ix-chart-selects"><span>Category / X-axis</span><select className="ix-select" value={selectedX} onChange={(e) => setXColumn(e.target.value)}>{columns.map((column) => <option key={column} value={column}>{column}</option>)}</select></label>
                  <label className="ix-chart-selects"><span>Numeric metric</span><select className="ix-select" value={selectedY} onChange={(e) => setYColumn(e.target.value)}>{numericColumns.map((column) => <option key={column} value={column}>{column}</option>)}</select></label>
                </div>
                {chartData.length && selectedY ? <div className="ix-chart-area" style={{ height: 380 }}>
                  <ResponsiveContainer width="100%" height="100%">
                    {chartType === "line" ? <LineChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 30 }}><CartesianGrid strokeDasharray="3 3" stroke="#edf0f6" /><XAxis dataKey="name" tick={{ fontSize: 11 }} angle={-20} textAnchor="end" interval={0} /><YAxis tick={{ fontSize: 11 }} /><Tooltip formatter={(value) => numberFormat(value)} /><Legend /><Line type="monotone" dataKey="value" name={selectedY} stroke="#6366f1" strokeWidth={3} dot={{ r: 4 }} /></LineChart> :
                      chartType === "pie" ? <PieChart><Tooltip formatter={(value) => numberFormat(value)} /><Legend /> <Pie data={chartData.slice(0, 10)} dataKey="value" nameKey="name" cx="50%" cy="48%" outerRadius={115} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>{chartData.slice(0, 10).map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}</Pie></PieChart> :
                        <BarChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 30 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#edf0f6" /><XAxis dataKey="name" tick={{ fontSize: 11 }} angle={-20} textAnchor="end" interval={0} /><YAxis tick={{ fontSize: 11 }} /><Tooltip formatter={(value) => numberFormat(value)} /><Legend /><Bar dataKey="value" name={selectedY} fill="#6366f1" radius={[5, 5, 0, 0]}>{chartData.map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}</Bar></BarChart>}
                  </ResponsiveContainer>
                </div> : <div className="ix-empty">Choose a dataset with at least one numeric column to generate a chart.</div>}
              </>}
            </section>}

            {activePage === "Data Cleaning" && <div className="ix-grid-2">
              <section className="ix-card"><div className="ix-card-head"><div><h2 className="ix-card-title">Data quality check</h2><p className="ix-card-sub">Review problems before preparing data for analysis.</p></div></div>
                {!currentRows.length ? <div className="ix-empty">Load a dataset to check its quality.</div> : <>
                  <div className="ix-kpi-list"><div className="ix-kpi"><span>Total rows</span><strong>{currentRows.length}</strong></div><div className="ix-kpi"><span>Missing cells</span><strong>{analysis.missingCells}</strong></div><div className="ix-kpi"><span>Exact duplicate rows</span><strong>{analysis.duplicateRows}</strong></div><div className="ix-kpi"><span>Columns with blanks</span><strong>{Object.values(analysis.missingByColumn).filter(Boolean).length}</strong></div></div>
                  <div style={{ marginTop: 18 }}><button className="ix-btn ix-btn-primary" onClick={cleanData}>Remove exact duplicate rows in working copy</button></div>
                  <p style={{ fontSize: 10, lineHeight: 1.7, color: "#9299aa", marginTop: 12 }}>This creates a working copy and does not alter your original file. Blank cells are reported, not automatically filled.</p>
                  {cleanedRows && <div className="ix-notice" style={{ marginTop: 12 }}>Cleaned working copy active: {cleanedRows.length} rows. Export it from Data Workspace.</div>}
                </>}
              </section>
              <section className="ix-card"><div className="ix-card-head"><div><h2 className="ix-card-title">Missing values by column</h2><p className="ix-card-sub">A count of blank cells in each field.</p></div></div>
                {columns.length ? <div className="ix-kpi-list">{columns.map((column) => <div className="ix-kpi" key={column}><span>{column}</span><strong style={{ color: analysis.missingByColumn[column] ? "#dc2626" : "#059669" }}>{analysis.missingByColumn[column]} missing</strong></div>)}</div> : <div className="ix-empty">No columns to inspect.</div>}
              </section>
            </div>}

            {activePage === "Reports" && <div className="ix-grid-2">
              <section className="ix-card"><div className="ix-card-head"><div><h2 className="ix-card-title">Export analysis report</h2><p className="ix-card-sub">Download a JSON report with the current analysis.</p></div><span className="ix-tag">JSON</span></div>
                <p style={{ fontSize: 12, lineHeight: 1.8, color: "#7d8599" }}>The report includes the source filename, generation timestamp, dataset dimensions, numeric summaries, missing values, duplicate counts, and automatic insights.</p>
                <button className="ix-btn ix-btn-primary" disabled={!currentRows.length} onClick={exportReport}>↓ Download analysis report</button>
              </section>
              <section className="ix-card"><div className="ix-card-head"><div><h2 className="ix-card-title">Export data</h2><p className="ix-card-sub">Download the current dataset as a CSV.</p></div><span className="ix-tag">CSV</span></div>
                <p style={{ fontSize: 12, lineHeight: 1.8, color: "#7d8599" }}>If you removed duplicates in the Data Cleaning page, this exports the cleaned working copy. Your original file on your computer is not changed.</p>
                <button className="ix-btn" disabled={!currentRows.length} onClick={exportCsv}>↓ Download current data</button>
              </section>
            </div>}

            <div className="ix-footer">Built with React, Vite, Recharts, PapaParse & SheetJS · <strong>InsightX AI</strong> · Data stays in your browser</div>
          </div>
        </main>
      </div>
    </div>
  );
}

function Table({ rows, columns }) {
  if (!rows.length) return <div className="ix-empty">No matching rows.</div>;
  return <div className="ix-table-wrap"><table className="ix-table"><thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index}>{columns.map((column) => <td key={column}>{String(row[column] ?? "") || <span style={{ color: "#c2c7d2" }}>—</span>}</td>)}</tr>)}</tbody></table></div>;
}
