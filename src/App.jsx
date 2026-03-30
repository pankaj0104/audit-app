import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const CANDIDATE_HEADERS = [
  "ess code",
  "employee code",
  "emp code",
  "employee id",
  "candidate code",
  "candidate id",
  "id",
  "code",
];

const MONTH_HEADERS = ["month", "salary month", "wage month", "period"];
const GROSS_HEADERS = ["gross", "gross salary", "earned gross", "total earnings"];
const NET_HEADERS = ["net", "net pay", "take home", "payable"];
const PT_HEADERS = ["pt", "professional tax", "prof tax"];
const PAID_HEADERS = ["paid days", "days paid", "present days", "attendance days", "working days"];
const NAME_HEADERS = ["name", "employee name", "candidate name", "staff name"];
const REMARK_HEADERS = ["remarks", "remark"];

function normalize(v) {
  return String(v ?? "").trim();
}

function lower(v) {
  return normalize(v).toLowerCase();
}

function num(v) {
  if (v === null || v === undefined || v === "") return null;
  const cleaned = String(v).replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function findHeaderIndex(headers, candidates) {
  const lowered = headers.map((h) => lower(h));
  return lowered.findIndex((header) => candidates.some((candidate) => header.includes(candidate)));
}

function detectColumns(headers) {
  return {
    code: findHeaderIndex(headers, CANDIDATE_HEADERS),
    name: findHeaderIndex(headers, NAME_HEADERS),
    month: findHeaderIndex(headers, MONTH_HEADERS),
    gross: findHeaderIndex(headers, GROSS_HEADERS),
    net: findHeaderIndex(headers, NET_HEADERS),
    pt: findHeaderIndex(headers, PT_HEADERS),
    paid: findHeaderIndex(headers, PAID_HEADERS),
    remark: findHeaderIndex(headers, REMARK_HEADERS),
  };
}

function safe(row, index) {
  if (!row || index < 0) return "";
  return row[index] ?? "";
}

function ensureRemarksColumn(rows, headers, cols) {
  const outRows = rows.map((row) => [...row]);
  const outHeaders = [...headers];
  const outCols = { ...cols };

  if (outCols.remark === -1) {
    outHeaders.push("Remarks");
    outRows[0] = outHeaders;
    for (let i = 1; i < outRows.length; i += 1) {
      outRows[i].push("");
    }
    outCols.remark = outHeaders.length - 1;
  }

  return { outRows, outHeaders, outCols };
}

function buildRowRemark(currentRow, previousRow, cols) {
  const notes = [];

  const currentGross = num(safe(currentRow, cols.gross));
  const previousGross = num(safe(previousRow, cols.gross));
  const currentNet = num(safe(currentRow, cols.net));
  const previousNet = num(safe(previousRow, cols.net));
  const currentPt = num(safe(currentRow, cols.pt));
  const previousPt = num(safe(previousRow, cols.pt));
  const currentPaid = num(safe(currentRow, cols.paid));
  const previousPaid = num(safe(previousRow, cols.paid));

  if (previousRow) {
    if (currentGross !== null && previousGross !== null && currentGross !== previousGross) {
      notes.push(`Gross changed ${previousGross} → ${currentGross}`);
    }
    if (currentNet !== null && previousNet !== null && currentNet !== previousNet) {
      notes.push(`Net changed ${previousNet} → ${currentNet}`);
    }
    if (currentPt !== null && previousPt !== null && currentPt !== previousPt) {
      notes.push(`PT changed ${previousPt} → ${currentPt}`);
    }
    if (currentPaid !== null && previousPaid !== null && currentPaid !== previousPaid) {
      notes.push(`Paid days changed ${previousPaid} → ${currentPaid}`);
    }
  }

  return notes.join(" | ") || "No major variance";
}

function buildChangeSummary(currentRow, previousRow, cols, sheetName, rowIndex) {
  const changes = [];
  if (!previousRow) return changes;

  const mapping = [
    { key: "Gross", index: cols.gross },
    { key: "Net", index: cols.net },
    { key: "PT", index: cols.pt },
    { key: "Paid Days", index: cols.paid },
  ];

  for (const field of mapping) {
    if (field.index < 0) continue;
    const currentValue = safe(currentRow, field.index);
    const previousValue = safe(previousRow, field.index);
    if (normalize(currentValue) !== normalize(previousValue)) {
      changes.push({
        sheet: sheetName,
        rowNumber: rowIndex + 1,
        field: field.key,
        from: normalize(previousValue),
        to: normalize(currentValue),
      });
    }
  }

  return changes;
}

function buildDuplicateReport(processedSheets) {
  const map = new Map();
  const duplicates = [];

  for (const sheet of processedSheets) {
    const codeIndex = sheet.cols.code;
    if (codeIndex < 0) continue;

    for (let i = 1; i < sheet.rows.length; i += 1) {
      const code = lower(safe(sheet.rows[i], codeIndex));
      if (!code) continue;
      if (!map.has(code)) map.set(code, []);
      map.get(code).push({
        sheet: sheet.name,
        rowNumber: i + 1,
        code: normalize(safe(sheet.rows[i], codeIndex)),
        name: sheet.cols.name >= 0 ? normalize(safe(sheet.rows[i], sheet.cols.name)) : "",
      });
    }
  }

  for (const [, entries] of map.entries()) {
    if (entries.length > 1) {
      duplicates.push(...entries.map((entry) => ({ ...entry, status: "Duplicate Entry" })));
    }
  }

  return duplicates;
}

function monthOrderScore(value) {
  const t = lower(value);
  const months = {
    jan: 1, january: 1,
    feb: 2, february: 2,
    mar: 3, march: 3,
    apr: 4, april: 4,
    may: 5,
    jun: 6, june: 6,
    jul: 7, july: 7,
    aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    oct: 10, october: 10,
    nov: 11, november: 11,
    dec: 12, december: 12,
  };
  const yearMatch = t.match(/20\d{2}/);
  const year = yearMatch ? Number(yearMatch[0]) : 0;
  const monthKey = Object.keys(months).find((m) => t.includes(m));
  const month = monthKey ? months[monthKey] : 0;
  return year * 100 + month;
}

function generateTimelineData(processedSheets, searchCode) {
  const timeline = [];
  const codeNeedle = lower(searchCode);
  if (!codeNeedle) return timeline;

  for (const sheet of processedSheets) {
    if (sheet.cols.code < 0) continue;

    for (let i = 1; i < sheet.rows.length; i += 1) {
      const row = sheet.rows[i];
      const code = lower(safe(row, sheet.cols.code));
      if (code === codeNeedle) {
        timeline.push({
          sheet: sheet.name,
          month: normalize(safe(row, sheet.cols.month)),
          gross: normalize(safe(row, sheet.cols.gross)),
          net: normalize(safe(row, sheet.cols.net)),
          pt: normalize(safe(row, sheet.cols.pt)),
          paidDays: normalize(safe(row, sheet.cols.paid)),
          rowNumber: i + 1,
        });
      }
    }
  }

  return timeline.sort((a, b) => monthOrderScore(a.month) - monthOrderScore(b.month));
}

function analyzeWorkbook(workbook) {
  const processedSheets = [];
  const changeReport = [];

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rawRows.length) continue;

    const headers = rawRows[0].map((h) => normalize(h));
    const cols = detectColumns(headers);
    const { outRows, outHeaders, outCols } = ensureRemarksColumn(rawRows, headers, cols);

    for (let i = 1; i < outRows.length; i += 1) {
      const currentRow = outRows[i];
      const previousRow = i > 1 ? outRows[i - 1] : null;
      outRows[i][outCols.remark] = buildRowRemark(currentRow, previousRow, outCols);
      changeReport.push(...buildChangeSummary(currentRow, previousRow, outCols, sheetName, i));
    }

    processedSheets.push({
      name: sheetName,
      rows: outRows,
      headers: outHeaders,
      cols: outCols,
    });
  }

  const duplicateReport = buildDuplicateReport(processedSheets);

  return {
    sheets: processedSheets,
    changeReport,
    duplicateReport,
  };
}

function cardStyle() {
  return {
    background: "#fff",
    borderRadius: 18,
    padding: 16,
    boxShadow: "0 2px 10px rgba(0,0,0,0.06)",
    border: "1px solid #e5e7eb",
  };
}

function buttonStyle(primary = true) {
  return {
    height: 44,
    borderRadius: 12,
    border: "none",
    padding: "0 16px",
    cursor: "pointer",
    background: primary ? "#111827" : "#e5e7eb",
    color: primary ? "#fff" : "#111827",
    fontWeight: 600,
  };
}

export default function App() {
  const fileRef = useRef(null);
  const [fileName, setFileName] = useState("");
  const [analysis, setAnalysis] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [searchCode, setSearchCode] = useState("");
  const [timeline, setTimeline] = useState([]);
  const [searchText, setSearchText] = useState("");
  const [auditRan, setAuditRan] = useState(false);
  const [workbook, setWorkbook] = useState(null);

  const activeSheet = useMemo(() => {
    if (!analysis?.sheets?.length) return null;
    return analysis.sheets.find((sheet) => sheet.name === selectedSheet) || analysis.sheets[0];
  }, [analysis, selectedSheet]);

  const visibleRows = useMemo(() => {
    if (!activeSheet?.rows?.length) return [];
    const q = lower(searchText);
    const body = activeSheet.rows.slice(1);
    if (!q) return body;
    return body.filter((row) => row.some((cell) => lower(cell).includes(q)));
  }, [activeSheet, searchText]);

  const stats = useMemo(() => {
    if (!analysis) return { sheets: 0, rows: 0, changes: 0, duplicates: 0 };
    return {
      sheets: analysis.sheets.length,
      rows: analysis.sheets.reduce((sum, s) => sum + Math.max(s.rows.length - 1, 0), 0),
      changes: analysis.changeReport.length,
      duplicates: analysis.duplicateReport.length,
    };
  }, [analysis]);

  const handleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    setWorkbook(wb);
    setFileName(file.name);
    setAnalysis(null);
    setSelectedSheet("");
    setTimeline([]);
    setAuditRan(false);
  };

  const runFullAudit = () => {
    if (!workbook) return;
    const result = analyzeWorkbook(workbook);
    setAnalysis(result);
    setSelectedSheet(result.sheets[0]?.name || "");
    setAuditRan(true);
  };

  const searchTimeline = () => {
    if (!analysis?.sheets?.length) return;
    setTimeline(generateTimelineData(analysis.sheets, searchCode));
  };

  const downloadWorkbook = () => {
    if (!analysis?.sheets?.length) return;
    const wb = XLSX.utils.book_new();
    analysis.sheets.forEach((sheet) => {
      const ws = XLSX.utils.aoa_to_sheet(sheet.rows);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(
      wb,
      fileName ? fileName.replace(/(\.xlsx|\.xls)$/i, "") + "_final_audit.xlsx" : "final_audit.xlsx"
    );
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f3f4f6", padding: 16, fontFamily: "Arial, sans-serif", color: "#111827" }}>
      <div style={{ maxWidth: 1200, margin: "0 auto" }}>
        <div style={{ ...cardStyle(), marginBottom: 16 }}>
          <h1 style={{ margin: 0, fontSize: 28 }}>Comparison Audit Dashboard</h1>
          <p style={{ color: "#6b7280", lineHeight: 1.6 }}>
            Workbook upload karo, phir <b>Run Full Audit</b> dabao. Update report, duplicate report,
            ESS timeline, preview aur final Excel sab yahin milega.
          </p>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 12 }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleUpload} />
            <button style={buttonStyle(true)} onClick={() => fileRef.current?.click()}>Upload Workbook</button>
            <button style={buttonStyle(false)} onClick={runFullAudit}>Run Full Audit</button>
            <button style={buttonStyle(true)} onClick={downloadWorkbook} disabled={!auditRan}>Download Final Excel</button>
          </div>

          <div style={{ marginTop: 12, color: "#6b7280" }}>
            {fileName ? `Selected file: ${fileName}` : "No workbook selected yet."}
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 12, marginBottom: 16 }}>
          <div style={cardStyle()}><b>Sheets</b><div style={{ fontSize: 24, marginTop: 8 }}>{stats.sheets}</div></div>
          <div style={cardStyle()}><b>Rows</b><div style={{ fontSize: 24, marginTop: 8 }}>{stats.rows}</div></div>
          <div style={cardStyle()}><b>Updates</b><div style={{ fontSize: 24, marginTop: 8 }}>{stats.changes}</div></div>
          <div style={cardStyle()}><b>Duplicates</b><div style={{ fontSize: 24, marginTop: 8 }}>{stats.duplicates}</div></div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 16, marginBottom: 16 }}>
          <div style={cardStyle()}>
            <h3 style={{ marginTop: 0 }}>Update Report</h3>
            <div style={{ maxHeight: 260, overflow: "auto", border: "1px solid #e5e7eb", borderRadius: 12 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
                <thead style={{ background: "#f9fafb" }}>
                  <tr>
                    <th style={{ padding: 10, textAlign: "left" }}>Sheet</th>
                    <th style={{ padding: 10, textAlign: "left" }}>Row</th>
                    <th style={{ padding: 10, textAlign: "left" }}>Field</th>
                    <th style={{ padding: 10, textAlign: "left" }}>From</th>
                    <th style={{ padding: 10, textAlign: "left" }}>To</th>
                  </tr>
                </thead>
                <tbody>
                  {(analysis?.changeReport || []).length ? analysis.changeReport.map((item, idx) => (
                    <tr key={idx} style={{ borderTop: "1px solid #e5e7eb" }}>
                      <td style={{ padding: 10 }}>{item.sheet}</td>
                      <td style={{ padding: 10 }}>{item.rowNumber}</td>
                      <td style={{ padding: 10 }}>{item.field}</td>
                      <td style={{ padding: 10 }}>{item.from}</td>
                      <td style={{ padding: 10 }}>{item.to}</td>
                    </tr>
                  )) : (
                    <tr><td style={{ padding: 10 }} colSpan={5}>No updates detected yet.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div style={cardStyle()}>
            <h3 style={{ marginTop: 0 }}>Timeline Search</h3>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
              <input
                value={searchCode}
                onChange={(e) => setSearchCode(e.target.value)}
                placeholder="Enter ESS code / employee code"
                style={{ flex: 1, minWidth: 220, height: 42, borderRadius: 12, border: "1px solid #d1d5db", padding: "0 12px" }}
              />
              <button style={buttonStyle(true)} onClick={searchTimeline}>Search Timeline</button>
            </div>

            <div style={{ maxHeight: 260, overflow: "auto", border: "1px solid #e5e7eb", borderRadius: 12 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
                <thead style={{ background: "#f9fafb" }}>
                  <tr>
                    <th style={{ padding: 10, textAlign: "left" }}>Month</th>
                    <th style={{ padding: 10, textAlign: "left" }}>Gross</th>
                    <th style={{ padding: 10, textAlign: "left" }}>Net</th>
                    <th style={{ padding: 10, textAlign: "left" }}>PT</th>
                    <th style={{ padding: 10, textAlign: "left" }}>Sheet</th>
                  </tr>
                </thead>
                <tbody>
                  {timeline.length ? timeline.map((item, idx) => (
                    <tr key={idx} style={{ borderTop: "1px solid #e5e7eb" }}>
                      <td style={{ padding: 10 }}>{item.month}</td>
                      <td style={{ padding: 10 }}>{item.gross}</td>
                      <td style={{ padding: 10 }}>{item.net}</td>
                      <td style={{ padding: 10 }}>{item.pt}</td>
                      <td style={{ padding: 10 }}>{item.sheet}</td>
                    </tr>
                  )) : (
                    <tr><td style={{ padding: 10 }} colSpan={5}>No timeline data yet.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div style={cardStyle()}>
          <h3 style={{ marginTop: 0 }}>Workbook Preview</h3>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
            <select
              value={activeSheet?.name || ""}
              onChange={(e) => setSelectedSheet(e.target.value)}
              style={{ height: 42, borderRadius: 12, border: "1px solid #d1d5db", padding: "0 12px", minWidth: 180 }}
            >
              <option value="">Select sheet</option>
              {(analysis?.sheets || []).map((sheet) => (
                <option key={sheet.name} value={sheet.name}>{sheet.name}</option>
              ))}
            </select>

            <input
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              placeholder="Search in sheet"
              style={{ flex: 1, minWidth: 220, height: 42, borderRadius: 12, border: "1px solid #d1d5db", padding: "0 12px" }}
            />
          </div>

          {!activeSheet ? (
            <div style={{ color: "#6b7280" }}>Run audit to preview workbook.</div>
          ) : (
            <div style={{ overflow: "auto", border: "1px solid #e5e7eb", borderRadius: 12, background: "#fff" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
                <thead style={{ background: "#f9fafb" }}>
                  <tr>
                    {activeSheet.headers.map((header, idx) => (
                      <th key={idx} style={{ padding: 10, textAlign: "left", whiteSpace: "nowrap" }}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.length ? visibleRows.map((row, ridx) => (
                    <tr key={ridx} style={{ borderTop: "1px solid #e5e7eb" }}>
                      {activeSheet.headers.map((_, cidx) => (
                        <td key={cidx} style={{ padding: 10, verticalAlign: "top" }}>{String(row[cidx] ?? "")}</td>
                      ))}
                    </tr>
                  )) : (
                    <tr><td style={{ padding: 10 }} colSpan={activeSheet.headers.length}>No matching rows found.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
