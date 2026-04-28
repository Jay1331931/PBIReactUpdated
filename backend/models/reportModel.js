const axios = require("axios");
const pool = require("../db/db");
const { fetchAzureToken } = require("../services/reportServices.js");
  const { uploadAndGetSasUrl } = require("../services/azureBlobService.js");
const XLSX = require("xlsx");
// ── Helper: fetch report mapping from DB ──────────────────────────────────────
async function getReportMapping(reportKey) {
  const result = await pool.query(
    `SELECT report_id, dataset_id, workspace_id
     FROM powerbi_report_mapping
     WHERE report_key = $1 AND is_active = TRUE`,
    [reportKey]
  );
  if (result.rows.length === 0) throw new Error(`Report not found for reportKey: ${reportKey}`);
  return result.rows[0];
}
// ── RFC 4180-compliant CSV parser ────────────────────────────────────────────
function parseCsvRow(line) {
  const cells = [];
  let i = 0;
  while (i < line.length) {
    if (line[i] === '"') {
      let cell = "";
      i++;
      while (i < line.length) {
        if (line[i] === '"' && line[i + 1] === '"') { cell += '"'; i += 2; }
        else if (line[i] === '"') { i++; break; }
        else { cell += line[i++]; }
      }
      cells.push(cell.trim());
      if (line[i] === ",") i++;
    } else {
      const end = line.indexOf(",", i);
      if (end === -1) { cells.push(line.slice(i).trim()); break; }
      cells.push(line.slice(i, end).trim());
      i = end + 1;
    }
  }
  return cells;
}

function parseCsv(rawCsv) {
  const rows = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < rawCsv.length; i++) {
    const ch = rawCsv[i];
    if (ch === '"') {
      if (inQuotes && rawCsv[i + 1] === '"') { current += '"'; i++; }
      else { inQuotes = !inQuotes; current += ch; }
    } else if (ch === "\n" && !inQuotes) {
      if (current.trim()) rows.push(parseCsvRow(current));
      current = "";
    } else {
      current += ch;
    }
  }
  if (current.trim()) rows.push(parseCsvRow(current));
  return rows;
}

// ── Chronological sort key for pivot column values ───────────────────────────
const MONTH_ORDER = { January:1,February:2,March:3,April:4,May:5,June:6,
                      July:7,August:8,September:9,October:10,November:11,December:12 };
const QUARTER_ORDER = { Q1:1, Q2:2, Q3:3, Q4:4 };
const DAY_ORDER = { Monday:1,Tuesday:2,Wednesday:3,Thursday:4,Friday:5,Saturday:6,Sunday:7 };

function pivotSortKey(val) {
  if (MONTH_ORDER[val])   return [0, MONTH_ORDER[val]];
  if (QUARTER_ORDER[val]) return [1, QUARTER_ORDER[val]];
  if (DAY_ORDER[val])     return [2, DAY_ORDER[val]];
  const n = parseFloat(val);
  if (!isNaN(n))          return [3, n];
  return [4, val];
}

// ── Main pivot function ───────────────────────────────────────────────────────
function pivotCsvData(rawCsv) {

  // ── 1. Parse CSV properly (handles quoted commas, newlines, escaped quotes) ─
  const rows = parseCsv(rawCsv);
  if (rows.length < 2) return rows;

  const headers  = rows[0];
  const dataRows = rows.slice(1);
  const totalRows = dataRows.length;

  // ── 2. Classify columns ─────────────────────────────────────────────────────
  const colData = headers.map((h, i) => {
    const values = dataRows.map(r => r[i]).filter(v => v !== "" && v != null);
    const isNumeric = values.length > 0 && values.every(v => !isNaN(parseFloat(v)) && isFinite(v));
    const uniqueCount = new Set(values).size;
    return { name: h, index: i, isNumeric, uniqueCount,
             cardinalityRatio: uniqueCount / totalRows };
  });

  const metricCols = colData.filter(c => c.isNumeric && c.cardinalityRatio > 0.15);
  const dimCols    = colData.filter(c => !metricCols.includes(c));

  if (!metricCols.length || !dimCols.length) return [headers, ...dataRows];

  // ── 3. Find pivot column (last named dim, skip null/empty headers) ──────────
  const pivotCol = [...dimCols].reverse().find(c => c.name && c.name.trim());
  if (!pivotCol) return [headers, ...dataRows];

  // ── 4. Split remaining dims into const (year label) vs row dims ─────────────
  const otherDims    = dimCols.filter(c => c !== pivotCol);
  const constDims    = otherDims.filter(c => c.uniqueCount === 1);
  const candidateDims = otherDims
    .filter(c => c.uniqueCount > 1)
    .sort((a, b) => a.uniqueCount - b.uniqueCount);  // ascending by cardinality

  // FIX BUG 2: drop sub-dimensions — any dim whose cardinality jumps ≥5x
  // from the previous one is a leaf-level detail not shown in the visual matrix
  const rowDims = [];
  for (let i = 0; i < candidateDims.length; i++) {
    if (i === 0) { rowDims.push(candidateDims[i]); continue; }
    const jump = candidateDims[i].uniqueCount / candidateDims[i - 1].uniqueCount;
    if (jump >= 5) break;  // everything from here is a sub-dimension
    rowDims.push(candidateDims[i]);
  }
  // Restore original column order
  rowDims.sort((a, b) => a.index - b.index);

  // ── 5. Pivot column values — sorted chronologically, not alphabetically ─────
  const pivotValSet = new Set(dataRows.map(r => r[pivotCol.index]).filter(v => v));
  const pivotVals   = [...pivotValSet].sort((a, b) => {
    const ka = pivotSortKey(a), kb = pivotSortKey(b);
    return ka[0] !== kb[0] ? ka[0] - kb[0] : ka[1] < kb[1] ? -1 : ka[1] > kb[1] ? 1 : 0;
  });

  // ── 6. Aggregate ────────────────────────────────────────────────────────────
  const lookup   = {};
  const rowOrder = [];

  for (const row of dataRows) {
    const rowKey  = rowDims.map(d => row[d.index]).join("\x00");
    const pivotVal = row[pivotCol.index];
    if (!pivotVal) continue;

    if (!lookup[rowKey]) { lookup[rowKey] = {}; rowOrder.push(rowKey); }
    if (!lookup[rowKey][pivotVal]) {
      lookup[rowKey][pivotVal] = {};
      for (const m of metricCols) lookup[rowKey][pivotVal][m.name] = 0;
    }
    for (const m of metricCols) {
      lookup[rowKey][pivotVal][m.name] += parseFloat(row[m.index]) || 0;
    }
  }

  // ── 7. Build output AOA ─────────────────────────────────────────────────────
  const aoa = [];
  const constLabel  = constDims.map(c => `${c.name}: ${dataRows[0][c.index]}`).join(", ");
  const pivotHeader = constLabel ? `${pivotCol.name} (${constLabel})` : pivotCol.name;

  // Row 0: pivot column header spanning all value columns
  aoa.push([...rowDims.map(() => ""), pivotHeader,
            ...Array(pivotVals.length - 1).fill(""), ""]);

  // Row 1: row dim names + pivot values + Total
  aoa.push([...rowDims.map(d => d.name), ...pivotVals, "Total"]);

  // Data rows + accumulators
  const grandTotals = Object.fromEntries(metricCols.map(m => [m.name, 0]));
  const colTotals   = Object.fromEntries(
    pivotVals.map(pv => [pv, Object.fromEntries(metricCols.map(m => [m.name, 0]))])
  );

  for (const rowKey of rowOrder) {
    const rowDimVals = rowKey.split("\x00");
    const rowTotals  = Object.fromEntries(metricCols.map(m => [m.name, 0]));
    const cells      = [];

    for (const pv of pivotVals) {
      if (lookup[rowKey][pv]) {
        for (const m of metricCols) {
          const v = lookup[rowKey][pv][m.name];
          rowTotals[m.name]      += v;
          colTotals[pv][m.name] += v;
          grandTotals[m.name]   += v;
        }
        const v = lookup[rowKey][pv][metricCols[0].name];
        cells.push(v !== 0 ? v : "");
      } else {
        cells.push("");
      }
    }

    aoa.push([...rowDimVals, ...cells, rowTotals[metricCols[0].name]]);
  }

  // Grand total row
  aoa.push([
    "Total", ...Array(rowDims.length - 1).fill(""),
    ...pivotVals.map(pv => colTotals[pv][metricCols[0].name]),
    grandTotals[metricCols[0].name]
  ]);

  return aoa;
}
const reportModel = {

  // ── Existing ────────────────────────────────────────────────────────────────
  generateEmbedTokenModel: async (reportKey, username, roles) => {
    const report = await getReportMapping(reportKey);
    const azureToken = await fetchAzureToken();

    const requestBody = {
      datasets: [{ id: report.dataset_id }],
      reports: [{ id: report.report_id, allowEdit: false }],
      targetWorkspaces: [{ id: report.workspace_id }],
      lifeSpanInMinutes: 60,
      identities: username && roles?.length > 0
        ? [{ username, roles, datasets: [report.dataset_id] }]
        : undefined
    };

    const tokenResponse = await axios.post(
      "https://api.powerbi.com/v1.0/myorg/GenerateToken",
      requestBody,
      { headers: { Authorization: `Bearer ${azureToken}`, "Content-Type": "application/json" } }
    ).catch(err => {
      console.error("Power BI API Error:", JSON.stringify(err.response?.data, null, 2));
      throw err;
    });

    const tokenData = tokenResponse.data;
    const embedUrl = `https://app.powerbi.com/reportEmbed?reportId=${report.report_id}&groupId=${report.workspace_id}`;

    return { embedToken: tokenData.token, tokenExpiry: tokenData.expiration, reportId: report.report_id, embedUrl };
  },

  // ── NEW: startExportModel ────────────────────────────────────────────────────
  startExportModel: async (reportKey, format) => {
    const report = await getReportMapping(reportKey);
    const azureToken = await fetchAzureToken();

    // ✅ Power BI Export API requires these exact format strings
    const formatMap = {
      "PDF":  "PDF",
      "XLSX": "EXCEL",      // ← some tenants need "EXCEL" instead, try swapping if still failing
      "PNG":  "PNG",
      "PPTX": "PPTX"
    };

    const pbiFormat = formatMap[format.toUpperCase()];
    if (!pbiFormat) throw new Error(`Unsupported format: ${format}`);

    const requestBody = {
      format: pbiFormat,
      powerBIReportConfiguration: {
        settings: { locale: "en-us" }
      }
    };

    console.log("📤 Sending export request:", JSON.stringify(requestBody, null, 2));
    console.log("📍 URL:", `https://api.powerbi.com/v1.0/myorg/groups/${report.workspace_id}/reports/${report.report_id}/ExportTo`);

    const response = await axios.post(
      `https://api.powerbi.com/v1.0/myorg/groups/${report.workspace_id}/reports/${report.report_id}/ExportTo`,
      requestBody,
      { headers: { Authorization: `Bearer ${azureToken}`, "Content-Type": "application/json" } }
    ).catch(err => {
      console.error("❌ Power BI ExportTo Error:", JSON.stringify(err.response?.data, null, 2));
      throw err;
    });

    return { exportId: response.data.id };
  },

  // ── NEW: getExportStatusModel ────────────────────────────────────────────────
  getExportStatusModel: async (reportKey, exportId, req) => {
    const report = await getReportMapping(reportKey);
    const azureToken = await fetchAzureToken();

    const response = await axios.get(
      `https://api.powerbi.com/v1.0/myorg/groups/${report.workspace_id}/reports/${report.report_id}/exports/${exportId}`,
      { headers: { Authorization: `Bearer ${azureToken}` } }
    ).catch(err => {
      console.error("❌ Power BI Export Status Error:", JSON.stringify(err.response?.data, null, 2));
      throw err;
    });

    const { status } = response.data;

    // ✅ Build absolute download URL pointing to our own proxy endpoint
    let downloadUrl = null;
    if (status === "Succeeded") {
      const baseUrl = `${req.protocol}://${req.get("host")}`;
      downloadUrl = `${baseUrl}/reports/downloadExport?reportKey=${encodeURIComponent(reportKey)}&exportId=${encodeURIComponent(exportId)}`;
    }

    return { status, downloadUrl };
  },

  // ── NEW: downloadExportModel ─────────────────────────────────────────────────
  downloadExportModel: async (reportKey, exportId) => {
    const report = await getReportMapping(reportKey);
    const azureToken = await fetchAzureToken();

    // Step 1: Check status to get file extension
    const statusResponse = await axios.get(
      `https://api.powerbi.com/v1.0/myorg/groups/${report.workspace_id}/reports/${report.report_id}/exports/${exportId}`,
      { headers: { Authorization: `Bearer ${azureToken}` } }
    );

    const { status, resourceFileExtension } = statusResponse.data;

    if (status !== "Succeeded") {
      throw new Error(`Export not ready. Current status: ${status}`);
    }

    // Step 2: Stream the file from Power BI → our server → client
    const fileResponse = await axios.get(
      `https://api.powerbi.com/v1.0/myorg/groups/${report.workspace_id}/reports/${report.report_id}/exports/${exportId}/file`,
      {
        headers: { Authorization: `Bearer ${azureToken}` },
        responseType: "stream" // ✅ stream directly, don't buffer
      }
    );

    // Step 3: Resolve MIME type
    const ext = (resourceFileExtension || "").toUpperCase();
    const mimeType = ext === ".PDF"
      ? "application/pdf"
      : ext === ".XLSX"
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : "application/octet-stream";

    const fileName = `report_${Date.now()}${resourceFileExtension || ".bin"}`;

    return { stream: fileResponse.data, mimeType, fileName };
  },

  // Add inside reportModel object:
  exportToAzureModel: async (reportKey, format, visualData, filters) => {
    const report = await getReportMapping(reportKey);

    let buffer, mimeType, filename;
    const timestamp = Date.now();

  //   if (format === "csv") {
  //     let csvContent = '';
  //     // if (filters && filters.length > 0) {
  //     //   csvContent += 'APPLIED FILTERS:\n';
  //     //   filters.forEach(filter => {
  //     //     const filterTarget = filter.target || 'Filter';
  //     //     const filterValue = filter.values || filter.conditions || 'Applied';
  //     //     const filterType = filter.filterType ? ` (${filter.filterType})` : '';
  //     //     csvContent += `${filterTarget}${filterType}: ${filterValue}\n`;
  //     //   });
  //     //   csvContent += '\n';
  //     // }
  //     csvContent += visualData;
  //     buffer   = Buffer.from(csvContent, "utf-8");
  //     mimeType = "text/csv";
  //     filename = `export_${timestamp}.csv`;

  //   } else if (format === "xlsx") {
  //     // visualData is the raw CSV string from Power BI exportData
  //     const dataRows = visualData
  //       .split("\n")
  //       .filter((row) => row.trim())
  //       .map((row) =>
  //         row.split(",").map((cell) => cell.replace(/^"|"$/g, "").trim())
  //       );
  // const headerRows = [];
  // //  if (filters && filters.length > 0) {
  // //       headerRows.push(['APPLIED FILTERS']);
  // //       headerRows.push(['Filter Target', 'Filter Values', 'Filter Type']);
        
  // //       filters.forEach(filter => {
  // //         headerRows.push([
  // //           filter.target || 'N/A',
  // //           filter.values || filter.conditions || 'All',
  // //           filter.filterType || 'Standard'
  // //         ]);
  // //       });
  // //       headerRows.push([]); // Empty row for spacing
  // //     }
  //     // headerRows.push([]);
  //     const allRows = [...headerRows, ...dataRows]
  //     const wb = XLSX.utils.book_new();
  //     const ws = XLSX.utils.aoa_to_sheet(allRows);
  //     XLSX.utils.book_append_sheet(wb, ws, "Export");

  //     // ✅ Write as buffer directly
  //     const xlsxBuffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
  //     buffer   = xlsxBuffer;
  //     mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  //     filename = `export_${timestamp}.xlsx`;

  //   } else if (format === "pdf") {
  //     // visualData is already a Buffer (from Power BI export API stream)
  //     buffer   = visualData;
  //     mimeType = "application/pdf";
  //     filename = `export_${timestamp}.pdf`;

  //   } else {
  //     throw new Error(`Unsupported format: ${format}`);
  //   }

  if (format === "csv") {
  const pivotedRows = pivotCsvData(visualData);
  const csvContent  = pivotedRows
    .map(row => row.map(cell => `"${cell}"`).join(","))
    .join("\n");
  buffer   = Buffer.from(csvContent, "utf-8");
  mimeType = "text/csv";
  filename = `export_${timestamp}.csv`;

} else if (format === "xlsx") {
  const allRows = pivotCsvData(visualData);
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(allRows);
  XLSX.utils.book_append_sheet(wb, ws, "Export");
  buffer   = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
  mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  filename = `export_${timestamp}.xlsx`;
} else if (format === "pdf") {
      // visualData is already a Buffer (from Power BI export API stream)
      buffer   = visualData;
      mimeType = "application/pdf";
      filename = `export_${timestamp}.pdf`;

    } else {
      throw new Error(`Unsupported format: ${format}`);
    }

    const { sasUrl, blobName } = await uploadAndGetSasUrl(buffer, filename, mimeType);
    return { sasUrl, filename, blobName };
  }
};

module.exports = reportModel;