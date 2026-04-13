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

    if (format === "csv") {
      let csvContent = '';
      // if (filters && filters.length > 0) {
      //   csvContent += 'APPLIED FILTERS:\n';
      //   filters.forEach(filter => {
      //     const filterTarget = filter.target || 'Filter';
      //     const filterValue = filter.values || filter.conditions || 'Applied';
      //     const filterType = filter.filterType ? ` (${filter.filterType})` : '';
      //     csvContent += `${filterTarget}${filterType}: ${filterValue}\n`;
      //   });
      //   csvContent += '\n';
      // }
      csvContent += visualData;
      buffer   = Buffer.from(csvContent, "utf-8");
      mimeType = "text/csv";
      filename = `export_${timestamp}.csv`;

    } else if (format === "xlsx") {
      // visualData is the raw CSV string from Power BI exportData
      const dataRows = visualData
        .split("\n")
        .filter((row) => row.trim())
        .map((row) =>
          row.split(",").map((cell) => cell.replace(/^"|"$/g, "").trim())
        );
  const headerRows = [];
  //  if (filters && filters.length > 0) {
  //       headerRows.push(['APPLIED FILTERS']);
  //       headerRows.push(['Filter Target', 'Filter Values', 'Filter Type']);
        
  //       filters.forEach(filter => {
  //         headerRows.push([
  //           filter.target || 'N/A',
  //           filter.values || filter.conditions || 'All',
  //           filter.filterType || 'Standard'
  //         ]);
  //       });
  //       headerRows.push([]); // Empty row for spacing
  //     }
      // headerRows.push([]);
      const allRows = [...headerRows, ...dataRows]
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(allRows);
      XLSX.utils.book_append_sheet(wb, ws, "Export");

      // ✅ Write as buffer directly
      const xlsxBuffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
      buffer   = xlsxBuffer;
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