const reportModel = require("../models/reportModel.js");
const { byodPool, byodPoolConnect, sql } = require("../config/byodDb.js");

const reportController = {

  // ── Existing ──────────────────────────────────────────────────────────────
  getEmbedToken: async (req, res) => {
    try {
      const { reportKey, d365User } = req.query;
      let username = null;
      let roles = [];

      if (d365User) {
        await byodPoolConnect;

        const result = await byodPool.request()
          .input("email", sql.NVarChar, d365User)
          .query(`
            SELECT TITLEID, NAME 
            FROM HcmEmployeeV2Staging 
            WHERE PRIMARYCONTACTEMAIL = @email
          `);

        const employee = result.recordset;

        if (employee && employee.length > 0) {
          username = d365User;
          roles = employee.flatMap(emp => {
            const name = emp.NAME?.trim();
            const titleId = emp.TITLEID?.trim();
            return [name, titleId].filter(Boolean);
          });
          roles = [...new Set(roles)];
          console.log(`✅ D365 User found: ${employee[0].NAME}, Role: ${employee[0].TITLEID}`);
        } else {
          console.warn(`⚠️ D365 User not found in BYOD: ${d365User}`);
          return res.status(403).json({ success: false, error: "User not authorized to access this report" });
        }
      }

      if (!reportKey) {
        return res.status(400).json({ success: false, error: "reportKey is required" });
      }

      const data = await reportModel.generateEmbedTokenModel(reportKey, username, roles);
      return res.json({ success: true, ...data });

    } catch (error) {
      console.error("❌ Controller Error:", error.message);
      return res.status(500).json({ success: false, error: error.message });
    }
  },

  // ── NEW: POST /reports/export ─────────────────────────────────────────────
  startExport: async (req, res) => {
    try {
      const { reportKey, d365User, format } = req.body;

      if (!reportKey) return res.status(400).json({ success: false, error: "reportKey is required" });
      if (!format)    return res.status(400).json({ success: false, error: "format is required (PDF or XLSX)" });

      const validFormats = ["PDF", "XLSX"];
      if (!validFormats.includes(format.toUpperCase())) {
        return res.status(400).json({ success: false, error: "Unsupported format. Use PDF or XLSX." });
      }

      const { exportId } = await reportModel.startExportModel(reportKey, format.toUpperCase());

      console.log(`✅ Export started: exportId=${exportId}, format=${format}, reportKey=${reportKey}`);
      return res.json({ success: true, exportId });

    } catch (error) {
      console.error("❌ startExport Error:", error.message);
      return res.status(500).json({ success: false, error: error.message });
    }
  },

  // ── NEW: GET /reports/exportStatus ────────────────────────────────────────
  getExportStatus: async (req, res) => {
    try {
      const { reportKey, exportId } = req.query;

      if (!reportKey || !exportId) {
        return res.status(400).json({ success: false, error: "reportKey and exportId are required" });
      }

      const { status, downloadUrl } = await reportModel.getExportStatusModel(
        reportKey,
        exportId,
        req // passed in so we can build an absolute download URL
      );

      return res.json({ success: true, status, ...(downloadUrl && { downloadUrl }) });

    } catch (error) {
      console.error("❌ getExportStatus Error:", error.message);
      return res.status(500).json({ success: false, error: error.message });
    }
  },

  // ── NEW: GET /reports/downloadExport ─────────────────────────────────────
  downloadExport: async (req, res) => {
    try {
      const { reportKey, exportId } = req.query;

      if (!reportKey || !exportId) {
        return res.status(400).json({ success: false, error: "reportKey and exportId are required" });
      }

      const { stream, mimeType, fileName } = await reportModel.downloadExportModel(reportKey, exportId);

      res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
      res.setHeader("Content-Type", mimeType);
      stream.pipe(res); // ✅ stream directly to response — no buffering in memory

    } catch (error) {
      console.error("❌ downloadExport Error:", error.message);
      return res.status(500).json({ success: false, error: error.message });
    }
  },

  // ── NEW: POST /reports/exportToAzure ─────────────────────────────────────────
exportToAzure: async (req, res) => {
  try {
    const { reportKey, format, visualData, filters } = req.body;

    if (!reportKey)   return res.status(400).json({ success: false, error: "reportKey is required" });
    if (!format)      return res.status(400).json({ success: false, error: "format is required" });
    if (!visualData)  return res.status(400).json({ success: false, error: "visualData is required" });

    const { sasUrl, filename } = await reportModel.exportToAzureModel(
      reportKey, format.toLowerCase(), visualData, filters
    );

    return res.json({ success: true, sasUrl, filename });

  } catch (error) {
    console.error("❌ exportToAzure Error:", error.message);
    return res.status(500).json({ success: false, error: error.message });
  }
},
};

module.exports = reportController;