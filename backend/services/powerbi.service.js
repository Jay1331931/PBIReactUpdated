const axios = require("axios");
const  pool  = require("../db/db");
const { generateKey } = require("../utils/generateKey");

// 🔐 Get Power BI Access Token
async function getAccessToken() {
  const response = await axios.post(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      grant_type: "client_credentials",
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: "https://analysis.windows.net/powerbi/api/.default"
    })
  );

  return response.data.access_token;
}
async function getAllWorkspaces() {
  const token = await getAccessToken();

  const res = await axios.get(
    `https://api.powerbi.com/v1.0/myorg/groups`,
    {
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  return res.data.value;
}
// 📊 Fetch reports from Power BI
async function getPowerBIReports() {
  const token = await getAccessToken();
  const workspaces = await getAllWorkspaces();

  const allReports = await Promise.all(
    workspaces.map(async (ws) => {
      const res = await axios.get(
        `https://api.powerbi.com/v1.0/myorg/groups/${ws.id}/reports`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      // Tag each report with its workspace
      return res.data.value.map(report => ({
        ...report,
        workspaceId: ws.id,
        workspaceName: ws.name
      }));
    })
  );

  return allReports.flat(); // single array of all reports across all workspaces
}

// 🔄 Sync reports into DB
async function syncReports() {
  const reports = await getPowerBIReports();

  const client = await pool.connect();

  try {
    await client.query("BEGIN");

    const reportIds = [];

    for (const rpt of reports) {
      const reportKey = generateKey(rpt.name);
      reportIds.push(rpt.id);

      await client.query(
        `
        INSERT INTO powerbi_report_mapping (
          report_key, report_name, report_id, dataset_id, workspace_id
        )
        VALUES ($1, $2, $3, $4, $5)
        ON CONFLICT (report_key)
        DO UPDATE SET
          report_name = EXCLUDED.report_name,
          report_id = EXCLUDED.report_id,
          dataset_id = EXCLUDED.dataset_id,
          workspace_id = EXCLUDED.workspace_id,
          is_active = TRUE,
          updated_at = CURRENT_TIMESTAMP;
        `,
        [
          reportKey,
          rpt.name,
          rpt.id,
          rpt.datasetId,
          rpt.workspaceId
        ]
      );
    }

    // 🧹 Mark deleted reports as inactive
    if (reportIds.length > 0) {
      await client.query(
        `
        UPDATE powerbi_report_mapping
        SET is_active = FALSE
        WHERE report_id NOT IN (${reportIds.map((_, i) => `$${i + 1}`).join(",")})
        `,
        reportIds
      );
    }

    await client.query("COMMIT");
    // console.log(`✅ Synced ${reports.length} reports`);

  } catch (err) {
    await client.query("ROLLBACK");
    // console.error("❌ Sync failed:", err);
  } finally {
    client.release();
  }
}

module.exports = {
  syncReports
};