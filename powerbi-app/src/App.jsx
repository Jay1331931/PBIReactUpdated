import { useEffect, useState, useCallback, useRef, useMemo } from "react";
import axios from "axios";
import { PowerBIEmbed } from "powerbi-client-react";
import { models } from "powerbi-client";
import * as XLSX from "xlsx";
import "./App.css";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL;

const EXPORTABLE_TYPES = [
  "tableEx", "pivotTable", "matrix", "table",
  "clusteredBarChart", "barChart", "lineChart",
  "columnChart", "clusteredColumnChart", "pieChart",
  "donutChart", "gauge", "cardVisual", "card",
  "stackedBarChart", "stackedColumnChart", "areaChart",
  "scatterChart", "waterfallChart", "funnel",
];

function App() {
  const [embedConfig, setEmbedConfig]               = useState(null);
  const [loading, setLoading]                       = useState(true);
  const [error, setError]                           = useState(null);
  const [exporting, setExporting]                   = useState(false);
  const [exportFormat, setExportFormat]             = useState(null);
  const [availableVisuals, setAvailableVisuals]     = useState([]);
  const [selectedVisualName, setSelectedVisualName] = useState(null);
  const [showExportPanel, setShowExportPanel]       = useState(false);
  const [blobPopup, setBlobPopup] = useState(null); // { url, filename }

  // ─── Page tab state ────────────────────────────────────────────────────────
  const [pages, setPages]                   = useState([]);        // all visible pages
  const [activePageName, setActivePageName] = useState(null);      // current page name
  const [isReportReady, setIsReportReady]           = useState(false); 
  const reportRef = useRef(null);

  const params    = useMemo(() => new URLSearchParams(window.location.search), []);
  const d365User  = params.get("d365User");
  const reportKey = params.get("reportKey");
  // Add state to track if arrows should be visible
const [showLeftArrow, setShowLeftArrow] = useState(false);
const [showRightArrow, setShowRightArrow] = useState(false);
const [scrollPosition, setScrollPosition] = useState(0);
const tabBarRef = useRef(null);

// Add scroll handlers
const scrollTabs = (direction) => {
  if (tabBarRef.current) {
    const scrollAmount = 200;
    
    if (direction === 'left') {
      tabBarRef.current.scrollBy({
        left: -scrollAmount,
        behavior: 'smooth'
      });
    } else {
      tabBarRef.current.scrollBy({
        left: scrollAmount,
        behavior: 'smooth'
      });
    }
  }
};
// Update scroll handler to check arrow visibility
const handleScroll = () => {
  if (tabBarRef.current) {
    const scrollLeft = tabBarRef.current.scrollLeft;
    const maxScroll = tabBarRef.current.scrollWidth - tabBarRef.current.clientWidth;
    
    setScrollPosition(scrollLeft);
    setShowLeftArrow(scrollLeft > 0);
    setShowRightArrow(scrollLeft < maxScroll - 5); // 5px threshold
  }
};

// Also check on resize and when pages change
useEffect(() => {
  const checkArrows = () => {
    if (tabBarRef.current) {
      const maxScroll = tabBarRef.current.scrollWidth - tabBarRef.current.clientWidth;
      setShowLeftArrow(tabBarRef.current.scrollLeft > 0);
      setShowRightArrow(tabBarRef.current.scrollLeft < maxScroll - 5);
    }
  };
  
  checkArrows();
  window.addEventListener('resize', checkArrows);
  
  return () => window.removeEventListener('resize', checkArrows);
}, [pages]);

  // ─── Load all report pages into tab bar ────────────────────────────────────
  const loadPages = useCallback(async () => {
    try {
      const report = reportRef.current;
      if (!report) return;

      const allPages = await report.getPages();

      // Filter out hidden pages
      const visiblePages = allPages.filter(
        (p) => p.visibility === models.SectionVisibility.AlwaysVisible
      );

      setPages(visiblePages);

      const active = allPages.find((p) => p.isActive);
      if (active) setActivePageName(active.name);

      console.log("📄 Pages loaded:", visiblePages.map((p) => p.displayName));
    } catch (err) {
      console.error("Failed to load pages:", err);
    }
  }, []);

  // ─── Load exportable visuals on active page ────────────────────────────────
  const loadExportableVisuals = useCallback(async () => {
    try {
      const report = reportRef.current;
      if (!report) return;

      const allPages    = await report.getPages();
      const currentPage = allPages.find((p) => p.isActive);
      if (!currentPage) return;

      setActivePageName(currentPage.name);

      const visuals    = await currentPage.getVisuals();
      const exportable = visuals.filter((v) => EXPORTABLE_TYPES.includes(v.type));

      setAvailableVisuals(exportable);
      setSelectedVisualName(exportable[0]?.name || null);
      console.log("✅ Exportable visuals:", exportable.map((v) => v.title));
    } catch (err) {
      console.error("Failed to load visuals:", err);
    }
  }, []);

  // ─── Navigate to a page when tab is clicked ────────────────────────────────
  const handleTabClick = useCallback(async (page) => {
    try {
      setActivePageName(page.name);
      setAvailableVisuals([]);
      setSelectedVisualName(null);
      setShowExportPanel(false);
      await page.setActive();           // ✅ Power BI API — switches the page
    } catch (err) {
      console.error("Failed to navigate to page:", err);
    }
  }, []);

  // ─── Event handlers ────────────────────────────────────────────────────────
  const eventHandlers = useMemo(() => new Map([
    ["loaded", async () => {
      await loadPages();
      await loadExportableVisuals();
    }],
    ["rendered", async () => {
      await loadExportableVisuals();
      setIsReportReady(true); // ✅ Mark report as ready when rendering completes
    }],
    ["pageChanged", async (event) => {
      const newPageName = event?.detail?.newPage?.name;
      console.log("📄 Page changed:", event?.detail?.newPage?.displayName);
      setActivePageName(newPageName);
      setAvailableVisuals([]);
      setSelectedVisualName(null);
      setShowExportPanel(false);
      await loadExportableVisuals();
    }],
    ["error", (event) => console.error("Embed error:", event.detail)],
  ]), [loadPages, loadExportableVisuals]);

// ─── Trigger file download (F&O iframe safe) ──────────────────────────────
const triggerDownload = (blob, filename) => {
  const url = URL.createObjectURL(blob);

  // ✅ window.open works inside F&O iframe — <a>.click() does not
  const newTab = window.open(url, "_blank");

  if (!newTab) {
    // Fallback if popup is blocked — try parent window
    window.parent?.open(url, "_blank");
  }

  // Revoke after a delay to allow the tab to load the blob
  setTimeout(() => URL.revokeObjectURL(url), 10000);
};

// ── After getting sasUrl from backend (both CSV/XLSX and PDF) ──────────────
const openDownloadWindow = (sasUrl, filename) => {
  // ✅ Open a new window with a simple auto-download HTML page
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Downloading ${filename}...</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
            margin: 0;
            background: #f4f4f4;
            color: #333;
          }
          .card {
            background: white;
            border-radius: 10px;
            padding: 40px 48px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            text-align: center;
            max-width: 480px;
          }
          h2 { margin: 0 0 8px; font-size: 20px; color: #0078d4; }
          p  { margin: 0 0 24px; font-size: 14px; color: #666; }
          a.btn {
            display: inline-block;
            padding: 10px 28px;
            background: #0078d4;
            color: white;
            border-radius: 6px;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
          }
          a.btn:hover { background: #005fa3; }
          .note { margin-top: 16px; font-size: 12px; color: #999; }
        </style>
      </head>
      <body>
        <div class="card">
          <h2>⬇ Your file is ready</h2>
          <p><strong>${filename}</strong></p>
          <a class="btn" href="${sasUrl}" download="${filename}">Download Now</a>
          <p class="note">Link expires in 30 minutes. You can close this tab after downloading.</p>
        </div>
        <script>
          // Auto-click download after 1 second
          setTimeout(() => {
            const a = document.getElementById('auto-dl');
            if (a) a.click();
          }, 1000);
        </script>
        <a id="auto-dl" href="${sasUrl}" download="${filename}" style="display:none"></a>
      </body>
    </html>
  `;

  const newWin = window.open("", "_blank");
  if (newWin) {
    newWin.document.write(html);
    newWin.document.close();
  } else {
    // Fallback if popup blocked — try parent
    const parentWin = window.parent?.open("", "_blank");
    if (parentWin) {
      parentWin.document.write(html);
      parentWin.document.close();
    }
  }
};
// Add this function to get all applied filters
const getActiveFilters = useCallback(async () => {
  try {
    const report = reportRef.current;
    if (!report) return [];

    // Get all pages
    const pages = await report.getPages();
    const activePage = pages.find(p => p.isActive);
    if (!activePage) return [];

    // Get filters from the active page
    const pageFilters = await activePage.getFilters();
    
    // Get visual-level filters from the target visual
    const targetVisual = availableVisuals.find((v) => v.name === selectedVisualName);
    let visualFilters = [];
    if (targetVisual) {
      visualFilters = await targetVisual.getFilters();
    }

    // Combine and format filters
    const allFilters = [...pageFilters, ...visualFilters];
    
    const formattedFilters = allFilters.map(filter => {
      if (filter.$schema === "http://powerbi.com/product/schema#basic") {
        // Basic filter
        return {
          target: filter.target?.table || filter.target?.column || filter.target?.measure,
          operator: filter.operator,
          values: filter.values?.join(', ') || 'All',
          filterType: 'Basic'
        };
      } else if (filter.$schema === "http://powerbi.com/product/schema#advanced") {
        // Advanced filter
        return {
          target: filter.target?.table || filter.target?.column,
          conditions: filter.logicalOperator,
          values: filter.conditions?.map(c => `${c.operator}: ${c.value}`).join('; '),
          filterType: 'Advanced'
        };
      } else if (filter.$schema === "http://powerbi.com/product/schema#slicer") {
        // Slicer
        return {
          target: filter.target?.column,
          values: filter.values?.join(', ') || filter.selectedValues?.map(v => v.value).join(', '),
          filterType: 'Slicer'
        };
      }
      return filter;
    });

    return formattedFilters;
  } catch (err) {
    console.error("Failed to get filters:", err);
    return [];
  }
}, [availableVisuals, selectedVisualName]);
const handleExport = async (format) => {
  setShowExportPanel(false);
  setExporting(true);
  setExportFormat(format);

  try {
    const report = reportRef.current;
    if (!report) throw new Error("Report not ready");

    // ── CSV / XLSX ─────────────────────────────────────────────────────────
    const targetVisual = availableVisuals.find((v) => v.name === selectedVisualName);
    if (!targetVisual) { alert("Please select a visual to export."); return; }

    const result = await targetVisual.exportData(models.ExportDataType.Summarized);
// Get current applied filters
    const activeFilters = await getActiveFilters();
    const allPages    = await report.getPages();
    const currentPage = allPages.find((p) => p.isActive);
    const pageName    = currentPage?.displayName || "export";
    const visualTitle = targetVisual.title || "data";
    const fileName    = `${pageName}-${visualTitle}.${format}`;

    const { data } = await axios.post(`${API_BASE_URL}/reports/exportToAzure`, {
      reportKey,
      format,
      visualData: result.data,
      filters: activeFilters,
    });

    // ✅ Open new window with auto-download page
    openDownloadWindow(data.sasUrl, fileName);

  } catch (err) {
    console.error("❌ Export failed:", err);
    alert(`Export failed: ${err.message}`);
  } finally {
    setExporting(false);
    setExportFormat(null);
  }
};

  // ─── Fetch embed token ─────────────────────────────────────────────────────
  const fetchEmbedToken = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);

      if (!reportKey) { setError("ReportKey missing in URL"); setLoading(false); return; }
      if (!d365User)  { setError("User identity not provided."); setLoading(false); return; }

      const response = await axios.get(`${API_BASE_URL}/reports/getEmbedToken`, {
        params: { d365User, reportKey },
      });

      const { reportId, embedUrl, embedToken } = response.data;
      if (!reportId || !embedUrl || !embedToken)
        throw new Error("Incomplete embed configuration received from server.");

      setEmbedConfig({
        type: "report",
        id: reportId,
        embedUrl: embedUrl,
        accessToken: embedToken,
        tokenType: models.TokenType.Embed,
        settings: {
          panes: {
            filters:        { visible: true, open: false },
            pageNavigation: { visible: true }, // ✅ Hide native nav — we use custom tabs
          },
          background:  models.BackgroundType.Transparent,
          layoutType:  models.LayoutType.Master,
        },
      });

    } catch (err) {
      console.error("❌ Failed to fetch embed token:", err);
      if      (err.response?.status === 401) setError("Session expired or unauthorized.");
      else if (err.response?.status === 403) setError("You do not have permission to view this report.");
      else if (err.response?.status === 500) setError("Server error occurred. Please try again later.");
      else setError(err.response?.data?.error || "Failed to load report.");
    } finally {
      setLoading(false);
    }
  }, [d365User, reportKey]);

  useEffect(() => { fetchEmbedToken(); }, [fetchEmbedToken]);

  if (loading) return (
    <div className="app-container">
      <div className="status-container">
        <div className="spinner" />
        <p className="status-message">Loading report, please wait...</p>
      </div>
    </div>
  );

  if (error) return (
    <div className="app-container">
      <div className="status-container">
        <div className="error-icon">⚠️</div>
        <p className="error-message">{error}</p>
        <button className="retry-button" onClick={fetchEmbedToken}>Retry</button>
      </div>
    </div>
  );

  return (
    <div className="app-container">

      {/* ── Top Export Toolbar ── */}
      {embedConfig && isReportReady && (
        <div className="export-toolbar">
          <button
            className="export-btn"
            onClick={() => setShowExportPanel((prev) => !prev)}
            disabled={exporting}
          >
            {exporting ? `Exporting ${exportFormat?.toUpperCase()}...` : "⬇ Export"}
          </button>

          {showExportPanel && (
            <div className="export-panel">
              <div className="export-panel-row">
                <label>Visual</label>
                <select
                  value={selectedVisualName || ""}
                  onChange={(e) => setSelectedVisualName(e.target.value)}
                >
                  {availableVisuals.length === 0
                    ? <option disabled value="">No exportable visuals</option>
                    : availableVisuals.map((v) => (
                        <option key={v.name} value={v.name}>
                          {v.title || v.name} ({v.type})
                        </option>
                      ))
                  }
                </select>
              </div>
              <div className="export-panel-row">
                <label>Format</label>
                <div className="export-format-btns">
                  <button onClick={() => handleExport("csv")}  disabled={!selectedVisualName}>CSV</button>
                  <button onClick={() => handleExport("xlsx")} disabled={!selectedVisualName}>Excel</button>
                  {/* <button onClick={() => handleExport("pdf")}>PDF</button> */}
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── Power BI Report ── */}
      <div className="powerbi-wrapper">
        {embedConfig && (
          <PowerBIEmbed
            embedConfig={embedConfig}
            eventHandlers={eventHandlers}
            cssClassName="report-container"
            getEmbeddedComponent={(report) => { reportRef.current = report; }}
          />
        )}
      </div>

      {/* ── Custom Bottom Page Tab Bar ── */}
      {/* {pages.length > 0 && (
  <div className="page-tab-container">
    {(
      <button 
        className="tab-scroll-btn tab-scroll-left"
        onClick={() => scrollTabs('left')}
        aria-label="Scroll left"
      >
        ◀
      </button>
    )}
    {(
      <button 
        className="tab-scroll-btn tab-scroll-right"
        onClick={() => scrollTabs('right')}
        aria-label="Scroll right"
      >
        ▶
      </button>
    )}
    
    <div className="page-tab-bar" ref={tabBarRef}>
      {pages.map((page) => (
        <button
          key={page.name}
          className={`page-tab ${activePageName === page.name ? "page-tab--active" : ""}`}
          onClick={() => handleTabClick(page)}
        >
          {page.displayName}
        </button>
      ))}
    </div>
  </div>
)} */}

    </div>
  );
}

export default App;