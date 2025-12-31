/**
 * PROJECT C002259 - PROCUREMENT SYSTEM (PLATINUM V16.0 - UI ENHANCED EDITION)
 *
 * UI & UX UPGRADES (Dec 30, 2025):
 * - "Command Center" Dashboard with Hero Metrics and Visual Status.
 * - Professional Styling (Material Design headers, alternating colors).
 * - "Traffic Light" Control Panel (🟢 RUNNING, ✅ DONE, ⚪ PENDING).
 *
 * STANDARD FEATURES (Audit Certified):
 * - Gemini 3 Pro/Flash auto-routing.
 * - Resumable PDF Extraction.
 * - AppSheet-safe Sync Logic.
 */

const CFG = {
    // --- USER CONFIGURATION ---
    INPUT_FOLDER_ID: '01_Input_Bids',
    PROCESSED_FOLDER_ID: '02_Processed_Bids',
    SKIPPED_FOLDER_ID: '03_Skipped_Files',

    // Optional: Master bids file import
    MASTERBIDS_FILE_ID: '1Q592ZuUQUHSmmjf3fqPuY7TZpL8S6bS2uL0nHMhwCtw',
    MASTERBIDS_FOLDER_ID: '',
    MASTERBIDS_FILENAME_HINT: 'master',
    MASTERBIDS_SHEET_NAME: '',
    MASTERBIDS_HEADER_ROW: 1,
    MASTERBIDS_DATA_START_ROW: 2,
    AUTO_IMPORT_MASTERBIDS_ON_SETUP: false,

    // --- SYSTEM LIMITS ---
    MAX_FILE_SIZE_BYTES: 45 * 1024 * 1024,
    MAX_PHASE_A_SECONDS: 280,
    MAX_PHASE_B_SECONDS: 300,
    MAX_PHASE_C_SECONDS: 320,
    MAX_PHASE_D_SECONDS: 320,
    MAX_PHASE_A_FILES_PER_RUN: 1,
    MAX_PHASE_B_PRODUCTS_PER_RUN: 15,
    MAX_PHASE_C_VENDORS_PER_RUN: 5,

    // --- MODEL STRATEGY ---
    MODEL_CANDIDATES: [
        'gemini-3-pro-preview',
        'gemini-3-flash-preview',
        'gemini-2.5-pro',
        'gemini-2.5-flash',
        'gemini-2.0-flash',
        'gemini-2.0-flash-001'
    ],

    // Phase-specific model order
    MODEL_CANDIDATES_BY_PHASE: {
        EXTRACT: ['gemini-3-pro-preview', 'gemini-3-flash-preview', 'gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
        RESEARCH: ['gemini-3-pro-preview', 'gemini-3-flash-preview', 'gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.0-flash'],
        VENDOR: ['gemini-3-pro-preview', 'gemini-3-flash-preview', 'gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
        RANK: ['gemini-3-pro-preview', 'gemini-3-flash-preview', 'gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash']
    },

    MEDIA_RESOLUTION: 'MEDIA_RESOLUTION_MEDIUM',
    HARD_STOP_MS: 350000,

    PHASE_GEN: {
        EXTRACT: { maxOutputTokens: 8192, temperature: 0.1 },
        RESEARCH: { maxOutputTokens: 8192, temperature: 0.2 },
        VENDOR: { maxOutputTokens: 8192, temperature: 0.2 },
        RANK: { maxOutputTokens: 8192, temperature: 0.4, thinkingBudget: 2048 },
    },
};

// Unified visual theme for sheets and dashboards
const THEME = {
    primary: '#0B4F6C',
    primaryDark: '#062F41',
    accent: '#F9A825',
    surface: '#F4F7FB',
    text: '#1F2933',
    success: '#2E7D32',
    warning: '#F57F17',
    danger: '#C62828',
    neutral: '#ECEFF1'
};

// =====================
// SHEET DEFINITIONS
// =====================
const SHEETS = {
    MASTER: {
        name: 'Master_Items',
        headers: ['Spec ID', 'Category', 'Description', 'Target Qty', 'Target Budget', 'Manufacturer', 'Model', 'Finish', 'Room Type', 'Notes']
    },
    VENDOR_SNAPSHOT: {
        name: 'Vendor_Snapshot',
        headers: ['Bid_ID', 'Vendor Name', 'Total Bid', 'Lead Time', 'Install Scope', 'Delivery Method', 'Sidewalk Feasible', 'Pkg Removal', 'Warranty', 'Support', 'Exclusions', 'Compliance', 'PDF Link', 'Timestamp']
    },
    VENDOR_ANALYSIS: {
        name: 'Vendor_Analysis',
        headers: ['Bid_ID', 'Timestamp', 'Vendor Name', 'Total Cost', 'Services Sum', 'Lead Time', 'Warranty', 'Warranty Rating', 'Delivery Plan', 'Exclusions', 'Compliance', 'Requires Dock', 'Receiving Fit', 'Scan Status', 'Truncated?', 'Evidence', 'File Link',
            'Lead_Time_Weeks', 'InstallPlan_Quality', 'InstallPlan_Summary', 'Install_SiteProtection', 'Install_GC_Coordination', 'Install_OSHA', 'Install_CertifiedInstallers', 'Install_PhasedApproach', 'Install_CleanupPlan',
            'Warranty_Years', 'Warranty_Scope', 'Warranty_Text',
            'DeliveryFit_RFIs', 'DeliveryFit_Risks', 'DeliveryFit_Positives']
    },
    LINE_ITEMS: {
        name: 'Line_Items_Raw',
        headers: ['Raw_ID', 'Bid_ID', 'Vendor Name', 'Spec ID (Extracted)', 'Vendor Description', 'Qty', 'Unit Price', 'Extended Price', 'Manufacturer', 'Model', 'Finish ', 'Is Alternate', 'Notes', 'Product_Key', 'Conf', 'RFI']
    },
    PRODUCTS: {
        name: 'Products',
        headers: ['Product_Key', 'Manufacturer', 'Model', 'Finish', 'Description', 'First Vendor', 'Research Status', 'Last Updated',
            'HC_Score', 'HC_RFIs', 'HC_Risks', 'HC_Positives']
    },
    PRODUCT_RESEARCH: {
        name: 'Product_Research',
        headers: ['Product_Key', 'Manufacturer', 'Model', 'Finish', 'Confidence', 'Durability', 'Support', 'Construction', 'Cleanability', 'Sentiment', 'Mkt Low', 'Mkt High', 'Failure Themes', 'Support Themes', 'Construction Notes', 'Review Summary', 'Sources', 'Last Researched',
            'Fire_Safety_Standard', 'Weight_Capacity_Lbs', 'Cleanable_Surfaces', 'Disinfectant_Compatible', 'Cleanability_Notes', 'Compliance_Notes']
    },
    DASHBOARD: {
        name: '\uD83D\uDCCA Command Center' // UI Upgrade: Enhanced Name
    },
    NOTIFICATIONS: {
        name: '\uD83D\uDD14 Notifications',
        headers: ['Timestamp', 'Type', 'Title', 'Message', 'Action Needed', 'Link']
    },
    RFIS: {
        name: 'RFIs',
        headers: ['Timestamp', 'Vendor', 'Scope', 'Item', 'Severity', 'Question', 'SourceSheet', 'Link']
    },
    RESEARCH_CLAIMS: {
        name: 'Research_Claims',
        headers: ['Timestamp', 'Claim_Label', 'Facet', 'Bid_ID', 'Vendor', 'Product_Key', 'Summary', 'Evidence', 'Sources', 'Why_Not_Scored', 'Linked_RFI']
    },
    VENDOR_RESEARCH: {
        name: 'Vendor_Research',
        headers: ['Vendor Name', 'Confidence', 'Execution', 'Install Cap', 'Healthcare Exp', 'Reputation', 'Red Flags', 'Service Rep', 'Sources', 'Last Researched', 'Supply Chain Score']
    },
    PRICING_RECONCILIATION: {
        name: 'Pricing_Reconciliation',
        headers: ['Bid_ID', 'Vendor Name', 'PDF Declared Total', 'Sum of Line Items', 'Delta (Lines-Declared)', 'Status', 'Notes', 'PDF Link', 'Timestamp']
    },
    RANKING: {
        name: 'Vendor_Ranking',
        headers: ['Rank', 'Vendor Name', 'Bid_ID', 'Overall Score', 'Confidence', 'Quality', 'Deal', 'Execution', 'Price', 'Delivery Fit', 'Quality Explain', 'Deal Explain', 'Exec Explain', 'Price Explain', 'Delivery Explain', 'Overall Explain', 'Risks', 'Positives', 'Prod Researched %', 'Prod Missing Key %', 'Sources', 'Scope Coverage %', 'Budget Score', 'Market Deal Score', 'Confidence Drivers', 'Budget Notes', 'RFI Flags', 'Missing Specs', 'Missing Product Keys', 'Math Status', 'Junk-Expensive Risk']
    },
    EXEC_SUMMARY: {
        name: 'Executive_Summary',
        headers: ['Section', 'Content']
    },
    SETTINGS: {
        name: 'Scoring_Settings',
        headers: ['Setting', 'Value'],
        defaults: [
            ['Weight_Quality', 0.35],
            ['Weight_Deal', 0.25],
            ['Weight_Execution', 0.25],
            ['Weight_Price', 0.10],
            ['Weight_DeliveryFit', 0.05],
            ['Weight_SupplyChain', 0.10],
            ['Deal_Cap', 120]
        ]
    },
    CONTROL_PANEL: {
        name: 'Control_Panel',
        headers: ['Action', 'Run', 'Last Run', 'Status', 'Message', 'Progress (%)'],
    },
    APP_METRICS: {
        name: 'App_Metrics_Feed',
        headers: ['MetricID', 'Title', 'Value', 'Category', 'Icon', 'Color', 'LastUpdated']
    },
    APP_VIEW_CARDS: {
        name: 'App_View_Cards',
        headers: ['Card_ID', 'Title', 'Subtitle', 'Value', 'Detail', 'Icon', 'Color_Start', 'Color_End', 'CTA', 'Link', 'LastUpdated']
    },
    APP_INSIGHTS: {
        name: 'App_Insights',
        headers: ['Section', 'Title', 'Detail', 'Badge', 'Link', 'LastUpdated']
    },
    COMPARISON_BRIEF: {
        name: 'Comparison_Brief',
        headers: ['Bid_ID', 'Vendor', 'Scorecard', 'Coverage', 'Deal Snapshot', 'Risks', 'Positives', 'RFIs/Open Items', 'Sources', 'Updated']
    },
    LOGS: {
        name: 'Logs',
        headers: ['Timestamp', 'Phase', 'Level', 'Message']
    },
    AUDIT: {
        name: 'Audit_Report',
        headers: ['Timestamp', 'Area', 'Severity', 'Finding', 'Action']
    }
};

// =====================
// MENU
// =====================
function onOpen() {
    SpreadsheetApp.getUi().createMenu('\u2728 Procurement Engine') // UI Upgrade: Sparkle Icon
        .addItem('0) One-Click Setup / Repair', 'setupAll_')
        .addItem('0.5) Import Master Bids', 'importMasterBids_')
        .addItem('0.9) Run System Audit', 'runAudit_')
        .addItem('I Install/Refresh Mobile Trigger', 'refreshControlPanelTrigger_')
        .addSeparator()
        .addItem('1) Phase A: Extract PDFs', 'runPhaseA_')
        .addItem('1.5) Phase A+: Auto-Audit Math', 'runPhase15_')
        .addItem('2) Phase B: Research Products', 'runPhaseB_')
        .addItem('3) Phase C: Research Vendors', 'runPhaseC_')
        .addItem('4) Rebuild Rankings', 'rebuildRankings_')
        .addItem('5) Generate Executive Summary', 'generateExecutiveSummary_')
        .addItem('6) Export Vendor Ranking to PDF', 'exportRankingToPdf_')
        .addItem('7) Generate Comparison Brief', 'generateComparisonBrief_')
        .addSeparator()
        .addItem('Maintenance: Retry ERROR Research', 'retryErrorResearch_')
        .addToUi();
}


// =====================
// UI ENHANCEMENTS
// =====================

/**
 * Applies a professional "Material Design" look to sheets.
 * Dark headers, white text, alternating rows, frozen top row.
 */
function styleSheetPro_(sheet, opts) {
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol < 1) return;

    const theme = opts && opts.theme ? opts.theme : THEME;

    // Header Style
    const headRange = sheet.getRange(1, 1, 1, lastCol);
    headRange.setBackground(theme.primary)
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontFamily('Roboto')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');

    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 45); // Taller header for readability

    // Data Body Style (Alternating Colors)
    if (lastRow > 1) {
        const bodyRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        bodyRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
        bodyRange.setFontFamily('Roboto').setVerticalAlignment('top');
        // Gentle zebra shading for readability
        const banding = bodyRange.getBandings();
        if (!banding || !banding.length) {
            bodyRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
        }
    }

    // Auto-resize judiciously (cap very wide columns)
    if (lastCol <= 50) {
        try { sheet.autoResizeColumns(1, lastCol); } catch (e) { }
    }
}

// Visual score highlighting for scorecard-style sheets
function styleScoreGradients_(sheet, cols) {
    if (!sheet || !cols || !cols.length) return;
    const rules = sheet.getConditionalFormatRules() || [];
    const lastRow = Math.max(2, sheet.getLastRow());
    cols.forEach(col => {
        const range = sheet.getRange(2, col, lastRow - 1, 1);
        rules.push(SpreadsheetApp.newConditionalFormatRule()
            .setGradientMinpointWithValue(THEME.danger, SpreadsheetApp.InterpolationType.NUMBER, '0')
            .setGradientMidpointWithValue('#FFB300', SpreadsheetApp.InterpolationType.NUMBER, '50')
            .setGradientMaxpointWithValue(THEME.success, SpreadsheetApp.InterpolationType.NUMBER, '100')
            .setRanges([range])
            .build());
    });
    sheet.setConditionalFormatRules(rules);
}

// Build an in-app link to a target sheet tab for AppSheet/desktop navigation
function makeSheetLink_(ss, sheetName) {
    try {
        const sh = ss.getSheetByName(sheetName);
        if (!sh) return '';
        return `${ss.getUrl()}#gid=${sh.getSheetId()}`;
    } catch (e) {
        return '';
    }
}

/**
 * Generates the "Command Center" Dashboard with Hero Metrics.
 */
function refreshDashboard_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ensureSheetSimple_(SHEETS.DASHBOARD.name);
    dash.clear().clearContents().clearFormats();

    // 1. Dashboard Title styling
    dash.getRange('A1:H1').merge().setValue('Procurement Command Center')
        .setFontSize(24).setFontWeight('bold').setFontFamily('Roboto')
        .setBackground('#1C313A').setFontColor('#FFFFFF')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.setRowHeight(1, 60);

    // 2. Metrics Calculation
    const count = (n) => ss.getSheetByName(n) ? Math.max(0, ss.getSheetByName(n).getLastRow() - 1) : 0;
    const bids = count(SHEETS.VENDOR_ANALYSIS.name);
    const vendors = count(SHEETS.VENDOR_RESEARCH.name);
    const products = count(SHEETS.PRODUCTS.name);
    const rankings = count(SHEETS.RANKING.name);

    // 3. Render Hero Cards
    const card = (r, c, title, val, color) => {
        // Title
        dash.getRange(r, c, 1, 2).merge().setValue(title)
            .setFontWeight('bold').setBackground('#ECEFF1').setHorizontalAlignment('center').setFontFamily('Roboto');
        // Value
        dash.getRange(r + 1, c, 2, 2).merge().setValue(val)
            .setFontSize(36).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setFontColor(color).setBackground('#FFFFFF').setBorder(true, true, true, true, null, null);
    };

    // Columns: [Spacer] [Card] [Spacer] [Card] ...
    card(3, 2, 'BIDS PROCESSED', bids, '#0D47A1');  // Blue
    card(3, 4, 'PRODUCTS', products, '#E65100');     // Orange
    card(3, 6, 'VENDORS RES.', vendors, '#1B5E20');    // Green

    // 4. System Health Status
    const startRow = 8;
    dash.getRange(startRow, 2).setValue('System Status').setFontWeight('bold').setFontSize(14).setFontFamily('Roboto');

    const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
    let status = '⚪ IDLE';
    let statusColor = '#9E9E9E'; // Grey
    if (cp) {
        const data = cp.getDataRange().getValues();
        if (data.some(r => r[3] === 'RUNNING')) { status = '🟢 SYSTEM RUNNING'; statusColor = '#2E7D32'; } // Green
        else if (data.some(r => String(r[3]).includes('ERROR'))) { status = '🔴 ATTENTION NEEDED'; statusColor = '#C62828'; } // Red
    }
    dash.getRange(startRow, 3).setValue(status).setFontWeight('bold').setFontSize(14);

    // Remove gridlines
    dash.setHiddenGridlines(true);

    // 5. Update Mobile Data Feed (AppSheet)
    const metrics = { bids, products, vendors, rankings, status, statusColor };
    const signals = computeExperienceSignals_(ss, metrics);
    syncAppSheetMetrics_(ss, metrics);
    buildAppViewCards_(ss, metrics, signals);
    buildAppInsights_(ss, signals);
}

// Apply scorecard visuals across ranking-style sheets
function styleRankingSheet_(sheet) {
    if (!sheet) return;
    styleSheetPro_(sheet, { theme: THEME });
    const scoreCols = [4, 6, 7, 8, 9, 10, 23, 24, 25]; // Overall + major sub-scores
    styleScoreGradients_(sheet, scoreCols);

    const rules = sheet.getConditionalFormatRules() || [];
    const lastRow = Math.max(2, sheet.getLastRow());

    // Highlight RFIs and missing data for quick scanning
    const rfiCol = 27; // RFI Flags
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('RFI')
        .setBackground(THEME.warning)
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([sheet.getRange(2, rfiCol, lastRow - 1, 1)])
        .build());

    // Celebrate top-ranked rows visually
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(1)
        .setBackground(THEME.primaryDark)
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([sheet.getRange(2, 1, Math.min(1, lastRow - 1), sheet.getLastColumn())])
        .build());

    sheet.setConditionalFormatRules(rules);
}

function styleComparisonBrief_(sheet) {
    if (!sheet) return;
    styleSheetPro_(sheet, { theme: THEME });
    const rules = sheet.getConditionalFormatRules() || [];
    const lastRow = Math.max(2, sheet.getLastRow());
    // Coverage/Deal snapshot quick colors
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('High')
        .setBackground('#E8F5E9')
        .setFontColor(THEME.success)
        .setRanges([sheet.getRange(2, 4, lastRow - 1, 2)])
        .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('Low')
        .setBackground('#FFF8E1')
        .setFontColor(THEME.warning)
        .setRanges([sheet.getRange(2, 4, lastRow - 1, 2)])
        .build());
    sheet.setConditionalFormatRules(rules);
}

// NEW: Data Feeder for "Stunning" AppSheet Mobile Dashboard (Card View)
function syncAppSheetMetrics_(ss, metrics) {
    const s = ensureSheetSimple_(SHEETS.APP_METRICS.name, SHEETS.APP_METRICS.headers);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();

    const now = new Date();
    const rows = [
        ['M1', 'Bids Processed', metrics.bids, 'Overview', 'https://img.icons8.com/color/48/check-file.png', '#0D47A1', now],
        ['M2', 'Products Scanned', metrics.products, 'Research', 'https://img.icons8.com/color/48/box.png', '#E65100', now],
        ['M3', 'Vendors Vetted', metrics.vendors, 'Research', 'https://img.icons8.com/color/48/truck.png', '#1B5E20', now],
        ['M4', 'Rankings Built', metrics.rankings, 'Output', 'https://img.icons8.com/color/48/trophy.png', '#F9A825', now],
        ['M5', 'System Status', metrics.status, 'System', 'https://img.icons8.com/color/48/server.png', metrics.statusColor, now]
    ];
    s.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function computeExperienceSignals_(ss, metrics) {
    const now = new Date();
    const rfisSheet = ss.getSheetByName(SHEETS.RFIS.name);
    const rfiCount = rfisSheet ? Math.max(0, rfisSheet.getLastRow() - 1) : 0;

    const rankingSheet = ss.getSheetByName(SHEETS.RANKING.name);
    const rankingRows = (rankingSheet && rankingSheet.getLastRow() > 1) ? rankingSheet.getDataRange().getValues().slice(1) : [];
    let topVendor = 'Pending Rankings', topScore = '--', topBid = '';
    let avgResearch = 0, researchCount = 0, avgScope = 0, scopeCount = 0;
    let lowestResearch = null, lowestScope = null, mathIssues = 0, mathStatus = 'Not Run';
    if (rankingRows.length) {
        const first = rankingRows[0];
        topVendor = String(first[1] || topVendor);
        topScore = Number(first[3] || 0).toFixed(1);
        topBid = String(first[2] || '');
    }

    rankingRows.forEach(r => {
        const researchPct = Number(r[18] || 0);
        const scopePct = Number(r[21] || 0);
        if (!isNaN(researchPct)) { avgResearch += researchPct; researchCount++; }
        if (!isNaN(scopePct)) { avgScope += scopePct; scopeCount++; }
        if (lowestResearch === null || researchPct < lowestResearch.value) lowestResearch = { bid: String(r[2] || ''), vendor: String(r[1] || ''), value: researchPct };
        if (lowestScope === null || scopePct < lowestScope.value) lowestScope = { bid: String(r[2] || ''), vendor: String(r[1] || ''), value: scopePct };
    });

    const mathSheet = ss.getSheetByName(SHEETS.PRICING_RECONCILIATION.name);
    if (mathSheet && mathSheet.getLastRow() > 1) {
        const rows = mathSheet.getDataRange().getValues().slice(1);
        mathIssues = rows.filter(r => String(r[5] || '').toUpperCase() !== 'MATCH').length;
        mathStatus = mathIssues ? 'Needs Review' : 'Match';
    }

    return {
        rfiCount,
        topVendor,
        topScore,
        topBid,
        avgResearch: researchCount ? Math.round(avgResearch / researchCount) : 0,
        avgScope: scopeCount ? Math.round(avgScope / scopeCount) : 0,
        lowestResearch,
        lowestScope,
        mathIssues,
        mathStatus,
        updated: now
    };
}

// VISUAL CARDS for AppSheet/App UI consumption (gradient-ready)
function buildAppViewCards_(ss, metrics, signals) {
    const s = ensureSheetSimple_(SHEETS.APP_VIEW_CARDS.name, SHEETS.APP_VIEW_CARDS.headers);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();

    const now = new Date();
    const sig = signals || computeExperienceSignals_(ss, metrics || {});
    const lowResearchDetail = sig.lowestResearch ? `Lowest: ${sig.lowestResearch.vendor || sig.lowestResearch.bid} (${sig.lowestResearch.value || 0}%)` : 'All bids covered';
    const lowScopeDetail = sig.lowestScope ? `Needs specs: ${sig.lowestScope.vendor || sig.lowestScope.bid} (${sig.lowestScope.value || 0}% coverage)` : 'Specs covered';
    const auditDetail = sig.mathIssues ? `${sig.mathIssues} pricing checks need review` : 'Math audit clean';
    const sheetLinks = {
        ranking: makeSheetLink_(ss, SHEETS.RANKING.name),
        rfis: makeSheetLink_(ss, SHEETS.RFIS.name),
        comparison: makeSheetLink_(ss, SHEETS.COMPARISON_BRIEF.name),
        researchClaims: makeSheetLink_(ss, SHEETS.RESEARCH_CLAIMS.name),
        pricing: makeSheetLink_(ss, SHEETS.PRICING_RECONCILIATION.name)
    };

    const cards = [
        ['CARD_TOP', 'Recommended Vendor', sig.topBid || 'Awaiting ranking', `${sig.topScore} / 100`, sig.topVendor, 'https://img.icons8.com/color/96/prize.png', '#2E7D32', '#A5D6A7', 'View Dossier', sheetLinks.comparison || sheetLinks.ranking, now],
        ['CARD_SCOPE', 'Spec Coverage', 'Unique master specs quoted', `${sig.avgScope}% avg`, lowScopeDetail, 'https://img.icons8.com/color/96/task.png', '#0B4F6C', '#90CAF9', 'View Scope Gaps', sheetLinks.comparison || sheetLinks.ranking, now],
        ['CARD_PRODUCTS', 'Research Depth', 'Product research coverage', `${sig.avgResearch}% avg`, lowResearchDetail, 'https://img.icons8.com/color/96/box.png', '#E65100', '#FFCC80', 'Review Products', sheetLinks.researchClaims || sheetLinks.ranking, now],
        ['CARD_VENDORS', 'Vendors Vetted', 'Execution & supply-chain', String(metrics.vendors), 'Confirm confidence + sources', 'https://img.icons8.com/color/96/truck.png', '#1B5E20', '#A5D6A7', 'Review Vendor Research', sheetLinks.ranking, now],
        ['CARD_RFIS', 'Open RFIs', 'Questions waiting on answers', String(sig.rfiCount), 'Close gaps before award', 'https://img.icons8.com/color/96/survey.png', '#C62828', '#EF9A9A', 'Open RFI Log', sheetLinks.rfis, now],
        ['CARD_RANKINGS', 'Audit Health', 'Pricing + ranking readiness', sig.mathStatus, auditDetail, 'https://img.icons8.com/color/96/leaderboard.png', '#6A1B9A', '#CE93D8', 'Run Math Audit', sheetLinks.pricing, now]
    ];

    s.getRange(2, 1, cards.length, cards[0].length).setValues(cards);
    styleSheetPro_(s);
}

function buildAppInsights_(ss, signals) {
    const s = ensureSheetSimple_(SHEETS.APP_INSIGHTS.name, SHEETS.APP_INSIGHTS.headers);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    const sig = signals || computeExperienceSignals_(ss, {});
    const rows = [];

    rows.push(['Scorecard', 'Research coverage', `Average product research completeness is ${sig.avgResearch}%`, sig.lowestResearch ? 'Action' : 'Good', '', sig.updated]);
    rows.push(['Scope', 'Spec coverage', `Average master spec coverage is ${sig.avgScope}%`, sig.lowestScope ? 'Action' : 'Good', '', sig.updated]);
    rows.push(['Integrity', 'Pricing audit', sig.mathIssues ? `${sig.mathIssues} pricing rows need review` : 'Pricing reconciliation clean', sig.mathIssues ? 'Review' : 'Ready', '', sig.updated]);
    rows.push(['RFIs', 'Open RFIs', `${sig.rfiCount} RFIs remain open. Close before award.`, sig.rfiCount ? 'Follow up' : 'Clear', '', sig.updated]);

    if (rows.length) s.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    styleSheetPro_(s, { theme: THEME });
}


// =====================
// CORE ACTIONS
// =====================
function setupAll_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create/Repair Sheets with PRO styling
    Object.keys(SHEETS).forEach(k => {
        const def = SHEETS[k];
        const s = ensureSheet_(ss, def.name, def.headers || []);
        styleSheetPro_(s); // Apply Visual Upgrade
    });

    patchHeaders_(SHEETS.VENDOR_ANALYSIS.name, SHEETS.VENDOR_ANALYSIS.headers);
    patchHeaders_(SHEETS.PRODUCTS.name, SHEETS.PRODUCTS.headers);
    patchHeaders_(SHEETS.PRODUCT_RESEARCH.name, SHEETS.PRODUCT_RESEARCH.headers);

    ensureSheetSimple_(SHEETS.DASHBOARD.name);
    ensureSheetSimple_(SHEETS.NOTIFICATIONS.name, SHEETS.NOTIFICATIONS.headers);
    ensureSheetSimple_(SHEETS.RFIS.name, SHEETS.RFIS.headers);
    ensureSheetSimple_(SHEETS.RESEARCH_CLAIMS.name, SHEETS.RESEARCH_CLAIMS.headers);

    const setSheet = ss.getSheetByName(SHEETS.SETTINGS.name);
    if (setSheet.getLastRow() < 2) setSheet.getRange(2, 1, SHEETS.SETTINGS.defaults.length, 2).setValues(SHEETS.SETTINGS.defaults);

    seedControlPanel_(ss);
    applyControlPanelFormatting_(); // Enhanced Traffic Light
    refreshDashboard_(); // Enhanced Dashboard
    ensureControlPanelTrigger_();
    ensureSheetSimple_(SHEETS.COMPARISON_BRIEF.name, SHEETS.COMPARISON_BRIEF.headers);

    if (CFG.AUTO_IMPORT_MASTERBIDS_ON_SETUP) {
        try { importMasterBids_(); } catch (e) { logToSheet_('MASTER_IMPORT', 'WARN', e.message); }
    }

    notify_('Info', 'V16 Upgrade Complete', 'UI Enhanced + Core Systems Ready. Check the Command Center!');
}

function retryErrorResearch_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const p = ss.getSheetByName(SHEETS.PRODUCTS.name);
    const pData = p.getDataRange().getValues();
    let prodChanged = 0;
    for (let i = 1; i < pData.length; i++) {
        const status = String(pData[i][6] || '').toUpperCase().trim();
        if (status === 'ERROR') { pData[i][6] = 'PENDING'; pData[i][7] = now; prodChanged++; }
    }
    if (prodChanged) p.getRange(1, 1, pData.length, pData[0].length).setValues(pData);

    const vr = ss.getSheetByName(SHEETS.VENDOR_RESEARCH.name);
    const vrData = vr.getDataRange().getValues();
    let vendChanged = 0;
    for (let i = 1; i < vrData.length; i++) {
        const name = vrData[i][0];
        if (!name) continue;
        const flags = String(vrData[i][6] || '');
        if (flags.toUpperCase().indexOf('ERROR:') !== -1 && String(name).indexOf('ERROR_ARCHIVE|') !== 0) {
            vrData[i][0] = 'ERROR_ARCHIVE|' + name + '|' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
            vendChanged++;
        }
    }
    if (vendChanged) vr.getRange(1, 1, vrData.length, vrData[0].length).setValues(vrData);
    ss.toast(`Retry queued: ${prodChanged} products, ${vendChanged} vendors archived.`);
}

// Comparison Brief: user-facing synthesis per bid for quick comprehension
function generateComparisonBrief_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ranking = ss.getSheetByName(SHEETS.RANKING.name);
    const brief = ensureSheetSimple_(SHEETS.COMPARISON_BRIEF.name, SHEETS.COMPARISON_BRIEF.headers);

    if (!ranking || ranking.getLastRow() < 2) {
        uiAlert_('No rankings available. Run "Rebuild Rankings" first.');
        return;
    }

    const data = ranking.getDataRange().getValues();
    const rows = [];
    const now = new Date();

    const fmtNum = (v) => {
        const n = Number(v);
        if (isNaN(n)) return 'N/A';
        return n.toFixed(1);
    };

    for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const bidId = r[2] || '';
        const vendor = r[1] || '';
        if (!bidId && !vendor) continue;

        const scorecard = `Overall ${fmtNum(r[3])} (Conf ${r[4] || 'n/a'}) | Quality ${fmtNum(r[5])} | Deal ${fmtNum(r[6])} | Exec ${fmtNum(r[7])} | Price ${fmtNum(r[8])} | Delivery ${fmtNum(r[9])}`;
        const coverage = `Scope ${fmtNum(r[21])}% | Researched ${fmtNum(r[18])}% | Missing Keys ${fmtNum(r[19])}% | Missing Specs ${r[27] || '0'} | Math ${r[29] || 'n/a'}`;
        const dealSnapshot = `Budget ${fmtNum(r[22])} | Market ${fmtNum(r[23])} | Junk Risk ${r[31] || 'None'} | Notes ${r[25] || '—'}`;

        const risks = r[16] || '—';
        const positives = r[17] || '—';
        const rfiFlags = [r[26], r[27], r[28]].filter(Boolean).join('; ');
        const rfis = rfiFlags || 'No open RFIs logged';
        const sources = r[20] || 'See research tabs';

        rows.push([bidId, vendor, scorecard, coverage, dealSnapshot, risks, positives, rfis, sources, now]);
    }

    if (rows.length) {
        if (brief.getLastRow() > 1) brief.getRange(2, 1, brief.getLastRow() - 1, brief.getLastColumn()).clearContent();
        brief.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        styleComparisonBrief_(brief);
        ss.toast(`Comparison brief generated for ${rows.length} bid(s).`);
    }
}

// Light-weight system audit for data/completeness checks
function runAudit_() {
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const audit = ensureSheetSimple_(SHEETS.AUDIT.name, SHEETS.AUDIT.headers);
        if (audit.getLastRow() > 1) audit.getRange(2, 1, audit.getLastRow() - 1, audit.getLastColumn()).clearContent();

        const now = new Date();
        const rows = [];
        const add = (area, severity, finding, action) => rows.push([now, area, severity, finding, action]);

        const key = getApiKey_();
        add('Config', key ? 'INFO' : 'ERROR', key ? 'Gemini API key present.' : 'Missing Gemini API key.', 'Set Script Property GEMINI_API_KEY.');

        try { resolveFolderId_(CFG.INPUT_FOLDER_ID); add('Config', 'INFO', 'Input folder reachable.', ''); }
        catch (e) { add('Config', 'WARN', 'Input folder not resolvable.', 'Verify CFG.INPUT_FOLDER_ID.'); }

        try { resolveFolderId_(CFG.PROCESSED_FOLDER_ID); add('Config', 'INFO', 'Processed folder reachable.', ''); }
        catch (e) { add('Config', 'WARN', 'Processed folder not resolvable.', 'Verify CFG.PROCESSED_FOLDER_ID.'); }

        // Dependency checks (advanced services and triggers)
        const driveReady = (typeof Drive !== 'undefined') && Drive.Files && typeof Drive.Files.copy === 'function';
        add('Dependencies', driveReady ? 'INFO' : 'WARN', driveReady ? 'Drive advanced service enabled for XLSX import.' : 'Drive advanced service not enabled (required for XLSX Master import).', 'Enable Drive API under Services if XLSX import is needed.');

        try {
            const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'onControlPanelChange_');
            const hasOnChange = triggers.some(t => t.getEventType && t.getEventType() === ScriptApp.EventType.ON_CHANGE);
            add('Dependencies', hasOnChange ? 'INFO' : 'WARN', hasOnChange ? 'Control Panel onChange trigger installed.' : 'Control Panel onChange trigger missing.', hasOnChange ? '' : 'Run refreshControlPanelTrigger_ to reinstall.');
        } catch (e) {
            add('Dependencies', 'WARN', 'Could not inspect project triggers.', 'Check triggers manually in Apps Script > Triggers.');
        }

        auditControlPanelActions_(ss, rows, now);

        const vaSheet = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
        const rSheet = ss.getSheetByName(SHEETS.RANKING.name);
        const liSheet = ss.getSheetByName(SHEETS.LINE_ITEMS.name);
        const pSheet = ss.getSheetByName(SHEETS.PRODUCTS.name);
        const prSheet = ss.getSheetByName(SHEETS.PRODUCT_RESEARCH.name);
        const vResSheet = ss.getSheetByName(SHEETS.VENDOR_RESEARCH.name);
        const masterSheet = ss.getSheetByName(SHEETS.MASTER.name);

        // Sheet readiness: confirm database tables align to UI needs
        auditSheetHeaders_(ss, SHEETS.RANKING, rows, 'UI/Outputs');
        auditSheetHeaders_(ss, SHEETS.COMPARISON_BRIEF, rows, 'UI/Outputs');
        auditSheetHeaders_(ss, SHEETS.RESEARCH_CLAIMS, rows, 'UI/Outputs');
        auditSheetHeaders_(ss, SHEETS.PRODUCT_RESEARCH, rows, 'DB/Core');
        auditSheetHeaders_(ss, SHEETS.VENDOR_RESEARCH, rows, 'DB/Core');

        // Logic audit: weights and scoring drivers
        const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS.name);
        if (settingsSheet && settingsSheet.getLastRow() > 1) {
            const settingsMap = new Map(settingsSheet.getDataRange().getValues().slice(1).map(r => [String(r[0] || '').trim(), Number(r[1]) || 0]));
            const weightKeys = ['Weight_Quality', 'Weight_Deal', 'Weight_Execution', 'Weight_Price', 'Weight_DeliveryFit', 'Weight_SupplyChain'];
            const missingWeights = weightKeys.filter(k => !settingsMap.has(k));
            if (missingWeights.length) add('Settings', 'WARN', `Missing weight(s): ${missingWeights.join(', ')}.`, 'Patch Scoring_Settings sheet or rerun setupAll_.');
            const sumWeights = weightKeys.reduce((sum, k) => sum + (settingsMap.get(k) || 0), 0);
            if (sumWeights < 0.95 || sumWeights > 1.05) {
                add('Settings', 'WARN', `Weights sum to ${sumWeights.toFixed(2)} (expected ~1.00).`, 'Normalize weights to avoid over/under weighting.');
            } else {
                add('Settings', 'INFO', `Weights normalized (total ${sumWeights.toFixed(2)}).`, '');
            }
        }

        const vaData = vaSheet && vaSheet.getLastRow() > 1 ? vaSheet.getDataRange().getValues() : [];
        const rData = rSheet && rSheet.getLastRow() > 1 ? rSheet.getDataRange().getValues() : [];
        const liData = liSheet && liSheet.getLastRow() > 1 ? liSheet.getDataRange().getValues() : [];
        const pData = pSheet && pSheet.getLastRow() > 1 ? pSheet.getDataRange().getValues() : [];
        const prData = prSheet && prSheet.getLastRow() > 1 ? prSheet.getDataRange().getValues() : [];
        const masterSpecs = new Set();
        if (masterSheet && masterSheet.getLastRow() > 1) {
            masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues().forEach(r => {
                const spec = String(r[0] || '').trim();
                if (spec) masterSpecs.add(spec);
            });
        }

        const bidVendorMap = new Map();
        if (vaData.length > 1) {
            vaData.slice(1).forEach(r => {
                if (r[0]) bidVendorMap.set(r[0], r[2]);
            });
        }

        let missingSpec = 0, missingProd = 0;
        let priceMismatches = 0;
        const priceMismatchExamples = [];
        const liVendorConflicts = new Set();
        const prodKeyConflicts = [];
        const productMap = new Map();
        pData.slice(1).forEach(r => {
            if (r[0]) productMap.set(r[0], { mfg: r[1] || '', model: r[2] || '' });
        });
        const bidSpecMap = new Map();

        if (liData.length > 1) {
            const bidSet = new Set();
            const bidLiVendors = new Map();
            const tolerancePct = 0.01;
            for (let i = 1; i < liData.length; i++) {
                const r = liData[i];
                if (!r[3]) missingSpec++;
                if (!r[13]) missingProd++;
                if (r[1]) bidSet.add(r[1]);
                const specId = String(r[3] || '').trim();
                if (specId) {
                    if (!bidSpecMap.has(r[1])) bidSpecMap.set(r[1], new Set());
                    bidSpecMap.get(r[1]).add(specId);
                }

                const bidVendor = bidVendorMap.get(r[1]);
                const liVendor = r[2];
                if (r[1]) {
                    if (!bidLiVendors.has(r[1])) bidLiVendors.set(r[1], new Set());
                    if (liVendor) bidLiVendors.get(r[1]).add(liVendor);
                }
                if (bidVendor && liVendor && bidVendor !== liVendor) liVendorConflicts.add(`${r[1]} (${bidVendor} vs ${liVendor})`);

                const qty = Number(r[5] || 0);
                const unit = Number(r[6] || 0);
                const ext = Number(r[7] || 0);
                const expected = qty * unit;
                const diff = Math.abs(ext - expected);
                if (expected && diff > Math.max(1, expected * tolerancePct)) {
                    priceMismatches++;
                    if (priceMismatchExamples.length < 5) priceMismatchExamples.push(String(r[0] || r[3] || 'row ' + (i + 1)));
                }

                const pk = String(r[13] || '').trim();
                if (pk && productMap.has(pk)) {
                    const prod = productMap.get(pk);
                    const mfg = String(r[8] || '').trim();
                    const model = String(r[9] || '').trim();
                    if ((prod.mfg && mfg && prod.mfg !== mfg) || (prod.model && model && prod.model !== model)) {
                        if (prodKeyConflicts.length < 5) prodKeyConflicts.push(pk);
                    }
                }
            }
            if (missingSpec) add('Line Items', 'WARN', `${missingSpec} line items missing Spec ID.`, 'Generate RFIs for scope alignment.');
            if (missingProd) add('Line Items', 'WARN', `${missingProd} line items missing Product_Key.`, 'Trigger Phase B research/RFIs.');
            if (priceMismatches) add('Line Items', 'WARN', `${priceMismatches} line items have unit*qty ≠ extended price.`, `Review Raw_IDs: ${priceMismatchExamples.join(', ')}`);
            if (prodKeyConflicts.length) add('Line Items', 'WARN', `Product_Key manufacturer/model conflicts detected (${prodKeyConflicts.length}).`, `Align Products sheet with line items (e.g., ${prodKeyConflicts.join(', ')}).`);
            const bidVendorDisagree = [...bidLiVendors.entries()].filter(([, set]) => set.size > 1).map(([bid]) => bid);
            if (bidVendorDisagree.length) add('Line Items', 'WARN', `Bid(s) with multiple vendor names in line items: ${bidVendorDisagree.join(', ')}.`, 'Normalize vendor names per Bid_ID.');
            if (liVendorConflicts.size) add('Line Items', 'WARN', `${liVendorConflicts.size} Bid_ID(s) have vendor name mismatches vs Vendor_Analysis.`, 'Correct vendor labels before scoring.');
            add('Line Items', 'INFO', `Bids represented: ${bidSet.size}.`, '');
            if (masterSpecs.size && bidSpecMap.size) {
                const missingByBid = [];
                const extraByBid = [];
                bidSpecMap.forEach((set, bid) => {
                    if (!bid) return;
                    const missingList = [...masterSpecs].filter(s => !set.has(s));
                    const extraList = [...set].filter(s => !masterSpecs.has(s));
                    if (missingList.length && missingByBid.length < 5) missingByBid.push(`${bid}: ${missingList.slice(0, 5).join(', ')}`);
                    if (extraList.length && extraByBid.length < 5) extraByBid.push(`${bid}: ${extraList.slice(0, 5).join(', ')}`);
                });
                if (missingByBid.length) add('Line Items', 'WARN', `Master specs missing in bid(s): ${missingByBid.length} bids.`, `Examples → ${missingByBid.join(' | ')}`);
                if (extraByBid.length) add('Line Items', 'WARN', `Specs not found in Master for bid(s): ${extraByBid.length} bids.`, `Examples → ${extraByBid.join(' | ')}`);
            }
        }

        if (pSheet && prSheet) {
            const products = new Set(pData.slice(1).map(r => r[0]).filter(Boolean));
            const researched = new Set(prData.slice(1).map(r => r[0]).filter(Boolean));
            const covered = [...products].filter(k => researched.has(k)).length;
            const pct = products.size ? Math.round((covered / products.size) * 100) : 0;
            const severity = pct >= 85 ? 'INFO' : 'WARN';
            add('Product Research', severity, `${covered}/${products.size || 0} product keys researched (${pct}%).`, 'Run Phase B and RFIs for missing keys.');

            if (liData.length > 1) {
                const byBid = new Map();
                liData.slice(1).forEach(r => {
                    const bidId = r[1];
                    const pk = r[13];
                    if (!bidId || !pk) return;
                    if (!byBid.has(bidId)) byBid.set(bidId, new Set());
                    byBid.get(bidId).add(pk);
                });
                byBid.forEach((set, bid) => {
                    const researchedCount = [...set].filter(k => researched.has(k)).length;
                    const pctBid = set.size ? Math.round((researchedCount / set.size) * 100) : 0;
                    if (pctBid < 70) add('Product Research', 'WARN', `Bid ${bid} has low researched coverage (${pctBid}% of ${set.size} keys).`, 'Run Phase B or issue RFIs for missing product keys.');
                });
            }

            const prMismatches = prData.slice(1).filter(r => {
                const key = r[0];
                if (!key || !productMap.has(key)) return false;
                const prod = productMap.get(key);
                return ((prod.mfg && r[1] && prod.mfg !== r[1]) || (prod.model && r[2] && prod.model !== r[2]));
            }).map(r => r[0]);
            if (prMismatches.length) add('Product Research', 'WARN', `${prMismatches.length} research rows conflict with Products manufacturer/model.`, `Align Product_Research with Products (e.g., ${prMismatches.slice(0, 5).join(', ')}).`);
        }

        if (vaSheet && vResSheet) {
            const vendors = new Set(vaData.slice(1).map(r => r[2]).filter(Boolean));
            const researchedVendors = new Set(vResSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(Boolean));
            const covered = [...vendors].filter(v => researchedVendors.has(v)).length;
            const pct = vendors.size ? Math.round((covered / vendors.size) * 100) : 0;
            const severity = pct >= 85 ? 'INFO' : 'WARN';
            add('Vendor Research', severity, `${covered}/${vendors.size || 0} vendors researched (${pct}%).`, 'Run Phase C to close gaps.');
        }

        const recon = ss.getSheetByName(SHEETS.PRICING_RECONCILIATION.name);
        if (recon && recon.getLastRow() > 1) {
            const recRows = recon.getDataRange().getValues().slice(1);
            const mismatches = recRows.filter(r => String(r[5]).toUpperCase() === 'MISMATCH').length;
            const missingTotals = recRows.filter(r => String(r[5]).toUpperCase() === 'MISSING TOTAL').length;
            if (mismatches) add('Pricing', 'WARN', `${mismatches} bids have math mismatches.`, 'Review Pricing_Reconciliation and request clarification.');
            if (missingTotals) add('Pricing', 'WARN', `${missingTotals} bids missing declared totals.`, 'Request updated cover sheet or addendum.');
        }

        // Connectivity: ensure bids flow from Vendor_Analysis to Ranking/Comparison and claims link back
        const bidSetVA = new Set();
        if (vaData.length > 1) {
            vaData.slice(1).forEach(r => { if (r[0]) bidSetVA.add(r[0]); });
        }

        const bidSetRank = new Set();
        if (rData.length > 1) {
            rData.slice(1).forEach(r => { if (r[2]) bidSetRank.add(r[2]); });
        }

        const missingInRanking = [...bidSetVA].filter(b => !bidSetRank.has(b));
        if (missingInRanking.length) {
            add('Ranking Connectivity', 'WARN', `${missingInRanking.length} Bid_IDs present in Vendor_Analysis are missing in Vendor_Ranking.`, 'Run Rebuild Rankings and investigate errors.');
        } else if (bidSetVA.size) {
            add('Ranking Connectivity', 'INFO', 'All Bid_IDs in Vendor_Analysis are represented in Vendor_Ranking.', '');
        }

        const claimsSheet = ss.getSheetByName(SHEETS.RESEARCH_CLAIMS.name);
        if (claimsSheet && claimsSheet.getLastRow() > 1) {
            const claimRows = claimsSheet.getDataRange().getValues().slice(1);
            const missingSources = claimRows.filter(r => !r[8]).length;
            const unlinkedBidClaims = claimRows.filter(r => !r[3]).length;
            const scoringWithoutSource = claimRows.filter(r => String(r[1]).toUpperCase().indexOf('SCORING') >= 0 && !r[8]).length;
            const orphanProductClaims = claimRows.filter(r => r[5] && !productMap.has(String(r[5]).trim())).map(r => r[5]);
            const unknownBidClaims = claimRows.filter(r => r[3] && !bidVendorMap.has(String(r[3]).trim())).map(r => r[3]);
            if (missingSources) add('Research Claims', 'WARN', `${missingSources} claims missing sources.`, 'Re-run research or attach grounding URLs.');
            if (unlinkedBidClaims) add('Research Claims', 'WARN', `${unlinkedBidClaims} claims missing Bid_ID linkage.`, 'Tie claims to Bid_ID so dossiers can surface them.');
            if (scoringWithoutSource) add('Research Claims', 'WARN', `${scoringWithoutSource} SCORING claims lack sources.`, 'Demote to CONTEXT/WATCHLIST or attach evidence.');
            if (orphanProductClaims.length) add('Research Claims', 'WARN', `${orphanProductClaims.length} claims reference unknown Product_Key values.`, `Align claims to Products: ${[...new Set(orphanProductClaims)].slice(0, 5).join(', ')}`);
            if (unknownBidClaims.length) add('Research Claims', 'WARN', `${unknownBidClaims.length} claims reference Bid_IDs not in Vendor_Analysis.`, 'Rebind claims to valid Bid_IDs before scoring.');
        }

        const brief = ss.getSheetByName(SHEETS.COMPARISON_BRIEF.name);
        if (brief && brief.getLastRow() > 1) {
            const briefRows = brief.getDataRange().getValues();
            const updatedDates = briefRows.slice(1).map(r => r[9]).filter(Boolean);
            if (!updatedDates.length) add('Comparison Brief', 'WARN', 'Brief exists but has no timestamps.', 'Run "Generate Comparison Brief" to refresh.');
        }

        // Ranking logic audit: uniqueness and score availability
        if (rSheet && rSheet.getLastRow() > 1) {
            const rankRows = rData.slice(1);
            const bidCounts = new Map();
            rankRows.forEach(r => {
                const bid = r[2];
                if (!bid) return;
                bidCounts.set(bid, (bidCounts.get(bid) || 0) + 1);
            });
            const dupBids = [...bidCounts.entries()].filter(([, c]) => c > 1).map(([b]) => b);
            if (dupBids.length) add('Vendor Ranking', 'WARN', `Duplicate Bid_ID rows detected: ${dupBids.join(', ')}.`, 'Ensure one row per Bid_ID before board review.');

            const missingScores = rankRows.filter(r => !r[3] || !r[5] || !r[6]).length;
            if (missingScores) add('Vendor Ranking', 'WARN', `${missingScores} ranking rows missing overall/quality/deal scores.`, 'Re-run rebuildRankings_ to populate.');

            const rankVendorByBid = new Map();
            rankRows.forEach(r => { if (r[2]) rankVendorByBid.set(r[2], r[1]); });
            const vendorNameMismatches = [];
            bidVendorMap.forEach((vendor, bid) => {
                const rankVendor = rankVendorByBid.get(bid);
                if (rankVendor && vendor && rankVendor !== vendor && vendorNameMismatches.length < 5) vendorNameMismatches.push(`${bid} (${vendor} vs ${rankVendor})`);
            });
            if (vendorNameMismatches.length) add('Vendor Ranking', 'WARN', `Vendor name conflicts between Vendor_Analysis and Vendor_Ranking (${vendorNameMismatches.length}).`, `Standardize names for Bid_IDs: ${vendorNameMismatches.join('; ')}`);
        }

        // RFI coverage: missing specs/product keys but no RFIs raised
        const rfiSheet = ss.getSheetByName(SHEETS.RFIS.name);
        if ((missingSpec || missingProd) && (!rfiSheet || rfiSheet.getLastRow() < 2)) {
            add('RFIs', 'WARN', 'Line-item gaps detected without RFI coverage.', 'Use RFI generator to request missing specs/product keys.');
        }

        if (rows.length) audit.getRange(2, 1, rows.length, audit.getLastColumn()).setValues(rows);
        styleSheetPro_(audit);
        ss.toast(`Audit complete: ${rows.length} findings.`);
    });
}

// Helper to validate sheet headers against the defined schema
function auditSheetHeaders_(ss, sheetDef, rows, area) {
    const add = (severity, finding, action) => rows.push([new Date(), area, severity, finding, action]);
    if (!sheetDef || !sheetDef.name || !sheetDef.headers) return;
    const s = ss.getSheetByName(sheetDef.name);
    if (!s) {
        add('ERROR', `${sheetDef.name} sheet missing.`, `Run setupAll_ to provision ${sheetDef.name}.`);
        return;
    }
    const lastCol = s.getLastColumn();
    if (lastCol === 0) {
        add('WARN', `${sheetDef.name} exists but has no headers.`, 'Run setupAll_ or patchHeaders_ to restore columns.');
        return;
    }
    const current = s.getRange(1, 1, 1, lastCol).getValues()[0];
    const missing = sheetDef.headers.filter(h => !current.includes(h));
    if (missing.length) {
        add('WARN', `${sheetDef.name} missing ${missing.length} header(s): ${missing.join(', ')}.`, 'Run setupAll_ or patchHeaders_ to repair.');
    } else {
        add('INFO', `${sheetDef.name} headers aligned.`, '');
    }

    // Guardrail: flag extra columns beyond the defined schema that can cause AppSheet/UI drift or text bleed.
    if (lastCol > sheetDef.headers.length) {
        const extras = lastCol - sheetDef.headers.length;
        add('WARN', `${sheetDef.name} has ${extras} extra column(s) beyond the defined schema.`, 'Trim unused columns or realign headers to prevent data spilling into adjacent sections.');
    }
}

// Helper to audit control panel actions for readiness and stuck runs
function auditControlPanelActions_(ss, rows, now) {
    const add = (severity, finding, action) => rows.push([now, 'Actions', severity, finding, action]);
    const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
    if (!cp) {
        add('ERROR', 'Control_Panel sheet missing.', 'Run setupAll_ to provision Control_Panel.');
        return;
    }
    if (cp.getLastRow() < 2) {
        add('WARN', 'Control_Panel has no actions seeded.', 'Run setupAll_ to seed standard actions.');
        return;
    }

    const data = cp.getDataRange().getValues();
    const actions = data.slice(1).map(r => ({
        name: String(r[0] || ''),
        last: r[2] instanceof Date ? r[2] : null,
        status: String(r[3] || ''),
        progress: parseFloat(String(r[5] || '').replace(/[^0-9.]/g, '')) || 0
    }));

    const errorActions = actions.filter(a => a.status.toUpperCase().indexOf('ERROR') >= 0);
    if (errorActions.length) {
        add('ERROR', `${errorActions.length} action(s) in ERROR: ${errorActions.map(a => a.name).join('; ')}.`, 'Open Control Panel and re-run after fixing data issues.');
    }

    const runningStale = actions.filter(a => a.status.toUpperCase() === 'RUNNING' && a.last && (now - a.last > 60 * 60 * 1000));
    if (runningStale.length) {
        add('WARN', `${runningStale.length} action(s) running for over 60 minutes: ${runningStale.map(a => a.name).join('; ')}.`, 'Confirm no concurrent locks and restart the action if stuck.');
    }

    const expected = ['0.5) Import Master Bids', '1) Phase A: Extract PDFs', '1.5) Phase A+: Auto-Audit Math', '2) Phase B: Research Products', '3) Phase C: Research Vendors', '4) Rebuild Rankings', '5) Generate Executive Summary', '6) Export Vendor Ranking to PDF'];
    const missingActions = expected.filter(e => !actions.some(a => a.name === e));
    if (missingActions.length) {
        add('WARN', `Control_Panel missing ${missingActions.length} standard action(s): ${missingActions.join(', ')}.`, 'Run setupAll_ to reseed or add rows manually.');
    }

    const li = ss.getSheetByName(SHEETS.LINE_ITEMS.name);
    const liRows = li ? Math.max(0, li.getLastRow() - 1) : 0;
    const va = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
    const vaRows = va ? Math.max(0, va.getLastRow() - 1) : 0;

    const needsMathAudit = liRows > 0 && !actions.some(a => a.name.startsWith('1.5)') && a.last);
    if (needsMathAudit) {
        add('WARN', 'Line items exist but Math Audit has not been run.', 'Run "1.5) Phase A+: Auto-Audit Math" to reconcile declared totals.');
    }

    const needsRanking = vaRows > 0 && !actions.some(a => a.name.startsWith('4)') && a.last);
    if (needsRanking) {
        add('WARN', 'Vendor analysis exists but rankings have not been generated.', 'Run "4) Rebuild Rankings" to create the board-ready view.');
    }

    const needsBrief = ss.getSheetByName(SHEETS.RANKING.name)?.getLastRow() > 1 && !actions.some(a => a.name.startsWith('7)'));
    if (needsBrief) {
        add('INFO', 'Comparison Brief action not present in Control_Panel.', 'Optionally add "7) Generate Comparison Brief" for quick summaries.');
    }
}

// ... [STANDARD IMPORT/PHASE A/B/C/RANKING LOGIC FOLLOWS] ...
// To safely keep the file size manageable for copy-paste, I will include the core logic in squeezed form where appropriate.

function importMasterBids_() {
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cpInfo = startAction_(ss, '0.5) Import Master Bids');
        try {
            validateConfigForMasterImport_();
            const fileId = findMasterBidsFileId_();
            if (!fileId) throw new Error(`Master bids file not found (hint: "${CFG.MASTERBIDS_FILENAME_HINT}"). Check Folder ID.`);
            const sheetId = ensureGoogleSheetFromFile_(fileId);
            const imported = readMasterBidsRows_(pickSourceSheet_(SpreadsheetApp.openById(sheetId)));
            const master = ensureSheet_(ss, SHEETS.MASTER.name, SHEETS.MASTER.headers);
            if (master.getLastRow() > 1) master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).clearContent();
            if (imported.rows.length) master.getRange(2, 1, imported.rows.length, SHEETS.MASTER.headers.length).setValues(imported.rows);
            // Re-apply style
            styleSheetPro_(master);
            if (cpInfo) finishAction_(cpInfo, 'OK', `Imported ${imported.rows.length} rows.`);
        } catch (err) {
            if (cpInfo) finishAction_(cpInfo, 'ERROR', err.message);
            uiAlert_('Import failed:\n' + err.message);
        }
    });
}

function runPhaseA_() {
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const t0 = Date.now();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
        const cpr = cp ? findControlRowByAction_(cp, '1) Phase A: Extract PDFs') : 0;
        if (cpr) updateProgress_(cp, cpr, 0);

        try {
            CFG.INPUT_FOLDER_ID = resolveFolderId_(CFG.INPUT_FOLDER_ID);
            CFG.PROCESSED_FOLDER_ID = resolveFolderId_(CFG.PROCESSED_FOLDER_ID);
            CFG.SKIPPED_FOLDER_ID = resolveFolderId_(CFG.SKIPPED_FOLDER_ID);
        } catch (e) { uiAlert_(e.message); return; }

        const inputFolder = DriveApp.getFolderById(CFG.INPUT_FOLDER_ID);
        const processedFolder = DriveApp.getFolderById(CFG.PROCESSED_FOLDER_ID);
        const files = inputFolder.getFilesByType(MimeType.PDF);

        // FIX: Build dedupe set from saved File Link column (17th col = index 16)
        const vSheet = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
        const seenFileIds = new Set();
        if (vSheet && vSheet.getLastRow() > 1) {
            const vals = vSheet.getRange(2, 17, vSheet.getLastRow() - 1, 1).getValues();
            vals.forEach(([u]) => {
                const m = String(u || '').match(/[-\w]{25,}/);
                if (m) seenFileIds.add(m[0]);
            });
        }

        let processed = 0;

        while (
            files.hasNext() &&
            processed < CFG.MAX_PHASE_A_FILES_PER_RUN &&
            (Date.now() - t0 < CFG.MAX_PHASE_A_SECONDS * 1000)
        ) {
            const file = files.next();

            if (seenFileIds.has(file.getId())) {
                moveFileTo_(file, inputFolder, processedFolder);
                continue;
            }

            if (file.getSize() > CFG.MAX_FILE_SIZE_BYTES) {
                logSkip_(ss, file, 'TOO LARGE');
                continue;
            }

            ss.toast(`Extracting: ${file.getName()}`);
            try {
                const up = uploadResumable_(file);     // { uri, name }
                waitForActive_(up);                   // poll using name/id
                if ((Date.now() + 90000) > (t0 + CFG.HARD_STOP_MS)) break;

                const data = extractBid_(up.uri, file.getMimeType());
                const bidId = makeBidId_((data.vendor_name || file.getName()), file);

                calculateAndSaveLineItems_(ss, bidId, data, file);
                saveExtraction_(ss, bidId, data, file, 'EXTRACTED');

                try { deleteGeminiFile_(up); } catch (e) { }
                moveFileTo_(file, inputFolder, processedFolder);

                processed++;
                if (cpr) updateProgress_(cp, cpr, processed / CFG.MAX_PHASE_A_FILES_PER_RUN);
            } catch (e) {
                logSkip_(ss, file, 'ERROR: ' + e.message);
            }
        }

        refreshDashboard_();
        if (cpr) updateProgress_(cp, cpr, 1);
        if (processed === 0) uiAlert_('No new PDFs processed.');
        else ss.toast('Phase A Complete.');
    });
}
// Helper for Phase A
function extractBid_(uri, mime) {
    const prompt = `Procurement Auditor. Extract bid data.
    RULES: Return ONLY raw JSON. No markdown.
    "requires_loading_dock": true only if explicit.
    "evidence_snippets": quote warranty/exclusions.`;

    // Condensed Schema for brevity (functional equivalent)
    const schema = {
        type: 'object', properties: {
            vendor_name: { type: ['string', 'null'] }, total_grand_sum: { type: ['number', 'null'] }, services_sum: { type: ['number', 'null'] },
            lead_time: { type: ['string', 'null'] }, warranty_rating: { type: ['string', 'null'], enum: ['High', 'Med', 'Low', 'Unknown', null] },
            delivery_plan: { type: ['string', 'null'] }, exclusions: { type: ['string', 'null'] }, healthcare_compliant: { type: ['string', 'null'] },
            requires_loading_dock: { type: ['boolean', 'null'] }, receiving_fit: { type: ['string', 'null'] }, truncated_items: { type: ['boolean', 'null'] },
            evidence_snippets: { type: ['string', 'null'] },
            line_items: {
                type: ['array', 'null'], items: {
                    type: 'object', properties: {
                        spec_id: { type: ['string', 'null'] }, description: { type: ['string', 'null'] }, qty: { type: ['number', 'null'] },
                        unit_price: { type: ['number', 'null'] }, extended_price: { type: ['number', 'null'] }, manufacturer: { type: ['string', 'null'] },
                        model: { type: ['string', 'null'] }, finish: { type: ['string', 'null'] }, is_alternate: { type: ['boolean', 'null'] }, notes: { type: ['string', 'null'] }
                    }
                }
            }
        }
    };
    return callGemini_(prompt, [{ file_data: { mime_type: mime, file_uri: uri } }], schema, false, { phase: 'EXTRACT' });
}

function runPhase15_() { // Math Audit
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
    const cpr = cp ? findControlRowByAction_(cp, '1.5) Phase A+: Auto-Audit Math') : 0;
    if (cpr) updateProgress_(cp, cpr, 0);

    const vSheet = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
    const lSheet = ss.getSheetByName(SHEETS.LINE_ITEMS.name);
    const rSheet = ss.getSheetByName(SHEETS.PRICING_RECONCILIATION.name);

    if (!vSheet || !lSheet || !rSheet) return;
    if (rSheet.getLastRow() > 1) rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 9).clearContent();

    const sumMap = {};
    lSheet.getDataRange().getValues().slice(1).forEach(r => {
        const bidId = r[1], ext = Number(r[7] || 0);
        if (bidId) sumMap[bidId] = (sumMap[bidId] || 0) + ext;
    });

    const rows = [];
    vSheet.getDataRange().getValues().slice(1).forEach(r => {
        const bidId = r[0], vend = r[2], decl = Number(r[3] || 0);
        const calc = sumMap[bidId] || 0;
        const delta = decl ? (calc - decl) : 0;
        let status = 'MATCH', notes = 'Math confirmed.';
        if (!decl) { status = 'MISSING TOTAL'; notes = 'No declared total.'; }
        else if (Math.abs(delta) > Math.max(1, decl * 0.001)) { status = 'MISMATCH'; notes = `Delta ${delta.toFixed(2)}`; }
        rows.push([bidId, vend, decl, calc, delta, status, notes, r[16] || '', new Date()]);
    });

    if (rows.length) rSheet.getRange(2, 1, rows.length, 9).setValues(rows);
    if (cpr) updateProgress_(cp, cpr, 1);
    ss.toast('Math Audit Complete.');
}

function runPhaseB_() { // Product Research
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
        const cpr = cp ? findControlRowByAction_(cp, '2) Phase B: Research Products') : 0;
        if (cpr) updateProgress_(cp, cpr, 0);

        reconcileWithMaster_(ss);

        const p = ss.getSheetByName(SHEETS.PRODUCTS.name);
        const res = ss.getSheetByName(SHEETS.PRODUCT_RESEARCH.name);
        const rfiSheet = ensureSheetSimple_(SHEETS.RFIS.name, SHEETS.RFIS.headers);
        const rfiRows = [];
        const rfiKeys = new Set();

        // Harvest new items
        const lSheet = ss.getSheetByName(SHEETS.LINE_ITEMS.name);
        const pMap = new Map();
        p.getDataRange().getValues().slice(1).forEach(r => pMap.set(String(r[0]), r));

        const newRows = [];
        lSheet.getDataRange().getValues().slice(1).forEach(r => {
            const mfg = r[8], mod = r[9];
            if (mfg && mod) {
                const key = (mfg + ' | ' + mod).toUpperCase().trim();
                if (!pMap.has(key)) {
                    const row = [key, mfg, mod, r[10] || '', r[4] || '', r[2] || '', 'PENDING', new Date()];
                    newRows.push(row);
                    pMap.set(key, row);
                }
            }
        });
        if (newRows.length) p.getRange(p.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);

        // Queue
        const pData = p.getDataRange().getValues();
        const queue = [];
        for (let i = 1; i < pData.length; i++) {
            if (['PENDING', 'PENDING_DEEP'].includes(pData[i][6])) queue.push({ key: pData[i][0], mfg: pData[i][1], mod: pData[i][2], desc: pData[i][4], idx: i + 1, firstVendor: pData[i][5] });
        }

        let cnt = 0;
        for (const item of queue.slice(0, CFG.MAX_PHASE_B_PRODUCTS_PER_RUN)) {
            const idx = findRowIndexByKey_(p, item.key, 1);
            if (idx < 1) continue;

            p.getRange(idx, 7).setValue('PROCESSING');
            try {
                const data = researchProduct_(item.mfg, item.mod, item.desc);
                const sources = (data._grounding_urls && data._grounding_urls.length) ? data._grounding_urls : (data.source_urls || []);
                let confidence = data.confidence;
                if (!sources || !sources.length || String(confidence || '').toUpperCase() === 'UNKNOWN') {
                    confidence = 'Low';
                    const rfiKey = `${item.key}|PRODUCT_RESEARCH|SOURCES`;
                    if (!rfiKeys.has(rfiKey)) {
                        rfiRows.push([new Date(), item.firstVendor || '', 'Product Research', item.key, 'MED', `Provide manufacturer specs/warranty/cleaning documentation for ${item.mfg} ${item.mod}. Marketplace links alone are insufficient.`, SHEETS.PRODUCT_RESEARCH.name, '']);
                        rfiKeys.add(rfiKey);
                    }
                }

                res.appendRow([item.key, item.mfg, item.mod, '',
                confidence, data.durability_score, data.support_score, data.construction_score, data.cleanability_score, data.sentiment_score,
                data.market_price_low, data.market_price_high, data.failure_themes, data.support_themes, data.construction_notes, data.review_summary,
                (sources || []).join(', '), new Date(), data.fire_safety_standard || '', data.weight_capacity_lbs || '', data.cleanable_surfaces || '', data.disinfectant_compatible || '', data.cleanability_notes || '', data.compliance_notes || '']);
                p.getRange(idx, 7).setValue('DONE');
            } catch (e) {
                p.getRange(idx, 7).setValue('ERROR');
            }
            cnt++;
        }

        if (rfiRows.length) rfiSheet.getRange(rfiSheet.getLastRow() + 1, 1, rfiRows.length, rfiRows[0].length).setValues(rfiRows);

        refreshDashboard_();
        if (cpr) updateProgress_(cp, cpr, 1);
        ss.toast(`Researched ${cnt} products.`);
    });
}

function researchProduct_(mfg, mod, desc) {
    const prompt = `Role: Procurement Director for a Behavioral Health Facility in NY.
    Task: Evaluate Product "${mfg} ${mod}" (${desc}).
    Guardrails: Prefer manufacturer spec sheets, warranty, cleaning/disinfectant guidance, and credible third-party reviews. Marketplace-only listings (Amazon/Wayfair) are insufficient evidence unless corroborated. If model/SKU cannot be confirmed or sources are weak, set confidence to "Low" and state that an RFI is needed instead of guessing.
    
    RESEARCH OBJECTIVES:
    1. CLINICAL SAFETY (The "Must Haves"):
       - Is this "High-Abuse/Healthcare Grade"?
       - Are there ligature risks, sharp edges, or tamper points?
       - Is it compliant (GREENGUARD, BIFMA)?
       
    2. USER EXPERIENCE (The "Likeability"):
       - Sentiment: Do users find it comfortable? Is it aesthetically pleasing?
       - Reviews: What are the common complaints vs praises?
       
    3. MAINTENANCE (The "Reality"):
       - Routine: Is it easy to just "wipe down" daily?
       - Deep Clean: Is it fluid-barrier/bleach-cleanable for accidents?
    
    OUTPUT REQUIREMENTS (JSON):
    - durability_score: 0-100 (Construction quality).
    - sentiment_score: 0-100 (Comfort, aesthetics, and general popularity).
    - cleanability_score: 0-100 (Weighted average of wipe-down ease AND clinical sanitization).
    - failure_themes: Specific mechanical or fabric failures.
    - review_summary: "People love the look but hate the firmness," etc.
    - construction_notes: Frame material, weight capacity, parts availability.
    - compliance_notes: Certifications.
    - source_urls: At least two manufacturer/credible URLs; avoid hallucinated or marketplace-only links.
    
    Return ONLY JSON matching the schema.`;

    const schema = {
        type: 'object', required: ['confidence', 'durability_score'], properties: {
            confidence: { type: 'string', enum: ['High', 'Med', 'Low'] },
            durability_score: { type: 'number' },
            support_score: { type: 'number' },
            construction_score: { type: 'number' },
            cleanability_score: { type: 'number' },
            sentiment_score: { type: 'number' },
            market_price_low: { type: 'number' },
            market_price_high: { type: 'number' },
            failure_themes: { type: 'string' },
            support_themes: { type: 'string' },
            construction_notes: { type: 'string' },
            review_summary: { type: 'string' },
            compliance_notes: { type: 'string' },
            fire_safety_standard: { type: ['string', 'null'] },
            weight_capacity_lbs: { type: ['number', 'null'] },
            cleanable_surfaces: { type: ['string', 'null'] },
            disinfectant_compatible: { type: ['string', 'null'] },
            cleanability_notes: { type: ['string', 'null'] },
            source_urls: { type: 'array', items: { type: 'string' } }
        }
    };
    return callGemini_(prompt, [], schema, true, { phase: 'RESEARCH' });
}

function runPhaseC_() { // Vendor Research
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
        const cpr = cp ? findControlRowByAction_(cp, '3) Phase C: Research Vendors') : 0;
        if (cpr) updateProgress_(cp, cpr, 0);

        const vAnal = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
        const vRes = ss.getSheetByName(SHEETS.VENDOR_RESEARCH.name);
        const rfiSheet = ensureSheetSimple_(SHEETS.RFIS.name, SHEETS.RFIS.headers);
        const rfiRows = [];
        const rfiKeys = new Set();

        const vendors = [...new Set(vAnal.getDataRange().getValues().slice(1).map(r => r[2]).filter(Boolean))];
        const done = new Set(vRes.getDataRange().getValues().slice(1).map(r => r[0]));

        const queue = vendors.filter(v => !done.has(v));
        let cnt = 0;

        for (const vend of queue.slice(0, CFG.MAX_PHASE_C_VENDORS_PER_RUN)) {
            let idx = findRowIndexByKey_(vRes, vend, 1);
            if (idx < 1) {
                vRes.appendRow([vend, 'PENDING', 0, '', '', '', '', '', '', new Date(), 0]);
                idx = vRes.getLastRow();
            }

            try {
                const data = researchVendor_(vend);
                const sources = (data._grounding_urls && data._grounding_urls.length) ? data._grounding_urls : (data.source_urls || []);
                let confidence = data.confidence || 'Low';
                if (!sources || !sources.length || String(confidence || '').toUpperCase() === 'UNKNOWN') {
                    confidence = 'Low';
                    const rfiKey = `${vend}|VENDOR_RESEARCH|SOURCES`;
                    if (!rfiKeys.has(rfiKey)) {
                        rfiRows.push([new Date(), vend, 'Vendor Research', vend, 'MED', `Provide verifiable project references, safety record, and logistics capability evidence for ${vend}. Marketplace-only links are insufficient.`, SHEETS.VENDOR_RESEARCH.name, '']);
                        rfiKeys.add(rfiKey);
                    }
                }
                // Overwrite the specific row instead of appending duplicate
                const rowData = [vend, confidence, Number(data.execution_score || 0), data.install_delivery_capability,
                    data.healthcare_experience || data.origin_country || '', data.reputation_summary || (data.origin_country ? `Origin ${data.origin_country}` : ''), data.red_flags, data.service_warranty_reputation || '',
                    (sources || []).join(', '), new Date(), Number(data.supply_chain_score || 0)];
                vRes.getRange(idx, 1, 1, rowData.length).setValues([rowData]);
            } catch (e) {
                vRes.getRange(idx, 1, 1, 11).setValues([[vend, 'Low', 50, '', '', '', 'ERROR: ' + e.message, '', '', new Date(), 0]]);
            }
            cnt++;
        }

        if (rfiRows.length) rfiSheet.getRange(rfiSheet.getLastRow() + 1, 1, rfiRows.length, rfiRows[0].length).setValues(rfiRows);

        refreshDashboard_();
        if (cpr) updateProgress_(cp, cpr, 1);
        ss.toast(`Researched ${cnt} vendors.`);
    });
}

function researchVendor_(name) {
    const prompt = `Role: COO vetting a furniture vendor for Syracuse, NY.
    Target: "${name}".
    Guardrails: Prefer manufacturer site, news, legal filings, and credible industry sources. Marketplace listings or unverifiable claims should not raise confidence. If the record is unclear, set confidence to "Low" and note that an RFI is required for proof of experience.
    
    INVESTIGATE:
    1. OPERATIONAL REALITY:
       - Execution: Track record in Northeast US? White Glove vs Tailgate?
       - Financials: Stability/Bankruptcies?
    2. SUPPLY CHAIN (Geopolitical Risk):
       - Manufacturing Origin: USA/North America vs Overseas?
       - Risk: Exposure to shipping delays (Red Sea/Tariffs)?
    3. SERVICE:
       - Local Representation in Upstate NY?
    
    OUTPUT REQUIREMENTS (JSON):
    - execution_score: 0-100 (Logistics + Support).
    - supply_chain_score: 0-100 (100 = Domestic/Secure, 0 = Volatile Import).
    - origin_country: "USA", "China", "Canada", etc.
    - red_flags: Specific operational risks.
    - install_delivery_capability: "Union Labor", "Non-Union", "Drop Ship".
    - source_urls: list of credible support; avoid hallucinated URLs.
    
    Return ONLY JSON matching schema.`;

    const schema = {
        type: 'object', properties: {
            confidence: { type: 'string', enum: ['High', 'Med', 'Low'] },
            execution_score: { type: 'number' },
            supply_chain_score: { type: 'number' }, // NEW FIELD
            origin_country: { type: 'string' },       // NEW FIELD
            install_delivery_capability: { type: 'string' },
            healthcare_experience: { type: 'string' },
            red_flags: { type: 'string' },
            source_urls: { type: 'array', items: { type: 'string' } }
        }
    };
    return callGemini_(prompt, [], schema, true, { phase: 'VENDOR' });
}

function buildResearchClaims_(ss, prodRes, vRes, productBidMap, bidVendorMap) {
    const sheet = ensureSheetSimple_(SHEETS.RESEARCH_CLAIMS.name, SHEETS.RESEARCH_CLAIMS.headers);
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, SHEETS.RESEARCH_CLAIMS.headers.length).clearContent();

    const rows = [];
    const now = new Date();

    prodRes.forEach(r => {
        const key = r[0];
        if (!key) return;
        const bidIds = productBidMap && productBidMap.has(key) ? Array.from(productBidMap.get(key)) : [];
        const vendors = bidIds.map(b => bidVendorMap && bidVendorMap[b] ? bidVendorMap[b] : '').filter(Boolean);
        const confidence = String(r[4] || '').toUpperCase();
        const sources = String(r[16] || '').split(/\s*,\s*/).filter(Boolean);
        const label = (['HIGH', 'MED'].includes(confidence) && sources.length) ? 'SCORING' : (sources.length ? 'CONTEXT' : 'WATCHLIST');
        const why = label === 'SCORING' ? '' : (sources.length ? 'Confidence limited; informational only' : 'No defensible sources attached');
        const linkedRfi = label === 'SCORING' ? '' : 'Sources RFI recommended';
        const summary = `Quality ${((Number(r[5]) + Number(r[6]) + Number(r[7]) + Number(r[8]) + Number(r[9])) / 5 || 0).toFixed(1)} | ${r[1]} ${r[2]}`;
        const evidence = truncate_(String(r[15] || r[14] || ''), 300);
        rows.push([now, label, 'Product Quality', bidIds.join('; '), vendors.join('; '), key, summary, evidence, sources.join(', '), why, linkedRfi]);
    });

    vRes.forEach(r => {
        const vendor = r[0];
        if (!vendor) return;
        const bidIds = productBidMap && vendor ? Array.from(new Set(Array.from(productBidMap.values()).flatMap(set => Array.from(set)).filter(b => bidVendorMap && bidVendorMap[b] === vendor))) : [];
        const confidence = String(r[1] || '').toUpperCase();
        const sources = String(r[8] || '').split(/\s*,\s*/).filter(Boolean);
        const label = (['HIGH', 'MED'].includes(confidence) && sources.length) ? 'SCORING' : (sources.length ? 'CONTEXT' : 'WATCHLIST');
        const why = label === 'SCORING' ? '' : (sources.length ? 'Confidence limited; informational only' : 'No defensible sources attached');
        const linkedRfi = label === 'SCORING' ? '' : 'Sources RFI recommended';
        const summary = `Execution ${r[2] || ''}, Supply ${r[10] || ''}`;
        const evidence = truncate_(String(r[6] || r[5] || ''), 300);
        rows.push([now, label, 'Vendor Capability', bidIds.join('; '), vendor, '', summary, evidence, sources.join(', '), why, linkedRfi]);
    });

    if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function rebuildRankings_() {
    return withScriptLock_('GLOBAL_PHASE_LOCK', () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
        const cpr = cp ? findControlRowByAction_(cp, '4) Rebuild Rankings') : 0;
        if (cpr) updateProgress_(cp, cpr, 0.1);

        patchHeaders_(SHEETS.RANKING.name, SHEETS.RANKING.headers);

        // 1. Load Weights from Settings Sheet (Platinum Feature)
        const setVals = ss.getSheetByName(SHEETS.SETTINGS.name).getDataRange().getValues();
        const W = {};
        setVals.forEach(r => {
            const key = String(r[0] || '');
            if (key.startsWith('Weight_')) W[key] = Number(r[1]) || 0;
            if (key === 'Deal_Cap') W.Deal_Cap = Number(r[1]) || 120;
        });

        const lData = ss.getSheetByName(SHEETS.LINE_ITEMS.name).getDataRange().getValues().slice(1);
        const vRes = ss.getSheetByName(SHEETS.VENDOR_RESEARCH.name).getDataRange().getValues().slice(1);
        const prodRes = ss.getSheetByName(SHEETS.PRODUCT_RESEARCH.name).getDataRange().getValues().slice(1);
        const vAnal = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name).getDataRange().getValues().slice(1);
        const master = ss.getSheetByName(SHEETS.MASTER.name);
        const pricingRecon = ss.getSheetByName(SHEETS.PRICING_RECONCILIATION.name);
        const rfiSheet = ensureSheetSimple_(SHEETS.RFIS.name, SHEETS.RFIS.headers);

        const masterSpecs = new Set();
        const masterBudgetMap = new Map();
        if (master && master.getLastRow() > 1) {
            const mVals = master.getDataRange().getValues().slice(1);
            mVals.forEach(r => {
                const spec = String(r[0] || '').trim();
                if (spec) {
                    masterSpecs.add(spec);
                    const budget = Number(r[4] || 0);
                    if (budget > 0) masterBudgetMap.set(spec, budget);
                }
            });
        }

        const deliveryMap = {};
        const bidVendorMap = {};
        vAnal.forEach(r => {
            let score = 70;
            const reqDock = String(r[11] || "").toUpperCase();
            const recFit = String(r[12] || "").toUpperCase();
            if (reqDock.includes("YES")) score -= 15;
            if (recFit.includes("GOOD") || recFit.includes("OK")) score += 10;
            else if (recFit.includes("POOR") || recFit.includes("TIGHT")) score -= 20;
            deliveryMap[r[0]] = Math.max(0, Math.min(100, score));
            bidVendorMap[r[0]] = r[2];
        });

        const productBidMap = new Map();
        lData.forEach(r => {
            const bidId = r[1];
            const vendor = r[2];
            const key = String(r[13] || '').toUpperCase();
            if (bidId && vendor && !bidVendorMap[bidId]) bidVendorMap[bidId] = vendor;
            if (key && bidId) {
                if (!productBidMap.has(key)) productBidMap.set(key, new Set());
                productBidMap.get(key).add(bidId);
            }
        });
        buildResearchClaims_(ss, prodRes, vRes, productBidMap, bidVendorMap);

        const vMap = {};
        vRes.forEach(r => {
            const supplyScore = (r.length > 10 && Number(r[10])) ? Number(r[10]) : 50;
            vMap[r[0]] = { exec: Number(r[2] || 50), conf: r[1], flags: r[6], supply: supplyScore, sources: String(r[8] || '').split(/\s*,\s*/).filter(Boolean) };
        });

        const pMap = {};
        prodRes.forEach(r => {
            const q = (Number(r[5]) + Number(r[6]) + Number(r[7]) + Number(r[8]) + Number(r[9])) / 5;
            const sources = String(r[16] || '').split(/\s*,\s*/).filter(Boolean);
            const low = Number(r[10] || 0), high = Number(r[11] || 0);
            pMap[r[0]] = { q: q || 50, low, high, sources };
        });

        const keyBenchmarks = {};
        lData.forEach(r => {
            const key = String(r[13] || '').toUpperCase();
            const unit = Number(r[6] || 0);
            if (key && unit > 0) {
                if (!keyBenchmarks[key]) keyBenchmarks[key] = [];
                keyBenchmarks[key].push(unit);
            }
        });

        const reconMap = {};
        if (pricingRecon && pricingRecon.getLastRow() > 1) {
            pricingRecon.getDataRange().getValues().slice(1).forEach(r => reconMap[r[0]] = { status: r[5], delta: r[4] });
        }

        const bids = new Map();
        lData.forEach(r => {
            const bidId = r[1];
            if (!bidId) return;
            if (!bids.has(bidId)) bids.set(bidId, []);
            bids.get(bidId).push(r);
        });

        const rfiRows = [];
        const rfiKeys = new Set();
        const benchmarkRfiKeys = new Set();
        const dealCap = Number(W.Deal_Cap || 120) || 120;
        const bidMetrics = [];

        bids.forEach((rows, bidId) => {
            const vendor = bidVendorMap[bidId] || (rows[0] ? rows[0][2] : '');
            let totalExt = 0;
            let qualitySum = 0, qualityWeight = 0;
            let budgetScoreSum = 0, budgetWeight = 0, marketScoreSum = 0, marketWeight = 0, budgetUnknown = 0;
            let dealScoreSum = 0, dealWeight = 0;
            const productKeys = new Set();
            const researchedKeys = new Set();
            const quotedSpecs = new Set();
            let missingProductKeys = 0;
            let junkRiskSpend = 0, spendTracked = 0;
            let dealBenchmarkSource = 'Unknown';

            rows.forEach(r => {
                const specId = String(r[3] || '').trim();
                const qty = Number(r[5] || 0);
                const unit = Number(r[6] || 0);
                const ext = Number(r[7] || 0);
                const key = String(r[13] || '').toUpperCase();
                if (specId) quotedSpecs.add(specId);
                if (key) productKeys.add(key); else missingProductKeys++;
                totalExt += ext;

                const prod = pMap[key];
                if (prod) {
                    researchedKeys.add(key);
                    qualitySum += prod.q * ext;
                    qualityWeight += ext;
                    const mid = (prod.low + prod.high) / 2;
                    if (!masterBudgetMap.has(specId) && mid > 0 && unit > 0) {
                        const mScore = Math.min(dealCap, (mid / unit) * 100);
                        marketScoreSum += mScore * ext;
                        marketWeight += ext;
                    }
                }

                const targetBudget = masterBudgetMap.get(specId);
                if (targetBudget > 0 && unit > 0) {
                    const bScore = Math.min(dealCap, (targetBudget / unit) * 100);
                    budgetScoreSum += bScore * ext;
                    budgetWeight += ext;
                } else if (!prod || !prod.low) {
                    budgetUnknown++;
                }

                const keyBenchmark = (key && keyBenchmarks[key] && keyBenchmarks[key].length)
                    ? median_(keyBenchmarks[key])
                    : null;
                let benchmark = keyBenchmark;
                if (!benchmark && prod && prod.low > 0 && prod.high > 0 && key) {
                    benchmark = (prod.low + prod.high) / 2;
                    dealBenchmarkSource = 'Market';
                } else if (benchmark) {
                    dealBenchmarkSource = 'Cross-bid median';
                }

                if (benchmark && unit > 0) {
                    const qualityFactor = (prod?.q || 50) / 100;
                    const valueIndex = (benchmark / unit) * qualityFactor;
                    const valueScore = Math.max(0, Math.min(dealCap, valueIndex * 100));
                    dealScoreSum += valueScore * ext;
                    dealWeight += ext;
                } else if (key && !benchmarkRfiKeys.has(key)) {
                    rfiRows.push([new Date(), vendor, 'Deal Benchmark', key, 'MED', `Provide defensible benchmark pricing or contract references for Product ${key} so value can be scored. Marketplace-only links are insufficient.`, SHEETS.LINE_ITEMS.name, '']);
                    benchmarkRfiKeys.add(key);
                }

                if (ext > 0) {
                    spendTracked += ext;
                    const midPrice = prod ? ((prod.low + prod.high) / 2) : 0;
                    if ((prod?.q || 50) < 60 && ((targetBudget && unit > targetBudget) || (!targetBudget && midPrice > 0 && unit > midPrice))) {
                        junkRiskSpend += ext;
                    }
                }
            });

            const qualityScore = qualityWeight ? (qualitySum / qualityWeight) : 50;
            let budgetScore = null;
            if (budgetWeight) {
                budgetScore = budgetScoreSum / budgetWeight;
            } else if (marketWeight) {
                budgetScore = marketScoreSum / marketWeight;
            }

            const marketDealScore = marketWeight ? (marketScoreSum / marketWeight) : '';
            const qualityAdjustedDeal = dealWeight ? (dealScoreSum / dealWeight) : null;
            const scopeCoverage = masterSpecs.size ? (quotedSpecs.size / masterSpecs.size) : 1;
            const prodResPct = productKeys.size ? (researchedKeys.size / productKeys.size) : 0;
            const missingKeyPct = rows.length ? (missingProductKeys / rows.length) : 0;
            const missingSpecs = [...masterSpecs].filter(s => !quotedSpecs.has(s));

            missingSpecs.forEach(ms => {
                const rfiKey = `${bidId}|SCOPE|${ms}`;
                if (!rfiKeys.has(rfiKey)) {
                    rfiRows.push([new Date(), vendor, 'Scope Gap', ms, 'HIGH', `Bid ${bidId} did not include Master Spec ${ms}. Please confirm quote or provide alternate.`, SHEETS.LINE_ITEMS.name, '']);
                    rfiKeys.add(rfiKey);
                }
            });
            if (missingProductKeys > 0) {
                const rfiKey = `${bidId}|PRODUCT_KEY|MISSING`;
                if (!rfiKeys.has(rfiKey)) {
                    rfiRows.push([new Date(), vendor, 'Product Key', bidId, 'MED', `Bid ${bidId} has ${missingProductKeys} line(s) without Product_Key / model. Please specify manufacturer + model.`, SHEETS.LINE_ITEMS.name, '']);
                    rfiKeys.add(rfiKey);
                }
            }

            const recon = reconMap[bidId] || {};
            const priceConfidencePenalty = recon.status === 'MISMATCH' ? 15 : (recon.status === 'MISSING TOTAL' ? 10 : 0);

            let confidence = 85;
            confidence -= (1 - scopeCoverage) * 25;
            confidence -= missingKeyPct * 30;
            confidence -= (1 - prodResPct) * 20;
            confidence -= priceConfidencePenalty;
            if (!budgetWeight && !marketWeight) confidence -= 10;
            if (qualityAdjustedDeal === null) confidence -= 8;
            confidence = Math.max(5, Math.min(100, confidence));

            const rfiFlags = [];
            if (missingSpecs.length) rfiFlags.push('Missing Specs');
            if (missingProductKeys) rfiFlags.push('Missing Product Keys');
            if (priceConfidencePenalty) rfiFlags.push('Pricing Audit');
            if (qualityAdjustedDeal === null) rfiFlags.push('Deal Benchmark Missing');

            const junkRiskPct = spendTracked ? (junkRiskSpend / spendTracked) * 100 : 0;

            const sources = new Set();
            (vMap[vendor]?.sources || []).forEach(s => sources.add(s));
            [...researchedKeys].slice(0, 5).forEach(k => {
                (pMap[k]?.sources || []).forEach(s => sources.add(s));
            });

            const deliveryScore = Math.max(0, Math.min(100, ((deliveryMap[bidId] || 60) * 0.6) + ((vMap[vendor]?.supply || 50) * 0.4)));

            bidMetrics.push({
                bidId,
                vendor,
                totalExt,
                qualityScore,
                budgetScore,
                marketDealScore,
                scopeCoverage,
                prodResPct,
                missingKeyPct,
                confidence,
                rfiFlags,
                missingSpecs: missingSpecs.join(', '),
                missingProductKeys,
                recon,
                junkRiskPct,
                delivery: deliveryScore,
                exec: vMap[vendor]?.exec ?? 50,
                supply: vMap[vendor]?.supply ?? 50,
                confLabel: vMap[vendor]?.conf || 'Low',
                flags: vMap[vendor]?.flags || '',
                sources: Array.from(sources).slice(0, 5),
                shadowPrice: totalExt * (1 + Math.max(0, (80 - qualityScore) / 100)),
                qualityAdjustedDeal,
                dealBenchmarkSource
            });
        });

            const priceSpace = bidMetrics.map(b => b.shadowPrice).filter(v => v > 0);
            const minShadow = priceSpace.length ? Math.min.apply(null, priceSpace) : 0;
            const maxShadow = priceSpace.length ? Math.max.apply(null, priceSpace) : 0;

        const weights = {
            quality: W.Weight_Quality || 0.30,
            exec: W.Weight_Execution || 0.20,
            price: W.Weight_Price || 0.10,
            deal: W.Weight_Deal || 0.25,
            supply: W.Weight_SupplyChain || 0.10,
            delivery: W.Weight_DeliveryFit || 0.05
        };
        const totalWeight = Object.values(weights).reduce((a, b) => a + b, 0) || 1;

        const output = bidMetrics.map((m) => {
            const priceScoreRaw = priceSpace.length ? ((priceSpace.length === 1) ? 100 : 100 * ((maxShadow - m.shadowPrice) / (maxShadow - minShadow || 1))) : 50;
            const priceScore = Math.max(0, Math.min(100, priceScoreRaw));
            const budgetScore = (m.budgetScore === null || m.budgetScore === undefined) ? null : m.budgetScore;
            const dealScore = m.qualityAdjustedDeal != null ? m.qualityAdjustedDeal : (budgetScore != null ? budgetScore : (m.marketDealScore || ''));
            const budgetForOverall = dealScore != null ? dealScore : priceScore;
            const overall = ((m.qualityScore * weights.quality) +
                ((m.exec) * weights.exec) +
                (priceScore * weights.price) +
                ((budgetForOverall) * weights.deal) +
                ((m.supply || 50) * weights.supply) +
                ((m.delivery || 60) * weights.delivery)) / totalWeight;

            const confidenceDrivers = [`Scope ${(m.scopeCoverage * 100).toFixed(0)}%`, `Research ${(m.prodResPct * 100).toFixed(0)}%`, `Missing Keys ${(m.missingKeyPct * 100).toFixed(0)}%`];
            if (m.recon?.status) confidenceDrivers.push(`Pricing ${m.recon.status}`);

            const budgetNote = dealScore != null ? '' : 'Deal benchmark unknown';
            const rfiFlagText = m.rfiFlags.join('; ');
            const risks = [m.flags || '', rfiFlagText].filter(Boolean).join('; ');
            const positives = [`Scope ${(m.scopeCoverage * 100).toFixed(0)}%`, `Research ${(m.prodResPct * 100).toFixed(0)}%`].join('; ');
            const qualityExplain = `Weighted product quality ${m.qualityScore.toFixed(1)}`;
            const dealExplain = m.qualityAdjustedDeal != null ? `${m.dealBenchmarkSource} value index (quality-adjusted)` : (budgetScore != null ? 'Budget vs target' : (m.marketDealScore ? 'Market price benchmark' : 'Deal not scored'));
            const execExplain = `Execution score ${m.exec.toFixed(1)}`;
            const priceExplain = `Risk-adjusted shadow cost $${m.shadowPrice.toFixed(0)}`;
            const deliveryExplain = `Delivery fit ${m.delivery.toFixed(1)}`;
            const overallExplain = 'Composite weighted by scoring settings';

            return [
                0,
                m.vendor,
                m.bidId,
                overall,
                m.confLabel || 'Low',
                m.qualityScore,
                dealScore,
                m.exec,
                priceScore,
                m.delivery,
                qualityExplain,
                dealExplain,
                execExplain,
                priceExplain,
                deliveryExplain,
                overallExplain,
                risks,
                positives,
                (m.prodResPct * 100),
                (m.missingKeyPct * 100),
                m.sources.join(', '),
                (m.scopeCoverage * 100),
                budgetScore == null ? '' : budgetScore,
                m.marketDealScore || '',
                confidenceDrivers.join('; '),
                budgetNote,
                rfiFlagText,
                m.missingSpecs,
                m.missingProductKeys,
                m.recon?.status || '',
                `${m.junkRiskPct.toFixed(1)}% of spend at risk`
            ];
        });

        output.sort((a, b) => b[3] - a[3]);
        output.forEach((r, i) => r[0] = i + 1);

        const rSheet = ss.getSheetByName(SHEETS.RANKING.name);
        if (rSheet.getLastRow() > 1) rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).clearContent();
        if (output.length) rSheet.getRange(2, 1, output.length, output[0].length).setValues(output);

        if (rfiRows.length) rfiSheet.getRange(rfiSheet.getLastRow() + 1, 1, rfiRows.length, rfiRows[0].length).setValues(rfiRows);

        styleRankingSheet_(rSheet);
        if (cpr) updateProgress_(cp, cpr, 1);
        ss.toast('Rankings Rebuilt Successfully.');
    });
}


// =====================
// UTILS
// =====================
function ensureSheet_(ss, name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (headers && headers.length) {
        if (s.getLastRow() === 0) s.getRange(1, 1, 1, headers.length).setValues([headers]);
        else if (s.getLastRow() < 2) s.getRange(1, 1, 1, headers.length).setValues([headers]); // Safe patch
    }
    return s;
}
function ensureSheetSimple_(name, headers) { return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), name, headers); }
function logToSheet_(ph, lvl, msg) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const logSheet = ss.getSheetByName(SHEETS.LOGS.name);
        if (logSheet) logSheet.appendRow([new Date(), ph, lvl, msg]);
    } catch (e) {
        console.error("Logging failed: " + e.message);
    }
}
function uiAlert_(msg) { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Alert', 10); } // Mobile safe
function getApiKey_() { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''; }
function withScriptLock_(key, fn) {
    const lock = LockService.getScriptLock();
    try { lock.waitLock(30000); return fn(); } finally { lock.releaseLock(); }
}
function notify_(type, title, msg) {
    try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.NOTIFICATIONS.name).appendRow([new Date(), type, title, msg, '', '']); } catch (e) { }
}
// Control Panel Utils
function findControlRowByAction_(s, axn) {
    const v = s.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] === axn) return i + 1;
    return 0;
}

// FIX: Robust Row Finder (TextFinder)
function findRowIndexByKey_(sheet, key, colIndex) {
    if (!sheet || !key) return -1;
    const tf = sheet.createTextFinder(key).matchEntireCell(true);
    const matches = tf.findAll();
    for (const m of matches) {
        if (m.getColumn() === colIndex) return m.getRow();
    }
    return -1;
}

function updateProgress_(s, r, p) { try { s.getRange(r, 6).setValue(p * 100 + '%'); SpreadsheetApp.flush(); } catch (e) { } }
function applyControlPanelFormatting_() {
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONTROL_PANEL.name);
    if (!s) return;
    s.clearConditionalFormatRules();
    const range = s.getRange('D:D');
    const r1 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('RUNNING').setBackground('#E8F5E9').setFontColor('#2E7D32').setRanges([range]).build();
    const r2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK').setBackground('#E3F2FD').setFontColor('#1565C0').setRanges([range]).build();
    const r3 = SpreadsheetApp.newConditionalFormatRule().whenTextContains('ERROR').setBackground('#FFEBEE').setFontColor('#C62828').setRanges([range]).build();
    s.setConditionalFormatRules([r1, r2, r3]);
    styleSheetPro_(s);
}
// Gemini API Utils with Model Routing, google_search fix, thinkingConfig wiring
function callGemini_(prompt, parts, schema, search, opts) {
    const key = getApiKey_();
    if (!key) throw new Error('Missing Gemini API Key in Script Properties (GEMINI_API_KEY).');

    const phase = (opts && opts.phase) ? String(opts.phase).toUpperCase() : 'EXTRACT';
    const modelList = (CFG.MODEL_CANDIDATES_BY_PHASE && CFG.MODEL_CANDIDATES_BY_PHASE[phase]) ?
        CFG.MODEL_CANDIDATES_BY_PHASE[phase] :
        (CFG.MODEL_CANDIDATES || ['gemini-2.0-flash']);

    const phaseGen = (CFG.PHASE_GEN && CFG.PHASE_GEN[phase]) ? CFG.PHASE_GEN[phase] : {};
    let generationConfig = Object.assign({}, phaseGen);

    if (generationConfig.thinkingBudget != null || generationConfig.thinkingLevel != null) {
        generationConfig.thinkingConfig = generationConfig.thinkingConfig || {};
        if (generationConfig.thinkingBudget != null) generationConfig.thinkingConfig.thinkingBudget = generationConfig.thinkingBudget;
        if (generationConfig.thinkingLevel != null) generationConfig.thinkingConfig.thinkingLevel = generationConfig.thinkingLevel;
        delete generationConfig.thinkingBudget;
        delete generationConfig.thinkingLevel;
    }

    if (schema) {
        generationConfig.responseMimeType = 'application/json';
        generationConfig.responseJsonSchema = schema;
    }

    const baseReq = {
        contents: [{ parts: [...(parts || []), { text: String(prompt || '') }] }]
    };

    if (Object.keys(generationConfig).length) baseReq.generationConfig = generationConfig;
    if (search) baseReq.tools = [{ google_search: {} }];

    let lastErr = null;

    for (const model of modelList) {
        const reqForModel = JSON.parse(JSON.stringify(baseReq));
        if (reqForModel.generationConfig && reqForModel.generationConfig.thinkingConfig && !model.toLowerCase().includes('pro')) {
            delete reqForModel.generationConfig.thinkingConfig;
        }

        const baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(reqForModel),
            muteHttpExceptions: true,
            headers: { 'x-goog-api-key': key }
        };

        try {
            let res = UrlFetchApp.fetch(baseUrl, options);
            if (res.getResponseCode() >= 300) {
                const body = res.getContentText() || '';
                if (schema && /responseSchema|schema|INVALID_ARGUMENT/i.test(body)) {
                    const retryReq = JSON.parse(JSON.stringify(baseReq));
                    if (retryReq.generationConfig) delete retryReq.generationConfig.responseJsonSchema;
                    res = UrlFetchApp.fetch(baseUrl, Object.assign({}, options, { payload: JSON.stringify(retryReq) }));
                    if (res.getResponseCode() >= 300) throw new Error(res.getContentText());
                    return parseGeminiResponse_(res, !!schema, !!search);
                }

                // Backward compatibility: retry with querystring key if header attempt fails
                const fallbackUrl = `${baseUrl}?key=${encodeURIComponent(key)}`;
                res = UrlFetchApp.fetch(fallbackUrl, Object.assign({}, options, { headers: {} }));
                if (res.getResponseCode() >= 300) throw new Error(body || res.getContentText());
            }

            return parseGeminiResponse_(res, !!schema, !!search);
        } catch (e) {
            lastErr = e;
            continue;
        }
    }

    throw new Error('Gemini failed for all candidate models. Last error: ' + (lastErr ? lastErr.message : 'Unknown'));
}

function parseGeminiResponse_(httpRes, expectJson, googleSearchEnabled) {
    const raw = JSON.parse(httpRes.getContentText() || '{}');
    const text = raw?.candidates?.[0]?.content?.parts?.map(p => p.text).filter(Boolean).join('\n') || '';
    if (!expectJson) return text;

    const parsed = safeJsonParse_(text);
    if (googleSearchEnabled && parsed && typeof parsed === 'object') {
        const groundingUrls = extractGroundingUrls_(raw);
        if (groundingUrls.length) parsed._grounding_urls = groundingUrls;
    }
    return parsed;
}

function extractGroundingUrls_(geminiRaw) {
    try {
        const urls = [];
        const meta = geminiRaw?.candidates?.[0]?.groundingMetadata ||
            geminiRaw?.candidates?.[0]?.grounding_metadata ||
            geminiRaw?.groundingMetadata ||
            geminiRaw?.grounding_metadata;

        const chunks = meta?.groundingChunks || meta?.grounding_chunks || [];
        for (const c of chunks) {
            const u = c?.web?.uri || c?.web?.url || c?.uri || c?.url;
            if (u && typeof u === 'string') urls.push(u);
        }
        return [...new Set(urls)];
    } catch (e) {
        return [];
    }
}

function safeJsonParse_(text) {
    let t = String(text || '').trim();
    t = t.replace(/^```(?:json)?/i, '').replace(/```$/i, '').trim();
    const objStart = t.indexOf('{');
    const arrStart = t.indexOf('[');
    const start = (objStart >= 0 && arrStart >= 0) ? Math.min(objStart, arrStart) : Math.max(objStart, arrStart);
    if (start > 0) t = t.slice(start).trim();
    try {
        return JSON.parse(t);
    } catch (e) {
        const lastObj = t.lastIndexOf('}');
        const lastArr = t.lastIndexOf(']');
        const end = Math.max(lastObj, lastArr);
        if (end > 0) return JSON.parse(t.slice(0, end + 1));
        throw e;
    }
}

function resolveFolderId_(val) {
    if (!val) throw new Error('Folder configuration missing.');
    const v = String(val).trim();
    if (!v) throw new Error('Empty folder configuration.');
    try { DriveApp.getFolderById(v); return v; } catch (e) { }
    const it = DriveApp.getRootFolder().getFoldersByName(v);
    if (it.hasNext()) return it.next().getId();
    const all = DriveApp.getFoldersByName(v);
    if (all.hasNext()) return all.next().getId();
    throw new Error('Could not find folder with ID or Name: "' + v + '". Please create it in your Drive.');
}

function uploadResumable_(file) {
    const key = getApiKey_();
    if (!key) throw new Error('Missing Gemini API Key in Script Properties (GEMINI_API_KEY).');

    const initUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${encodeURIComponent(key)}`;
    const numBytes = file.getSize();
    const mimeType = file.getMimeType() || MimeType.PDF;

    const startRes = UrlFetchApp.fetch(initUrl, {
        method: 'post',
        muteHttpExceptions: true,
        contentType: 'application/json',
        headers: {
            'X-Goog-Upload-Protocol': 'resumable',
            'X-Goog-Upload-Command': 'start',
            'X-Goog-Upload-Header-Content-Length': String(numBytes),
            'X-Goog-Upload-Header-Content-Type': mimeType
        },
        payload: JSON.stringify({ file: { display_name: file.getName() } })
    });

    if (startRes.getResponseCode() >= 300) {
        throw new Error('Gemini upload start failed: ' + startRes.getContentText());
    }

    const hdrs = startRes.getAllHeaders() || {};
    const uploadUrl = hdrs['x-goog-upload-url'] || hdrs['X-Goog-Upload-URL'] || hdrs['X-Goog-Upload-Url'];
    if (!uploadUrl) throw new Error('Upload URL missing from response headers.');

    const upRes = UrlFetchApp.fetch(uploadUrl, {
        method: 'post',
        muteHttpExceptions: true,
        headers: {
            'X-Goog-Upload-Offset': '0',
            'X-Goog-Upload-Command': 'upload, finalize',
            'Content-Length': String(numBytes)
        },
        payload: file.getBlob()
    });

    if (upRes.getResponseCode() >= 300) {
        throw new Error('Gemini upload finalize failed: ' + upRes.getContentText());
    }

    const json = JSON.parse(upRes.getContentText() || '{}');
    const uri = json?.file?.uri;
    const name = json?.file?.name;
    if (!uri) throw new Error('Upload response missing file.uri');

    return { uri: uri, name: name || uri };
}
function deleteGeminiFile_(fileRef) {
    const key = getApiKey_();
    const ref = (fileRef && (fileRef.name || fileRef.uri)) ? (fileRef.name || fileRef.uri) : String(fileRef || '');
    const id = extractGeminiFileId_(ref);
    const url = `https://generativelanguage.googleapis.com/v1beta/files/${encodeURIComponent(id)}?key=${encodeURIComponent(key)}`;
    try { UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true }); } catch (e) { }
}
function waitForActive_(fileRef) {
    const key = getApiKey_();
    const ref = (fileRef && (fileRef.name || fileRef.uri)) ? (fileRef.name || fileRef.uri) : String(fileRef || '');
    const id = extractGeminiFileId_(ref);

    const url = `https://generativelanguage.googleapis.com/v1beta/files/${encodeURIComponent(id)}?key=${encodeURIComponent(key)}`;

    const maxMs = 240000;
    const t0 = Date.now();
    while (Date.now() - t0 < maxMs) {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (res.getResponseCode() >= 300) throw new Error('Gemini file status error: ' + res.getContentText());

        const json = JSON.parse(res.getContentText() || '{}');
        if (json.state === 'ACTIVE') return;
        if (json.state === 'FAILED') throw new Error('Gemini processing failed.');

        Utilities.sleep(2000);
    }
    throw new Error('Gemini file timeout waiting for ACTIVE.');
}

function extractGeminiFileId_(ref) {
    const s = String(ref || '');
    const m = s.match(/\/files\/([^\/\s\?]+)|\bfiles\/([^\s\?]+)\b/);
    if (m) return (m[1] || m[2]);
    const parts = s.split('/').filter(Boolean);
    return parts.length ? parts[parts.length - 1].split('?')[0] : s;
}

function moveFileTo_(file, srcFolder, destFolder) {
    try {
        destFolder.addFile(file);
        srcFolder.removeFile(file);
    } catch (e) {
        try { file.moveTo(destFolder); } catch (e2) { }
    }
}
function makeBidId_(v, f) { return (v + f.getId()).replace(/\W/g, ''); }
function saveExtraction_(ss, bidId, data, file, status) {
    const va = ss.getSheetByName(SHEETS.VENDOR_ANALYSIS.name);
    va.appendRow([
        bidId,
        new Date(),
        data.vendor_name || '',
        data.total_grand_sum ?? '',
        data.services_sum ?? '',
        data.lead_time || '',
        '',
        data.warranty_rating || '',
        data.delivery_plan || '',
        data.exclusions || '',
        data.healthcare_compliant || '',
        (data.requires_loading_dock === true) ? 'YES' : '',
        data.receiving_fit || '',
        status,
        data.truncated_items ? 'YES' : '',
        data.evidence_snippets || '',
        file.getUrl()
    ]);
}

function calculateAndSaveLineItems_(ss, bidId, data, file) {
    const li = ss.getSheetByName(SHEETS.LINE_ITEMS.name);
    const vendor = data.vendor_name || "Unknown";
    const rows = [];

    (data.line_items || []).forEach((it, i) => {
        const mfg = String(it.manufacturer || "").trim();
        const mod = String(it.model || "").trim();
        const prodKey = (mfg && mod) ? (mfg + ' | ' + mod).toUpperCase() : "";

        rows.push([
            bidId + ':' + i,
            bidId,
            vendor,
            it.spec_id,
            it.description,
            it.qty,
            it.unit_price,
            it.extended_price,
            mfg,
            mod,
            it.finish || '',
            it.is_alternate || false,
            it.notes || '',
            prodKey,
            '100%',
            ''
        ]);
    });

    if (rows.length > 0) {
        const startRow = Math.max(li.getLastRow() + 1, 2);
        li.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }
}
function seedControlPanel_(ss) {
    const cp = ensureSheetSimple_(SHEETS.CONTROL_PANEL.name, SHEETS.CONTROL_PANEL.headers);
    if (cp.getLastRow() >= 2) return;

    const actions = [
        '0.5) Import Master Bids',
        '1) Phase A: Extract PDFs',
        '1.5) Phase A+: Auto-Audit Math',
        '2) Phase B: Research Products',
        '3) Phase C: Research Vendors',
        '4) Rebuild Rankings',
        '5) Generate Executive Summary',
        '6) Export Vendor Ranking to PDF'
    ];

    const rows = actions.map(a => [a, false, '', 'PENDING', 'Waiting...', '0%']);
    cp.getRange(2, 1, rows.length, 6).setValues(rows);

    cp.getRange(2, 2, rows.length, 1).insertCheckboxes();
    styleSheetPro_(cp);
}
function ensureControlPanelTrigger_() {
    const ss = SpreadsheetApp.getActive();
    if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onControlPanelChange_')) return;
    ScriptApp.newTrigger('onControlPanelChange_').forSpreadsheet(ss).onChange().create();
}
function refreshControlPanelTrigger_() {
    const t = ScriptApp.getProjectTriggers();
    t.forEach(ct => ScriptApp.deleteTrigger(ct));
    ensureControlPanelTrigger_();
    SpreadsheetApp.getActive().toast('AppSheet-Compatible Triggers Installed.');
}
function onControlPanelChange_(e) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
        if (!cp) return;

        const data = cp.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            if (data[i][1] === true || String(data[i][1]).toUpperCase() === 'TRUE') {
                const action = String(data[i][0]).trim();

                cp.getRange(i + 1, 2).setValue(false);
                SpreadsheetApp.flush();

                if (action.startsWith('0.5)')) importMasterBids_();
                else if (action.startsWith('1.5)')) runPhase15_();
                else if (action.startsWith('1)')) runPhaseA_();
                else if (action.startsWith('2)')) runPhaseB_();
                else if (action.startsWith('3)')) runPhaseC_();
                else if (action.startsWith('4)')) rebuildRankings_();
                else if (action.startsWith('5)')) generateExecutiveSummary_();
                else if (action.startsWith('6)')) exportRankingToPdf_();

                return;
            }
        }
    } catch (err) {
        logToSheet_('CONTROL_PANEL', 'ERROR', err.message);
    }
}
function patchHeaders_(n, h) {
    const ss = SpreadsheetApp.getActive();
    const s = ss.getSheetByName(n);
    if (!s || s.getLastColumn() === 0) return;
    const cur = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0] || [];
    const missing = h.filter(x => !cur.includes(x));
    if (missing.length) s.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]);
}
function validateConfigForMasterImport_() {
    const missing = [];
    if (isMissingConfigValue_(CFG.INPUT_FOLDER_ID) && isMissingConfigValue_(CFG.MASTERBIDS_FOLDER_ID)) missing.push('CFG.INPUT_FOLDER_ID (or CFG.MASTERBIDS_FOLDER_ID)');
    if (missing.length) throw new Error('Missing configuration for Master Import: ' + missing.join(', '));
}

function findMasterBidsFileId_() {
    if (CFG.MASTERBIDS_FILE_ID && String(CFG.MASTERBIDS_FILE_ID).trim()) return String(CFG.MASTERBIDS_FILE_ID).trim();
    const folderId = (CFG.MASTERBIDS_FOLDER_ID && String(CFG.MASTERBIDS_FOLDER_ID).trim()) ? String(CFG.MASTERBIDS_FOLDER_ID).trim() : String(CFG.INPUT_FOLDER_ID).trim();
    const hint = String(CFG.MASTERBIDS_FILENAME_HINT || 'master').toLowerCase().trim();
    const folder = DriveApp.getFolderById(folderId);
    let newestId = '';
    let newestTs = 0;
    const mimeCandidates = [MimeType.GOOGLE_SHEETS, MimeType.MICROSOFT_EXCEL, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    mimeCandidates.forEach(mime => {
        const it = folder.getFilesByType(mime);
        while (it.hasNext()) {
            const f = it.next();
            const name = String(f.getName() || '').toLowerCase();
            if (hint && name.indexOf(hint) === -1) continue;
            const ts = f.getLastUpdated().getTime();
            if (ts > newestTs) { newestTs = ts; newestId = f.getId(); }
        }
    });
    return newestId;
}

function ensureGoogleSheetFromFile_(fileId) {
    const f = DriveApp.getFileById(fileId);
    const mime = String(f.getMimeType() || '');
    if (mime === MimeType.GOOGLE_SHEETS || mime === 'application/vnd.google-apps.spreadsheet') return fileId;
    if (mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mime === MimeType.MICROSOFT_EXCEL) {
        const cacheKey = 'MASTERBIDS_CONVERTED_' + fileId;
        const cached = getRunProp_(cacheKey);
        if (cached) { try { DriveApp.getFileById(cached); return cached; } catch (e) { } }
        if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.copy) {
            throw new Error('To import an XLSX master bids file, enable Advanced Google Services: Drive API (Services → +).');
        }
        const resource = { title: f.getName() + ' (Converted)', mimeType: MimeType.GOOGLE_SHEETS };
        if (CFG.PROCESSED_FOLDER_ID) resource.parents = [{ id: String(CFG.PROCESSED_FOLDER_ID).trim() }];
        const copy = Drive.Files.copy(resource, fileId, { convert: true });
        if (!copy || !copy.id) throw new Error('XLSX conversion failed.');
        setRunProp_(cacheKey, copy.id);
        return copy.id;
    }
    throw new Error('Unsupported master bids file type: ' + mime);
}

function pickSourceSheet_(ss) {
    if (CFG.MASTERBIDS_SHEET_NAME) {
        const named = ss.getSheetByName(String(CFG.MASTERBIDS_SHEET_NAME).trim());
        if (named) return named;
    }
    const sheets = ss.getSheets();
    if (!sheets.length) throw new Error('No sheets found.');
    const preferred = ['MasterBidTable', 'Master Bid Table', 'MasterBid', 'Master Bid', 'Master', 'Bids'];
    for (let name of preferred) { const sh = ss.getSheetByName(name); if (sh) return sh; }
    for (let sh of sheets) {
        const h = sh.getRange(1, 1, 1, Math.min(50, sh.getLastColumn())).getValues()[0] || [];
        const n = h.map(x => normalizeHeaderKey_(x));
        if (n.includes('specid') || n.includes('spec id')) return sh;
    }
    return sheets[0];
}

function readMasterBidsRows_(srcSheet) {
    const headerRow = Math.max(1, toNumber_(CFG.MASTERBIDS_HEADER_ROW) || 1);
    const dataStart = Math.max(headerRow + 1, toNumber_(CFG.MASTERBIDS_DATA_START_ROW) || (headerRow + 1));
    const lastRow = srcSheet.getLastRow(), lastCol = srcSheet.getLastColumn();
    if (lastRow < dataStart || lastCol < 1) return { rows: [], sourceName: srcSheet.getName() };
    const headers = srcSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(String);
    const idx = buildMasterBidsHeaderIndex_(headers);
    if (idx.specId < 0) throw new Error('Could not find "Spec ID" column.');
    const data = srcSheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getValues();
    const out = [];
    for (let row of data) {
        const spec = String(row[idx.specId] || '').trim();
        if (!spec) continue;
        const qty = (idx.qty >= 0) ? toNumber_(row[idx.qty]) : '';
        const desc = (idx.desc >= 0) ? String(row[idx.desc] || '').trim() : '';
        const cat = (idx.category >= 0) ? String(row[idx.category] || '').trim() : '';
        const mfg = (idx.mfg >= 0) ? String(row[idx.mfg] || '').trim() : '';
        const model = (idx.model >= 0) ? String(row[idx.model] || '').trim() : '';
        const finish = (idx.finish >= 0) ? String(row[idx.finish] || '').trim() : '';
        const room = (idx.room >= 0) ? String(row[idx.room] || '').trim() : '';
        const notes = (idx.notes >= 0) ? String(row[idx.notes] || '').trim() : '';
        let budget = '';
        if (idx.budget >= 0) budget = parseCurrency_(row[idx.budget]);
        else if (idx.totalBudget >= 0 && qty > 0) budget = parseCurrency_(row[idx.totalBudget]) / qty;
        out.push([spec, cat, desc, qty, budget, mfg, model, finish, room, notes]);
    }
    return { rows: out, sourceName: srcSheet.getName() };
}

function buildMasterBidsHeaderIndex_(headers) {
    const norm = headers.map(normalizeHeaderKey_);
    const pick = (syns) => { for (let i = 0; i < norm.length; i++) if (syns.includes(norm[i])) return i; return -1; };
    return {
        specId: pick(['specid', 'spec id', 'itemid', 'item id', 'item#', 'id']),
        category: pick(['category', 'cat', 'type', 'division']),
        desc: pick(['description', 'desc', 'item description']),
        qty: pick(['targetqty', 'qty', 'quantity', 'count']),
        budget: pick(['targetbudget', 'unitbudget', 'budget', 'unitprice', 'unit cost']),
        totalBudget: pick(['totalbudget', 'extendedbudget', 'total', 'extended price']),
        mfg: pick(['manufacturer', 'mfg', 'brand']),
        model: pick(['model', 'model#', 'sku', 'part#']),
        finish: pick(['finish', 'color', 'material']),
        room: pick(['roomtype', 'room']),
        notes: pick(['notes', 'comments', 'remarks'])
    };
}
function normalizeHeaderKey_(h) { return String(h || '').toLowerCase().replace(/[^\w\s]/g, '').trim(); }
function parseCurrency_(v) { if (typeof v === 'number') return v; return parseFloat(String(v).replace(/[$,]/g, '')) || 0; }
function toNumber_(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function isMissingConfigValue_(v) { return !v || String(v).trim().length === 0 || String(v).includes('PASTE_'); }
function getRunProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function setRunProp_(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); }
function reconcileWithMaster_(ss) {
    const m = ss.getSheetByName(SHEETS.MASTER.name);
    if (!m || m.getLastRow() < 2) return;
    const mData = m.getDataRange().getValues();
    const mHead = mData[0].map(s => String(s).toLowerCase());
    const idx = {
        mfg: mHead.indexOf('manufacturer'), model: mHead.indexOf('model'), finish: mHead.indexOf('finish'),
        desc: mHead.indexOf('description'), cat: mHead.indexOf('category')
    };
    if (idx.mfg < 0 || idx.model < 0) return;

    const p = ss.getSheetByName(SHEETS.PRODUCTS.name);
    const pData = p.getDataRange().getValues();
    const pMap = new Map();
    pData.slice(1).forEach(r => pMap.set(String(r[0]), true));

    const newRows = [];
    mData.slice(1).forEach(r => {
        const mfg = String(r[idx.mfg] || '').trim();
        const mod = String(r[idx.model] || '').trim();
        if (!mfg || !mod) return;
        const key = (mfg + ' | ' + mod).toUpperCase();
        if (!pMap.has(key)) {
            newRows.push([key, mfg, mod, r[idx.finish] || '', r[idx.desc] || '', 'MASTER', 'PENDING', new Date(), '', '', '', '']);
            pMap.set(key, true);
        }
    });

    if (newRows.length) {
        p.getRange(p.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        ss.toast(`Synced ${newRows.length} items from Master.`);
    }
}

function logSkip_(ss, f, r) {
    logToSheet_('PHASE_A', 'SKIP', `Skipped ${f.getName()}: ${r}`);
    try {
        const sf = DriveApp.getFolderById(CFG.SKIPPED_FOLDER_ID);
        f.moveTo(sf);
    } catch (e) { logToSheet_('PHASE_A', 'WARN', 'Could not move skipped file: ' + e.message); }
}

function generateExecutiveSummary_() {
    const ss = SpreadsheetApp.getActive();
    const cp = startAction_(ss, '5) Generate Executive Summary');
    const es = ensureSheetSimple_(SHEETS.EXEC_SUMMARY.name, SHEETS.EXEC_SUMMARY.headers);
    es.clearContents();
    es.appendRow(SHEETS.EXEC_SUMMARY.headers);

    const rankings = ss.getSheetByName(SHEETS.RANKING.name).getDataRange().getValues();
    if (rankings.length < 2) {
        if (cp) finishAction_(cp, 'ERROR', 'No rankings found.');
        return;
    }

    const top = rankings[1];
    const runnerUp = rankings[2] || null;

    const sections = [
        ['Primary Recommendation', `Based on weighted analysis, **${top[1]}** is the recommended vendor with an overall score of ${(Number(top[3]) || 0).toFixed(1)}/100.`],
        ['Quality Analysis', top[10] || "Quality meets or exceeds facility standards based on product research."],
        ['Financial Impact', `This bid represents a ${top[6] > 80 ? 'highly competitive' : 'standard'} value. Price score: ${Number(top[8]).toFixed(1)}/100.`],
        ['Risk Mitigation', top[16] || "No significant red flags identified in vendor research."],
        ['Comparison', runnerUp ? `${top[1]} outperformed ${runnerUp[1]} primarily on ${top[5] > runnerUp[5] ? 'Quality' : 'Execution'}.` : "No viable second-place candidates identified."]
    ];

    sections.forEach(row => es.appendRow(row));

    es.getRange("A:A").setFontWeight("bold");
    styleSheetPro_(es);

    if (cp) finishAction_(cp, 'OK', 'Executive Narrative Generated');
}

function exportRankingToPdf_() {
    const ss = SpreadsheetApp.getActive();
    const cp = startAction_(ss, '6) Export Vendor Ranking to PDF');
    const rs = ss.getSheetByName(SHEETS.RANKING.name);

    try {
        const folder = DriveApp.getFolderById(CFG.PROCESSED_FOLDER_ID);
        const name = 'Vendor_Ranking_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.pdf';

        const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?exportFormat=pdf&format=pdf' +
            '&size=letter&portrait=false&fitw=true&gridlines=false&gid=' + rs.getSheetId();

        const token = ScriptApp.getOAuthToken();
        const blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob().setName(name);
        const file = folder.createFile(blob);

        if (cp) finishAction_(cp, 'OK', 'PDF Saved: ' + file.getUrl());
        ss.toast('PDF Exported: ' + name);
    } catch (e) {
        if (cp) finishAction_(cp, 'ERROR', e.message);
        uiAlert_('PDF Export Failed: ' + e.message);
    }
}

function startAction_(ss, actionName) {
    const cp = ss.getSheetByName(SHEETS.CONTROL_PANEL.name);
    if (!cp) return null;
    const row = findControlRowByAction_(cp, actionName);
    if (row > 0) {
        cp.getRange(row, 3).setValue(new Date());
        cp.getRange(row, 4).setValue('RUNNING');
        cp.getRange(row, 5).setValue('Starting...');
        cp.getRange(row, 6).setValue('0%');
        SpreadsheetApp.flush();
        return { sheet: cp, row: row };
    }
    return null;
}

function finishAction_(cpInfo, status, msg) {
    if (!cpInfo || !cpInfo.sheet) return;
    const s = cpInfo.sheet;
    const r = cpInfo.row;
    try {
        s.getRange(r, 4).setValue(status);
        s.getRange(r, 5).setValue(msg);
        if (status === 'OK') s.getRange(r, 6).setValue('100%');
        SpreadsheetApp.flush();
    } catch (e) { }
}

function truncate_(str, len) {
    if (!str) return '';
    if (str.length <= len) return str;
    return str.substring(0, len) + '...';
}

function median_(arr) {
    if (!arr || !arr.length) return 0;
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    if (sorted.length % 2 === 0) return (sorted[mid - 1] + sorted[mid]) / 2;
    return sorted[mid];
}
