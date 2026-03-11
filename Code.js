/**
 * Code.gs
 * Google Apps Script for Jira Sheets Sync
 */

const APP_TITLE = "Sync Sheets for JIRA";
const APP_VERSION = "2.1.0";
const BUILD_DATE = "2026-03-09 12:37:43";
const CLOUDFLARE_WORKER_URL = "https://jira-proxy.railflow.workers.dev"; // Bakes the proxy URL into the addon
const PROXY_SECRET = "jira-sheets-secret-2026"; // Must match the secret in cloudflare_worker.js

function onOpen() {
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('⚙︎ JIRA Configuration', 'showConfig')
        .addItem('▦ JIRA Issues', 'showIssues')
        .addItem('✚ Bulk Create Issues', 'showBulkCreate')
        .addSeparator()
        .addItem('▣ Dashboards', 'showDashboard')
        .addItem('📈 Metrics', 'showMetrics')
        .addItem('📅 Roadmap Planning', 'showRoadmap')
        .addItem('👥 Capacity Planning', 'showCapacity')
        .addItem('⊞ Pivot Table', 'showPivot')
        .addItem('✏️ JQL Update', 'showJqlUpdate')
        .addItem('🏷️ Label Management', 'showLabels')
        .addSeparator()
        .addItem('ⓘ License Info', 'showLicense')
        .addItem('☰ Logs', 'showRefreshLogs')
        .addItem('ℹ️ About', 'showAbout')
        .addToUi();
}

/**
 * Tracks user edits to identify rows for bulk updates.
 */
function onEdit(e) {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const range = e.range;
    const startRow = range.getRow();
    const numRows = range.getNumRows();

    // If strictly header, return. If range starts at 1 but spans more, we process.
    if (startRow === 1 && numRows === 1) return;

    // Get all headers once to find Key column
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    // Optimization: Check if we already know Key column index (cached)
    // But simple triggers can't use Properties reliably for high frequency? 
    // Actually PropertiesService is fine.

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const keyIndex = headers.findIndex(h => h && h.toString().trim().toLowerCase() === 'key');

    if (keyIndex === -1) return; // No Key column found

    // VISUAL FEEDBACK: Mark edited cell(s) with Orange background
    range.setBackground('#ff9900');

    // Get keys for ALL edited rows
    // Adjust if range includes header (startRow 1) - unlikely due to check above, but mostly safe
    // But if selection starts at 1 and goes down? 
    // Trigger says `startRow` is top. If startRow=1, we return line 39. 
    // Wait, if I edit A1:A10, startRow is 1. I return. I miss rows 2-10!
    // FIX: logic should handle intersection.

    let effectiveStartRow = startRow;
    let effectiveNumRows = numRows;

    if (startRow === 1) {
        if (numRows === 1) return; // Only header
        effectiveStartRow = 2;
        effectiveNumRows = numRows - 1;
    }

    if (effectiveNumRows <= 0) return;

    const keys = sheet.getRange(effectiveStartRow, keyIndex + 1, effectiveNumRows, 1).getValues();

    // Store in Document Properties as "Modified"
    const props = PropertiesService.getDocumentProperties();
    let modifiedMap = {};
    let saveNeeded = false;

    try {
        const json = props.getProperty('modifiedKeys');
        if (json) modifiedMap = JSON.parse(json);
    } catch (err) { }

    keys.forEach(row => {
        const keyVal = row[0];
        if (keyVal && keyVal.toString().trim() !== '') {
            if (!modifiedMap[keyVal]) {
                modifiedMap[keyVal] = true;
                saveNeeded = true;
            }
        }
    });

    if (saveNeeded) {
        props.setProperty('modifiedKeys', JSON.stringify(modifiedMap));
    }
}

/**
 * Creates the homepage card for the add-on.
 * Required for Workspace Marketplace add-ons.
 */
function onHomepage(e) {
    return createHomepageCard();
}

/**
 * Creates the homepage card UI.
 */
function createHomepageCard() {
    var card = CardService.newCardBuilder();

    var header = CardService.newCardHeader()
        .setTitle('Jira Sync for Sheets')
        .setSubtitle('Sync your Jira issues with Google Sheets');
    card.setHeader(header);

    var section = CardService.newCardSection();

    section.addWidget(
        CardService.newTextParagraph()
            .setText('Connect your Jira account and sync issues directly to your spreadsheet.')
    );

    section.addWidget(
        CardService.newTextButton()
            .setText('Open Sidebar')
            .setOnClickAction(
                CardService.newAction().setFunctionName('openSidebarFromCard')
            )
    );

    card.addSection(section);
    return card.build();
}

/**
 * Opens the sidebar from the homepage card.
 */
function openSidebarFromCard() {
    showSidebar('config');
}
/**
 * Returns a list of all sheet names in the active spreadsheet
 */
function getJiraSheets() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/**
 * Validates if the spreadsheet has required Jira columns (Key, Summary, Status)
 */
function validateJiraHeaders(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sheet) return { valid: false, error: 'Sheet not found' };

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { valid: false, error: 'Sheet is empty' };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).toLowerCase().trim());
    const required = ['key', 'summary', 'status'];
    const missing = required.filter(r => !headers.includes(r));

    return {
        valid: missing.length === 0,
        missing: missing
    };
}


/**
 * Returns available columns for a specific sheet.
 */
function getSheetColumns(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sheet) return [];

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return [];

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    return headers.map(h => {
        const val = String(h).trim();
        if (!val) return null;
        return {
            id: val,
            name: val,
            key: val.toLowerCase().replace(/[^a-z0-9]/g, '')
        };
    })
        .filter(Boolean);
}

function showConfig() {
    showSidebar('config');
}

function showIssues() {
    showSidebar('issues');
}

function showBulkCreate() {
    showSidebar('bulkCreate');
}

function showDashboard() {
    showSidebar('dashboard');
}

function showMetrics() {
    showSidebar('metrics');
}

function showRoadmap() {
    showSidebar('roadmap');
}

function showCapacity() {
    showSidebar('capacity');
}

function showPivot() {
    showSidebar('pivot');
}

function showJqlUpdate() {
    showSidebar('jqlUpdate');
}

function showLabels() {
    showSidebar('labels');
}

function showLicense() {
    showSidebar('license');
}

function showAbout() {
    const html = HtmlService.createHtmlOutput(
        '<div style="font-family:Google Sans,Arial,sans-serif;padding:20px;">' +
        '<h2 style="margin:0 0 16px;">' + APP_TITLE + '</h2>' +
        '<p style="margin:6px 0;color:#555;">Version: <strong>' + APP_VERSION + '</strong></p>' +
        '<p style="margin:6px 0;color:#555;">Build Date: <strong>' + BUILD_DATE + '</strong></p>' +
        '</div>'
    )
        .setWidth(400)
        .setHeight(220);
    SpreadsheetApp.getUi().showModalDialog(html, 'About');
}

function showSidebar(mode) {
    const html = HtmlService.createTemplateFromFile('Sidebar');
    html.initialMode = mode || 'config';

    // Try to get email robustly
    let email = Session.getActiveUser().getEmail();
    if (!email) email = Session.getEffectiveUser().getEmail();

    html.userEmail = email || '';

    const output = html.evaluate()
        .setTitle(APP_TITLE)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Saves user configuration to Document Properties
 */
function saveConfig(config) {
    let props;
    try {
        props = PropertiesService.getDocumentProperties();
        // Trigger a read to verify access before writing
        props.getKeys();
    } catch (e) {
        // Retry once after a short delay (intermittent GAS document context issue)
        Utilities.sleep(1000);
        try {
            props = PropertiesService.getDocumentProperties();
            props.getKeys();
        } catch (e2) {
            throw new Error('Could not access document properties. Make sure you have edit access to this spreadsheet. (' + e2.message + ')');
        }
    }
    const sheetId = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();

    // Save Credentials Globally
    if (config.domain && config.email && config.token) {
        props.setProperty('jira_creds', JSON.stringify({
            domain: config.domain,
            email: config.email,
            token: config.token,
            useCloudflare: config.useCloudflare,
            cloudflareWorkerUrl: config.cloudflareWorkerUrl
        }));
    }

    // Save Sheet-Specific Config
    const sheetConfig = {
        jql: config.jql,
        columns: config.columns,
        projectKey: config.projectKey
    };
    props.setProperty(`config_${sheetId}`, JSON.stringify(sheetConfig));

    return { success: true };
}

/**
 * Loads user configuration
 */
function loadConfig() {
    try {
        const props = PropertiesService.getDocumentProperties();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        const sheetId = sheet.getSheetId();

        // Load Global Creds
        const credsJson = props.getProperty('jira_creds');
        let creds = {};
        if (credsJson) {
            try { creds = JSON.parse(credsJson); } catch (e) { }
        }

        // Load Sheet Config
        const sheetJson = props.getProperty(`config_${sheetId}`);
        let sheetConfig = {};
        if (sheetJson) {
            try { sheetConfig = JSON.parse(sheetJson); } catch (e) { }
        }

        // Load Sheet Schedule
        const schedJson = props.getProperty(`scheduledConfig_${sheetId}`);

        const result = { ...creds, ...sheetConfig };
        if (schedJson) {
            result.scheduledConfig = schedJson;
        }

        return result;
    } catch (e) {
        console.error("Storage access failed: " + e.message);
        if (e.message.includes("PERMISSION_DENIED")) {
            throw new Error("Google Storage Error: You may be logged into multiple accounts. Please try Incognito mode or log out of other accounts.");
        }
        throw e;
    }
}


/**
 * Returns the column headers from the active sheet
 */
function getSheetHeaders(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sheet) return [];

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return [];

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Known custom field labels
    const CUSTOM_FIELD_LABELS = {
        'customfield_10015': 'Start Date',
        'customfield_10016': 'Story Points',
        'customfield_10017': 'Epic Link',
        'customfield_10018': 'Epic Name',
        'customfield_10019': 'Sprint'
    };

    // Convert to "Header (A)", "Header (B)" etc to handle duplicates
    return headers.map((h, i) => {
        if (!h || h.toString().trim() === '') return null;
        const letter = columnToLetter(i + 1);
        let label = h.toString().trim();
        if (label.startsWith('customfield_')) {
            label = CUSTOM_FIELD_LABELS[label] || label;
        }
        return `${label} (${letter})`;
    }).filter(h => h !== null);
}

/**
 * Converts a column number to its letter equivalent (e.g. 1 -> A, 27 -> AA)
 */
function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

/**
 * Converts a column letter to its number equivalent (e.g. A -> 1, AA -> 27)
 */
function letterToColumn(letter) {
    let column = 0;
    for (let i = 0; i < letter.length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
    }
    return column;
}

/**
 * Tests the connection to Jira
 */
function testConnection(config) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing credentials");
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const url = `https://${cleanDomain}/rest/api/3/myself`;

    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token)
        },
        muteHttpExceptions: true
    };

    try {
        // Always test directly — proxy bypass not needed for credential validation
        const response = fetchWithRetry(url, options, { ...config, useCloudflare: false });
        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            return { success: true, message: `Connected as ${data.displayName}` };
        } else {
            let msg = response.getContentText();
            try { msg = JSON.parse(msg).errorMessages.join(', '); } catch (e) { }
            throw new Error(`Connection failed (${response.getResponseCode()}): ${msg}`);
        }
    } catch (e) {
        throw new Error(e.message);
    }
}

/**
 * Validates JQL and gets the count of matching issues
 */
function getJqlCount(config) {
    if (!config.domain || !config.email || !config.token) {
        return { valid: false, count: 0 };
    }
    if (!config.jql || config.jql.trim() === '') {
        return { valid: true, count: 0 };
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const url = `https://${cleanDomain}/rest/api/3/search/approximate-count`;

    const options = {
        method: 'POST',
        contentType: 'application/json',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token),
            "Accept": "application/json"
        },
        payload: JSON.stringify({ jql: config.jql }),
        muteHttpExceptions: true
    };

    try {
        const response = fetchWithRetry(url, options, config);
        const text = response.getContentText();
        const code = response.getResponseCode();

        if (code === 200) {
            const data = JSON.parse(text);
            return { valid: true, count: data.count || 0 };
        } else {
            let errorMsg = "Invalid query";
            try {
                const errData = JSON.parse(text);
                if (errData.errorMessages && errData.errorMessages.length > 0) {
                    errorMsg = errData.errorMessages.join(', ');
                } else if (errData.message) {
                    errorMsg = errData.message;
                }
            } catch (e) {
                if (text.length < 200) errorMsg = text;
            }
            return { valid: false, error: `${errorMsg} [${code}]` };
        }
    } catch (e) {
        return { valid: false, error: `Request failed: ${e.message}` };
    }
}

/**
 * Fetches data from Jira and populates the sheet
 * @param {Object} params - {domain, email, token, jql, columns}
 */
function fetchJiraData(params) {
    if (!params.domain || !params.email || !params.token) {
        try { logRefreshAttempt('FAILED', 'Missing credentials', params.targetSheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()); } catch (e) { }
        throw new Error("Missing credentials");
    }

    // Initialize proxy counter for this fetch
    params.proxyCallCount = 0;

    // Determine Pro status if not provided (e.g. background trigger)
    if (typeof params.isPro === 'undefined') {
        const license = checkLicense();
        params.isPro = license.allowed && license.plan !== 'free';
    }

    // Default columns if empty
    const rawColumns = params.columns || 'key, summary, status, priority, assignee';
    const columns = rawColumns.split(',').map(c => c.trim().toLowerCase());

    // Ensure Key is always fetched
    if (!columns.includes('key')) {
        columns.unshift('key');
    }

    // Always fetch 'updated' field for change detection (but don't display it unless requested)
    const fetchColumns = [...columns];
    if (!fetchColumns.includes('updated')) {
        fetchColumns.push('updated');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet;

    if (params.targetSheetName) {
        sheet = ss.getSheetByName(params.targetSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(params.targetSheetName);
        }
    }

    // Fallback to active if not specified
    if (!sheet) {
        sheet = ss.getActiveSheet();
    } else if (params.targetSheetName) {
        sheet.activate();
    }

    const spreadsheet = ss;

    // Flash header green to indicate refresh
    const existingLastCol = sheet.getLastColumn();

    const cleanDomain = params.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const jql = params.jql || 'assignee = currentUser() AND resolution = Unresolved order by updated DESC';
    const jqlEncoded = encodeURIComponent(jql);
    const fieldsEncoded = encodeURIComponent(fetchColumns.join(','));
    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(params.email + ":" + params.token),
            "Accept": "application/json"
        },
        muteHttpExceptions: true
    };

    let issues = [];
    let startAt = 0;
    let total = -1;
    let nextPageToken = null;

    do {
        let url = `https://${cleanDomain}/rest/api/3/search/jql?jql=${jqlEncoded}&fields=${fieldsEncoded}&maxResults=100`;
        if (nextPageToken) {
            url += `&nextPageToken=${encodeURIComponent(nextPageToken)}`;
        } else {
            url += `&startAt=${startAt}`;
        }

        const response = fetchWithRetry(url, options, params);
        const code = response.getResponseCode();
        const text = response.getContentText();

        if (code !== 200) {
            // Restore header on error
            if (existingLastCol > 0) {
                const headerRange = sheet.getRange(1, 1, 1, existingLastCol);
                headerRange.setBackground('#f4f5f7');
            }
            let msg = text;
            try {
                const err = JSON.parse(text);
                if (err.errorMessages) msg = err.errorMessages.join(', ');
                else if (err.message) msg = err.message;
            } catch (e) { }

            try { logRefreshAttempt('FAILED', `Jira Error (${code}): ${msg}`, sheet ? sheet.getName() : 'Unknown'); } catch (e) { }
            throw new Error(`Jira Error (${code}): ${msg}`);
        }

        const data = JSON.parse(text);
        if (total === -1 && typeof data.total !== 'undefined') {
            total = data.total;
        }

        // Update token for next iteration
        nextPageToken = data.nextPageToken;

        // Warning Logic
        if (issues.length === 0 && total > 5000 && !params.force) {
            return { warning: true, total: total };
        }

        const batch = data.issues || [];
        if (batch.length === 0) break;

        // Duplicate Detection
        if (issues.length > 0 && batch[0].key === issues[0].key) {
            try { logRefreshAttempt('WARN', 'Pagination loop detected (duplicates). Stopping.', sheet ? sheet.getName() : 'Unknown'); } catch (e) { }
            break;
        }

        issues = issues.concat(batch);
        startAt += batch.length;

        // FREE VERSION LIMIT: 150 Issues
        if (!params.isPro && issues.length >= 150) {
            issues = issues.slice(0, 150);
            params.limited = true; // Mark as limited for UI
            break;
        }

        // Break if done
        if (!nextPageToken && total !== -1 && issues.length >= total) break;
        if (!nextPageToken && total === -1) break; // End of results assumed

    } while (true);

    // Get previous update timestamps to detect changes
    // Get previous update timestamps to detect changes
    const props = PropertiesService.getDocumentProperties();
    const rawPrev = props.getProperty('issueTimestamps');
    const isFirstRun = rawPrev === null;
    const prevTimestamps = rawPrev ? JSON.parse(rawPrev) : {};

    // Build new timestamps map and count changes
    const newTimestamps = {};
    let changedCount = 0;
    const changedKeys = new Set();

    issues.forEach(issue => {
        const key = issue.key;
        const updated = issue.fields?.updated || '';
        newTimestamps[key] = updated;

        // Count as changed only if not first run AND (new issue OR updated timestamp differs)
        if (!isFirstRun && (!prevTimestamps[key] || prevTimestamps[key] !== updated)) {
            changedCount++;
            changedKeys.add(key);
        }
    });

    // Store new timestamps for next comparison
    props.setProperty('issueTimestamps', JSON.stringify(newTimestamps));

    // Load field name mapping for friendly column headers
    let fieldNameMapJson = props.getProperty('fieldNameMap');
    let fieldNameMap = fieldNameMapJson ? JSON.parse(fieldNameMapJson) : {};

    // Proactive check: If we have customfields in columns but NOT in map, refresh mapping
    const hasUnmappedCustomFields = columns.some(col =>
        col.startsWith('customfield_') && !fieldNameMap[col]
    );

    if (hasUnmappedCustomFields || Object.keys(fieldNameMap).length === 0) {
        try {
            getJiraFields(params);
            fieldNameMapJson = props.getProperty('fieldNameMap');
            fieldNameMap = fieldNameMapJson ? JSON.parse(fieldNameMapJson) : {};
        } catch (e) {
            console.warn("Failed to refresh fieldNameMap: " + e.message);
        }
    }

    // Create friendly header names (map customfield_xxx to their display names)
    const friendlyHeaders = columns.map(col => {
        // If we have a mapping, use it; otherwise use the original column name
        let name = fieldNameMap[col] || col;
        // Capitalize standard fields if they are keys
        if (name === col && !col.startsWith('customfield_')) {
            name = col.charAt(0).toUpperCase() + col.slice(1);
        }
        return name;
    });

    // Build a map of issue key -> row data for quick lookup
    const issueMap = {};
    issues.forEach(issue => {
        const row = columns.map(col => {
            if (col === 'key') {
                return `=HYPERLINK("https://${cleanDomain}/browse/${issue.key}", "${issue.key}")`;
            }
            const val = issue.fields ? issue.fields[col] : null;
            if (val === null || val === undefined) return "";
            if (typeof val === 'object') {
                return val.name || val.displayName || val.value || JSON.stringify(val);
            }
            return val;
        });
        issueMap[issue.key] = row;
    });

    // Preserve existing row order: read current keys from sheet, update in place, append new
    const output = [friendlyHeaders];
    const rowIsChanged = [false]; // First row is header, not changed
    const processedKeys = new Set();

    const existingLastRow = sheet.getLastRow();
    const keyColIndex = columns.indexOf('key');
    if (existingLastRow > 1 && keyColIndex !== -1) {
        const existingKeys = sheet.getRange(2, keyColIndex + 1, existingLastRow - 1, 1).getDisplayValues();
        existingKeys.forEach(row => {
            const existingKey = row[0] ? row[0].toString().trim() : '';
            if (existingKey && issueMap[existingKey]) {
                output.push(issueMap[existingKey]);
                rowIsChanged.push(changedKeys.has(existingKey));
                processedKeys.add(existingKey);
            }
        });
    }

    // Append any new issues not already on the sheet
    issues.forEach(issue => {
        if (!processedKeys.has(issue.key)) {
            output.push(issueMap[issue.key]);
            rowIsChanged.push(changedKeys.has(issue.key));
        }
    });

    // Collect unique users for mapping and dropdown
    const userMap = {};
    const uniqueUsersArray = [];

    // 1. Users from current issues
    issues.forEach(issue => {
        if (issue.fields && issue.fields.assignee) {
            const user = issue.fields.assignee;
            const name = user.displayName;
            const id = user.accountId;
            if (name && id && !userMap[name]) {
                userMap[name] = id;
                uniqueUsersArray.push(name);
            }
        }
    });

    // 2. Supplement with assignable users from projects in this batch
    const projectsInBatch = new Set();
    if (params.projectKey) projectsInBatch.add(params.projectKey);
    issues.forEach(issue => {
        if (issue.fields && issue.fields.project && issue.fields.project.key) {
            projectsInBatch.add(issue.fields.project.key);
        }
    });

    // Fetch from up to 3 projects to stay responsive
    const projectKeys = Array.from(projectsInBatch).slice(0, 3);
    if (projectKeys.length === 0) projectKeys.push(null);

    projectKeys.forEach(pKey => {
        const extraUsers = getAssignableUsers(params, pKey);
        extraUsers.forEach(u => {
            if (u.displayName && u.accountId && !userMap[u.displayName]) {
                userMap[u.displayName] = u.accountId;
                uniqueUsersArray.push(u.displayName);
            }
        });
    });

    // Store user map for updates
    // Helper to merge with existing map if needed, but fresh map is safer to avoid stale users
    // For robustness, maybe we load existing first? Let's just store current batch + persist.
    // Actually, simple overwrite is ok if we assume the sheet reflects current state.
    // But if we want to update a row to a user NOT in the current query, we'd lose them.
    // Let's fallback: try to read existing map.
    const existUserMapJson = props.getProperty('userMap') || '{}';
    const existUserMap = JSON.parse(existUserMapJson);
    const finalUserMap = { ...existUserMap, ...userMap };
    props.setProperty('userMap', JSON.stringify(finalUserMap));

    // Cache issue type by key for later use (e.g., when auto-adding Issue Type column)
    const issueTypeByKey = {};
    issues.forEach(issue => {
        if (issue.key && issue.fields && issue.fields.issuetype && issue.fields.issuetype.name) {
            issueTypeByKey[issue.key] = issue.fields.issuetype.name;
        }
    });
    props.setProperty('issueTypeByKey', JSON.stringify(issueTypeByKey));

    // Sort unique users for dropdown
    uniqueUsersArray.sort();

    // Update Sheet
    // Clear contents AND formats to start fresh (wipe out), but preserve column widths
    sheet.clearContents();
    sheet.clearFormats();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

    if (output.length > 0) {
        const range = sheet.getRange(1, 1, output.length, output[0].length);
        range.setValues(output);

        // Style header - restore normal styling
        const headerRange = sheet.getRange(1, 1, 1, output[0].length);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f4f5f7'); // Jira light grey
        headerRange.setBorder(null, null, true, null, null, null, '#dfe1e6', SpreadsheetApp.BorderStyle.SOLID);

        // Highlight ONLY changed rows in light yellow (keep existing highlights for unchanged)


        // Apply Dropdown to Assignee column
        const assigneeColIndex = columns.indexOf('assignee');
        if (assigneeColIndex !== -1 && uniqueUsersArray.length > 0) {
            // Apply to all data rows
            const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(uniqueUsersArray)
                .setAllowInvalid(true) // Allow existing values not in list (e.g. unassigned)
                .build();
            // Range: row 2, col Index+1, numRows, 1
            if (issues.length > 0) {
                sheet.getRange(2, assigneeColIndex + 1, issues.length, 1).setDataValidation(rule);
            }
        }

        // Apply Dropdown to Priority column
        const priorityColIndex = columns.indexOf('priority');
        if (priorityColIndex !== -1) {
            const uniquePriorities = new Set();
            issues.forEach(issue => {
                if (issue.fields && issue.fields.priority && issue.fields.priority.name) {
                    uniquePriorities.add(issue.fields.priority.name);
                }
            });
            const prioritiesArray = Array.from(uniquePriorities).sort();

            if (prioritiesArray.length > 0) {
                const rule = SpreadsheetApp.newDataValidation()
                    .requireValueInList(prioritiesArray)
                    .setAllowInvalid(true)
                    .build();
                if (issues.length > 0) {
                    sheet.getRange(2, priorityColIndex + 1, issues.length, 1).setDataValidation(rule);
                }
            }
        }

        // Apply Dropdown to Issue Type column
        const issueTypeColIndex = columns.indexOf('issuetype');
        if (issueTypeColIndex !== -1) {
            const uniqueTypes = new Set();
            issues.forEach(issue => {
                if (issue.fields && issue.fields.issuetype && issue.fields.issuetype.name) {
                    uniqueTypes.add(issue.fields.issuetype.name);
                }
            });
            const typesArray = Array.from(uniqueTypes).sort();

            if (typesArray.length > 0) {
                const rule = SpreadsheetApp.newDataValidation()
                    .requireValueInList(typesArray)
                    .setAllowInvalid(true)
                    .build();
                if (issues.length > 0) {
                    sheet.getRange(2, issueTypeColIndex + 1, issues.length, 1).setDataValidation(rule);
                }
            }
            // Cache types for later use (e.g. creating new columns)
            PropertiesService.getDocumentProperties().setProperty('cachedIssueTypes', JSON.stringify(typesArray));
        }

        // Apply Date Validation (Calendar Control) to Date Columns
        // User Request: Exclude system timestamps (created, updated). Only for Due Dates, Start/End dates.
        const dateKeywords = ['date', 'due', 'start', 'end', 'target'];
        const excludeKeywords = ['created', 'updated', 'resolved'];

        const dateColIndices = columns
            .map((col, index) => ({ col: col.toLowerCase(), index }))
            .filter(item => {
                const isDate = dateKeywords.some(k => item.col.includes(k));
                const isExcluded = excludeKeywords.some(k => item.col.includes(k));
                return isDate && !isExcluded;
            })
            .map(item => item.index);

        if (dateColIndices.length > 0 && issues.length > 0) {
            const dateRule = SpreadsheetApp.newDataValidation()
                .requireDate()
                .setAllowInvalid(false)
                .build();

            dateColIndices.forEach(idx => {
                // Apply to all rows for this column, extending to the bottom of the sheet
                // This ensures empty rows below are also ready for data entry
                const maxRows = sheet.getMaxRows();
                if (maxRows > 1) {
                    sheet.getRange(2, idx + 1, maxRows - 1, 1).setDataValidation(dateRule);
                    sheet.getRange(2, idx + 1, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd');
                }
            });
        }
    }
    const type = params.triggerType || 'Manual';
    const proxyInfo = params.proxyCallCount > 0 ? ` (via Worker x${params.proxyCallCount})` : '';
    const logDetails = `[${type}] ${issues.length || 0} issues fetched. ${changedCount || 0} changed.${proxyInfo}`;
    try { logRefreshAttempt('OK', logDetails, sheet.getName()); } catch (e) { }

    return {
        success: true,
        count: issues.length,
        changedCount: changedCount,
        proxyCalls: params.proxyCallCount || 0,
        limited: !!params.limited
    };
}

/**
 * Executes a bulk update based on JQL or JQLU syntax
 * @param {Object} config - {domain, email, token, jql}
 * @param {String} commandInput - JSON string OR JQLU syntax string
 */
function executeBulkUpdate(config, commandInput) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing Jira configuration.");
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');

    let updatePayload = {};
    let finalJql = config.jql; // Default to UI JQL

    // Detect Syntax: JSON vs JQLU
    const trimmedCommand = commandInput.trim();
    if (trimmedCommand.toUpperCase().startsWith('UPDATE ')) {
        try {
            const parsed = parseJqluCommand(trimmedCommand);
            updatePayload = parsed.updatePayload;
            finalJql = parsed.jql;
        } catch (e) {
            throw new Error("JQLU Error: " + e.message);
        }
    } else {
        // JSON Fallback
        try {
            updatePayload = JSON.parse(commandInput);
        } catch (e) {
            throw new Error("Invalid format: Must be valid JSON or 'UPDATE... WHERE...' syntax.");
        }
    }

    // 1. Fetch Issues to Update
    // Using POST for search to handle long JQL and avoid 410 errors
    const searchUrl = `https://${cleanDomain}/rest/api/3/search/jql`;
    const authHeader = "Basic " + Utilities.base64Encode(config.email + ":" + config.token);

    // Initial fetch to find issues - Paginate to avoid 1000 limit error
    let issues = [];
    let nextPageToken = null;
    try {
        while (true) {
            const response = fetchJira(searchUrl, {
                method: 'POST',
                headers: {
                    "Authorization": authHeader,
                    "Content-Type": "application/json"
                },
                payload: JSON.stringify({
                    jql: finalJql,
                    nextPageToken: nextPageToken,
                    maxResults: 100,
                    fields: ["key"]
                }),
                muteHttpExceptions: true
            }, config);

            if (response.getResponseCode() !== 200) {
                throw new Error(`Search failed (${response.getResponseCode()}): ${response.getContentText()}`);
            }

            const data = JSON.parse(response.getContentText());
            const batch = data.issues || [];
            if (batch.length === 0) break;

            issues = issues.concat(batch);
            nextPageToken = data.nextPageToken;

            if (!nextPageToken || batch.length < 100 || issues.length >= 5000) break; // Cap at 5000 for safety
        }
    } catch (e) {
        return { success: false, error: e.message };
    }

    if (issues.length === 0) {
        return { success: true, updated: 0, message: "No issues found to update." };
    }

    // 2. Perform Updates in batches
    // UrlFetchApp.fetchAll allows parallel requests.
    // We'll process in chunks to be safe with execution time limits.
    let successCount = 0;
    let failureCount = 0;
    let firstError = null;

    const batchSize = 20;

    // We only process up to 50 batches (1000 issues) max due to execution time limits in Apps Script
    // The search already limited to 1000.

    for (let i = 0; i < issues.length; i += batchSize) {
        const batch = issues.slice(i, i + batchSize);

        const requests = batch.map(issue => ({
            url: `https://${cleanDomain}/rest/api/3/issue/${issue.key}`,
            method: 'PUT',
            headers: {
                "Authorization": authHeader,
                "Content-Type": "application/json"
            },
            payload: JSON.stringify(updatePayload),
            muteHttpExceptions: true
        }));

        try {
            const responses = fetchAllJira(requests, config);
            responses.forEach((res, idx) => {
                if (res.getResponseCode() === 204 || res.getResponseCode() === 200) {
                    successCount++;
                } else {
                    failureCount++;
                    if (!firstError) {
                        try {
                            const err = JSON.parse(res.getContentText());
                            firstError = `${batch[idx].key}: ${JSON.stringify(err.errors || err.errorMessages)}`;
                        } catch (e) {
                            firstError = `${batch[idx].key}: ${res.getResponseCode()}`;
                        }
                    }
                }
            });
        } catch (e) {
            failureCount += batch.length;
            if (!firstError) firstError = "Batch failed: " + e.message;
        }

        // Brief pause between batches
        Utilities.sleep(100);
    }

    // 3. Refresh the sheet if updates were successful
    let refreshResult = null;
    if (successCount > 0) {
        try {
            // We use the CONFIG's JQL (the main sheet query), not the Update JQL.
            refreshResult = fetchJiraData(config);
        } catch (e) {
            console.error("Auto-refresh failed: " + e.message);
        }
    }

    return {
        success: failureCount === 0,
        updated: successCount,
        failed: failureCount,
        error: firstError,
        refreshed: refreshResult ? refreshResult.count : 0
    };
}

/**
 * Previews a bulk update to get match count
 */
function previewBulkUpdate(config, commandInput) {
    if (!config.domain || !config.email || !config.token) {
        return { error: "Missing Jira configuration." };
    }
    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const authHeader = "Basic " + Utilities.base64Encode(config.email + ":" + config.token);

    let finalJql = config.jql;

    const trimmedCommand = commandInput.trim();
    if (trimmedCommand.toUpperCase().startsWith('UPDATE ')) {
        try {
            const parsed = parseJqluCommand(trimmedCommand);
            finalJql = parsed.jql;
        } catch (e) {
            return { error: "JQLU Error: " + e.message };
        }
    }

    // Fetch count using approximate-count endpoint
    const countUrl = `https://${cleanDomain}/rest/api/3/search/approximate-count`;

    try {
        const response = fetchWithRetry(countUrl, {
            method: 'POST',
            headers: {
                "Authorization": authHeader,
                "Content-Type": "application/json",
                "Accept": "application/json"
            },
            payload: JSON.stringify({ jql: finalJql }),
            muteHttpExceptions: true
        }, config);

        if (response.getResponseCode() !== 200) {
            return { error: `Preview failed (${response.getResponseCode()}): ${response.getContentText()}` };
        }

        const data = JSON.parse(response.getContentText());

        return {
            jql: finalJql,
            count: data.count || 0
        };

    } catch (e) {
        return { error: "Preview execution failed: " + e.message };
    }
}

/**
 * Resets all issue row backgrounds to white and clears stored timestamps
 */
function resetIssueColors() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow > 1 && lastCol > 0) {
        // Reset data rows to white (skip header row)
        const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        dataRange.setBackground('#FFFFFF');
    }

    // Clear stored timestamps so next refresh shows all as unchanged
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty('issueTimestamps');

    return { success: true };
}

/**
 * Highlights specific rows with a color, then clears after a delay
 * @param {number[]} rows - Array of row numbers to highlight
 * @param {string} color - Hex color (default light green)
 * @param {number} delaySeconds - Seconds before clearing (default 10)
 */
function highlightRowsTemporarily(rows, color, delaySeconds) {
    if (!rows || rows.length === 0) return;
    color = color || '#d1fae5';
    delaySeconds = delaySeconds || 10;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastCol = sheet.getLastColumn();

    rows.forEach(row => {
        sheet.getRange(row, 1, 1, lastCol).setBackground(color);
    });
    SpreadsheetApp.flush();

    Utilities.sleep(delaySeconds * 1000);

    rows.forEach(row => {
        sheet.getRange(row, 1, 1, lastCol).setBackground(null);
    });
    SpreadsheetApp.flush();
}

/**
 * Fetches all available fields from Jira
 */
function getJiraFields(config) {
    if (!config.domain || !config.email || !config.token) {
        return [];
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const url = `https://${cleanDomain}/rest/api/3/field`;

    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token),
            "Accept": "application/json"
        },
        muteHttpExceptions: true
    };

    try {
        const response = fetchWithRetry(url, options, config);
        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());

            // Build and cache field name mapping (id/key -> friendly name)
            const fieldNameMap = {};
            data.forEach(f => {
                fieldNameMap[f.id] = f.name;
                if (f.key && f.key !== f.id) {
                    fieldNameMap[f.key] = f.name;
                }
            });
            // Cache the mapping for later use in fetchJiraData
            PropertiesService.getDocumentProperties().setProperty('fieldNameMap', JSON.stringify(fieldNameMap));

            // Return id and name, sorted by name.
            return data
                .map(f => ({ id: f.id, key: f.key, name: f.name, custom: f.custom }))
                .sort((a, b) => a.name.localeCompare(b.name));
        }
        return [];
    } catch (e) {
        console.error("getJiraFields failed: " + e.message);
        return [];
    }
}

/**
 * Pushes updates to Jira
 * @param {Object} params - config params
 * @param {Array} updates - Array of {row, key, updates: {field: value}}
 */
function updateJiraIssues(params, updates) {
    if (!updates || updates.length === 0) return { success: true, count: 0 };

    const cleanDomain = params.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');

    const requests = updates.map(update => {
        const url = `https://${cleanDomain}/rest/api/3/issue/${update.key}`;
        return {
            url: url,
            method: 'PUT',
            contentType: 'application/json',
            headers: {
                "Authorization": "Basic " + Utilities.base64Encode(params.email + ":" + params.token)
            },
            payload: JSON.stringify({ fields: update.fields }),
            muteHttpExceptions: true
        };
    });

    const responses = fetchAllJira(requests, params);
    const errors = [];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const successfulKeys = [];
    const rangesToClear = [];
    let updatedCells = 0;

    // Load Modified Keys to start (so we can remove successful ones)
    const props = PropertiesService.getDocumentProperties();
    let modifiedMap = {};
    try {
        const json = props.getProperty('modifiedKeys');
        if (json) modifiedMap = JSON.parse(json);
    } catch (err) { }

    responses.forEach((res, i) => {
        const update = updates[i];
        if (res.getResponseCode() === 204) {
            // Success
            successfulKeys.push(update.key);
            delete modifiedMap[update.key];

            // Count cells updated
            if (update.fields) {
                updatedCells += Object.keys(update.fields).length;
            }

            // Clear background for updated fields
            if (update.fieldCols) {
                Object.values(update.fieldCols).forEach(colIndex => {
                    // efficient batching? rangesToClear.push(...)
                    // colIndex is 0-based from getSelectedRowsData?
                    // Let's check getSelectedRowsData. 
                    // It says: fieldCols[fieldId] = colIndex; where colIndex comes from forEach((header, colIndex) ...
                    // colIndex is 0-based index of headers array.
                    // The sheet column is colIndex + 1.
                    rangesToClear.push(sheet.getRange(update.lineNumber, colIndex + 1).getA1Notation());
                });
            }
        } else {
            // Error handling
            const startRow = update.lineNumber; // Variable rename from 'row' to avoid confusion, though original used 'row'
            // let's stick to 'row' if local variable, but scope is inside loop.
            const row = update.lineNumber; // Re-declare for clarity as per original text
            const content = res.getContentText();
            let errorMessage = content;

            try {
                const errJson = JSON.parse(content);
                // Handle "errors" object { fieldId: message }
                if (errJson.errors) {
                    Object.entries(errJson.errors).forEach(([field, msg]) => {
                        // Find column index for this field
                        const colIndex = update.fieldCols ? update.fieldCols[field] : -1;
                        if (colIndex !== -1) {
                            // Highlight cell
                            sheet.getRange(row, colIndex + 1).setBackground('#ffcccc')
                                .setNote(`Error: ${msg}`); // Add note for detail
                        }
                    });
                    errorMessage = JSON.stringify(errJson.errors);
                }

                // Handle "errorMessages" array [ message, ... ]
                if (errJson.errorMessages && errJson.errorMessages.length > 0) {
                    errorMessage = errJson.errorMessages.join(', ');
                }
            } catch (e) {
                // Raw text error
            }

            errors.push(`Row ${row}: ${errorMessage}`);
        }
    });

    // Save updated Modified Keys
    if (successfulKeys.length > 0) {
        props.setProperty('modifiedKeys', JSON.stringify(modifiedMap));

        // Clear backgrounds in batch if possible
        if (rangesToClear.length > 0) {
            sheet.getRangeList(rangesToClear).setBackground(null); // OR 'white'
        }
    }

    if (errors.length > 0) {
        throw new Error(`Some updates failed. Check pink highlighted cells for details.\n\n${errors.slice(0, 5).join('\n')}${errors.length > 5 ? '\n...' : ''}`);
    }

    return { success: true, count: updates.length - errors.length, cellsUpdated: updatedCells };
}

/**
 * Fetches user's favorite filters from Jira
 */
function getJiraFilters(config) {
    if (!config.domain || !config.email || !config.token) {
        return [];
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    // Use search endpoint to get all visible filters, not just favorites
    // maxResults=100 should cover most users' immediate needs
    const url = `https://${cleanDomain}/rest/api/2/filter/search?maxResults=100&expand=jql`;

    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token)
        },
        muteHttpExceptions: true
    };

    try {
        const response = fetchJira(url, options, config);
        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            // Search endpoint returns { values: [...] }
            // We sort by name alphabetically for better UX
            const filters = (data.values || []).sort((a, b) => a.name.localeCompare(b.name));
            return filters.map(f => ({ id: f.id, name: f.name, jql: f.jql }));
        }
        return [];
    } catch (e) {
        return [];
    }
}

/**
 * Helper to get selected data for update
 */
function getSelectedRowsData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rangeList = sheet.getSelection().getActiveRangeList();
    let ranges = rangeList ? rangeList.getRanges() : [];

    // If no selection or selection is a single cell (default state), auto-detect all data rows
    const isSingleCell = ranges.length === 1 && ranges[0].getNumRows() === 1 && ranges[0].getNumColumns() === 1;
    if (ranges.length === 0 || isSingleCell) {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { updates: [], creates: [] };
        ranges = [sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())];
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = headers.map(h => h.toString().toLowerCase().trim()); // Normalize for index lookup
    const keyIndex = headerMap.indexOf('key');

    if (keyIndex === -1) {
        throw new Error("Column 'key' not found in header row (row 1).");
    }

    // Load Maps
    const props = PropertiesService.getDocumentProperties();
    const userMapJson = props.getProperty('userMap');
    const userMap = userMapJson ? JSON.parse(userMapJson) : {};

    // Load Field IDs (Reverse Map: Name -> ID)
    const fieldNameMapJson = props.getProperty('fieldNameMap');
    const fieldNameMap = fieldNameMapJson ? JSON.parse(fieldNameMapJson) : {};
    const fieldIdMap = {};

    // Populate standard fallbacks first (Case-sensitive names from Jira usually Title Case)
    fieldIdMap['Summary'] = 'summary';
    fieldIdMap['Project'] = 'project';
    fieldIdMap['Issue Type'] = 'issuetype';
    fieldIdMap['Priority'] = 'priority';
    fieldIdMap['Assignee'] = 'assignee';
    fieldIdMap['Status'] = 'status';
    fieldIdMap['Description'] = 'description';
    // Add dynamic fields
    Object.entries(fieldNameMap).forEach(([id, name]) => {
        if (!fieldIdMap[name]) fieldIdMap[name] = id;
    });

    // Identify mandatory columns for creation (use let since we may add missing columns)
    let summaryIndex = headerMap.findIndex(h => h === 'summary' || h === 'issue summary');
    let issueTypeIndex = headerMap.findIndex(h => h === 'issuetype' || h === 'issue type' || h === 'type');
    let projectIndex = headerMap.findIndex(h => h === 'project' || h === 'project key' || h === 'project name');

    // Helper to add a missing column
    const addMissingColumn = (colName, displayName) => {
        const newColIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, newColIndex).setValue(displayName);
        sheet.getRange(1, newColIndex).setFontWeight('bold').setBorder(null, null, true, null, null, null, '#dfe1e6', SpreadsheetApp.BorderStyle.SOLID);
        return newColIndex - 1; // Return 0-indexed position
    };

    // Auto-add mandatory columns if missing
    // Summary
    if (summaryIndex === -1) {
        summaryIndex = addMissingColumn('summary', 'Summary');
    }

    // Project  
    if (projectIndex === -1) {
        projectIndex = addMissingColumn('project', 'Project');
    }

    // Issue Type (with special handling for cached data)
    if (issueTypeIndex === -1) {
        const newColIndex = sheet.getLastColumn() + 1;

        // Add header
        sheet.getRange(1, newColIndex).setValue("Issue Type");
        sheet.getRange(1, newColIndex).setFontWeight('bold').setBorder(null, null, true, null, null, null, '#dfe1e6', SpreadsheetApp.BorderStyle.SOLID);

        // Populate values for rows with keys
        const issueTypeByKeyJson = props.getProperty('issueTypeByKey');
        const lastRow = sheet.getLastRow();

        if (issueTypeByKeyJson && lastRow > 1) {
            const issueTypeByKey = JSON.parse(issueTypeByKeyJson);
            const allKeys = sheet.getRange(2, keyIndex + 1, lastRow - 1, 1).getValues();
            const issueTypeData = allKeys.map(row => {
                const key = row[0];
                return [key && issueTypeByKey[key] ? issueTypeByKey[key] : ''];
            });
            sheet.getRange(2, newColIndex, issueTypeData.length, 1).setValues(issueTypeData);
        }

        // Apply dropdown validation
        const cachedTypesJson = props.getProperty('cachedIssueTypes');
        if (cachedTypesJson) {
            const types = JSON.parse(cachedTypesJson);
            if (types && types.length > 0) {
                const rule = SpreadsheetApp.newDataValidation()
                    .requireValueInList(types)
                    .setAllowInvalid(true)
                    .build();
                sheet.getRange(2, newColIndex, Math.max(lastRow - 1, 100), 1).setDataValidation(rule);
            }
        }

        // Update index since we added the column
        issueTypeIndex = newColIndex - 1; // Convert to 0-indexed for headerMap
    }

    const rowsToUpdate = [];
    const rowsToCreate = [];

    // Track validation errors
    let validationFailed = false;
    const missingFields = new Set(); // Track which fields are missing across all rows
    const rangesToHighlight = [];

    // Load Modified Keys for Smart Update
    const modifiedKeysJson = props.getProperty('modifiedKeys');
    const modifiedKeys = modifiedKeysJson ? JSON.parse(modifiedKeysJson) : {};
    const hasModifications = Object.keys(modifiedKeys).length > 0;

    // Check if selection is effectively "All Data" or "Large Range"
    let useSmartFilter = false;
    let totalSelectedRows = 0;
    ranges.forEach(r => totalSelectedRows += r.getNumRows());

    if (totalSelectedRows > 50) {
        useSmartFilter = true;
    }

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();

        // Skip header row if selected alone. If mixed, we handle inside loop.
        if (startRow === 1 && numRows === 1) return;

        // Effective data range (skip row 1 if included)
        const effectiveStartRow = startRow === 1 ? 2 : startRow;
        const effectiveNumRows = startRow === 1 ? numRows - 1 : numRows;

        if (effectiveNumRows <= 0) return;

        const dataRange = sheet.getRange(effectiveStartRow, 1, effectiveNumRows, sheet.getLastColumn());
        const values = dataRange.getValues();

        // Capture column range of selection
        const rangeStartCol = range.getColumn();
        const rangeEndCol = range.getLastColumn();
        const isSingleCell = (effectiveNumRows === 1 && rangeStartCol === rangeEndCol);

        values.forEach((rowValues, i) => {
            const absoluteRow = effectiveStartRow + i;
            const key = rowValues[keyIndex];

            // Common Read-Only Fields relevant for both operations (mostly) -> status CANNOT be set on create
            const readOnlyFields = ['status', 'created', 'updated', 'creator', 'resolution', 'resolutiondate', 'lastviewed', 'votes', 'watches', 'attachment', 'comment', 'worklog', 'issuelinks'];

            // LOGIC SPLIT: Update vs Create
            if (key) {
                // --- UPDATE LOGIC ---

                // SMART FILTER CHECK: If enabled, skip if key not modified
                // Exception: If user selected a SINGLE CELL, we always update it (explicit intent)
                if (useSmartFilter && !modifiedKeys[key] && !isSingleCell) {
                    return;
                }

                const fields = {};
                const fieldCols = {}; // Map fieldId -> colIndex

                headers.forEach((header, colIndex) => {
                    // Check if current column (1-based) is in selected range
                    const currentAbsCol = colIndex + 1;

                    // Optimization: If single cell selected, ONLY process that column
                    // If multi-selection, process all columns in selection
                    if (currentAbsCol < rangeStartCol || currentAbsCol > rangeEndCol) return;

                    if (colIndex === keyIndex) return; // Don't update key
                    if (!header) return;

                    // Resolve Field ID
                    const fieldId = fieldIdMap[header] || header.toString().toLowerCase().trim();
                    if (readOnlyFields.includes(fieldId)) return;

                    let val = rowValues[colIndex];

                    // Check undefined to allow empty string updates (clearing fields)
                    if (val === undefined) return;

                    // Common Value Logic
                    const addField = (v) => {
                        fields[fieldId] = v;
                        fieldCols[fieldId] = colIndex;
                    };

                    if (fieldId === 'assignee') {
                        const mappedId = userMap[val];
                        addField({ accountId: mappedId || val });
                    } else if (['priority', 'issuetype'].includes(fieldId)) {
                        addField({ name: val });
                    } else if (fieldId === 'project') {
                        addField({ key: val });
                    } else if (val instanceof Date) {
                        // Format Date objects to YYYY-MM-DD
                        const dateStr = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                        addField(dateStr);
                    } else {
                        // Check if value looks like a date string and needs formatting? 
                        // Google Sheets often returns Date objects for date cells, so the above covers it.
                        // If it's a string, we assume it's in correct format or plain text.
                        if (val === "") val = null;
                        addField(val);
                    }
                });

                if (Object.keys(fields).length > 0) {
                    rowsToUpdate.push({ lineNumber: absoluteRow, key: key, fields: fields, fieldCols: fieldCols });
                }

            } else {
                // --- CREATE LOGIC --- (Empty Key)

                // Check if row is effectively empty (No Summary AND No Type)
                // If both are missing, we skip validation and assume it's just an empty row in the selection
                const summaryValCheck = rowValues[summaryIndex];
                const issueTypeValCheck = (issueTypeIndex !== -1 && issueTypeIndex < rowValues.length) ? rowValues[issueTypeIndex] : '';

                const isSummaryEmpty = !summaryValCheck || summaryValCheck.toString().trim() === '';
                // Note: We check rowValues first. If we just added the column, it might be empty string in array.
                // We double check against current sheet value if needed, but for skipping, rowValues is good enough proxy for "did user type anything".
                // If user typed nothing, rowValues has empty strings.

                // However, we must be careful. issueTypeIndex might be the NEW column we just added.
                // If so, rowValues might not have it if we fetched values before adding column?
                // dataRange (line 857) is fetched inside loop, AFTER addMissingColumn. So rowValues includes it.
                // And since it's a new column, it's empty unless we populated it from cache.

                // So if Summary is empty, and Type is empty (likely), we SKIP.
                // If Type is populated (from cache/user) but Summary empty -> ERROR (missing summary).

                // Let's rely on Summary primarily? No, create requires both.
                // If both are empty -> Skip.

                // Re-read strictly for check
                const currentTypeVal = (issueTypeIndex !== -1) ? sheet.getRange(absoluteRow, issueTypeIndex + 1).getValue() : '';
                const isTypeEmpty = !currentTypeVal || currentTypeVal.toString().trim() === '';

                if (isSummaryEmpty && isTypeEmpty) {
                    return; // SKIP this row
                }

                // Validate Mandatory Fields
                let rowValid = true;

                // Check Summary
                const summaryVal = rowValues[summaryIndex];
                if (summaryIndex === -1 || !summaryVal || summaryVal.toString().trim() === '') {
                    rowValid = false;
                    missingFields.add('Summary');
                    if (summaryIndex !== -1) {
                        sheet.getRange(absoluteRow, summaryIndex + 1).setBackground('#ffcccc');
                    }
                }

                // Check Issue Type (column should exist now since we add it early if missing)
                // But need to check if the VALUE is empty for this row
                // Note: rowValues was cached before we added the column, so we need to re-read
                const currentIssueTypeVal = sheet.getRange(absoluteRow, issueTypeIndex + 1).getValue();
                if (!currentIssueTypeVal || currentIssueTypeVal.toString().trim() === '') {
                    rowValid = false;
                    missingFields.add('Issue Type');
                    sheet.getRange(absoluteRow, issueTypeIndex + 1).setBackground('#ffcccc');
                }

                if (!rowValid) {
                    validationFailed = true;
                } else {
                    // Build fields for creation
                    const fields = {};
                    headers.forEach((header, colIndex) => {
                        if (colIndex === keyIndex) return;
                        if (!header) return;

                        // Resolve Field ID
                        const fieldId = fieldIdMap[header] || header.toString().toLowerCase().trim();
                        if (readOnlyFields.includes(fieldId)) return;

                        let val = rowValues[colIndex];
                        if (val === "" || val === null || val === undefined) return; // SKIP EMPTY field values

                        // Reuse same mapping logic
                        if (fieldId === 'assignee') {
                            const mappedId = userMap[val];
                            // Ensure we have a valid ID or value
                            if (mappedId || val) {
                                fields[fieldId] = { accountId: mappedId || val };
                            }
                        } else if (['priority', 'issuetype'].includes(fieldId)) {
                            fields[fieldId] = { name: val };
                        } else if (fieldId === 'project') {
                            fields[fieldId] = { key: val };
                        } else {
                            fields[fieldId] = val;
                        }
                    });
                    rowsToCreate.push({ lineNumber: absoluteRow, fields: fields });
                }
            }
        });
    });

    if (validationFailed) {
        const fieldsList = Array.from(missingFields).join(', ');
        throw new Error(`Validation Failed: Missing required field(s): ${fieldsList}. Cells highlighted in pink.`);
    }

    return { updates: rowsToUpdate, creates: rowsToCreate, sheetName: sheet.getName() };
}

/**
 * Creates new issues in Jira
 */
function createJiraIssues(params, creates) {
    if (!creates || creates.length === 0) return { success: true, count: 0, types: {} };

    // Determine Pro status if not provided
    if (params && typeof params.isPro === 'undefined') {
        const license = checkLicense();
        params.isPro = license.allowed && license.plan !== 'free';
    }

    let isLimited = false;
    if (!params.isPro && creates.length > 150) {
        creates = creates.slice(0, 150);
        isLimited = true;
    }

    if (!params || !params.domain) {
        throw new Error("Configuration Error: Jira Domain is missing. Please check your settings.");
    }

    const cleanDomain = params.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    if (!cleanDomain) {
        throw new Error("Configuration Error: Jira Domain is invalid.");
    }

    const url = `https://${cleanDomain}/rest/api/3/issue`;
    const authHeader = "Basic " + Utilities.base64Encode(params.email + ":" + params.token);

    const requests = creates.map(create => {
        // Fallback to configured projectKey/projectId if not in fields
        if (!create.fields.project && params.projectKey) {
            create.fields.project = { key: params.projectKey };
        } else if (!create.fields.project && params.projectId) {
            create.fields.project = { id: params.projectId };
        }

        return {
            url: url,
            method: 'POST',
            contentType: 'application/json',
            headers: { "Authorization": authHeader },
            payload: JSON.stringify({ fields: create.fields }),
            muteHttpExceptions: true
        };
    });

    const responses = fetchAllJira(requests, params);
    const errors = [];
    const createdTypes = {};
    const createdRows = [];
    let successCount = 0;

    // Find key column to write back created keys
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const keyColIndex = headers.map(h => h.toString().toLowerCase().trim()).indexOf('key');

    responses.forEach((res, i) => {
        if (res.getResponseCode() === 201) {
            successCount++;
            const data = JSON.parse(res.getContentText());
            // Try to record type if available in payload
            const typeName = creates[i].fields.issuetype?.name || 'Issue';
            createdTypes[typeName] = (createdTypes[typeName] || 0) + 1;
            if (creates[i].lineNumber) {
                createdRows.push(creates[i].lineNumber);
                // Write the Jira key back to the sheet so rows stay in place on next refresh
                if (keyColIndex !== -1 && data.key) {
                    const link = `=HYPERLINK("https://${cleanDomain}/browse/${data.key}", "${data.key}")`;
                    sheet.getRange(creates[i].lineNumber, keyColIndex + 1).setFormula(link);
                }
            }
        } else {
            errors.push(`Row ${creates[i].lineNumber}: ${res.getContentText()}`);
        }
    });

    if (errors.length > 0) {
        throw new Error(`Some creations failed:\n${errors.join('\n')}`);
    }

    // Highlight created rows in light green
    const lastCol = sheet.getLastColumn();
    createdRows.forEach(row => {
        sheet.getRange(row, 1, 1, lastCol).setBackground('#d1fae5');
    });
    SpreadsheetApp.flush();

    // Clear highlights after 10 seconds
    Utilities.sleep(10000);
    createdRows.forEach(row => {
        sheet.getRange(row, 1, 1, lastCol).setBackground(null);
    });
    SpreadsheetApp.flush();

    return { success: true, count: successCount, types: createdTypes, limited: isLimited, createdRows: createdRows };
}

/**
 * Fetches available projects from Jira
 */
function getJiraProjects(config) {
    if (!config.domain || !config.email || !config.token) {
        return [];
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const url = `https://${cleanDomain}/rest/api/3/project`;

    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token),
            "Accept": "application/json"
        },
        muteHttpExceptions: true
    };

    try {
        const response = fetchJira(url, options, config);
        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            return data.map(p => ({ id: p.id, key: p.key, name: p.name }));
        }
        return [];
    } catch (e) {
        return [];
    }
}

/**
 * Generates a Dashboard sheet with charts/**
 * Generates a dashboard with charts on the active sheet
 * @param {string} chartStyle - 'pie', 'donut', or 'bar'
 * @param {string[]} selectedColumns - Array of column names to visualize
 */
function generateDashboard(chartStyle, selectedColumns, sourceSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = sourceSheetName ? ss.getSheetByName(sourceSheetName) : ss.getActiveSheet();
    if (!dataSheet) throw new Error("Source sheet not found: " + sourceSheetName);

    // Validate this looks like a Jira sheet
    const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const normalizedHeaders = headers.map(h => String(h).toLowerCase().trim());

    // Check for key column (case-insensitive)
    if (!normalizedHeaders.includes('key')) {
        throw new Error("Active sheet does not appear to contain Jira data (missing 'Key' column).");
    }

    // Get Data
    const data = dataSheet.getDataRange().getValues();
    const rows = data.slice(1); // skip header

    if (rows.length === 0) throw new Error("No data to visualize.");

    // Default columns if none selected
    const userColumns = (selectedColumns && selectedColumns.length > 0)
        ? selectedColumns
        : ['status', 'issue type', 'priority', 'assignee'];

    console.log(`[Dashboard] Generating for columns: ${JSON.stringify(userColumns)}`);

    // Helper for fuzzy matching headers
    const findHeaderIndex = (name) => {
        const cleanName = String(name).toLowerCase().replace(/\s/g, '').trim();
        return headers.findIndex(h => {
            const cleanH = String(h).toLowerCase().replace(/\s/g, '').trim();
            return cleanH === cleanName || cleanH.includes(cleanName);
        });
    };

    // Deduplicate and validate
    const uniqueIndices = new Set();
    const columnsToVisualize = [];

    userColumns.forEach(col => {
        const idx = findHeaderIndex(col);
        if (idx !== -1 && !uniqueIndices.has(idx)) {
            uniqueIndices.add(idx);
            columnsToVisualize.push(headers[idx]);
        }
    });

    const skipped = [];
    const MAX_UNIQUE_VALUES = 50;

    const getCounts = (colName) => {
        const idx = findHeaderIndex(colName);
        if (idx === -1) return null;

        const counts = {};
        rows.forEach(r => {
            let val = r[idx];
            if (val === "" || val === null || val === undefined) val = '(Empty)';
            if (Array.isArray(val)) val = val.join(', ') || '(Empty)';
            counts[val] = (counts[val] || 0) + 1;
        });
        return counts;
    };

    // Use Active Sheet
    let dashSheet = dataSheet;
    const lastDataCol = dashSheet.getLastColumn();
    // We store the dashboard in a safe zone to the right of data
    const startX = lastDataCol + 2;

    // Layout constants
    const CHARTS_PER_ROW = 2;
    const CHART_HEIGHT_ROWS = 20;

    const ADDON_TAG = "JiraAddonDashboard";

    // SURGICAL CLEAR: Only remove what we created
    try {
        // 1. Remove Addon Charts
        const existingCharts = dashSheet.getCharts();
        existingCharts.forEach(chart => {
            const description = chart.getOptions().get('description');
            if (description === ADDON_TAG) {
                dashSheet.removeChart(chart);
            }
        });

        // 2. Clear Addon Tables (starting from Dashboard heading)
        const headerCell = dashSheet.getRange(1, startX);
        if (headerCell.getValue() === "Dashboard") {
            const colsToClear = CHARTS_PER_ROW * 6;
            dashSheet.getRange(1, startX, dashSheet.getMaxRows(), colsToClear).clear();
        }

        // Write fresh heading
        dashSheet.getRange(1, startX).setValue("Dashboard").setFontSize(24).setFontWeight("bold");
    } catch (e) {
        console.warn("Could not selectively clear dashboard area: " + e.message);
    }

    let currentRow = 3;
    let chartsCreated = 0;

    const createSection = (title, counts, colOffset) => {
        if (!counts) return false;
        const keys = Object.keys(counts);
        if (keys.length > MAX_UNIQUE_VALUES) return false;

        const startCol = startX + colOffset;
        dashSheet.getRange(currentRow, startCol).setValue(title).setFontWeight("bold").setFontSize(14);
        const tableData = keys.map(k => [k, counts[k]]).sort((a, b) => b[1] - a[1]);

        // Table
        dashSheet.getRange(currentRow + 2, startCol, tableData.length, 2).setValues(tableData);
        dashSheet.getRange(currentRow + 1, startCol, 1, 2).setValues([["Category", "Count"]]).setFontWeight("bold").setBackground("#f4f5f7");

        // Chart
        const range = dashSheet.getRange(currentRow + 1, startCol, tableData.length + 1, 2);
        let chartType = chartStyle === 'bar' ? Charts.ChartType.BAR : Charts.ChartType.PIE;

        let chartBuilder = dashSheet.newChart()
            .setChartType(chartType)
            .addRange(range)
            .setPosition(currentRow, startCol + 2, 0, 0)
            .setOption('title', title + ' Distribution')
            .setOption('description', ADDON_TAG) // TAG for surgical clearing
            .setOption('width', 400)
            .setOption('height', 350);

        if (chartStyle === 'donut') chartBuilder = chartBuilder.setOption('pieHole', 0.4);

        dashSheet.insertChart(chartBuilder.build());
        return true;
    };

    columnsToVisualize.forEach((colName) => {
        const counts = getCounts(colName);
        if (!counts) {
            skipped.push(colName + ' (not found)');
            return;
        }

        if (Object.keys(counts).length > MAX_UNIQUE_VALUES) {
            skipped.push(colName + ' (too many values)');
            return;
        }

        const colOffset = (chartsCreated % CHARTS_PER_ROW) * 6;
        if (chartsCreated > 0 && chartsCreated % CHARTS_PER_ROW === 0) {
            currentRow += CHART_HEIGHT_ROWS;
        }

        const displayName = colName.toString().split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
        if (createSection(displayName, counts, colOffset)) {
            chartsCreated++;
        }
    });

    if (chartsCreated === 0) {
        throw new Error("No columns could be visualized. " + (skipped.length > 0 ? "Skipped: " + skipped.join(", ") : ""));
    }

    return { success: true, chartsCreated: chartsCreated, skipped: skipped };
}



/**
 * Disables any scheduled background refreshes
 */
function disableScheduledRefresh() {
    const props = PropertiesService.getDocumentProperties();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetId = sheet.getSheetId().toString();

    // 1. Delete Trigger for this sheet
    const triggerMapJson = props.getProperty('triggerMap');
    let triggerMap = triggerMapJson ? JSON.parse(triggerMapJson) : {};

    const triggerId = Object.keys(triggerMap).find(key => triggerMap[key] === sheetId);

    if (triggerId) {
        const triggers = ScriptApp.getProjectTriggers();
        const trigger = triggers.find(t => t.getUniqueId() === triggerId);
        if (trigger) {
            ScriptApp.deleteTrigger(trigger);
        }
        delete triggerMap[triggerId];
        props.setProperty('triggerMap', JSON.stringify(triggerMap));
    }

    // 2. Clear Sheet Schedule Config
    props.deleteProperty(`scheduledConfig_${sheetId}`);
    props.deleteProperty(`scheduledSheetName_${sheetId}`);
    // But keep config_${sheetId} (JQL/Columns)

    return { success: true };
}

/**
 * Helper to clear all scheduled refresh properties
 */
function clearScheduledConfig(props) {
    props.deleteProperty('scheduledConfig');
    props.deleteProperty('scheduledRefreshTime');
    props.deleteProperty('scheduledSheetName');
    props.deleteProperty('scheduledRetryCount');
}

/**
 * Generates a visual Roadmap sheet with Gantt-style timeline
 */
/**
 * Generates a Roadmap sheet
 * @param {Object} options - Configuration options { durationMode: 'smart'|'resolution'|'duedate' }
 */
function generateRoadmap(options, sourceSheetName) {
    const durationMode = (options && options.durationMode) || 'smart';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSheetName ? ss.getSheetByName(sourceSheetName) : ss.getActiveSheet();
    if (!sourceSheet) throw new Error("Source sheet not found: " + sourceSheetName);

    const sheetName = sourceSheet.getName();
    if (sheetName === 'Jira Roadmap' || sheetName === 'Jira Capacity Plan') {
        throw new Error(`Cannot generate roadmap from the "${sheetName}" sheet. Please switch to your Jira Issues data sheet first.`);
    }

    // Get data from the current sheet
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow < 2) {
        throw new Error("No data found in the current sheet. Please fetch Jira data first.");
    }

    const headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const dataDisplay = dataRange.getDisplayValues();

    // Find relevant column indices
    // Normalize headers: remove special chars, lowercase
    const normalize = (s) => s.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
    const headerMap = headers.map(normalize);

    // Helper to find best column index
    const findBestColumn = (primaryKeyword, aliases = []) => {
        const searchTerms = [primaryKeyword, ...aliases].map(normalize);
        const indices = [];

        headerMap.forEach((h, i) => {
            if (searchTerms.includes(h)) indices.push(i);
        });

        // Also try 'contains' if exact match fails? 
        // Let's stick to exact match against normalized string first to avoid false positives (e.g. 'status' in 'subtask status')
        // But 'issue status' normalized is 'issuestatus'.

        if (indices.length === 0) return -1;

        // Prefer column with data in first 10 rows
        for (const idx of indices) {
            const hasData = data.slice(0, 50).some(row => row[idx] && row[idx].toString().trim() !== '');
            if (hasData) return idx;
        }
        return indices[0];
    };

    const keyIndex = findBestColumn('key', ['issuekey', 'issueid', 'issue']);
    const summaryIndex = findBestColumn('summary', ['issuesummary', 'title']);
    const statusIndex = findBestColumn('status', ['issuestatus', 'currentstatus', 'jirastatus', 'state', 'stat']);
    const createdIndex = findBestColumn('created', ['createddate', 'datecreated', 'createdate']);
    const resolvedIndex = findBestColumn('resolutiondate', ['resolved', 'resolveddate']);
    const issueTypeIndex = findBestColumn('issuetype', ['type', 'issuetype']);
    const priorityIndex = findBestColumn('priority');
    const assigneeIndex = findBestColumn('assignee', ['assignedto', 'assigned']);

    // Map user-selected date columns
    let startDateIndex = -1;
    let dueDateIndex = -1;

    if (options && options.startDateCol) {
        startDateIndex = headers.indexOf(options.startDateCol);
    }
    if (options && options.endDateCol) {
        dueDateIndex = headers.indexOf(options.endDateCol);
    }

    // Fallback logic if selection failed or not provided
    if (startDateIndex === -1) {
        startDateIndex = findBestColumn('startdate', ['start', 'customfield_10015']);
    }
    if (dueDateIndex === -1) {
        dueDateIndex = findBestColumn('duedate', ['due', 'duedate']);
    }

    // Auto-Fix: If Required or Date Columns are Missing, Add them and Refetch
    const isMissingCore = (keyIndex === -1 || summaryIndex === -1 || statusIndex === -1);
    const isMissingDates = (dueDateIndex === -1 || startDateIndex === -1);

    if ((isMissingCore || isMissingDates) && (!options || !options.isRetry)) {
        const props = PropertiesService.getDocumentProperties();
        const sheetId = sourceSheet.getSheetId();
        const configKey = `config_${sheetId}`;
        const sheetJson = props.getProperty(configKey);

        if (sheetJson) {
            let sheetConfig = JSON.parse(sheetJson);
            let currentCols = sheetConfig.columns ? sheetConfig.columns.split(',').map(c => c.trim()) : [];
            let added = false;
            const normalizedCurrentCols = currentCols.map(normalize);

            // 1. Core Columns
            if (keyIndex === -1 && !normalizedCurrentCols.includes('key')) {
                currentCols.push('key');
                added = true;
            }
            if (summaryIndex === -1 && !normalizedCurrentCols.includes('summary')) {
                currentCols.push('summary');
                added = true;
            }
            if (statusIndex === -1 && !normalizedCurrentCols.includes('status')) {
                currentCols.push('status');
                added = true;
            }

            // 2. Date Columns
            if (dueDateIndex === -1 && !normalizedCurrentCols.includes('duedate')) {
                currentCols.push('duedate');
                added = true;
            }
            if (startDateIndex === -1 && !normalizedCurrentCols.includes('customfield10015') && !normalizedCurrentCols.includes('startdate')) {
                currentCols.push('customfield_10015');
                added = true;
            }

            if (added) {
                // Save new config
                sheetConfig.columns = currentCols.join(', ');
                props.setProperty(configKey, JSON.stringify(sheetConfig));

                // Fetch Data with new columns
                const credsJson = props.getProperty('jira_creds');
                if (credsJson) {
                    const creds = JSON.parse(credsJson);
                    const fetchParams = {
                        ...creds,
                        jql: sheetConfig.jql,
                        columns: sheetConfig.columns,
                        targetSheetName: sourceSheet.getName(),
                        triggerType: 'RoadmapAutoFix'
                    };
                    try {
                        fetchJiraData(fetchParams);
                        // Recursive call with retry flag to avoid infinite loops
                        return generateRoadmap({ ...options, isRetry: true }, sourceSheet.getName());
                    } catch (e) {
                        console.warn("Auto-fetch failed: " + e.message);
                    }
                }
            }
        }
    }

    if (keyIndex === -1 || summaryIndex === -1 || statusIndex === -1) {
        throw new Error("Required columns (Key, Summary, Status) not found in the active sheet. Please fetch these columns first.");
    }

    // Debug Feedback for User
    const foundStatus = statusIndex !== -1 ? headers[statusIndex] : 'N/A';
    const sampleStatus = data[0] && statusIndex !== -1 ? data[0][statusIndex] : 'Empty';
    const foundResolved = resolvedIndex !== -1 ? headers[resolvedIndex] : 'N/A';
    const sampleResolved = data[0] && resolvedIndex !== -1 ? data[0][resolvedIndex] : 'Empty';

    // Debug Feedback for User - Removed as per request
    // ss.toast(`Status: '${foundStatus}' | Resolved: '${foundResolved}' (${sampleResolved})`, 'Roadmap Debug', 8);

    // Create or get Roadmap sheet
    let roadmapSheet = ss.getSheetByName('Jira Roadmap');
    if (roadmapSheet) {
        roadmapSheet.setFrozenRows(0);
        roadmapSheet.setFrozenColumns(0);
        roadmapSheet.clear();
    } else {
        roadmapSheet = ss.insertSheet('Jira Roadmap');
    }

    // Calculate date range for timeline based on data
    const { startDate, endDate, totalWeeks, msPerWeek } = calculateRoadmapDateRange(data, {
        createdIndex, startDateIndex, dueDateIndex, resolvedIndex
    });

    // Fixed columns: Key, Summary, Status, Type, Assignee, Resolved
    const fixedCols = 6;

    // Ensure enough columns exist before setting widths
    const requiredCols = fixedCols + totalWeeks;
    const currentCols = roadmapSheet.getMaxColumns();
    if (requiredCols > currentCols) {
        roadmapSheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
    }

    // Set column widths
    roadmapSheet.setColumnWidth(1, 100);  // Key
    roadmapSheet.setColumnWidth(2, 280);  // Summary
    roadmapSheet.setColumnWidth(3, 100);  // Status
    roadmapSheet.setColumnWidth(4, 90);   // Type
    roadmapSheet.setColumnWidth(5, 120);  // Assignee
    roadmapSheet.setColumnWidth(6, 110);  // Resolved

    // Timeline column widths
    for (let i = 0; i < totalWeeks; i++) {
        roadmapSheet.setColumnWidth(fixedCols + 1 + i, 25);
    }

    // Status colors
    const statusColors = {
        'to do': '#3b82f6',       // Blue
        'open': '#3b82f6',        // Blue
        'backlog': '#6b7280',     // Gray
        'in progress': '#f59e0b', // Amber
        'in review': '#8b5cf6',   // Purple
        'done': '#10b981',        // Green
        'closed': '#10b981',      // Green
        'resolved': '#10b981'     // Green
    };

    // Build header row
    const headerRow = ['Key', 'Summary', 'Status', 'Type', 'Assignee', 'Resolved'];

    // Add month markers
    const monthHeaders = [];
    let currentDate = new Date(startDate);
    for (let i = 0; i < totalWeeks; i++) {
        const weekStart = new Date(startDate.getTime() + i * msPerWeek);
        // Show month name on first week of month
        if (weekStart.getDate() <= 7) {
            monthHeaders.push(weekStart.toLocaleString('default', { month: 'short' }));
        } else {
            monthHeaders.push('');
        }
    }
    headerRow.push(...monthHeaders);

    // Write header row
    roadmapSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

    // Style header row
    const headerRange = roadmapSheet.getRange(1, 1, 1, headerRow.length);
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // Add week number sub-header (row 2)
    const weekRow = ['', '', '', '', '', ''];
    for (let i = 0; i < totalWeeks; i++) {
        const weekNum = Math.floor(i % 4) + 1;
        weekRow.push(`W${weekNum}`);
    }
    roadmapSheet.getRange(2, 1, 1, weekRow.length).setValues([weekRow]);
    roadmapSheet.getRange(2, 1, 1, weekRow.length).setBackground('#334155');
    roadmapSheet.getRange(2, 1, 1, weekRow.length).setFontColor('#94a3b8');
    roadmapSheet.getRange(2, 1, 1, weekRow.length).setFontSize(8);

    // Today marker column
    const today = new Date();
    const todayWeekOffset = Math.floor((today - startDate) / msPerWeek);
    if (todayWeekOffset >= 0 && todayWeekOffset < totalWeeks) {
        roadmapSheet.getRange(1, fixedCols + 1 + todayWeekOffset, 2, 1).setBackground('#ef4444');
    }

    // Retrieve domain for hyperlinking
    const props = PropertiesService.getDocumentProperties();
    const credsJson = props.getProperty('jira_creds');
    let domain = '';
    if (credsJson) {
        try {
            const config = JSON.parse(credsJson);
            domain = config.domain ? config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '') : '';
        } catch (e) {
            console.error("Failed to parse credentials for domain linking");
        }
    }

    // Process each issue
    let rowNum = 3;
    console.log(`[Roadmap Debug] KeyIndex: ${keyIndex}, StatusIndex: ${statusIndex}`);

    data.forEach((row, idx) => {
        const rowDisplay = dataDisplay[idx];
        const key = row[keyIndex];
        if (!key) return;

        // Use DISPLAY values for text fields
        let status = String(rowDisplay[statusIndex] || '').trim();

        // Debug first 5 rows
        if (idx < 5) {
            console.log(`[Roadmap Debug] Row ${idx}: Key=${key}, StatusRaw='${row[statusIndex]}', StatusDisplay='${rowDisplay[statusIndex]}', FinalStatus='${status}'`);
        }

        // Fallback: If display value is empty, try raw value
        if (!status) {
            status = String(row[statusIndex] || '').trim();
        }

        // Handle if status is an object string (e.g. {"name":"To Do"...}) - Just in case
        if (status.startsWith('{') && status.includes('name')) {
            try {
                const sObj = JSON.parse(status);
                if (sObj.name) status = sObj.name;
            } catch (e) { }
        }

        const summary = String(rowDisplay[summaryIndex] || '').trim();
        const issueType = issueTypeIndex !== -1 ? String(rowDisplay[issueTypeIndex] || '').trim() : '';
        const assignee = assigneeIndex !== -1 ? String(rowDisplay[assigneeIndex] || '').trim() : '';
        const created = createdIndex !== -1 ? row[createdIndex] : null;
        const dueDate = dueDateIndex !== -1 ? row[dueDateIndex] : null;
        const resolved = resolvedIndex !== -1 ? row[resolvedIndex] : null;

        // Write Fixed Columns
        const keyCell = roadmapSheet.getRange(rowNum, 1);
        if (domain) {
            const url = `https://${domain}/browse/${key}`;
            const richText = SpreadsheetApp.newRichTextValue()
                .setText(key)
                .setLinkUrl(url)
                .build();
            keyCell.setRichTextValue(richText);
        } else {
            keyCell.setValue(key);
        }

        roadmapSheet.getRange(rowNum, 2).setValue(summary.toString().substring(0, 60) + (summary.length > 60 ? '...' : ''));
        roadmapSheet.getRange(rowNum, 3).setValue(status);
        roadmapSheet.getRange(rowNum, 4).setValue(issueType);
        roadmapSheet.getRange(rowNum, 5).setValue(assignee);

        // Format resolved date
        if (resolved) {
            const resolvedDate = new Date(resolved);
            if (!isNaN(resolvedDate.getTime())) {
                roadmapSheet.getRange(rowNum, 6).setValue(resolvedDate).setNumberFormat('yyyy-mm-dd');
            } else {
                roadmapSheet.getRange(rowNum, 6).setValue('');
            }
        } else {
            roadmapSheet.getRange(rowNum, 6).setValue('');
        }

        // Determine bar color based on status
        const statusLower = status.toString().toLowerCase().trim();
        let barColor = '#64748b'; // Default Slate-500 (visible gray)
        let matchFound = false;

        // Exact match first
        if (statusColors[statusLower]) {
            barColor = statusColors[statusLower];
            matchFound = true;
        } else {
            // Fuzzy match
            for (const [key, color] of Object.entries(statusColors)) {
                if (statusLower.includes(key) || key.includes(statusLower)) {
                    barColor = color;
                    matchFound = true;
                    break;
                }
            }
        }

        // If still no match and status is not empty, generate a consistent color hash?
        // For now, stick to default Gray if unknown.

        if (idx < 5) console.log(`[Roadmap Debug] Row ${idx} Status: '${status}', Color: ${barColor}`);

        // Color status pill (Column 3)
        const statusRange = roadmapSheet.getRange(rowNum, 3);
        statusRange.setValue(status); // Ensure value is set
        statusRange.setBackground(null); // Remove background color
        statusRange.setFontColor('#000000'); // Set text to black
        statusRange.setFontWeight('normal'); // Match other columns
        statusRange.setHorizontalAlignment('left'); // Match other columns

        let barStart = 0;
        let barEnd = 2; // Default 2 weeks

        const startDateValue = startDateIndex !== -1 ? row[startDateIndex] : null;

        if (startDateValue) {
            const derivedStart = new Date(startDateValue);
            barStart = Math.max(0, Math.floor((derivedStart - startDate) / msPerWeek));
        } else if (created) {
            const createdDate = new Date(created);
            barStart = Math.max(0, Math.floor((createdDate - startDate) / msPerWeek));
        }

        // Bar ends at: Resolved Date > Due Date > Created + 3 weeks (fallback)
        // Calculate End Position based on Mode
        const mode = (options && options.durationMode) || 'smart';
        let targetEnd = null;

        if (mode === 'smart') {
            if (resolved) targetEnd = new Date(resolved);
            else if (dueDate) targetEnd = new Date(dueDate);
        } else if (mode === 'resolution') {
            if (resolved) targetEnd = new Date(resolved);
        } else if (mode === 'duedate') {
            if (dueDate) targetEnd = new Date(dueDate);
        }

        if (targetEnd && !isNaN(targetEnd.getTime())) {
            barEnd = Math.min(totalWeeks - 1, Math.floor((targetEnd - startDate) / msPerWeek));
        } else {
            // Fallback: Show bar for 3 weeks from created if no end date
            barEnd = Math.min(totalWeeks - 1, barStart + 3);
        }

        // Make sure barEnd >= barStart
        if (barEnd < barStart) barEnd = barStart;

        // Draw the bar
        if (barStart < totalWeeks && barEnd >= 0) {
            const startCol = fixedCols + 1 + Math.max(0, barStart);
            const endCol = fixedCols + 1 + Math.min(totalWeeks - 1, barEnd);
            const barWidth = endCol - startCol + 1;

            if (barWidth > 0) {
                roadmapSheet.getRange(rowNum, startCol, 1, barWidth).setBackground(barColor);
            }
        }

        // Alternate row colors for readability
        if (idx % 2 === 0) {
            roadmapSheet.getRange(rowNum, 1, 1, fixedCols).setBackground('#f8fafc');
        } else {
            roadmapSheet.getRange(rowNum, 1, 1, fixedCols).setBackground('#ffffff');
        }

        rowNum++;
    });

    // Freeze header rows and fixed columns
    roadmapSheet.setFrozenRows(2);
    roadmapSheet.setFrozenColumns(fixedCols);

    // ── LEGEND ──
    const legendStartRow = rowNum + 2;

    // Legend title
    roadmapSheet.getRange(legendStartRow, 1, 1, 2).merge().setValue('📊  ROADMAP LEGEND').setFontWeight('bold').setFontSize(11);

    // ── 1. Status Colors ──
    const colorLegendRow = legendStartRow + 2;
    roadmapSheet.getRange(colorLegendRow, 1).setValue('Status Colors:').setFontWeight('bold');

    const legendItems = [
        { label: 'To Do / Open', color: '#3b82f6' },
        { label: 'Backlog', color: '#6b7280' },
        { label: 'In Progress', color: '#f59e0b' },
        { label: 'In Review', color: '#8b5cf6' },
        { label: 'Done / Resolved / Closed', color: '#10b981' },
    ];
    legendItems.forEach((item, i) => {
        const r = colorLegendRow + 1 + i;
        roadmapSheet.getRange(r, 1).setValue('').setBackground(item.color);
        roadmapSheet.getRange(r, 2, 1, 2).merge().setValue(item.label).setFontColor('#334155');
    });

    // ── 2. Timeline Bars ──
    const barsRow = colorLegendRow + 1 + legendItems.length + 1;
    roadmapSheet.getRange(barsRow, 1).setValue('Timeline Bars:').setFontWeight('bold');
    // Define bar explanation based on mode
    const startName = (options && options.startDateCol) || 'Start Date';
    const endName = (options && options.endDateCol) || 'End Date';

    const barExplanations = [
        'Each colored square = 1 week on the timeline.',
        `Bar spans from '${startName}' to '${endName}'.`,
        'Bar color matches the issue\'s current status (see above).',
    ];
    barExplanations.forEach((text, i) => {
        roadmapSheet.getRange(barsRow + 1 + i, 1).setValue('  •').setFontColor('#94a3b8');
        roadmapSheet.getRange(barsRow + 1 + i, 2, 1, 4).merge().setValue(text).setFontColor('#475569').setFontStyle('italic');
    });

    // ── 3. Today Marker ──
    const todayRow = barsRow + 1 + barExplanations.length + 1;
    roadmapSheet.getRange(todayRow, 1).setValue('Today Marker:').setFontWeight('bold');
    roadmapSheet.getRange(todayRow + 1, 1).setValue('').setBackground('#ef4444');
    roadmapSheet.getRange(todayRow + 1, 2, 1, 3).merge().setValue('The red column in the timeline marks the current week.').setFontColor('#475569').setFontStyle('italic');

    // ── 4. Columns ──
    const colsRow = todayRow + 3;
    roadmapSheet.getRange(colsRow, 1).setValue('Columns:').setFontWeight('bold');
    const colExplanations = [
        'Key — Jira issue key (clickable link to Jira).',
        'Summary — Issue title (truncated to 60 chars).',
        'Status — Current workflow status (color-coded).',
        'Type — Issue type (Bug, Story, Task, etc.).',
        'Assignee — Person assigned to the issue.',
        'Resolved — Date the issue was resolved (empty if unresolved).',
    ];
    colExplanations.forEach((text, i) => {
        roadmapSheet.getRange(colsRow + 1 + i, 1).setValue('  •').setFontColor('#94a3b8');
        roadmapSheet.getRange(colsRow + 1 + i, 2, 1, 4).merge().setValue(text).setFontColor('#475569').setFontStyle('italic');
    });

    // ── 5. EXECUTIVE SUMMARY ──
    const summaryRow = colsRow + colExplanations.length + 3;
    const roadmapIndices = { keyIndex, summaryIndex, statusIndex, createdIndex, dueDateIndex, resolvedIndex, startDateIndex };
    addRoadmapAnalysisSection(roadmapSheet, summaryRow, data, roadmapIndices);

    // Activate the roadmap sheet
    roadmapSheet.activate();

    return { success: true };
}

/**
 * Generates a Capacity Planning sheet with team workload visualization
 * @param {Object} capacityConfig - Configuration for capacity calculation
 */
function generateCapacityPlan(capacityConfig, sourceSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSheetName ? ss.getSheetByName(sourceSheetName) : ss.getActiveSheet();
    if (!sourceSheet) throw new Error("Source sheet not found: " + sourceSheetName);

    const sheetName = sourceSheet.getName();
    if (sheetName === 'Jira Capacity Plan' || sheetName === 'Jira Roadmap') {
        throw new Error(`Cannot generate capacity plan from the "${sheetName}" sheet. Please switch to your Jira Issues data sheet first.`);
    }

    // Unfreeze to avoid merged cell conflicts during regeneration
    sourceSheet.setFrozenRows(0);
    sourceSheet.setFrozenColumns(0);

    // Get data from the current sheet
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow < 2) {
        throw new Error("No data found. Please fetch Jira data first.");
    }

    const headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Normalize headers: remove special chars, lowercase
    const normalize = (s) => (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
    const headerMap = headers.map(normalize);
    const estimationType = capacityConfig.estimationType || 'storypoints';

    // ── COLUMN MAPPING ──
    // Helper to find index from "Name (Letter)" format
    const findColIndex = (selectedName) => {
        if (!selectedName) return -1;
        const match = selectedName.match(/\(([A-Z]+)\)$/);
        if (match) {
            return letterToColumn(match[1]) - 1;
        }
        return headerMap.indexOf(normalize(selectedName));
    };

    let assigneeIndex = findColIndex(capacityConfig.assigneeCol);
    if (assigneeIndex === -1) assigneeIndex = headerMap.indexOf('assignee');

    let estimateIndex = findColIndex(capacityConfig.estimateCol);
    if (estimateIndex === -1) {
        const storyPointsIndex = headerMap.findIndex(h => h.includes('storypoint') || ['estimate', 'points', 'storypoints'].includes(h));
        const timeEstimateIndex = headerMap.indexOf('timeoriginalestimate');
        estimateIndex = estimationType === 'storypoints' ? storyPointsIndex : timeEstimateIndex;
    }

    let dueDateIndex = findColIndex(capacityConfig.dueDateCol);
    if (dueDateIndex === -1) dueDateIndex = headerMap.indexOf('duedate');

    let startDateIndex = findColIndex(capacityConfig.startDateCol);

    // Other non-selectable columns (static auto-detect)
    const statusIndex = headerMap.indexOf('status');
    const summaryIndex = headerMap.indexOf('summary');
    const keyIndex = headerMap.indexOf('key');

    if (assigneeIndex === -1) {
        throw new Error(`Assignee column not found. Selected: "${capacityConfig.assigneeCol}". Found headers: [${headers.join(', ')}]. Please ensure your Jira data is loaded on the active sheet.`);
    }

    if (estimateIndex === -1) {
        throw new Error(`${estimationType === 'storypoints' ? 'Story Points' : 'Time Estimate'} column not found. Selected: "${capacityConfig.estimateCol}". Please ensure the column is present.`);
    }

    // Parse config
    const hoursPerPoint = capacityConfig.hoursPerPoint || 1;
    const hoursPerWeek = capacityConfig.hoursPerWeek || 40;
    const capacityPercent = (capacityConfig.capacityPercent || 80) / 100;
    const weeksAhead = capacityConfig.weeksAhead || 8;

    // Parse holidays
    const holidays = [];
    if (capacityConfig.holidays) {
        capacityConfig.holidays.split(',').forEach(d => {
            const date = new Date(d.trim());
            if (!isNaN(date)) holidays.push(date);
        });
    }

    // Calculate week boundaries
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of day for current week calc

    const weeks = [];
    for (let i = 0; i < weeksAhead; i++) {
        const weekStart = new Date(today);
        weekStart.setDate(today.getDate() + (i * 7) - today.getDay() + 1);
        weekStart.setHours(0, 0, 0, 0);

        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);
        weekEnd.setHours(23, 59, 59, 999);

        weeks.push({ start: weekStart, end: weekEnd, label: `W${i + 1}` });
    }

    // Helper: check if date falls in a week
    const getWeekIndex = (date) => {
        if (!date) return 0;
        const d = new Date(date);
        if (isNaN(d.getTime())) return 0;

        // If before range, it counts as Week 1 (overdue/current work)
        if (d < weeks[0].start) return 0;

        // If after range, it counts as the last visible week
        if (d > weeks[weeks.length - 1].end) return weeks.length - 1;

        for (let i = 0; i < weeks.length; i++) {
            if (d >= weeks[i].start && d <= weeks[i].end) return i;
        }

        return 0;
    };

    // Helper: count holidays in a week
    const countHolidaysInWeek = (weekIdx) => {
        let count = 0;
        holidays.forEach(h => {
            if (h >= weeks[weekIdx].start && h <= weeks[weekIdx].end) count++;
        });
        return count;
    };

    // Build team data structure
    const teamData = {}; // { assignee: { weekIdx: hours } }
    const unassignedHours = new Array(weeksAhead).fill(0);

    const excludedUsers = capacityConfig.excludedUsers || [];

    let processedCount = 0;
    let totalWorkFound = 0;

    data.forEach((row, idx) => {
        const assignee = row[assigneeIndex] || 'Unassigned';

        // If user is excluded, ignore their work entirely
        if (excludedUsers.includes(assignee)) {
            return;
        }

        // Skip completed issues (only forward-looking)
        if (statusIndex !== -1) {
            const status = (row[statusIndex] || '').toString().toLowerCase();
            if (['done', 'resolved', 'completed', 'closed'].includes(status)) {
                return;
            }
        }
        let estimate = row[estimateIndex];

        // Convert estimate to hours
        let hours = 0;
        if (estimationType === 'storypoints') {
            hours = parseFloat(estimate) * hoursPerPoint;
        } else {
            // Jira returns time in seconds
            hours = parseFloat(estimate) / 3600;
        }

        if (isNaN(hours) || hours <= 0) return;

        processedCount++;
        totalWorkFound += hours;

        // ── LOAD SPREADING LOGIC ──
        const rawDue = dueDateIndex !== -1 ? row[dueDateIndex] : null;
        const rawStart = startDateIndex !== -1 ? row[startDateIndex] : null;

        const parseDate = (val) => {
            if (!val || val === "" || val === undefined) return null;
            const d = new Date(val);
            return isNaN(d.getTime()) ? null : d;
        };

        const dueDateObj = parseDate(rawDue);
        const startDateObj = parseDate(rawStart);

        let targetWeeks = [];
        if (startDateObj && dueDateObj && dueDateObj >= startDateObj) {
            let startIdx = getWeekIndex(startDateObj);
            const endIdx = getWeekIndex(dueDateObj);

            // If Start is in the past, limit visualization to current window
            if (startDateObj < weeks[0].start) {
                startIdx = 0;
            }

            if (endIdx >= startIdx) {
                for (let i = startIdx; i <= endIdx; i++) {
                    targetWeeks.push(i);
                }
            } else {
                targetWeeks.push(endIdx);
            }
        } else if (dueDateObj) {
            targetWeeks.push(getWeekIndex(dueDateObj));
        } else if (startDateObj) {
            targetWeeks.push(getWeekIndex(startDateObj));
        } else {
            targetWeeks.push(0);
        }

        const hoursPerWeekSpread = hours / (targetWeeks.length > 0 ? targetWeeks.length : 1); // Ensure no division by zero

        targetWeeks.forEach(weekIdx => {
            if (assignee === 'Unassigned' || !assignee) {
                unassignedHours[weekIdx] += hoursPerWeekSpread;
            } else {
                if (!teamData[assignee]) {
                    teamData[assignee] = new Array(weeksAhead).fill(0);
                }
                teamData[assignee][weekIdx] += hoursPerWeekSpread;
            }
        });
    });

    console.log(`[Capacity Plan] Processed ${processedCount} issues with work. Total hours: ${totalWorkFound}`);

    // Create or get Capacity sheet
    let capSheet = ss.getSheetByName('Jira Capacity Plan');
    if (capSheet) {
        capSheet.setFrozenRows(0);
        capSheet.setFrozenColumns(0);
        capSheet.clear();
    } else {
        capSheet = ss.insertSheet('Jira Capacity Plan');
    }

    // Calculate available hours per week (accounting for holidays and capacity %)
    const availableHours = weeks.map((w, i) => {
        const holidayDays = countHolidaysInWeek(i);
        const workDays = 5 - holidayDays;
        return (hoursPerWeek / 5) * workDays * capacityPercent;
    });

    // Set column widths
    capSheet.setColumnWidth(1, 150); // Name
    capSheet.setColumnWidth(2, 70);  // Total %
    for (let i = 0; i < weeksAhead; i++) {
        capSheet.setColumnWidth(3 + i, 65);
    }

    // Build header row
    const headerRow = ['Team Member', 'Utilization'];
    weeks.forEach((w, i) => {
        const monthDay = `${w.start.getMonth() + 1}/${w.start.getDate()}`;
        headerRow.push(monthDay);
    });

    // Week labels sub-header
    const weekLabelRow = ['', ''];
    weeks.forEach((w, i) => weekLabelRow.push(w.label));

    // Write headers
    capSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    capSheet.getRange(2, 1, 1, weekLabelRow.length).setValues([weekLabelRow]);

    // Style headers
    capSheet.getRange(1, 1, 1, headerRow.length)
        .setBackground('#0f172a')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    capSheet.getRange(2, 1, 1, weekLabelRow.length)
        .setBackground('#1e293b')
        .setFontColor('#94a3b8')
        .setFontSize(9)
        .setHorizontalAlignment('center');

    // Color function for utilization
    const getUtilColor = (percent) => {
        if (percent <= 50) return '#86efac';      // Light green
        if (percent <= 80) return '#4ade80';      // Green
        if (percent <= 100) return '#fbbf24';     // Yellow
        if (percent <= 120) return '#fb923c';     // Orange
        return '#f87171';                          // Red
    };

    // Write team data
    let rowNum = 3;
    const teamTotals = new Array(weeksAhead).fill(0);
    const sortedTeam = Object.keys(teamData).sort();

    sortedTeam.forEach(assignee => {
        const weekHours = teamData[assignee];
        let totalHours = 0;
        let totalAvailable = 0;

        const rowData = [assignee, ''];

        weekHours.forEach((hours, weekIdx) => {
            rowData.push(Math.round(hours) + 'h');
            totalHours += hours;
            totalAvailable += availableHours[weekIdx];
            teamTotals[weekIdx] += hours;
        });

        // Calculate utilization %
        const utilPercent = totalAvailable > 0 ? Math.round((totalHours / totalAvailable) * 100) : 0;
        rowData[1] = utilPercent + '%';

        capSheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);

        // Color the utilization cell
        capSheet.getRange(rowNum, 2).setBackground(getUtilColor(utilPercent)).setFontWeight('bold');

        // Color each week cell based on that week's utilization
        weekHours.forEach((hours, weekIdx) => {
            const weekUtil = availableHours[weekIdx] > 0 ? (hours / availableHours[weekIdx]) * 100 : 0;
            capSheet.getRange(rowNum, 3 + weekIdx)
                .setBackground(getUtilColor(weekUtil))
                .setHorizontalAlignment('center');
        });

        // Alternate row styling
        if (rowNum % 2 === 0) {
            capSheet.getRange(rowNum, 1).setBackground('#f8fafc');
        }

        rowNum++;
    });

    // Add Unassigned row if there's unassigned work
    const totalUnassigned = unassignedHours.reduce((a, b) => a + b, 0);
    if (totalUnassigned > 0) {
        const unassignedRow = ['⚠️ Unassigned', Math.round(totalUnassigned) + 'h'];
        unassignedHours.forEach(h => unassignedRow.push(Math.round(h) + 'h'));
        capSheet.getRange(rowNum, 1, 1, unassignedRow.length).setValues([unassignedRow]);
        capSheet.getRange(rowNum, 1, 1, unassignedRow.length)
            .setBackground('#fef3c7')
            .setFontStyle('italic');
        rowNum++;
    }

    // Add Team Total row
    rowNum++;
    const totalRow = ['📊 TEAM TOTAL', ''];
    let grandTotalHours = 0;
    let grandTotalAvailable = 0;

    teamTotals.forEach((hours, weekIdx) => {
        totalRow.push(Math.round(hours) + 'h');
        grandTotalHours += hours;
        grandTotalAvailable += availableHours[weekIdx] * sortedTeam.length;
    });

    const teamUtilPercent = grandTotalAvailable > 0 ? Math.round((grandTotalHours / grandTotalAvailable) * 100) : 0;
    totalRow[1] = teamUtilPercent + '%';

    capSheet.getRange(rowNum, 1, 1, totalRow.length).setValues([totalRow]);
    capSheet.getRange(rowNum, 1, 1, totalRow.length)
        .setBackground('#0f172a')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    capSheet.getRange(rowNum, 2).setBackground(getUtilColor(teamUtilPercent));

    // Add available hours row
    rowNum++;
    const availRow = ['Available Hours', ''];
    availableHours.forEach(h => availRow.push(Math.round(h) + 'h'));
    capSheet.getRange(rowNum, 1, 1, availRow.length).setValues([availRow]);
    capSheet.getRange(rowNum, 1, 1, availRow.length)
        .setBackground('#e2e8f0')
        .setFontStyle('italic')
        .setFontSize(9);

    // Add Legend
    rowNum += 2;
    capSheet.getRange(rowNum, 1).setValue('Legend:').setFontWeight('bold');
    capSheet.getRange(rowNum, 2).setValue('0-50%').setBackground('#86efac');
    capSheet.getRange(rowNum, 3).setValue('51-80%').setBackground('#4ade80');
    capSheet.getRange(rowNum, 4).setValue('81-100%').setBackground('#fbbf24');
    capSheet.getRange(rowNum, 5).setValue('101-120%').setBackground('#fb923c');
    capSheet.getRange(rowNum, 6).setValue('120%+').setBackground('#f87171');

    // Add config info
    rowNum += 2;
    const configInfo = [
        `Estimation: ${estimationType === 'storypoints' ? 'Story Points (' + hoursPerPoint + 'h/pt)' : 'Hours'}`,
        `Capacity: ${capacityPercent * 100}% of ${hoursPerWeek}h/week`,
        `Holidays: ${holidays.length} days`
    ];
    capSheet.getRange(rowNum, 1, 1, 3).setValues([configInfo]);
    capSheet.getRange(rowNum, 1, 1, 3).setFontColor('#64748b').setFontSize(9);

    // Add Capacity Analysis & Recommendations
    const analysisRow = rowNum + 2;
    addCapacityAnalysisSection(capSheet, analysisRow, teamData, teamTotals, availableHours, sortedTeam, unassignedHours, weeks);

    // Add Interpretation Guide
    rowNum += 15;
    addCapacityInterpretationGuide(capSheet, rowNum);

    // Final Formatting
    capSheet.setFrozenRows(2);
    capSheet.setFrozenColumns(2);
    capSheet.activate();

    return { success: true };
}

/**
 * Adds a "How to Read This Plan" section to the Capacity sheet
 */
function addCapacityInterpretationGuide(sheet, startRow) {
    const guide = [
        ["💡 HOW TO READ THIS PLAN", "", ""],
        ["Row 1 (Dates)", "Represents the Monday of each week.", ""],
        ["Utilization %", "Formula: (Total Hours Assigned) ÷ (Total Available Capacity).", ""],
        ["Available Capacity", "Formula: (Weekly Hours ÷ 5) × (Working Days) × (Capacity Select %).", ""],
        ["Example", "If 40h/week at 80% with no holidays: 32h available per week.", ""],
        ["Weekly Totals", "Sum of all issues assigned to that person falling within that week (based on Due Date).", ""],
        ["Week 1 Logic", "IMPORTANT: Issues with NO due date are automatically assigned to Week 1.", ""],
        ["", "", ""],
        ["📊 COLOR THRESHOLDS", "", ""],
        ["🟢 Green (0-80%)", "Healthy. The person has room for more work or unexpected tasks.", ""],
        ["🟡 Yellow (81-100%)", "Full. This person is at their target capacity. High risk of slippage if new work is added.", ""],
        ["🟠 Orange (101-120%)", "Strained. Person is working overtime or ignoring non-ticket work. Burnout risk.", ""],
        ["🔴 Red (120%+)", "Critical Overload. Deadlines will likely be missed. Immediate load balancing required.", ""]
    ];

    const range = sheet.getRange(startRow, 1, guide.length, 3);
    range.setValues(guide);

    sheet.getRange(startRow, 1).setFontWeight('bold').setFontSize(11).setFontColor('#1e293b');
    sheet.getRange(startRow + 1, 1, guide.length - 1, 1).setFontWeight('bold').setFontColor('#475569').setFontSize(9);
    sheet.getRange(startRow + 1, 2, guide.length - 1, 2).setFontColor('#64748b').setFontSize(9);

    // Header for color thresholds
    sheet.getRange(startRow + 6, 1).setFontWeight('bold').setFontColor('#1e293b').setFontSize(10);
}

/**
 * Generates an automated analysis and recommendation report on the Capacity sheet
 */
function addCapacityAnalysisSection(sheet, startRow, teamData, teamTotals, availableHours, sortedTeam, unassignedHours, weeks) {
    const today = new Date();
    const headers = [['📋 CAPACITY ANALYSIS & RECOMMENDATIONS', '', '', '', '', '']];

    sheet.getRange(startRow, 1, 1, 6).setValues(headers)
        .setBackground('#f1f5f9')
        .setFontWeight('bold')
        .setFontSize(11)
        .setBorder(true, true, true, true, false, false);

    const observations = [];
    const recommendations = [];

    // 1. Analyze Week 1 Overload (Data Quality Check)
    const w1Hours = teamTotals[0];
    const avgOtherWeeksHours = teamTotals.slice(1).length > 0 ? teamTotals.slice(1).reduce((a, b) => a + b, 0) / (weeks.length - 1) : 0;

    if (w1Hours > avgOtherWeeksHours * 3 && w1Hours > 0) {
        observations.push("WEEK 1 SPIKE: Significant workload concentration in the current week.");
        recommendations.push("ACTION: Verify 'Due Dates' for all issues. Tasks without deadlines default to Week 1, potentially skewing your real capacity.");
    }

    // 2. Identify Overloaded Individuals
    const overloaded = [];
    sortedTeam.forEach(name => {
        const data = teamData[name];
        const maxUtil = Math.max(...data.map((h, i) => availableHours[i] > 0 ? (h / availableHours[i]) * 100 : 0));
        if (maxUtil > 110) overloaded.push(name);
    });

    if (overloaded.length > 0) {
        observations.push(`BOTTLENECKS: ${overloaded.join(', ')} are significantly over-capacity in certain weeks.`);
        recommendations.push("ACTION: Attempt to 'Load Balance' by moving tasks from these individuals to team members with green indicators below.");
    }

    // 3. Analyze Unassigned Risk
    const totalUnassigned = unassignedHours.reduce((a, b) => a + b, 0);
    if (totalUnassigned > 40) {
        observations.push(`UNASSIGNED BACKLOG: There are roughly ${Math.round(totalUnassigned)}h of unassigned work in the current view.`);
        recommendations.push("ACTION: Assign these orphan tasks immediately to prevent them from becoming critical blockers as deadlines approach.");
    }

    // 4. Team Health
    const totalTeamHours = teamTotals.reduce((a, b) => a + b, 0);
    const totalTeamAvailable = availableHours.reduce((a, b) => a + b, 0) * sortedTeam.length;
    const teamUtil = totalTeamAvailable > 0 ? (totalTeamHours / totalTeamAvailable) * 100 : 0;

    if (teamUtil > 85) {
        observations.push("CRITICAL UTILIZATION: The entire team is operating at high intensity (>85%).");
        recommendations.push("ACTION: Review roadmap scope. This level of utilization leaves no room for bugs, meetings, or emergent risks.");
    } else if (teamUtil < 40 && totalTeamHours > 0) {
        observations.push("UNDER-UTILIZATION: The team appears to have significant head-room.");
        recommendations.push("ACTION: Consider pulling forward high-priority items from the backlog to maximize throughput.");
    }

    // Fallback if everything looks perfect
    if (observations.length === 0) {
        observations.push("HEALTHY LOAD: Workload is well-distributed and within capacity bounds.");
        recommendations.push("ACTION: Monitor for emergent blockers and maintain current scheduling precision.");
    }

    let currentRow = startRow + 1;

    // Write Observations
    sheet.getRange(currentRow, 1).setValue("OBSERVED RISKS").setFontWeight('bold').setFontColor('#e11d48').setFontSize(9);
    observations.forEach(obs => {
        currentRow++;
        sheet.getRange(currentRow, 1).setValue("• " + obs).setFontSize(9).setFontColor('#475569');
    });

    // Write Recommendations
    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue("EXECUTIVE RECOMMENDATION").setFontWeight('bold').setFontColor('#2563eb').setFontSize(9);
    recommendations.forEach(rec => {
        currentRow++;
        sheet.getRange(currentRow, 1).setValue("➜ " + rec).setFontSize(9).setFontWeight('bold').setFontColor('#1e293b');
    });

    // Formatting
    sheet.getRange(startRow + 1, 1, currentRow - startRow, 6).setWrap(true);
}

/**
 * Generates an automated executive summary for the Roadmap sheet
 */
function addRoadmapAnalysisSection(sheet, startRow, data, indices) {
    const { statusIndex, dueDateIndex, resolvedIndex, startDateIndex } = indices;

    const headers = [['📋 ROADMAP EXECUTIVE SUMMARY', '', '', '', '', '']];
    sheet.getRange(startRow, 1, 1, 6).setValues(headers)
        .setBackground('#f1f5f9')
        .setFontWeight('bold')
        .setFontSize(11)
        .setBorder(true, true, true, true, false, false);

    const observations = [];
    const recommendations = [];

    let doneCount = 0;
    let inProgressCount = 0;
    let toDoCount = 0;
    let overdueCount = 0;
    let missingDatesCount = 0;
    const today = new Date();

    data.forEach(row => {
        const status = (row[statusIndex] || '').toString().toLowerCase();
        const due = row[dueDateIndex] ? new Date(row[dueDateIndex]) : null;
        const start = row[startDateIndex] ? new Date(row[startDateIndex]) : null;

        if (status.includes('done') || status.includes('resolved') || status.includes('closed')) {
            doneCount++;
        } else if (status.includes('progress') || status.includes('review')) {
            inProgressCount++;
        } else {
            toDoCount++;
        }

        if (!status.includes('done') && due && due < today && !isNaN(due.getTime())) {
            overdueCount++;
        }

        if (!due && !start) {
            missingDatesCount++;
        }
    });

    const total = data.length;
    const progressPercent = total > 0 ? Math.round((doneCount / total) * 100) : 0;

    observations.push(`COMPLETION VELOCITY: ${progressPercent}% of the current roadmap is marked as 'Done'.`);
    observations.push(`PIPELINE: ${toDoCount} issues in Backlog, ${inProgressCount} currently active.`);

    if (overdueCount > 0) {
        observations.push(`SLIPPAGE RISK: ${overdueCount} issues are currently past their due dates but not finished.`);
        recommendations.push(`ACTION: Review the ${overdueCount} overdue items. Close completed tasks or move deadlines to reflect realistic delivery.`);
    }

    if (missingDatesCount > 0) {
        observations.push(`VISIBILITY GAP: ${missingDatesCount} issues are missing both Start and End dates.`);
        recommendations.push("ACTION: Populate date fields for orphan issues to include them in the timeline visualization.");
    }

    if (inProgressCount > (total * 0.5) && total > 5) {
        observations.push("WIP OVERLOAD: More than 50% of the team's work is 'In Progress' simultaneously.");
        recommendations.push("ACTION: Focus on 'stopping starting and starting finishing'. Limit Work-In-Progress to drive items to 'Done'.");
    }

    if (progressPercent < 20 && total > 10) {
        observations.push("EARLY STAGE: Roadmap is heavily weighted towards future work.");
        recommendations.push("ACTION: Ensure immediate milestones are clear and high-priority items are correctly sequenced.");
    }

    // Default recommendation
    if (recommendations.length === 0) {
        recommendations.push("ACTION: Maintain current execution pace. Monitor 'In Progress' tickets for transition to 'Done'.");
    }

    let currentRow = startRow + 1;
    sheet.getRange(currentRow, 1).setValue("KEY OBSERVATIONS").setFontWeight('bold').setFontColor('#e11d48').setFontSize(9);
    observations.forEach(obs => {
        currentRow++;
        sheet.getRange(currentRow, 1).setValue("• " + obs).setFontSize(9).setFontColor('#475569');
    });

    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue("ACTIONABLE RECOMMENDATIONS").setFontWeight('bold').setFontColor('#2563eb').setFontSize(9);
    recommendations.forEach(rec => {
        currentRow++;
        sheet.getRange(currentRow, 1).setValue("➜ " + rec).setFontSize(9).setFontWeight('bold').setFontColor('#1e293b');
    });

    sheet.getRange(startRow + 1, 1, currentRow - startRow, 6).setWrap(true);
}

/**
 * Returns a unique list of assignees found in the current sheet
 */
function getAssigneesFromSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) return [];

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerMap = headers.map(h => h.toString().toLowerCase().trim());
    const assigneeIndex = headerMap.indexOf('assignee');

    if (assigneeIndex === -1) return [];

    const data = sheet.getRange(2, assigneeIndex + 1, lastRow - 1, 1).getValues();
    const assignees = new Set();
    data.forEach(row => {
        const val = row[0];
        if (val && val.toString().trim() !== '') {
            assignees.add(val.toString().trim());
        }
    });

    return Array.from(assignees).sort();
}

/**
 * Deletes selected Jira issues from both Jira and the Sheet
 */
function deleteSelectedJiraIssues(config) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing Jira configuration.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = headers.map(h => h.toString().toLowerCase().trim());
    const keyIndex = headerMap.indexOf('key');

    if (keyIndex === -1) {
        throw new Error("Column 'Key' not found in header.");
    }

    // Get selected ranges
    const selection = sheet.getSelection();
    const activeRangeList = selection.getActiveRangeList();
    if (!activeRangeList) {
        throw new Error("No rows selected.");
    }
    const ranges = activeRangeList.getRanges();

    // Collect all rows to delete first, to avoid messing up indices 
    // We store { key: string, rowIndex: number }
    let issuesToDelete = [];

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        const values = sheet.getRange(startRow, keyIndex + 1, numRows, 1).getValues();

        for (let i = 0; i < numRows; i++) {
            const rowIndex = startRow + i;
            // Skip header row if selected
            if (rowIndex === 1) continue;

            const key = values[i][0];
            if (key && key.toString().trim() !== '') {
                issuesToDelete.push({ key: key.toString().trim(), rowIndex: rowIndex });
            }
        }
    });

    if (issuesToDelete.length === 0) {
        throw new Error("No issue keys found in selection.");
    }

    // Sort by rowIndex descending so we can delete from bottom up without affect indices
    issuesToDelete.sort((a, b) => b.rowIndex - a.rowIndex);

    // Dedup keys just in case, though deleting same row twice is handled by index method
    // Actually, distinct rows might have same key if data is dupe? 
    // Usually standard practice is one row per issue. 
    // We will attempt to delete from Jira for each unique Key, and delete row for each row.

    const uniqueKeys = [...new Set(issuesToDelete.map(item => item.key))];
    const results = {
        deleted: 0,
        failed: 0,
        errors: []
    };

    // We track which keys were successfully deleted to know if we should delete the row
    const deletedKeys = new Set();

    const authHeader = 'Basic ' + Utilities.base64Encode(config.email + ':' + config.token);
    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');

    // Delete from Jira
    uniqueKeys.forEach(key => {
        try {
            const url = `https://${cleanDomain}/rest/api/2/issue/${key}?deleteSubtasks=true`;
            const options = {
                method: 'delete',
                headers: {
                    'Authorization': authHeader,
                    'Accept': 'application/json'
                },
                muteHttpExceptions: true
            };

            const response = fetchJira(url, options, config);
            const code = response.getResponseCode();

            if (code === 204) {
                results.deleted++;
                deletedKeys.add(key);
            } else {
                results.failed++;
                let msg = `Failed to delete ${key}: ${code}`;
                try {
                    const err = JSON.parse(response.getContentText());
                    if (err.errorMessages) msg += " - " + err.errorMessages.join(", ");
                } catch (e) { }
                results.errors.push(msg);
            }
        } catch (e) {
            results.failed++;
            results.errors.push(`Error deleting ${key}: ${e.message}`);
        }
    });

    // Delete rows from sheet if their key was deleted
    // (Or if the key isn't in Jira anymore, maybe we should force delete row? 
    //  Safest is only delete row if API said 204 or user accepts)
    // For now, only delete row if API success or maybe 404 (already gone).
    // Let's stick to strict success.

    issuesToDelete.forEach(item => {
        if (deletedKeys.has(item.key)) {
            try {
                sheet.deleteRow(item.rowIndex);
            } catch (e) {
                // Row might have shifted if we didn't sort correctly?
                // We sorted desc, so it should be fine.
            }
        }
    });
    return results;
}

// --- Trend Chart ---
function generateTrendChart(type, period, sourceSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = sourceSheetName ? ss.getSheetByName(sourceSheetName) : ss.getActiveSheet();
    if (!dataSheet) throw new Error("Source sheet not found");

    const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
    const dateCol = type === 'created' ? 'created' : (type === 'resolved' ? 'resolutiondate' : 'duedate');

    // Placeholder for trend chart logic
    // This function would typically process data from dataSheet based on type and period
    // and then create or update a chart.
    // For now, it just returns a placeholder message.
    return `Trend chart for ${type} issues over ${period} from sheet ${dataSheet.getName()} would be generated here.`;
}

/**
 * Gets the count of selected Jira issues (rows with keys)
 */
function getSelectedIssueCount() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = headers.map(h => h.toString().toLowerCase().trim());
    const keyIndex = headerMap.indexOf('key');

    if (keyIndex === -1) return 0;

    const selection = sheet.getSelection();
    const activeRangeList = selection.getActiveRangeList();
    if (!activeRangeList) return 0;

    const ranges = activeRangeList.getRanges();
    let count = 0;

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        const values = sheet.getRange(startRow, keyIndex + 1, numRows, 1).getValues();

        for (let i = 0; i < numRows; i++) {
            const rowIndex = startRow + i;
            if (rowIndex === 1) continue; // Skip header
            const key = values[i][0];
            if (key && key.toString().trim() !== '') {
                count++;
            }
        }
    });

    return count;
}

/**
 * Shows the log sidebar
 */
function showRefreshLogs() {
    const html = HtmlService.createHtmlOutputFromFile('Logs')
        .setTitle('Logs');
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets the logs from properties
 */
function getRefreshLogs() {
    const props = PropertiesService.getDocumentProperties();
    const logsJson = props.getProperty('refreshLogs');
    return logsJson ? JSON.parse(logsJson) : [];
}

/**
 * Clears the refresh logs
 */
function clearRefreshLogs() {
    PropertiesService.getDocumentProperties().deleteProperty('refreshLogs');
    return { success: true };
}

/**
 * Logs a refresh attempt
 */
function logRefreshAttempt(status, details, sheetName = '') {
    const props = PropertiesService.getDocumentProperties();
    const logsJson = props.getProperty('refreshLogs');
    let logs = logsJson ? JSON.parse(logsJson) : [];

    // Add new log
    logs.unshift({
        timestamp: new Date().toISOString(),
        status: status,
        details: details,
        sheetName: sheetName
    });

    // Cap at 50
    if (logs.length > 50) {
        logs = logs.slice(0, 50);
    }

    props.setProperty('refreshLogs', JSON.stringify(logs));
}

/**
 * Returns basic info to ping for sheet changes
 */
function getPingInfo() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return {
        sheetId: sheet.getSheetId(),
        range: sheet.getActiveRange().getA1Notation()
    };
}

/**
 * Returns the exact count of issues matching a JQL query or Filter ID.
 * Usage: =JIRA_COUNT("resolution = Unresolved") OR =JIRA_COUNT(12345)
 *
 * @param {string|number} input The JQL query string OR a specific Filter ID.
 * @return {number} The count of matching issues.
 * @customfunction
 */
function JIRA_COUNT(input) {
    if (!input) throw new Error("Input is required (JQL query or Filter ID)");

    const props = PropertiesService.getDocumentProperties();
    const credsJson = props.getProperty('jira_creds');

    if (!credsJson) {
        throw new Error("Jira credentials not configured. Open the Jira Add-on sidebar to configure.");
    }

    let config;
    try {
        config = JSON.parse(credsJson);
    } catch (e) {
        throw new Error("Invalid configuration.");
    }

    const domain = config.domain;
    const email = config.email;
    const token = config.token;

    if (!domain || !email || !token) {
        throw new Error("Incomplete credentials.");
    }

    const cleanDomain = domain.replace(/^https?:\/\//, '').replace(/\/$/, '');

    let jql = input;

    // Check if input is a Filter ID (numeric)
    if (/^\d+$/.test(input.toString().trim())) {
        const filterId = input.toString().trim();
        const filterUrl = `https://${cleanDomain}/rest/api/3/filter/${filterId}`;
        const filterOptions = {
            method: 'GET',
            headers: {
                "Authorization": "Basic " + Utilities.base64Encode(email + ":" + token),
                "Accept": "application/json"
            },
            muteHttpExceptions: true
        };

        try {
            const fRes = fetchJira(filterUrl, filterOptions, config);
            if (fRes.getResponseCode() !== 200) {
                // Show Error in Toaster
                try { SpreadsheetApp.getActiveSpreadsheet().toast(`Error: Invalid Filter ID ${filterId}`, "Jira Error", 5); } catch (e) { }
                // Return Error to Cell
                return `Error: Filter ${filterId} invalid or not found.`;
            }
            const fData = JSON.parse(fRes.getContentText());
            if (!fData.jql) {
                return `Error: Filter ${filterId} has no JQL.`;
            }
            jql = fData.jql;
        } catch (e) {
            return `Error checking filter: ${e.message}`;
        }
    }

    // Switch to approximate-count endpoint as /search/jql does not return 'total'
    const url = `https://${cleanDomain}/rest/api/3/search/approximate-count`;

    const options = {
        method: 'POST',
        contentType: 'application/json',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(email + ":" + token),
            "Accept": "application/json"
        },
        payload: JSON.stringify({ jql: jql }),
        muteHttpExceptions: true
    };

    try {
        const response = fetchJira(url, options, config);
        const code = response.getResponseCode();
        const text = response.getContentText();

        if (code !== 200) {
            let msg = text;
            try {
                const err = JSON.parse(text);
                if (err.errorMessages) msg = err.errorMessages.join(', ');
                else if (err.message) msg = err.message;
            } catch (e) { }
            throw new Error(`Jira API Error (${code}): ${msg}`);
        }

        const data = JSON.parse(text);

        if (typeof data.count !== 'undefined') {
            return data.count;
        }

        return "Error: Unknown count";
    } catch (e) {
        throw new Error(e.message);
    }
}

// =============================================================================
// LICENSE MANAGEMENT
// =============================================================================

/**
 * Cloudflare Worker endpoints for license and checkout
 */
const LICENSE_API_URL = CLOUDFLARE_WORKER_URL + '/api/license/check';
const CHECKOUT_API_URL = CLOUDFLARE_WORKER_URL + '/api/stripe/create-session';
const PORTAL_API_URL = CLOUDFLARE_WORKER_URL + '/api/stripe/portal'; // Future use

// Stripe Price IDs - UPDATE THESE with your Stripe price IDs
const STRIPE_PRICES = {
    PRO_MONTHLY: 'price_0T06lCPaYAnupERfYwY8d95p',
    PRO_YEARLY: 'price_0T06jnPaYAnupERfQTenhkHV',
    ENTERPRISE_MONTHLY: 'price_1T6PvIDhvP6DurKSCrisUWHI',
    ENTERPRISE_YEARLY: 'price_1T6PvVDhvP6DurKSHcVjbbC2'
};

/**
 * Checks user's license status by calling Firebase API
 * @returns {Object} License info including allowed, plan, features
 */
function checkLicense(forceRefresh = false) {
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail) {
        return { allowed: false, plan: 'free', message: 'Unable to determine user email' };
    }

    try {
        const response = UrlFetchApp.fetch(LICENSE_API_URL, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify({ email: userEmail, forceRefresh: forceRefresh }),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
            return JSON.parse(response.getContentText());
        } else {
            return { allowed: true, plan: 'Trial', message: 'API Offline' }; // Graceful fallback
        }
    } catch (e) {
        console.error('License check error:', e);
        return { allowed: true, plan: 'Trial', message: 'Connection Error' };
    }
}

/**
 * Gets license information for display in the UI
 * @returns {Object} License details formatted for UI display
 */
function getLicenseInfo(forceRefresh = false) {
    const userEmail = Session.getActiveUser().getEmail();
    const license = checkLicense(forceRefresh);

    return {
        email: userEmail,
        allowed: license.allowed,
        plan: license.plan || 'free',
        status: license.status || 'none',
        licenseType: license.licenseType || 'none',
        domain: license.domain || null,
        expiresAt: license.expiresAt || null,
        features: license.features || [],
        seats: license.seats || null,
        usedSeats: license.usedSeats || null,
        message: license.message || null
    };
}

/**
 * Checks if a specific feature is available
 * @param {string} featureName - Feature to check
 * @returns {boolean} Whether feature is available
 */
function hasFeature(featureName) {
    const license = checkLicense();
    return license.allowed && license.features && license.features.includes(featureName);
}

/**
 * Creates a checkout session (Mock for now, returns Web App URL)
 */
function createCheckoutSession(priceId, licenseType, domain, seats) {
    const userEmail = Session.getActiveUser().getEmail();
    const scriptUrl = ScriptApp.getService().getUrl();

    // Construct mock checkout URL
    const params = [
        `page=checkout`,
        `email=${encodeURIComponent(userEmail)}`,
        `plan=${encodeURIComponent(licenseType)}`,
        `priceId=${encodeURIComponent(priceId)}`
    ];

    if (domain) params.push(`domain=${encodeURIComponent(domain)}`);
    if (seats) params.push(`seats=${encodeURIComponent(seats)}`);

    return {
        url: `${scriptUrl}?${params.join('&')}`
    };
}

/**
 * Creates a Stripe customer portal session for managing subscription
 * @returns {Object} Portal session URL
 */
function createManageSubscriptionSession() {
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail) {
        throw new Error('Unable to determine user email');
    }

    try {
        const response = UrlFetchApp.fetch(PORTAL_API_URL, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify({
                email: userEmail,
                returnUrl: 'https://docs.google.com/spreadsheets'
            }),
            muteHttpExceptions: true
        });

        const code = response.getResponseCode();
        const text = response.getContentText();

        if (code === 200) {
            return JSON.parse(text);
        } else {
            let errorMsg = 'Failed to create portal session';
            try {
                const errorData = JSON.parse(text);
                if (errorData.error) errorMsg = errorData.error;
            } catch (pErr) { }
            console.error('Portal session failed:', code, text);
            throw new Error(errorMsg);
        }
    } catch (e) {
        throw new Error('Portal error: ' + e.message);
    }
}

/**
 * Gets available pricing info for display
 * @returns {Object} Pricing tiers
 */
function getPricingInfo() {
    return {
        individual: {
            name: 'Personal License',
            description: 'Individual Pro license',
            monthly: { price: '$9/month', priceId: STRIPE_PRICES.PRO_MONTHLY, value: 9 },
            yearly: { price: '$89/year', priceId: STRIPE_PRICES.PRO_YEARLY, value: 89 },
            features: [
                'Unlimited syncs',
                'Priority support',
                'Advanced filters',
                'Custom fields',
                'Cloudflare Proxy Support'
            ]
        },
        enterprise: {
            name: 'Enterprise License',
            description: 'Best for large orgs',
            monthly: { price: '$345/month', priceId: STRIPE_PRICES.ENTERPRISE_MONTHLY, value: 345 },
            yearly: { price: '$2900/year', priceId: STRIPE_PRICES.ENTERPRISE_YEARLY, value: 2900 },
            features: [
                'Unlimited users',
                'Domain-wide access',
                'Centralized billing',
                'Dedicated support',
                'SLA guarantees',
                'Cloudflare Proxy Support'
            ]
        }
    };
}

/**
 * Web App Handler
 * (Simplified for Free Version)
 */
function doGet(e) {
    return HtmlService.createHtmlOutput('<h1>Jira Sync Add-on</h1><p>Running.</p>');
}

/**
 * Payment processing disabled
 */
function processMockPayment() {
    return { success: false, message: 'Payments disabled' };
}

/**
 * Generates a Trend Chart for issue metrics
 * @param {string} type - 'created', 'resolved', or 'updated'
 * @param {string} period - 'daily', 'weekly', or 'monthly'
 */
function generateTrendChart(type, period) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getActiveSheet();

    if (dataSheet.getName().includes('Trend')) {
        throw new Error("Cannot generate trend charts from the Trends sheet itself. Please switch to your Jira Data sheet first.");
    }
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = headers.map(h => String(h).toLowerCase().trim()); // Helper map

    // 1. Find Date Column
    let possibleNames = ['created'];
    if (type === 'resolved') possibleNames = ['resolutiondate', 'resolution date', 'resolved', 'resolved date', 'resolve date'];
    if (type === 'updated') possibleNames = ['updated', 'updated date', 'date updated'];

    // Fuzzy search for Date Column
    let colIdx = headerMap.findIndex(h => possibleNames.includes(h) || possibleNames.some(p => h.includes(p)));

    // Find Issue Type Column
    const typeIdx = headerMap.findIndex(h => ['issuetype', 'issue type', 'type'].includes(h));

    if (colIdx === -1) {
        const displayColName = type === 'resolved' ? 'Resolution Date' : possibleNames[0];
        throw new Error(`Column '${displayColName}' not found in current sheet. Please fetch '${displayColName}' field first.`);
    }

    // 2. Aggregate Data
    const counts = {}; // Key: DateStr, Value: { Bug: 5, Story: 2, ... }
    const issueTypes = new Set();
    const rows = data.slice(1);

    rows.forEach(row => {
        const dateVal = row[colIdx];
        if (!dateVal) return;

        let dateObj = new Date(dateVal);
        if (isNaN(dateObj.getTime())) return;

        let key;
        if (period === 'daily') {
            key = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } else if (period === 'weekly') {
            const day = dateObj.getDay();
            const diff = dateObj.getDate() - day + (day == 0 ? -6 : 1);
            dateObj.setDate(diff);
            key = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } else { // monthly
            key = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM");
        }

        if (!counts[key]) counts[key] = {};

        // Determine Type
        let iType = 'Total';
        if (typeIdx !== -1) {
            iType = row[typeIdx] ? String(row[typeIdx]) : 'Unspecified';
        }

        counts[key][iType] = (counts[key][iType] || 0) + 1;
        issueTypes.add(iType); // Track unique types
    });

    // 3. Prepare Chart Data
    const sortedDates = Object.keys(counts).sort();

    // Filter by period duration
    let finalDates = sortedDates;
    if (period === 'daily') finalDates = sortedDates.slice(-30);
    if (period === 'weekly') finalDates = sortedDates.slice(-12);
    if (period === 'monthly') finalDates = sortedDates.slice(-12);

    if (finalDates.length === 0) {
        throw new Error("No data found for the selected period.");
    }

    // Chart Header: ['Date', 'Bug', 'Story'...]
    // If no types found (typeIdx == -1), issueTypes has {'Total'}
    const sortedTypes = Array.from(issueTypes).sort();
    const chartHeader = ['Date', ...sortedTypes];
    const chartData = [chartHeader];

    finalDates.forEach(date => {
        const row = [date];
        sortedTypes.forEach(t => {
            row.push(counts[date][t] || 0); // Add count or 0
        });
        chartData.push(row);
    });

    // 4. Create Sheet
    let trendSheet = ss.getSheetByName('Issue Trends');
    if (trendSheet) ss.deleteSheet(trendSheet);
    trendSheet = ss.insertSheet('Issue Trends');

    // Write Data 
    trendSheet.getRange(1, 1, chartData.length, chartData[0].length).setValues(chartData);
    trendSheet.getRange(1, 1, 1, chartData[0].length).setFontWeight("bold").setBackground("#f3f4f6");

    // 5. Create Chart
    const range = trendSheet.getRange(1, 1, chartData.length, chartData[0].length);

    let chartTitle = `Issues ${type} over time (${period}) by Type`;

    let chartBuilder = trendSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setNumHeaders(1)
        .setPosition(2, chartData[0].length + 2, 0, 0)
        .setOption('title', chartTitle)
        .setOption('hAxis', { title: 'Date' })
        .setOption('vAxis', { title: 'Issues' })
        .setOption('isStacked', true) // Enables Stacking
        .setOption('width', 800)
        .setOption('height', 400);

    trendSheet.insertChart(chartBuilder.build());

    return { success: true };
}



/**
 * Generates a Pivot Table configuration
 * @param {Object} config - { row: 'assignee', col: 'status', val: 'count', filterUser: '' }
 */
function generatePivotTable(config, sourceSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSheetName ? ss.getSheetByName(sourceSheetName) : ss.getActiveSheet();
    if (!sourceSheet) throw new Error("Source sheet not found: " + sourceSheetName);

    if (sourceSheet.getName().includes('Jira Pivot')) {
        throw new Error("Cannot generate from existing Pivot sheet. Please switch to your Jira Data sheet.");
    }

    const dataRange = sourceSheet.getDataRange();
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    const headerMap = headers.map(h => h.toString().toLowerCase().trim());

    // Helper
    const findIndex = (aliases) => {
        const found = headerMap.findIndex(h => aliases.some(a => h === a || h.includes(a))); // Fuzzy match
        return found; // returns -1 if not found
    };

    // Resolve Fields
    const rowField = (config.row || 'assignee').toLowerCase();
    const colField = (config.col || 'status').toLowerCase();
    const valField = (config.val || 'count').toLowerCase();

    // Map logic to real aliases
    let aliasMap = {
        'assignee': ['assignee', 'assigned to', 'assigned'],
        'status': ['status', 'state', 'issue status'],
        'priority': ['priority', 'issue priority'],
        'issuetype': ['issuetype', 'issue type', 'type'],
        'sprint': ['sprint', 'sprint name'],
        'project': ['project', 'project name', 'project key'],
        'resolution': ['resolution', 'resolution date', 'resolved', 'resolutiondate'],
        'created': ['created', 'date created'],
        'updated': ['updated', 'date updated']
    };

    const determineIndex = (field) => {
        if (field === 'none') return -1;
        const aliases = aliasMap[field] || [field];
        return findIndex(aliases);
    };

    const rowIndex = determineIndex(rowField);
    const colIndex = determineIndex(colField);

    // Value: Count (Key) or Sum (Story Points)
    let valIndex = -1;
    let summarizeFunc = SpreadsheetApp.PivotTableSummarizeFunction.COUNTA;

    if (valField === 'points') {
        const spAliases = ['story points', 'story point', 'points', 'estimates', 'story_points'];
        valIndex = findIndex(spAliases);
        summarizeFunc = SpreadsheetApp.PivotTableSummarizeFunction.SUM;
        if (valIndex === -1) throw new Error("Story Points column not found.");
    } else {
        // Count: Use Key
        valIndex = findIndex(['key', 'issue key', 'issue id']);
        if (valIndex === -1) throw new Error("Key column not found.");
    }

    if (rowIndex === -1 && colIndex === -1) {
        throw new Error("Must select at least a Row or Column dimension.");
    }

    // Create Sheet
    let pivotSheet = ss.getSheetByName('Jira Pivot');
    if (pivotSheet) ss.deleteSheet(pivotSheet);
    pivotSheet = ss.insertSheet('Jira Pivot');

    // Create Pivot
    // PivotTable API requires valid range. sourceSheet must have data.
    const pivotAnchor = pivotSheet.getRange('A3');
    const pivotTable = pivotAnchor.createPivotTable(dataRange);

    // Row Group
    if (rowIndex !== -1) {
        const group = pivotTable.addRowGroup(rowIndex + 1);
        group.showTotals(true).sortAscending();
    }

    // Column Group
    if (colIndex !== -1) {
        const group = pivotTable.addColumnGroup(colIndex + 1);
        group.showTotals(true).sortAscending();
        // If col is date, maybe group by Month? Auto-grouping is complex in API.
    }

    // Value
    if (valIndex !== -1) {
        pivotTable.addPivotValue(valIndex + 1, summarizeFunc);
    }

    // Filter by User
    if (config.filterUser && config.filterUser.trim() !== '') {
        const assigneeIdx = determineIndex('assignee');
        if (assigneeIdx !== -1) {
            const criteria = SpreadsheetApp.newFilterCriteria()
                .setVisibleValues([config.filterUser])
                .build();
            pivotTable.addFilter(assigneeIdx + 1, criteria);
        }
    }

    // Header
    const title = `Pivot: ${config.val} by ${config.row} & ${config.col}`;
    pivotSheet.getRange('A1').setValue(title).setFontSize(14).setFontWeight('bold');

    return { success: true };
}

/**
 * Refreshes all sheets that have a Jira configuration
 */
function refreshAllJiraSheets(params) {
    if (!params.domain || !params.email || !params.token) {
        throw new Error("Missing credentials");
    }

    const props = PropertiesService.getDocumentProperties();
    const allProps = props.getProperties();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    let totalCount = 0;
    let totalChanged = 0;
    const refreshedSheets = [];
    params.proxyCallCount = 0; // Initialize counter for the entire batch

    // Determine Pro status if not provided (e.g. background trigger)
    if (typeof params.isPro === 'undefined') {
        const license = checkLicense();
        params.isPro = license.allowed && license.plan !== 'free';
    }

    // Filter properties for sheet configs
    Object.keys(allProps).forEach(key => {
        if (key.startsWith('config_')) {
            const sheetId = key.split('_')[1];
            // Find sheet by ID
            const targetSheet = sheets.find(s => s.getSheetId().toString() === sheetId);

            if (targetSheet) {
                const sheetConfig = JSON.parse(allProps[key]);
                try {
                    // Highlight tab to show activity (Orange)
                    targetSheet.setTabColor('#F59E0B');
                    SpreadsheetApp.flush(); // Force UI update

                    // Merge global creds with sheet-specific JQL/columns
                    const fetchParams = {
                        ...params,
                        jql: sheetConfig.jql,
                        columns: sheetConfig.columns,
                        targetSheetName: targetSheet.getName(),
                        triggerType: params.triggerType || 'GlobalInterval'
                    };

                    const res = fetchJiraData(fetchParams);
                    totalCount += (res.count || 0);
                    totalChanged += (res.changedCount || 0);
                    if (res.limited) params.anyLimited = true;
                    // Carry forward proxy calls from fetchJiraData
                    if (res.proxyCalls) {
                        params.proxyCallCount = (params.proxyCallCount || 0) + res.proxyCalls;
                    }
                    refreshedSheets.push(targetSheet.getName());

                    // Clear tab color on success
                    targetSheet.setTabColor(null);
                } catch (e) {
                    // unexpected error: highlight Red
                    targetSheet.setTabColor('#EF4444');
                    console.error(`Failed to refresh sheet ${targetSheet.getName()}: ${e.message}`);
                    try { logRefreshAttempt('FAILED', `Bulk refresh error: ${e.message}`, targetSheet.getName()); } catch (logErr) { }
                }
            }
        }
    });

    // Log the entire global sync outcome
    const proxyInfo = params.proxyCallCount > 0 ? ` (via Worker x${params.proxyCallCount})` : '';
    const logDetails = `[Global Sync] Completed for ${refreshedSheets.length} sheets. ${totalCount} issues total.${proxyInfo}`;
    if (params.anyLimited) {
        try { logRefreshAttempt('WARN', 'Global sync reached 150-issue limit on some sheets.', 'GLOBAL SYNC'); } catch (e) { }
    } else {
        try { logRefreshAttempt('OK', logDetails, 'GLOBAL SYNC'); } catch (e) { }
    }

    return {
        count: totalCount,
        changed: totalChanged,
        sheetCount: refreshedSheets.length,
        sheets: refreshedSheets,
        limited: !!params.anyLimited
    };
}

/**
 * Fetches labels and their issue counts from issues on the active sheet
 */
function getLabelStats(config) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing Jira configuration.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = config.sourceSheet ? ss.getSheetByName(config.sourceSheet) : ss.getActiveSheet();
    if (!sheet) throw new Error("No active sheet found.");

    // Find the Key column (contains issue keys like PROJ-123)
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) throw new Error("Sheet is empty.");
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const keyColIndex = headers.findIndex(h => h && h.toString().trim().toLowerCase() === 'key');
    if (keyColIndex === -1) throw new Error("No 'Key' column found in this sheet.");

    // Get all issue keys from the sheet
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const keys = sheet.getRange(2, keyColIndex + 1, lastRow - 1, 1).getValues()
        .map(r => r[0] ? r[0].toString().trim() : '')
        .filter(k => k && k.match(/^[A-Z]+-\d+$/));

    if (keys.length === 0) return {};

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const authHeader = "Basic " + Utilities.base64Encode(config.email + ":" + config.token);
    const searchUrl = `https://${cleanDomain}/rest/api/3/search/jql`;
    const stats = {};

    // Query in batches of 100 keys
    const BATCH_SIZE = 100;
    try {
        for (let i = 0; i < keys.length; i += BATCH_SIZE) {
            const batch = keys.slice(i, i + BATCH_SIZE);
            const jql = `key in (${batch.join(',')}) AND labels is not EMPTY`;

            let nextPageToken = null;
            while (true) {
                const response = fetchJira(searchUrl, {
                    method: 'POST',
                    headers: {
                        "Authorization": authHeader,
                        "Content-Type": "application/json",
                        "Accept": "application/json"
                    },
                    payload: JSON.stringify({
                        jql: jql,
                        nextPageToken: nextPageToken,
                        maxResults: 100,
                        fields: ["labels"]
                    }),
                    muteHttpExceptions: true
                }, config);

                if (response.getResponseCode() !== 200) {
                    throw new Error(`Jira API Error: ${response.getContentText()}`);
                }

                const data = JSON.parse(response.getContentText());
                const issues = data.issues || [];

                issues.forEach(issue => {
                    if (issue.fields && issue.fields.labels) {
                        issue.fields.labels.forEach(label => {
                            stats[label] = (stats[label] || 0) + 1;
                        });
                    }
                });

                nextPageToken = data.nextPageToken;
                if (!nextPageToken || issues.length === 0) break;
            }
        }

        return stats;

    } catch (e) {
        throw new Error("Failed to fetch labels: " + e.message);
    }
}

/**
 * Removes a specific label from all issues matching the JQL + label
 */
function removeLabelFromIssues(config, labelToRemove) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing Jira configuration.");
    }
    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const authHeader = "Basic " + Utilities.base64Encode(config.email + ":" + config.token);

    // 1. Find issues that satisfy current JQL AND have this label
    let baseJql = config.jql || "";
    // Strip ORDER BY if present as it cannot be inside parentheses
    baseJql = baseJql.replace(/\s+ORDER BY.*/i, "").trim();

    const jql = (baseJql ? `(${baseJql}) AND ` : '') + `labels = "${labelToRemove}"`;
    const searchUrl = `https://${cleanDomain}/rest/api/3/search/jql`;

    let issues = [];
    let nextPageToken = null;
    let pagesProcessed = 0;

    try {
        while (true) {
            const body = {
                jql: jql,
                maxResults: 100,
                fields: ["key"]
            };
            if (nextPageToken) body.nextPageToken = nextPageToken;

            const response = fetchJira(searchUrl, {
                method: 'POST',
                headers: {
                    "Authorization": authHeader,
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                },
                payload: JSON.stringify(body),
                muteHttpExceptions: true
            }, config);

            if (response.getResponseCode() !== 200) {
                throw new Error(response.getContentText());
            }

            const data = JSON.parse(response.getContentText());
            const batch = data.issues || [];
            if (batch.length === 0) break;

            issues = issues.concat(batch);
            nextPageToken = data.nextPageToken;
            pagesProcessed++;

            // Break if no more tokens or safety limit (50 pages = 5000 issues)
            if (!nextPageToken || batch.length < 100 || pagesProcessed >= 50) break;
        }
    } catch (e) {
        throw new Error("Search failed: " + e.message);
    }

    if (issues.length === 0) {
        return { success: true, updated: 0, message: "No issues found with this label." };
    }

    // 2. Batch Update to Remove Label
    let successCount = 0;
    let failureCount = 0;
    let firstError = null;
    const batchSize = 20; // Use 20 for updates (safer)

    for (let i = 0; i < issues.length; i += batchSize) {
        const batch = issues.slice(i, i + batchSize);

        const requests = batch.map(issue => ({
            url: `https://${cleanDomain}/rest/api/3/issue/${issue.key}`,
            method: 'PUT',
            headers: {
                "Authorization": authHeader,
                "Content-Type": "application/json"
            },
            payload: JSON.stringify({
                "update": {
                    "labels": [
                        { "remove": labelToRemove }
                    ]
                }
            }),
            muteHttpExceptions: true
        }));

        try {
            const responses = fetchAllJira(requests, config);
            responses.forEach((res, idx) => {
                const code = res.getResponseCode();
                if (code === 204 || code === 200) {
                    successCount++;
                } else {
                    failureCount++;
                    if (!firstError) {
                        try {
                            const err = JSON.parse(res.getContentText());
                            firstError = `${batch[idx].key}: ${JSON.stringify(err.errors || err.errorMessages)}`;
                        } catch (e) {
                            firstError = `${batch[idx].key}: ${code}`;
                        }
                    }
                }
            });
        } catch (e) {
            failureCount += batch.length;
            if (!firstError) firstError = e.message;
        }
        Utilities.sleep(100);
    }

    return {
        success: failureCount === 0,
        updated: successCount,
        failed: failureCount,
        error: firstError
    };
}

/**
 * Renames a specific label to a new name on all issues matching the JQL + label
 */
function renameLabelInIssues(config, oldLabel, newLabel) {
    if (!config.domain || !config.email || !config.token) {
        throw new Error("Missing Jira configuration.");
    }
    if (!newLabel || newLabel.trim() === "") {
        throw new Error("New label name cannot be empty.");
    }
    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const authHeader = "Basic " + Utilities.base64Encode(config.email + ":" + config.token);

    // 1. Find issues that satisfy current JQL AND have the old label
    let baseJql = config.jql || "";
    baseJql = baseJql.replace(/\s+ORDER BY.*/i, "").trim();
    const jql = (baseJql ? `(${baseJql}) AND ` : '') + `labels = "${oldLabel}"`;
    const searchUrl = `https://${cleanDomain}/rest/api/3/search/jql`;

    let issues = [];
    let nextPageToken = null;
    let pagesProcessed = 0;

    try {
        while (true) {
            const body = {
                jql: jql,
                maxResults: 100,
                fields: ["key"]
            };
            if (nextPageToken) body.nextPageToken = nextPageToken;

            const response = UrlFetchApp.fetch(searchUrl, {
                method: 'POST',
                headers: {
                    "Authorization": authHeader,
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                },
                payload: JSON.stringify(body),
                muteHttpExceptions: true
            });

            if (response.getResponseCode() !== 200) {
                throw new Error(response.getContentText());
            }

            const data = JSON.parse(response.getContentText());
            const batch = data.issues || [];
            if (batch.length === 0) break;

            issues = issues.concat(batch);
            nextPageToken = data.nextPageToken;
            pagesProcessed++;

            if (!nextPageToken || batch.length < 100 || pagesProcessed >= 50) break;
        }
    } catch (e) {
        throw new Error("Search failed: " + e.message);
    }

    if (issues.length === 0) {
        return { success: true, updated: 0, message: "No issues found with this label." };
    }

    // 2. Batch Update to Rename Label (Add new, Remove old)
    let successCount = 0;
    let failureCount = 0;
    let firstError = null;
    const batchSize = 20;

    for (let i = 0; i < issues.length; i += batchSize) {
        const batch = issues.slice(i, i + batchSize);

        const requests = batch.map(issue => ({
            url: `https://${cleanDomain}/rest/api/3/issue/${issue.key}`,
            method: 'PUT',
            headers: {
                "Authorization": authHeader,
                "Content-Type": "application/json"
            },
            payload: JSON.stringify({
                "update": {
                    "labels": [
                        { "add": newLabel.trim() },
                        { "remove": oldLabel }
                    ]
                }
            }),
            muteHttpExceptions: true
        }));

        try {
            const responses = fetchAllJira(requests, config);
            responses.forEach((res, idx) => {
                const code = res.getResponseCode();
                if (code === 204 || code === 200) {
                    successCount++;
                } else {
                    failureCount++;
                    if (!firstError) {
                        try {
                            const err = JSON.parse(res.getContentText());
                            firstError = `${batch[idx].key}: ${JSON.stringify(err.errors || err.errorMessages)}`;
                        } catch (e) {
                            firstError = `${batch[idx].key}: ${code}`;
                        }
                    }
                }
            });
        } catch (e) {
            failureCount += batch.length;
            if (!firstError) firstError = e.message;
        }
        Utilities.sleep(100);
    }

    return {
        success: failureCount === 0,
        updated: successCount,
        failed: failureCount,
        error: firstError
    };
}


/**
 * Fetches assignable users for the specified project or globally
 */
function getAssignableUsers(config, projectKey) {
    if (!config.domain || !config.email || !config.token) {
        return [];
    }

    const cleanDomain = config.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');
    const pKey = projectKey || config.projectKey;

    // Try to find if we've already cached users for this project recently
    const props = PropertiesService.getDocumentProperties();
    const cacheKey = `users_${pKey || 'global'}`;
    const cached = props.getProperty(cacheKey);
    if (cached) {
        try {
            const cachedData = JSON.parse(cached);
            // 1 hour cache
            if (new Date().getTime() - cachedData.ts < 3600000) {
                return cachedData.users;
            }
        } catch (e) { }
    }

    let url = `https://${cleanDomain}/rest/api/3/user/assignable/search?maxResults=1000`;
    if (pKey) {
        url += `&project=${pKey}`;
    } else {
        // Fallback: search for active users
        url = `https://${cleanDomain}/rest/api/3/user/search?query=%&maxResults=1000`;
    }

    const options = {
        method: 'GET',
        headers: {
            "Authorization": "Basic " + Utilities.base64Encode(config.email + ":" + config.token),
            "Accept": "application/json"
        },
        muteHttpExceptions: true
    };

    try {
        const response = fetchJira(url, options, config);
        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            const users = data
                .filter(u => u.accountType === 'atlassian' && u.active)
                .map(u => ({
                    accountId: u.accountId,
                    displayName: u.displayName || u.name
                }));

            // Cache it
            props.setProperty(cacheKey, JSON.stringify({ ts: new Date().getTime(), users: users }));
            return users;
        }
    } catch (e) {
        console.warn("Failed to fetch assignable users: " + e.message);
    }
    return [];
}

function getPingInfo() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return {
        sheetId: sheet.getSheetId(),
        sheetName: sheet.getName()
    };
}

/**
 * Logic extracted for testability
 * @param {Array} data - 2D data array from sheet (excluding header)
 * @param {Object} indices - { createdIndex, startDateIndex, dueDateIndex, resolvedIndex }
 */
function calculateRoadmapDateRange(data, indices) {
    const { createdIndex, startDateIndex, dueDateIndex, resolvedIndex } = indices;

    const today = new Date();
    // Default range: 1 month back, 4 months forward
    let minDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    let maxDate = new Date(today.getFullYear(), today.getMonth() + 4, 0);

    // Scan data for actual date range
    data.forEach(row => {
        const created = createdIndex !== -1 ? (row[createdIndex] ? new Date(row[createdIndex]) : null) : null;
        const start = startDateIndex !== -1 ? (row[startDateIndex] ? new Date(row[startDateIndex]) : null) : null;
        const due = dueDateIndex !== -1 ? (row[dueDateIndex] ? new Date(row[dueDateIndex]) : null) : null;
        const resolved = resolvedIndex !== -1 ? (row[resolvedIndex] ? new Date(row[resolvedIndex]) : null) : null;

        // Update Min (Start or Created)
        if (start && !isNaN(start.getTime())) {
            if (start < minDate) minDate = start;
        } else if (created && !isNaN(created.getTime())) {
            if (created < minDate) minDate = created;
        }

        // Update Max (Due, Resolved, or Created + est)
        if (due && !isNaN(due.getTime())) {
            if (due > maxDate) maxDate = due;
        }
        if (resolved && !isNaN(resolved.getTime())) {
            if (resolved > maxDate) maxDate = resolved;
        }
    });

    // Add buffer
    const startDate = new Date(minDate);
    startDate.setDate(1); // Start at beginning of that month
    const endDate = new Date(maxDate);
    endDate.setMonth(endDate.getMonth() + 1); // Add 1 month buffer
    endDate.setDate(0); // End of that month

    // Calculate number of weeks for timeline columns
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    const totalWeeks = Math.ceil((endDate - startDate) / msPerWeek);

    return { startDate, endDate, totalWeeks, msPerWeek };
}

/**
 * Parses a JQLU command string into an update payload and JQL
 * @param {String} commandString - e.g. "UPDATE assignee='bob' WHERE project='A'"
 * @returns {Object} { updatePayload, jql }
 */
function parseJqluCommand(commandString) {
    const trimmedCommand = commandString.trim();

    // 1. Split UPDATE vs WHERE
    const whereIndex = trimmedCommand.toUpperCase().indexOf(' WHERE ');
    if (whereIndex === -1) {
        throw new Error("Syntax Error: Missing 'WHERE' clause. Usage: UPDATE field=value WHERE jql_condition");
    }

    const updatePart = trimmedCommand.substring(7, whereIndex).trim(); // Remove "UPDATE "
    const wherePart = trimmedCommand.substring(whereIndex + 7).trim(); // Remove " WHERE "

    if (!updatePart || !wherePart) {
        throw new Error("Syntax Error: Incomplete command.");
    }

    // 2. Parse Updates (Simple k=v parser)
    // We need to handle commas inside quotes.
    // Regex matches a comma only if it's not followed by an odd number of quotes (roughly).
    // Better approach: simple state machine or split by regex looking ahead.
    const kvPairs = updatePart.split(/,(?=(?:[^']*'[^']*')*[^']*$)/);

    const updatePayload = { fields: {} };

    kvPairs.forEach(pair => {
        const parts = pair.split('=');
        if (parts.length < 2) return; // Skip invalid pairs

        const key = parts[0].trim();
        const val = parts.slice(1).join('=').trim(); // Re-join in case value has =
        const cleanVal = val.replace(/^['"]|['"]$/g, ''); // Remove surrounding quotes

        // Map common aliases/structure
        if (key.toLowerCase() === 'assignee') {
            // Assignee needs object { name: "..." } or { accountId: "..." }
            if (cleanVal.includes(':')) {
                updatePayload.fields[key] = { accountId: cleanVal };
            } else {
                updatePayload.fields[key] = { name: cleanVal };
            }
        } else if (['priority', 'issuetype', 'status', 'resolution'].includes(key.toLowerCase())) {
            updatePayload.fields[key.toLowerCase()] = { name: cleanVal };
        } else {
            // Standard field
            // Try to detect numbers?
            if (!isNaN(cleanVal) && cleanVal !== '') {
                // keeping as string for safety unless sure
                updatePayload.fields[key] = cleanVal;
            } else {
                updatePayload.fields[key] = cleanVal;
            }
        }
    });

    return {
        updatePayload: updatePayload,
        jql: wherePart
    };
}

/**
 * Centralized fetch function to handle Jira requests.
 * Uses Cloudflare Worker proxy if enabled.
 */
function fetchJira(url, options, config) {
    let fetchUrl = url;
    let fetchOptions = { ...options };

    if (config && config.useCloudflare) {
        console.log(`[Proxy] Routing request through Cloudflare: ${url}`);
        // Use the baked-in Cloudflare Proxy URL
        const workerUrlString = CLOUDFLARE_WORKER_URL + (CLOUDFLARE_WORKER_URL.includes('?') ? '&' : '?') + "target=" + encodeURIComponent(url);
        fetchUrl = workerUrlString;

        // Add Proxy Secret to headers
        fetchOptions.headers = fetchOptions.headers || {};
        fetchOptions.headers["X-Proxy-Secret"] = PROXY_SECRET;

        // Track calls
        if (config && typeof config.proxyCallCount === 'number') {
            config.proxyCallCount++;
        }
    }

    return UrlFetchApp.fetch(fetchUrl, fetchOptions);
}

/**
 * Batched fetch function to handle Jira requests.
 * Uses Cloudflare Worker proxy if enabled.
 */
function fetchAllJira(requests, config) {
    if (config && config.useCloudflare) {
        console.log(`[Proxy] Batch routing ${requests.length} requests through Cloudflare`);
        const proxiedRequests = requests.map(req => {
            const workerUrlString = CLOUDFLARE_WORKER_URL + (CLOUDFLARE_WORKER_URL.includes('?') ? '&' : '?') + "target=" + encodeURIComponent(req.url);
            const newHeaders = { ...(req.headers || {}) };
            newHeaders["X-Proxy-Secret"] = PROXY_SECRET;

            return {
                ...req,
                url: workerUrlString,
                headers: newHeaders
            };
        });

        // Track batched calls
        if (config && typeof config.proxyCallCount === 'number') {
            config.proxyCallCount += proxiedRequests.length;
        }

        return UrlFetchApp.fetchAll(proxiedRequests);
    }
    return UrlFetchApp.fetchAll(requests);
}

/**
 * Helper to fetch with retry logic for "Service invoked too many times" errors
 */
function fetchWithRetry(url, options, config) {
    const maxRetries = 3;
    let attempt = 0;

    while (attempt < maxRetries) {
        try {
            return fetchJira(url, options, config);
        } catch (e) {
            const errorStr = e.toString();
            // Check for specific Google Apps Script rate limit exception
            if (errorStr.includes("Service invoked too many times") || errorStr.includes("Invoke limit exceeded")) {
                console.warn(`UrlFetchApp rate limit hit. Sleeping... Attempt ${attempt + 1}/${maxRetries}`);
                Utilities.sleep(1000 * Math.pow(2, attempt)); // 1s, 2s, 4s
                attempt++;
            } else if (errorStr.includes("PERMISSION_DENIED") && attempt < maxRetries - 1) {
                console.warn(`UrlFetchApp PERMISSION_DENIED (GAS transient error). Retrying... Attempt ${attempt + 1}/${maxRetries}`);
                Utilities.sleep(1000);
                attempt++;
            } else {
                throw e;
            }
        }
    }
    throw new Error(`Request failed after ${maxRetries} retries due to rate limiting.`);
}

/**
 * Batched version of updateJiraIssues to handle large volumes and rate limiting.
 */
function updateJiraIssuesBatched(params, updates, context = {}) {
    if (!updates || updates.length === 0) return { success: true, count: 0 };

    let isLimited = false;
    if (!params.isPro && updates.length > 150) {
        updates = updates.slice(0, 150);
        isLimited = true;
    }

    const cleanDomain = params.domain.replace(/^https?:\/\//, '').replace(/\/$/, '');

    const requests = updates.map(update => {
        const url = `https://${cleanDomain}/rest/api/3/issue/${update.key}`;
        return {
            url: url,
            method: 'PUT',
            contentType: 'application/json',
            headers: {
                "Authorization": "Basic " + Utilities.base64Encode(params.email + ":" + params.token)
            },
            payload: JSON.stringify({ fields: update.fields }),
            muteHttpExceptions: true
        };
    });

    const errors = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = context.sheetName ? ss.getSheetByName(context.sheetName) : ss.getActiveSheet();
    const successfulKeys = [];
    const rangesToClear = [];
    const errorHighlights = []; // Array of [{row, col, note}]
    let updatedCells = 0;

    // Load Modified Keys
    const props = PropertiesService.getDocumentProperties();
    let modifiedMap = {};
    try {
        const json = props.getProperty('modifiedKeys');
        if (json) modifiedMap = JSON.parse(json);
    } catch (err) { }

    // Determine Pro status if not provided
    if (params && typeof params.isPro === 'undefined') {
        const license = checkLicense();
        params.isPro = license.allowed && license.plan !== 'free';
    }

    // Dynamic Batching Configuration based on Proxy status
    const isProxy = params && params.useCloudflare && params.isPro;
    const BATCH_SIZE = isProxy ? 50 : 5; // Aggressive for proxy, conservative for direct
    const sleepBetweenBatches = isProxy ? 200 : 1000;
    const sleepBetweenRequests = isProxy ? 0 : 250;

    const batches = [];
    for (let i = 0; i < requests.length; i += BATCH_SIZE) {
        batches.push(requests.slice(i, i + BATCH_SIZE));
    }

    const batchId = context.batchId;
    const cache = batchId ? CacheService.getScriptCache() : null;

    batches.forEach((batchRequests, batchIndex) => {
        try {
            // Update Progress in Cache if we have a batchId
            if (cache) {
                const progressMsg = `Updating Batch ${batchIndex + 1}/${batches.length} (${batchIndex * BATCH_SIZE}-${Math.min((batchIndex + 1) * BATCH_SIZE, updates.length)} of ${updates.length})`;
                const currentMeta = JSON.parse(cache.get(batchId + "_meta") || "{}");
                cache.put(batchId + "_meta", JSON.stringify({ ...currentMeta, progress: progressMsg, status: 'RUNNING' }), 21600);
            }

            if (batchIndex > 0) Utilities.sleep(sleepBetweenBatches);

            let batchResponses = [];
            if (isProxy) {
                // PARALLEL PROCESSING (Super Fast via Cloudflare)
                batchResponses = fetchAllJira(batchRequests, params);
            } else {
                // SEQUENTIAL PROCESSING (Safe fallback for direct Google Apps Script)
                for (let i = 0; i < batchRequests.length; i++) {
                    const req = batchRequests[i];
                    try {
                        batchResponses.push(fetchJira(req.url, req, params));
                    } catch (e) {
                        console.warn(`Row request failed: ${e.message}`);
                        batchResponses.push({
                            getResponseCode: () => 500,
                            getContentText: () => JSON.stringify({ errorMessages: ["System Error: " + e.message] })
                        });
                    }
                    if (i < batchRequests.length - 1 && sleepBetweenRequests > 0) {
                        Utilities.sleep(sleepBetweenRequests);
                    }
                }
            }

            batchResponses.forEach((res, i) => {
                const globalIndex = (batchIndex * BATCH_SIZE) + i;
                const update = updates[globalIndex];

                if (res.getResponseCode() === 204) {
                    // Success
                    successfulKeys.push(update.key);
                    delete modifiedMap[update.key];

                    if (update.fields) {
                        updatedCells += Object.keys(update.fields).length;
                    }

                    if (update.fieldCols) {
                        Object.values(update.fieldCols).forEach(colIndex => {
                            rangesToClear.push(sheet.getRange(update.lineNumber, colIndex + 1).getA1Notation());
                        });
                    }
                } else {
                    // Error Handling
                    const row = update.lineNumber;
                    const content = res.getContentText();
                    let errorMessage = content;

                    try {
                        const errJson = JSON.parse(content);
                        if (errJson.errors) {
                            Object.entries(errJson.errors).forEach(([field, msg]) => {
                                const colIndex = update.fieldCols ? update.fieldCols[field] : -1;
                                if (colIndex !== -1) {
                                    errorHighlights.push({ row, col: colIndex + 1, note: `Error: ${msg}` });
                                }
                            });
                            errorMessage = JSON.stringify(errJson.errors);
                        }
                        if (errJson.errorMessages && errJson.errorMessages.length > 0) {
                            errorMessage = errJson.errorMessages.join(', ');
                        }
                    } catch (e) { }

                    errors.push(`Row ${row}: ${errorMessage}`);
                }
            });
        } catch (e) {
            // Mark batch as failed
            const startIdx = batchIndex * BATCH_SIZE;
            const endIdx = Math.min(startIdx + BATCH_SIZE, updates.length);
            for (let j = startIdx; j < endIdx; j++) {
                errors.push(`Row ${updates[j].lineNumber}: Batch failed - ${e.message}`);
            }

            // Log partial failure
            console.error(`Batch ${batchIndex + 1} failed: ${e.message}`);
            if (cache) {
                const currentMeta = JSON.parse(cache.get(batchId + "_meta") || "{}");
                cache.put(batchId + "_meta", JSON.stringify({ ...currentMeta, lastError: e.message }), 21600);
            }

            // Larger backoff if we hit a global error
            Utilities.sleep(3000);
        }
    });

    if (cache) {
        const currentMeta = JSON.parse(cache.get(batchId + "_meta") || "{}");
        cache.put(batchId + "_meta", JSON.stringify({ ...currentMeta, progress: 'Finalizing sheet updates...' }), 21600);
    }

    // Save updated Modified Keys
    if (successfulKeys.length > 0) {
        props.setProperty('modifiedKeys', JSON.stringify(modifiedMap));
        if (rangesToClear.length > 0) {
            const CLEAR_CHUNK = 1000;
            for (let i = 0; i < rangesToClear.length; i += CLEAR_CHUNK) {
                try {
                    const rangeList = sheet.getRangeList(rangesToClear.slice(i, i + CLEAR_CHUNK));
                    rangeList.setBackground('#b6d7a8'); // Light Green
                } catch (e) { }
            }
            SpreadsheetApp.flush();
            Utilities.sleep(800);
            for (let i = 0; i < rangesToClear.length; i += CLEAR_CHUNK) {
                try {
                    sheet.getRangeList(rangesToClear.slice(i, i + CLEAR_CHUNK)).setBackground(null);
                } catch (e) { }
            }
        }
    }

    // Apply Error Highlights in Bulk
    if (errorHighlights.length > 0) {
        errorHighlights.forEach(h => {
            try {
                sheet.getRange(h.row, h.col).setBackground('#ffcccc').setNote(h.note);
            } catch (e) { }
        });
    }

    if (errors.length > 0) {
        throw new Error(`Some updates failed. Check pink highlighted cells for details.\n\n${errors.slice(0, 5).join('\n')}${errors.length > 5 ? '\n...' : ''}`);
    }

    return { success: true, count: updates.length - errors.length, cellsUpdated: updatedCells, limited: isLimited };
}

/**
 * Direct entry point for the sidebar to run updates instantly.
 * This avoids the 60-second startup delay of background triggers.
 */
function updateJiraIssuesDirect(params, updates, context = {}) {
    try {
        return updateJiraIssuesBatched(params, updates, context);
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * Triggers a background update process.
 * Stores the update payload in CacheService/PropertiesService and spawns a trigger.
 */
function triggerBackgroundUpdate(params, updates) {
    if (!updates || updates.length === 0) return { success: true, message: 'No updates found.' };

    // Store the large payload. CacheService is best for short-term, but has size limits (100KB).
    // PropertiesService has size limits too (9KB/val).
    // For 2500 rows, payload is huge.
    // Solution: Store in a hidden sheet or chunk into multiple Cache entries?
    // Splitting into CacheService is safer.

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const BATCH_ID = Utilities.getUuid();
    const payload = JSON.stringify({
        params,
        updates,
        context: {
            ssId: ss.getId(),
            sheetName: sheet.getName(),
            batchId: BATCH_ID
        }
    });

    const CHUNK_SIZE = 90000; // 90KB safety limit
    const totalChunks = Math.ceil(payload.length / CHUNK_SIZE);
    const cache = CacheService.getScriptCache();

    try {
        for (let i = 0; i < totalChunks; i++) {
            const chunk = payload.substr(i * CHUNK_SIZE, CHUNK_SIZE);
            cache.put(BATCH_ID + "_chunk_" + i, chunk, 21600); // 6 hours
        }
        cache.put(BATCH_ID + "_meta", JSON.stringify({ totalChunks, status: 'QUEUED' }), 21600);
    } catch (e) {
        return { success: false, message: "Payload too large for background processing. Try fewer rows." };
    }

    // Store the batch ID uniquely for this user to avoid collisions
    const userEmail = Session.getActiveUser() ? Session.getActiveUser().getEmail() : 'ANON';
    PropertiesService.getDocumentProperties().setProperty('BG_BATCH_' + userEmail, BATCH_ID);

    ScriptApp.newTrigger('processBackgroundUpdate')
        .timeBased()
        .after(100)
        .create();

    return { success: true, batchId: BATCH_ID, message: "Update started in background." };
}

/**
 * The function triggered by the time-driven trigger.
 */
function processBackgroundUpdate() {
    // Delete any pending triggers to avoid duplicate runs
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === 'processBackgroundUpdate') {
            try { ScriptApp.deleteTrigger(t); } catch (e) { }
        }
    });

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return;

    const props = PropertiesService.getDocumentProperties();
    const userEmail = Session.getActiveUser() ? Session.getActiveUser().getEmail() : 'ANON';
    const batchId = props.getProperty('BG_BATCH_' + userEmail);

    if (!batchId) {
        lock.releaseLock();
        return;
    }

    const cache = CacheService.getScriptCache();
    const metaJson = cache.get(batchId + "_meta");
    if (!metaJson) {
        console.error(`[BG] Meta missing for batch ${batchId}`);
        props.deleteProperty('BG_BATCH_' + userEmail);
        lock.releaseLock();
        return;
    }

    const meta = JSON.parse(metaJson);
    // Mark as running
    cache.put(batchId + "_meta", JSON.stringify({ ...meta, status: 'RUNNING', progress: 'Initializing worker...' }), 21600);

    try {
        let fullPayload = "";
        for (let i = 0; i < meta.totalChunks; i++) {
            const chunk = cache.get(batchId + "_chunk_" + i);
            if (!chunk) throw new Error("Missing payload chunk " + i);
            fullPayload += chunk;
        }

        const data = JSON.parse(fullPayload);
        const { params, updates, context } = data;

        console.log(`[BG] Starting update for ${updates.length} issues`);

        const result = updateJiraIssuesBatched(params, updates, context);

        // Final Status Update
        cache.put(batchId + "_meta", JSON.stringify({ ...meta, status: 'COMPLETED', result: result, progress: 'Completed' }), 21600);

        try {
            const type = params.triggerType || 'Background';
            logRefreshAttempt('OK', `[${type}] ${result.count} issues updated.`, context.sheetName || 'Unknown');
        } catch (e) { }

    } catch (e) {
        console.error(`[BG] Critical failure: ${e.message}`);
        cache.put(batchId + "_meta", JSON.stringify({ ...meta, status: 'FAILED', error: e.message }), 21600);
        try { logRefreshAttempt('FAILED', `Background Update Failed: ${e.message}`, 'SYSTEM'); } catch (err) { }
    } finally {
        // Clean up props
        props.deleteProperty('BG_BATCH_' + userEmail);
        lock.releaseLock();
    }
}

function checkBatchStatus(batchId) {
    const cache = CacheService.getScriptCache();
    const metaJson = cache.get(batchId + "_meta");
    if (!metaJson) return { status: 'UNKNOWN' };
    return JSON.parse(metaJson);
}


