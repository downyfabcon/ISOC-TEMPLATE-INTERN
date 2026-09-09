/**
 * Web App Template - Configuration Management
 * Handles PropertiesService for storing spreadsheet IDs and app configuration
 *
 * Main DB contains three sheets:
 * - USERS: User authentication (Email, Name, Role)
 * - AUDIT LOGS: Audit trail (Timestamp, User Email, User Name, Action, Entity Type, Entity ID, Changes, Summary)
 * - QUICK LINKS: quick links (Name, Link, Category, Icon)
 */

// ==================== APP BRANDING CONFIG ====================
// Modify these values to customize the app branding
const APP_CONFIG = {
    // App name displayed in sidebar header and mobile navbar
    name: 'ISOC Template',

    // Subtitle displayed below the app name in sidebar
    subtitle: 'Template for Web Apps',

    // Page title shown in browser tab
    pageTitle: 'ISOC Template',

    // App icon (SVG path from QUICK_LINK_ICON_SVGS - e.g., 'cog', 'rocket', 'sparkles')
    // Leave empty or remove to use the default cog icon
    icon: 'folder',

    // Optional favicon URL for the web app (e.g., https://.../favicon.png)
    // Leave blank to keep the default.
    faviconUrl: 'https://i.ibb.co/HLV9Cr0g/globe-b.png',

    // Default theme (must match one of availableThemes values)
    defaultTheme: 'dark',

    // Primary color class (for custom styling if needed)
    primaryColor: 'primary',

    // App version (optional, for display purposes)
    version: '1.0.0',

    // Footer text (optional)
    footerText: '🧩 Sophisticated Simplicity | By Louis Lagare'
};

/**
 * Gets the app configuration for client-side use
 * @returns {Object} App configuration object
 */
function getAppConfig() {
    return APP_CONFIG;
}

// Property keys
const PROP_KEYS = {
    MAIN_DB_ID: 'MAIN_DB_ID'
};

/**
 * Gets a property value from Script Properties
 * @param {string} key - The property key
 * @returns {string|null} The property value or null if not found
 */
function getProperty(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * Sets a property value in Script Properties
 * @param {string} key - The property key
 * @param {string} value - The property value
 */
function setProperty(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * Deletes a property from Script Properties
 * @param {string} key - The property key to delete
 */
function deleteProperty(key) {
    PropertiesService.getScriptProperties().deleteProperty(key);
}

/**
 * Gets the Main Database Spreadsheet ID
 * @returns {string|null} The spreadsheet ID or null if not configured
 */
function getMainDbId() {
    return getProperty(PROP_KEYS.MAIN_DB_ID);
}

/**
 * Sets the Main Database Spreadsheet ID
 * @param {string} spreadsheetId - The spreadsheet ID
 */
function setMainDbId(spreadsheetId) {
    setProperty(PROP_KEYS.MAIN_DB_ID, spreadsheetId);
    updateDataVersion();
}

/**
 * Extracts the Spreadsheet ID from a URL or returns the ID if already an ID
 * Accepts formats:
 * - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 * - Just the ID: SPREADSHEET_ID
 * @param {string} input - The spreadsheet URL or ID
 * @returns {string|null} The extracted spreadsheet ID or null if invalid
 */
function extractSpreadsheetId(input) {
    if (!input || typeof input !== 'string') {
        return null;
    }

    input = input.trim();

    // Check if it's a URL
    const urlPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
    const match = input.match(urlPattern);

    if (match) {
        return match[1];
    }

    // Check if it's already a valid ID format (alphanumeric with dashes and underscores)
    const idPattern = /^[a-zA-Z0-9-_]+$/;
    if (idPattern.test(input) && input.length > 10) {
        return input;
    }

    return null;
}

/**
 * Validates that a spreadsheet ID is accessible
 * @param {string} spreadsheetId - The spreadsheet ID to validate
 * @returns {Object} Result object with success status and message
 */
function validateSpreadsheetAccess(spreadsheetId) {
    try {
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const name = ss.getName();
        return {
            success: true,
            message: 'Spreadsheet accessible: ' + name,
            name: name
        };
    } catch (error) {
        return {
            success: false,
            message: 'Cannot access spreadsheet: ' + error.message,
            name: null
        };
    }
}

/**
 * Gets the current database configuration status
 * @returns {Object} Configuration status object
 */
function getDbConfig() {
    const mainDbId = getMainDbId();

    return {
        success: true,
        mainDbUrl: mainDbId || '',
        mainDb: {
            id: mainDbId,
            configured: !!mainDbId
        }
    };
}

/**
 * Saves database configuration (called from client)
 * @param {Object} config - Configuration object with mainDbInput
 * @returns {Object} Result object with success status and messages
 */
function saveDbConfig(config) {
    const results = {
        mainDb: { success: false, message: '' }
    };

    const oldConfig = getDbConfig();
    const configChanges = {};

    // Process Main DB
    if (config.mainDbInput) {
        const mainDbId = extractSpreadsheetId(config.mainDbInput);
        if (mainDbId) {
            const validation = validateSpreadsheetAccess(mainDbId);
            if (validation.success) {
                const oldId = oldConfig.mainDb.id;
                setMainDbId(mainDbId);
                results.mainDb = { success: true, message: 'Main DB configured: ' + validation.name, id: mainDbId };

                if (oldId !== mainDbId) {
                    configChanges.mainDbId = { old: oldId || '(not set)', new: mainDbId };
                }
            } else {
                results.mainDb = { success: false, message: validation.message };
            }
        } else {
            results.mainDb = { success: false, message: 'Invalid spreadsheet URL or ID' };
        }
    }

    // Log audit event if any config changed
    if (Object.keys(configChanges).length > 0) {
        try {
            const summary = generateChangeSummary('UPDATE', 'DB_CONFIG', configChanges, {});
            logAuditEvent('UPDATE', 'DB_CONFIG', 'settings', configChanges, summary);
        } catch (e) {
            Logger.log('Failed to log audit event: ' + e.message);
        }
    }

    return results;
}

/**
 * Verifies Main Database structure
 * Checks for required sheets: USERS, AUDIT LOGS, QUICK LINKS
 * @returns {Object} Verification result
 */
function verifyMainDb() {
    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return {
            success: false,
            status: 'not_configured',
            message: 'Main DB not configured'
        };
    }

    try {
        const ss = SpreadsheetApp.openById(mainDbId);

        // Check for required sheets
        const usersSheet = ss.getSheetByName('USERS');
        const auditSheet = ss.getSheetByName('AUDIT LOGS');
        const quickLinksSheet = ss.getSheetByName('QUICK LINKS');

        const missingSheets = [];
        if (!usersSheet) missingSheets.push('USERS');
        if (!auditSheet) missingSheets.push('AUDIT LOGS');
        if (!quickLinksSheet) missingSheets.push('QUICK LINKS');

        if (missingSheets.length > 0) {
            const availableSheets = ss.getSheets().map(s => s.getName()).slice(0, 5);
            return {
                success: false,
                status: 'missing_sheet',
                message: 'Missing sheet(s): "' + missingSheets.join('", "') + '"',
                spreadsheetName: ss.getName(),
                availableSheets: availableSheets
            };
        }

        // Check USERS sheet for required columns
        const headers = usersSheet.getRange(1, 1, 1, 4).getValues()[0];
        const requiredHeaders = ['Email', 'Name', 'Role'];
        const foundHeaders = headers.map(h => String(h).toLowerCase().trim());

        const missingHeaders = requiredHeaders.filter(req =>
            !foundHeaders.some(found => found.includes(req.toLowerCase()))
        );

        if (missingHeaders.length > 0) {
            return {
                success: false,
                status: 'missing_columns',
                message: 'Missing columns in USERS: ' + missingHeaders.join(', '),
                spreadsheetName: ss.getName()
            };
        }

        const userCount = Math.max(0, usersSheet.getLastRow() - 1);
        const auditCount = Math.max(0, auditSheet.getLastRow() - 1);
        const quickLinksCount = Math.max(0, quickLinksSheet.getLastRow() - 1);

        return {
            success: true,
            status: 'ok',
            message: `Valid - ${userCount} user(s), ${auditCount} audit log(s), ${quickLinksCount} quick link(s)`,
            spreadsheetName: ss.getName(),
            sheets: {
                users: userCount,
                auditLogs: auditCount,
                quickLinks: quickLinksCount
            }
        };

    } catch (error) {
        return {
            success: false,
            status: 'error',
            message: 'Access error: ' + error.message
        };
    }
}

/**
 * Creates missing sheets in the Main DB
 * @returns {Object} Result of the operation
 */
function createMissingSheets() {
    if (!isCurrentUserAdmin()) {
        return { success: false, message: 'Unauthorized: Admin access required' };
    }

    const mainDbId = getMainDbId();
    if (!mainDbId) {
        return { success: false, message: 'Main DB not configured' };
    }

    try {
        const ss = SpreadsheetApp.openById(mainDbId);
        const createdSheets = [];

        // Create USERS sheet if missing
        if (!ss.getSheetByName('USERS')) {
            const sheet = ss.insertSheet('USERS');
            sheet.getRange(1, 1, 1, 3).setValues([['Email', 'Name', 'Role']]);
            sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
            sheet.setFrozenRows(1);
            createdSheets.push('USERS');
        }

        // Create AUDIT LOGS sheet if missing
        if (!ss.getSheetByName('AUDIT LOGS')) {
            const sheet = ss.insertSheet('AUDIT LOGS');
            sheet.appendRow([
                'Timestamp',
                'User Email',
                'User Name',
                'Action',
                'Entity Type',
                'Entity ID',
                'Changes',
                'Summary'
            ]);
            sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
            sheet.setFrozenRows(1);
            createdSheets.push('AUDIT LOGS');
        }

        // Create QUICK LINKS sheet if missing
        if (!ss.getSheetByName('QUICK LINKS')) {
            const sheet = ss.insertSheet('QUICK LINKS');
            sheet.appendRow(['ID', 'Name', 'Link', 'Category', 'Icon']);
            sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f0f0f0');
            sheet.setFrozenRows(1);
            createdSheets.push('QUICK LINKS');
        }

        if (createdSheets.length > 0) {
            return {
                success: true,
                message: 'Created sheets: ' + createdSheets.join(', ')
            };
        } else {
            return {
                success: true,
                message: 'All required sheets already exist'
            };
        }

    } catch (error) {
        return {
            success: false,
            message: 'Error creating sheets: ' + error.message
        };
    }
}
