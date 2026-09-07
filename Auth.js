/**
 * Web App Template - Authentication Module
 * Handles user authentication, role-based access, and first-time setup
 */

/**
 * Gets the current user's email address
 * @returns {string} The user's email address
 */
function getCurrentUserEmail() {
    return Session.getActiveUser().getEmail();
}

/**
 * Gets the current user's info from the database
 * @returns {Object|null} User info object or null
 */
function getCurrentUserInfo() {
    const email = getCurrentUserEmail();
    return findUserByEmail(email);
}

/**
 * Gets the authentication status and user info
 * This is the main authentication entry point called on app load
 * @returns {Object} Authentication result object
 */
function authenticate() {
    const email = getCurrentUserEmail();
    const mainDbId = getMainDbId();

    // Check if Main DB is configured
    if (!mainDbId) {
        return {
            status: 'SETUP_REQUIRED',
            email: email,
            message: 'Main database not configured. Initial setup required.'
        };
    }

    // Look up user in the database
    try {
        const user = findUserByEmail(email);

        if (!user) {
            return {
                status: 'ACCESS_DENIED',
                email: email,
                message: 'Your email is not registered in the system. Please contact an administrator.'
            };
        }

        return {
            status: 'AUTHENTICATED',
            email: user.email,
            name: user.name,
            role: user.role,
            isAdmin: user.role === 'admin'
        };
    } catch (error) {
        return {
            status: 'ERROR',
            email: email,
            message: 'Authentication error: ' + error.message
        };
    }
}

/**
 * Finds a user by email in the Users sheet
 * @param {string} email - The email to search for
 * @returns {Object|null} User object or null if not found
 */
function findUserByEmail(email) {
    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return null;
    }

    try {
        const ss = SpreadsheetApp.openById(mainDbId);
        const sheet = ss.getSheetByName('USERS');

        if (!sheet) {
            // USERS sheet is required - no fallback to prevent data leaks
            Logger.log('USERS sheet not found in main database');
            return null;
        }

        return findUserInSheet(sheet, email);
    } catch (error) {
        Logger.log('Error finding user: ' + error.message);
        throw error;
    }
}

/**
 * Helper function to find user in a specific sheet
 * @param {Sheet} sheet - The sheet to search
 * @param {string} email - The email to find
 * @returns {Object|null} User object with rowIndex or null
 */
function findUserInSheet(sheet, email) {
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        const rowEmail = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';

        if (rowEmail === email.toLowerCase()) {
            return {
                email: data[i][0],
                name: data[i][1] || '',
                role: data[i][2] || 'user',
                rowIndex: i + 1
            };
        }
    }

    return null;
}

/**
 * Performs initial setup - creates first admin user and sets Main DB
 * @param {Object} setupData - Setup data with spreadsheetInput, email, name
 * @returns {Object} Result object with success status
 */
function performInitialSetup(setupData) {
    const { spreadsheetInput, email, name } = setupData;

    // Extract and validate spreadsheet ID
    const spreadsheetId = extractSpreadsheetId(spreadsheetInput);

    if (!spreadsheetId) {
        return {
            success: false,
            message: 'Invalid spreadsheet URL or ID'
        };
    }

    // Validate access to spreadsheet
    const validation = validateSpreadsheetAccess(spreadsheetId);

    if (!validation.success) {
        return {
            success: false,
            message: validation.message
        };
    }

    try {
        const ss = SpreadsheetApp.openById(spreadsheetId);

        // Check if USERS sheet exists, if not create it
        let sheet = ss.getSheetByName('USERS');
        if (!sheet) {
            sheet = ss.insertSheet('USERS');
            sheet.getRange(1, 1, 1, 3).setValues([['Email', 'Name', 'Role']]);
            sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
            sheet.setFrozenRows(1);
        }

        // Check if headers exist
        const headers = sheet.getRange(1, 1, 1, 3).getValues()[0];
        if (!headers[0] || headers[0].toString().trim() === '') {
            sheet.getRange(1, 1, 1, 3).setValues([['Email', 'Name', 'Role']]);
        }

        // Check if user already exists
        const existingUser = findUserInSheet(sheet, email);

        if (existingUser) {
            // Update existing user to admin
            sheet.getRange(existingUser.rowIndex, 3).setValue('admin');
        } else {
            // Add new admin user
            const lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[email, name, 'admin']]);
        }

        // Create AUDIT LOGS sheet if it doesn't exist
        let auditSheet = ss.getSheetByName('AUDIT LOGS');
        if (!auditSheet) {
            auditSheet = ss.insertSheet('AUDIT LOGS');
            auditSheet.appendRow([
                'Timestamp',
                'User Email',
                'User Name',
                'Action',
                'Entity Type',
                'Entity ID',
                'Changes',
                'Summary'
            ]);
            auditSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
            auditSheet.setFrozenRows(1);
        }

        // Create QUICK LINKS sheet if it doesn't exist
        let quickLinksSheet = ss.getSheetByName('QUICK LINKS');
        if (!quickLinksSheet) {
            quickLinksSheet = ss.insertSheet('QUICK LINKS');
            quickLinksSheet.appendRow(['Name', 'Link', 'Category', 'Icon']);
            quickLinksSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
            quickLinksSheet.setFrozenRows(1);
        }

        // Save the Main DB ID
        setMainDbId(spreadsheetId);

        // Update cache version
        updateDataVersion();

        return {
            success: true,
            message: 'Initial setup completed successfully',
            spreadsheetName: validation.name
        };
    } catch (error) {
        return {
            success: false,
            message: 'Setup failed: ' + error.message
        };
    }
}

/**
 * Checks if the current user is an admin
 * @returns {boolean} True if user is admin
 */
function isCurrentUserAdmin() {
    const email = getCurrentUserEmail();
    const user = findUserByEmail(email);
    return user && user.role === 'admin';
}
