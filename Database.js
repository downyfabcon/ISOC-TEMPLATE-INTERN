/**
 * Web App Template - Database Operations
 * CRUD functions for User Management with LockService concurrency control
 */

/**
 * Gets all users from the Main DB USERS sheet
 * @returns {Object} Result object with users array or error
 */
function getUsers() {
    if (!isCurrentUserAdmin()) {
        return {
            success: false,
            message: 'Access denied. Admin privileges required.'
        };
    }

    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return {
            success: false,
            message: 'Main database not configured'
        };
    }

    try {
        const ss = SpreadsheetApp.openById(mainDbId);
        const sheet = ss.getSheetByName('USERS') || ss.getSheets()[0];
        const data = sheet.getDataRange().getValues();

        const users = [];

        // Skip header row
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[0] && row[0].toString().trim() !== '') {
                users.push({
                    id: i + 1, // Row number as ID
                    email: row[0] || '',
                    name: row[1] || '',
                    role: row[2] || 'user'
                });
            }
        }

        return {
            success: true,
            users: users
        };
    } catch (error) {
        return {
            success: false,
            message: 'Error fetching users: ' + error.message
        };
    }
}

/**
 * Creates a new user in the USERS sheet
 * @param {Object} userData - User data with email, name, role
 * @returns {Object} Result object with success status
 */
function createUser(userData) {
    if (!isCurrentUserAdmin()) {
        return {
            success: false,
            message: 'Access denied. Admin privileges required.'
        };
    }

    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return {
            success: false,
            message: 'Main database not configured'
        };
    }

    const lock = LockService.getScriptLock();

    try {
        // Wait up to 10 seconds for lock
        if (!lock.tryLock(10000)) {
            return {
                success: false,
                message: 'Could not obtain lock. Please try again.'
            };
        }

        const ss = SpreadsheetApp.openById(mainDbId);
        const sheet = ss.getSheetByName('USERS') || ss.getSheets()[0];

        // Check if email already exists
        const existingUser = findUserInSheet(sheet, userData.email);

        if (existingUser) {
            return {
                success: false,
                message: 'A user with this email already exists'
            };
        }

        // Validate role
        const validRoles = ['admin', 'user'];
        const role = validRoles.includes(userData.role) ? userData.role : 'user';

        // Add new user
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, 1, 3).setValues([
            [userData.email.trim(), userData.name.trim(), role]
        ]);

        // Update cache version
        updateDataVersion();

        // Log audit event
        try {
            const newUser = {
                email: userData.email.trim(),
                name: userData.name.trim(),
                role: role
            };
            const summary = generateChangeSummary('CREATE', 'USER', null, { identifier: userData.email.trim() });
            logAuditEvent('CREATE', 'USER', userData.email.trim(), newUser, summary);
        } catch (e) {
            Logger.log('Failed to log audit event: ' + e.message);
        }

        return {
            success: true,
            message: 'User created successfully',
            user: {
                id: lastRow + 1,
                email: userData.email.trim(),
                name: userData.name.trim(),
                role: role
            }
        };
    } catch (error) {
        return {
            success: false,
            message: 'Error creating user: ' + error.message
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Updates an existing user in the USERS sheet
 * @param {Object} userData - User data with id, email, name, role
 * @returns {Object} Result object with success status
 */
function updateUser(userData) {
    if (!isCurrentUserAdmin()) {
        return {
            success: false,
            message: 'Access denied. Admin privileges required.'
        };
    }

    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return {
            success: false,
            message: 'Main database not configured'
        };
    }

    const lock = LockService.getScriptLock();

    try {
        if (!lock.tryLock(10000)) {
            return {
                success: false,
                message: 'Could not obtain lock. Please try again.'
            };
        }

        const ss = SpreadsheetApp.openById(mainDbId);
        const sheet = ss.getSheetByName('USERS') || ss.getSheets()[0];
        const rowIndex = userData.id;

        // Verify row exists
        const currentData = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];

        if (!currentData[0] || currentData[0].toString().trim() === '') {
            return {
                success: false,
                message: 'User not found'
            };
        }

        // Check if changing email to one that already exists
        if (currentData[0].toString().toLowerCase() !== userData.email.toLowerCase()) {
            const existingUser = findUserInSheet(sheet, userData.email);
            if (existingUser && existingUser.rowIndex !== rowIndex) {
                return {
                    success: false,
                    message: 'A user with this email already exists'
                };
            }
        }

        // Validate role
        const validRoles = ['admin', 'user'];
        const role = validRoles.includes(userData.role) ? userData.role : 'user';

        // Capture old values for audit
        const oldValues = {
            email: currentData[0].toString(),
            name: currentData[1].toString(),
            role: currentData[2].toString()
        };
        const newValues = {
            email: userData.email.trim(),
            name: userData.name.trim(),
            role: role
        };

        // Update user
        sheet.getRange(rowIndex, 1, 1, 3).setValues([
            [userData.email.trim(), userData.name.trim(), role]
        ]);

        // Update cache version
        updateDataVersion();

        // Log audit event
        try {
            const changes = calculateChanges(oldValues, newValues);
            if (Object.keys(changes).length > 0) {
                const summary = generateChangeSummary('UPDATE', 'USER', changes, { identifier: userData.email.trim() });
                logAuditEvent('UPDATE', 'USER', userData.email.trim(), changes, summary);
            }
        } catch (e) {
            Logger.log('Failed to log audit event: ' + e.message);
        }

        return {
            success: true,
            message: 'User updated successfully',
            user: {
                id: rowIndex,
                email: userData.email.trim(),
                name: userData.name.trim(),
                role: role
            }
        };
    } catch (error) {
        return {
            success: false,
            message: 'Error updating user: ' + error.message
        };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Deletes a user from the USERS sheet
 * @param {number} userId - The row ID of the user to delete
 * @returns {Object} Result object with success status
 */
function deleteUser(userId) {
    if (!isCurrentUserAdmin()) {
        return {
            success: false,
            message: 'Access denied. Admin privileges required.'
        };
    }

    const mainDbId = getMainDbId();

    if (!mainDbId) {
        return {
            success: false,
            message: 'Main database not configured'
        };
    }

    const lock = LockService.getScriptLock();

    try {
        if (!lock.tryLock(10000)) {
            return {
                success: false,
                message: 'Could not obtain lock. Please try again.'
            };
        }

        const ss = SpreadsheetApp.openById(mainDbId);
        const sheet = ss.getSheetByName('USERS') || ss.getSheets()[0];

        const rowId = Number(userId);
        if (!Number.isFinite(rowId) || rowId < 2) {
            return {
                success: false,
                message: 'Invalid user id'
            };
        }

        const lastRow = sheet.getLastRow();
        if (rowId > lastRow) {
            return {
                success: false,
                message: 'User not found'
            };
        }

        // Verify row exists
        const currentData = sheet.getRange(rowId, 1, 1, 3).getValues()[0];

        if (!currentData[0] || currentData[0].toString().trim() === '') {
            return {
                success: false,
                message: 'User not found'
            };
        }

        // Prevent deleting yourself
        const currentUserEmail = getCurrentUserEmail();
        if (currentData[0].toString().toLowerCase() === currentUserEmail.toLowerCase()) {
            return {
                success: false,
                message: 'You cannot delete your own account'
            };
        }

        // Capture user info for audit before deletion
        const deletedUser = {
            email: currentData[0].toString(),
            name: currentData[1].toString(),
            role: currentData[2].toString()
        };

        // Delete the row
        sheet.deleteRow(rowId);

        // Update cache version
        updateDataVersion();

        // Log audit event
        try {
            const summary = generateChangeSummary('DELETE', 'USER', null, { identifier: deletedUser.email });
            logAuditEvent('DELETE', 'USER', deletedUser.email, deletedUser, summary);
        } catch (e) {
            Logger.log('Failed to log audit event: ' + e.message);
        }

        return {
            success: true,
            message: 'User deleted successfully'
        };
    } catch (error) {
        return {
            success: false,
            message: 'Error deleting user: ' + error.message
        };
    } finally {
        lock.releaseLock();
    }
}
