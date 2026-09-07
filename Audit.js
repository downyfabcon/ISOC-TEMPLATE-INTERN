/**
 * Web App Template - Audit Trail Functions
 * Handles logging and retrieval of audit events
 * 
 * Audit logs are stored in the Main DB "AUDIT LOGS" sheet
 * 
 * Entity Types (extensible - add new types as needed):
 * - USER: User management operations
 * - DB_CONFIG: Database configuration changes
 * - QUICK_LINK: Quick link operations
 * 
 * To add new entity types in the future, simply use them in logAuditEvent()
 * and add them to the filter dropdown in AuditTrail.html
 */

const AUDIT_SHEET_NAME = 'AUDIT LOGS';
const AUDIT_RETENTION_DAYS_KEY = 'AUDIT_RETENTION_DAYS';
const DEFAULT_RETENTION_DAYS = 90; // 3 months
const CLEANUP_PROBABILITY = 0.01; // ~1% chance = roughly once per 100 events

/**
 * Get the audit retention period in days
 */
function getAuditRetentionDays() {
    const props = PropertiesService.getScriptProperties();
    const days = props.getProperty(AUDIT_RETENTION_DAYS_KEY);
    return days ? parseInt(days, 10) : DEFAULT_RETENTION_DAYS;
}

/**
 * Set the audit retention period in days
 */
function setAuditRetentionDays(days) {
    if (!isCurrentUserAdmin()) {
        return { success: false, message: 'Unauthorized: Admin access required' };
    }

    const numDays = parseInt(days, 10);
    if (isNaN(numDays) || numDays < 1 || numDays > 365) {
        return { success: false, message: 'Retention days must be between 1 and 365' };
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty(AUDIT_RETENTION_DAYS_KEY, numDays.toString());

    return { success: true, message: `Audit retention set to ${numDays} days` };
}

/**
 * Get the audit log sheet, creating headers if needed
 */
function getAuditSheet() {
    const mainDbId = getMainDbId();
    if (!mainDbId) {
        throw new Error('Main Database not configured');
    }

    const ss = SpreadsheetApp.openById(mainDbId);
    let sheet = ss.getSheetByName(AUDIT_SHEET_NAME);

    if (!sheet) {
        // Create the sheet with headers
        sheet = ss.insertSheet(AUDIT_SHEET_NAME);
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
        // Format header row
        sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
        sheet.setFrozenRows(1);
    }

    return sheet;
}

/**
 * Log an audit event
 * @param {string} action - The action performed (CREATE, UPDATE, DELETE, etc.)
 * @param {string} entityType - Type of entity (USER, DB_CONFIG, QUICK_LINK, or custom types)
 * @param {string} entityId - Identifier for the entity
 * @param {Object} changes - Object containing changed fields {field: {old, new}}
 * @param {string} summary - Human-readable summary of the change
 */
function logAuditEvent(action, entityType, entityId, changes, summary) {
    try {
        const user = Session.getActiveUser();
        const userEmail = user.getEmail();

        // Get user's display name from the user database
        let userName = userEmail;
        try {
            const userInfo = getCurrentUserInfo();
            if (userInfo && userInfo.name) {
                userName = userInfo.name;
            }
        } catch (e) {
            // Fall back to email if we can't get the name
        }

        const sheet = getAuditSheet();
        const timestamp = new Date().toISOString();

        // Convert changes object to JSON string for storage
        const changesJson = changes ? JSON.stringify(changes) : '';

        // Insert at row 2 (after header) to keep newest first
        sheet.insertRowAfter(1);
        const targetRange = sheet.getRange(2, 1, 1, 8);
        targetRange.setValues([[
            timestamp,
            userEmail,
            userName,
            action,
            entityType,
            entityId || '',
            changesJson,
            summary || ''
        ]]);

        // New rows inherit formatting from the row above (the header), which makes
        // the inserted audit entry look like a header. Normalize formatting by
        // copying the previous data row's format (now at row 3) when available.
        const lastRow = sheet.getLastRow();
        if (lastRow >= 3) {
            sheet.getRange(3, 1, 1, 8).copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        } else {
            // If this is the first data row, explicitly reset basic header styling.
            targetRange.setFontWeight('normal').setBackground('#ffffff');
        }

        // Only run cleanup occasionally (roughly once per day based on random chance)
        // This prevents cleanup from running on every single audit event
        if (Math.random() < CLEANUP_PROBABILITY) {
            cleanupOldAuditLogs();
        }

    } catch (error) {
        Logger.log('Error logging audit event: ' + error.message);
        // Don't throw - audit logging failure shouldn't break the main operation
    }
}

/**
 * Calculate the changes between old and new values
 * @param {Object} oldValues - The original values
 * @param {Object} newValues - The new values
 * @param {Array} fieldsToTrack - Optional array of field names to track
 * @returns {Object} Object with changed fields: {fieldName: {old: value, new: value}}
 */
function calculateChanges(oldValues, newValues, fieldsToTrack) {
    const changes = {};
    const fields = fieldsToTrack || Object.keys(newValues);

    for (const field of fields) {
        const oldVal = oldValues[field];
        const newVal = newValues[field];

        // Handle arrays
        if (Array.isArray(oldVal) || Array.isArray(newVal)) {
            const oldArr = Array.isArray(oldVal) ? oldVal : [];
            const newArr = Array.isArray(newVal) ? newVal : [];

            if (JSON.stringify(oldArr) !== JSON.stringify(newArr)) {
                changes[field] = { old: oldArr, new: newArr };
            }
        } else if (oldVal !== newVal) {
            // Only track if values are actually different
            changes[field] = { old: oldVal || '', new: newVal || '' };
        }
    }

    return changes;
}

/**
 * Generate a human-readable summary of changes
 * @param {string} action - The action performed
 * @param {string} entityType - Type of entity
 * @param {Object} changes - The changes object
 * @param {Object} context - Additional context (e.g., identifier, name, email)
 * @returns {string} Human-readable summary
 */
function generateChangeSummary(action, entityType, changes, context) {
    const contextStr = context ? (context.identifier || context.name || context.email || '') : '';

    if (action === 'CREATE') {
        return `Created ${entityType.toLowerCase().replace('_', ' ')}${contextStr ? ': ' + contextStr : ''}`;
    }

    if (action === 'DELETE') {
        return `Deleted ${entityType.toLowerCase().replace('_', ' ')}${contextStr ? ': ' + contextStr : ''}`;
    }

    if (action === 'UPDATE') {
        const changedFields = Object.keys(changes || {});
        if (changedFields.length === 0) {
            return `Updated ${entityType.toLowerCase().replace('_', ' ')}${contextStr ? ': ' + contextStr : ''} (no changes detected)`;
        }

        const fieldList = changedFields.slice(0, 3).join(', ');
        const moreCount = changedFields.length > 3 ? ` +${changedFields.length - 3} more` : '';
        return `Updated ${entityType.toLowerCase().replace('_', ' ')}${contextStr ? ' ' + contextStr : ''}: ${fieldList}${moreCount}`;
    }

    return `${action} ${entityType.toLowerCase().replace('_', ' ')}${contextStr ? ': ' + contextStr : ''}`;
}

/**
 * Clean up audit logs older than the retention period
 * Optimized to delete rows in batches for better performance
 */
function cleanupOldAuditLogs() {
    try {
        const retentionDays = getAuditRetentionDays();
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

        const sheet = getAuditSheet();
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) return; // Only header, nothing to clean

        // Get only timestamp column for efficiency
        const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

        // Find the first row that's too old (since data is newest-first, old rows are at bottom)
        let firstOldRow = -1;
        for (let i = timestamps.length - 1; i >= 0; i--) {
            const timestamp = new Date(timestamps[i][0]);
            if (timestamp < cutoffDate) {
                firstOldRow = i + 2; // Convert to 1-indexed (accounting for header)
            } else {
                break; // Since data is sorted newest-first, we can stop when we hit recent data
            }
        }

        // Delete all old rows in one batch operation
        if (firstOldRow > 1) {
            const rowsToDelete = lastRow - firstOldRow + 1;
            if (rowsToDelete > 0) {
                sheet.deleteRows(firstOldRow, rowsToDelete);
                Logger.log(`Cleaned up ${rowsToDelete} old audit log entries`);
            }
        }
    } catch (error) {
        Logger.log('Error cleaning up audit logs: ' + error.message);
    }
}

/**
 * Get audit logs with filtering and pagination
 * @param {Object} filters - Filter options {action, entityType, userEmail, dateFrom, dateTo, searchQuery}
 * @param {Object} pagination - Pagination options {page, pageSize}
 * @returns {Object} {success, entries, total, page, pageSize, totalPages}
 */
function getAuditLogs(filters, pagination) {
    if (!isCurrentUserAdmin()) {
        return { success: false, message: 'Unauthorized: Admin access required' };
    }

    try {
        const sheet = getAuditSheet();
        const data = sheet.getDataRange().getValues();

        if (data.length <= 1) {
            return {
                success: true,
                entries: [],
                total: 0,
                page: 1,
                pageSize: pagination?.pageSize || 25,
                totalPages: 0,
                retentionDays: getAuditRetentionDays()
            };
        }

        // Parse all entries (skip header)
        let entries = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            let changes = {};
            try {
                changes = row[6] ? JSON.parse(row[6]) : {};
            } catch (e) {
                changes = {};
            }

            entries.push({
                id: i,
                timestamp: row[0],
                userEmail: row[1],
                userName: row[2],
                action: row[3],
                entityType: row[4],
                entityId: row[5],
                changes: changes,
                summary: row[7]
            });
        }

        // Apply filters
        if (filters) {
            if (filters.action) {
                entries = entries.filter(e => e.action === filters.action);
            }
            if (filters.entityType) {
                entries = entries.filter(e => e.entityType === filters.entityType);
            }
            if (filters.userEmail) {
                entries = entries.filter(e => e.userEmail === filters.userEmail);
            }
            if (filters.dateFrom) {
                const fromDate = new Date(filters.dateFrom);
                entries = entries.filter(e => new Date(e.timestamp) >= fromDate);
            }
            if (filters.dateTo) {
                const toDate = new Date(filters.dateTo);
                toDate.setHours(23, 59, 59, 999); // Include the entire day
                entries = entries.filter(e => new Date(e.timestamp) <= toDate);
            }
            if (filters.searchQuery) {
                const query = filters.searchQuery.toLowerCase();
                entries = entries.filter(e =>
                    (e.entityId || '').toString().toLowerCase().includes(query) ||
                    (e.summary || '').toString().toLowerCase().includes(query) ||
                    (e.userName || '').toString().toLowerCase().includes(query) ||
                    (e.userEmail || '').toString().toLowerCase().includes(query)
                );
            }
        }

        const total = entries.length;
        const pageSize = pagination?.pageSize || 25;
        const totalPages = Math.ceil(total / pageSize) || 1;
        const page = Math.min(Math.max(pagination?.page || 1, 1), totalPages);

        // Paginate
        const startIndex = (page - 1) * pageSize;
        const paginatedEntries = entries.slice(startIndex, startIndex + pageSize);

        // Get unique users for filter dropdown
        const uniqueUsers = [...new Set(data.slice(1).map(row => row[1]))].filter(Boolean);

        return {
            success: true,
            entries: paginatedEntries,
            total: total,
            page: page,
            pageSize: pageSize,
            totalPages: totalPages,
            uniqueUsers: uniqueUsers,
            retentionDays: getAuditRetentionDays()
        };
    } catch (error) {
        Logger.log('Error getting audit logs: ' + error.message);
        return { success: false, message: 'Error loading audit logs: ' + error.message };
    }
}

/**
 * Get audit log settings
 */
function getAuditSettings() {
    if (!isCurrentUserAdmin()) {
        return { success: false, message: 'Unauthorized: Admin access required' };
    }

    return {
        success: true,
        retentionDays: getAuditRetentionDays()
    };
}

/**
 * Manually trigger cleanup of old audit logs (admin only)
 */
function runAuditCleanup() {
    if (!isCurrentUserAdmin()) {
        return { success: false, message: 'Unauthorized: Admin access required' };
    }

    try {
        const sheet = getAuditSheet();
        const beforeCount = sheet.getLastRow() - 1; // Exclude header

        cleanupOldAuditLogs();

        const afterCount = sheet.getLastRow() - 1;
        const deletedCount = beforeCount - afterCount;

        return {
            success: true,
            message: `Cleanup complete. Removed ${deletedCount} entries older than ${getAuditRetentionDays()} days.`
        };
    } catch (error) {
        return { success: false, message: 'Error during cleanup: ' + error.message };
    }
}
