/**
 * Web App Template - Quick Links Operations
 * Manages Quick Links CRUD operations using the Main Database spreadsheet
 * Sheet name: "QUICK LINKS"
 * Columns: ID, Name, Link, Category, Icon
 */

// Constants
const LOCK_TIMEOUT_MS = 30000; // 30 seconds timeout for script lock

/**
 * Finds the first matching header from a list of candidates
 * @param {Array} headers - Array of header strings to search
 * @param {Array} candidates - Array of candidate header names to match
 * @returns {number} Index of first match, or -1 if none found
 */
function findFirstHeaderIndex(headers, candidates) {
    for (const candidate of candidates) {
        const idx = headers.indexOf(candidate);
        if (idx !== -1) return idx;
    }
    return -1;
}

function ensureQuickLinkIdColumn(sheet) {
    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let headerMap = headers.map(h => String(h).trim().toUpperCase());
    let idIndex = headerMap.indexOf('ID');

    if (idIndex === -1) {
        idIndex = headers.length;
        sheet.getRange(1, idIndex + 1).setValue('ID');
        headers.push('ID');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const idRange = sheet.getRange(2, idIndex + 1, lastRow - 1, 1);
        const ids = idRange.getValues();
        let changed = false;
        ids.forEach(row => {
            if (!String(row[0] || '').trim()) {
                row[0] = Utilities.getUuid();
                changed = true;
            }
        });
        if (changed) idRange.setValues(ids);
    }

    return { headers, headerMap: headers.map(h => String(h).trim().toUpperCase()), idIndex };
}

function getFavoriteIds_() {
    const raw = PropertiesService.getUserProperties().getProperty('QUICK_LINK_FAVORITES');
    if (!raw) return [];
    try {
        const ids = JSON.parse(raw);
        return Array.isArray(ids) ? [...new Set(ids.map(id => String(id).trim()).filter(Boolean))] : [];
    } catch (error) {
        return [];
    }
}

function getQuickLinkFavorites() {
    return { success: true, favoriteIds: getFavoriteIds_() };
}

function toggleQuickLinkFavorite(linkId) {
    const id = String(linkId || '').trim();
    if (!id) return { success: false, error: 'Link ID is required' };

    const spreadsheetId = getMainDbId();
    if (!spreadsheetId) return { success: false, error: 'Main database not configured' };

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('QUICK LINKS');
    if (!sheet) return { success: false, error: 'Sheet "QUICK LINKS" not found' };

    const schema = ensureQuickLinkIdColumn(sheet);
    const lastRow = sheet.getLastRow();
    const ids = lastRow > 1
        ? sheet.getRange(2, schema.idIndex + 1, lastRow - 1, 1).getValues().flat().map(value => String(value || '').trim())
        : [];
    if (!ids.includes(id)) return { success: false, error: 'Quick Link not found' };

    const favoriteIds = getFavoriteIds_();
    const isFavorite = favoriteIds.includes(id);
    const nextFavoriteIds = isFavorite ? favoriteIds.filter(value => value !== id) : [...favoriteIds, id];
    PropertiesService.getUserProperties().setProperty('QUICK_LINK_FAVORITES', JSON.stringify(nextFavoriteIds));
    return { success: true, isFavorite: !isFavorite, favoriteIds: nextFavoriteIds };
}

/**
 * Get all Quick Links from the database
 * @returns {Object} { success: boolean, items: Array, version: string, error: string }
 */
function getQuickLinks() {
    let spreadsheetId;

    try {
        spreadsheetId = getMainDbId();

        if (!spreadsheetId) {
            return {
                success: false,
                configured: false,
                items: [],
                error: 'Main database not configured'
            };
        }

        const ss = SpreadsheetApp.openById(spreadsheetId);
        const sheet = ss.getSheetByName('QUICK LINKS');

        if (!sheet) {
            return {
                success: true,
                configured: true,
                items: [],
                version: Utilities.getUuid(),
                message: 'Sheet "QUICK LINKS" not found. Please create a "QUICK LINKS" sheet in the Main Database spreadsheet.'
            };
        }

        ensureQuickLinkIdColumn(sheet);
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();

        if (values.length === 0) {
            return {
                success: true,
                configured: true,
                items: [],
                version: Utilities.getUuid()
            };
        }

        // Get headers from first row
        const headers = values[0].map(h => String(h).trim().toUpperCase());

        // Validate required columns
        const missingColumns = [];
        if (!headers.includes('NAME')) missingColumns.push('NAME');
        if (!headers.includes('CATEGORY')) missingColumns.push('CATEGORY');

        const linkIndexCandidate = findFirstHeaderIndex(headers, ['LINK', 'LINKS', 'URL', 'LINK URL', 'LINK_URL']);
        if (linkIndexCandidate === -1) missingColumns.push('LINK (or LINKS)');

        if (missingColumns.length > 0) {
            return {
                success: false,
                configured: true,
                items: [],
                error: 'Missing required columns: ' + missingColumns.join(', ')
            };
        }

        // Get column indices
        const nameIndex = headers.indexOf('NAME');
        const linkIndex = linkIndexCandidate;
        const categoryIndex = headers.indexOf('CATEGORY');
        const iconIndex = headers.indexOf('ICON');
        const idIndex = headers.indexOf('ID');

        // Parse data rows
        const items = [];
        for (let i = 1; i < values.length; i++) {
            const row = values[i];

            // Skip empty rows
            if (!row[nameIndex] || String(row[nameIndex]).trim() === '') {
                continue;
            }

            items.push({
                id: String(row[idIndex] || '').trim(),
                rowIndex: i + 1, // 1-based row number for updates/deletes
                name: String(row[nameIndex] || '').trim(),
                link: String(row[linkIndex] || '').trim(),
                category: String(row[categoryIndex] || '').trim(),
                icon: iconIndex !== -1 ? String(row[iconIndex] || '').trim() : '',
                iconKey: iconIndex !== -1 ? String(row[iconIndex] || '').trim() : ''
            });
        }

        return {
            success: true,
            configured: true,
            items: items,
            version: Utilities.getUuid()
        };

    } catch (error) {
        Logger.log('Error in getQuickLinks: ' + error.message);
        return {
            success: false,
            configured: !!spreadsheetId,
            items: [],
            error: 'Failed to load Quick Links: ' + error.message
        };
    }
}

/**
 * Create a new Quick Link
 * @param {Object} data - { name, link, category, iconKey }
 * @returns {Object} { success: boolean, error: string }
 */
function createQuickLink(data) {
    if (!isCurrentUserAdmin()) {
        return { success: false, error: 'Access denied. Admin privileges required.' };
    }

    const lock = LockService.getScriptLock();

    try {
        lock.waitLock(LOCK_TIMEOUT_MS);

        const spreadsheetId = getMainDbId();
        if (!spreadsheetId) {
            return { success: false, error: 'Main database not configured' };
        }

        const ss = SpreadsheetApp.openById(spreadsheetId);
        const sheet = ss.getSheetByName('QUICK LINKS');

        if (!sheet) {
            return { success: false, error: 'Sheet "QUICK LINKS" not found' };
        }

        // Validate required fields
        if (!data.name || String(data.name).trim() === '') {
            return { success: false, error: 'Name is required' };
        }
        if (!data.link || String(data.link).trim() === '') {
            return { success: false, error: 'Link is required' };
        }

        const schema = ensureQuickLinkIdColumn(sheet);
        const headers = schema.headers;
        const headerMap = headers.map(h => String(h).trim().toUpperCase());

        const nameIndex = headerMap.indexOf('NAME');
        const linkIndex = findFirstHeaderIndex(headerMap, ['LINK', 'LINKS', 'URL', 'LINK URL', 'LINK_URL']);
        const categoryIndex = headerMap.indexOf('CATEGORY');
        const iconIndex = headerMap.indexOf('ICON');
        const idIndex = headerMap.indexOf('ID');

        if (nameIndex === -1 || linkIndex === -1 || categoryIndex === -1) {
            return { success: false, error: 'Required columns not found in sheet' };
        }

        // Prepare new row
        const newRow = new Array(headers.length).fill('');
        newRow[idIndex] = Utilities.getUuid();
        newRow[nameIndex] = String(data.name).trim();
        newRow[linkIndex] = String(data.link).trim();
        newRow[categoryIndex] = String(data.category || '').trim();
        if (iconIndex !== -1 && data.iconKey && String(data.iconKey).trim() !== '') {
            newRow[iconIndex] = String(data.iconKey).trim();
        }

        // Append row
        sheet.appendRow(newRow);

        // Log to audit
        try {
            const fieldsToTrack = ['name', 'link', 'category'];
            if (iconIndex !== -1) fieldsToTrack.push('iconKey');

            const oldValues = {
                name: '',
                link: '',
                category: '',
                iconKey: ''
            };

            const newValues = {
                name: String(data.name || '').trim(),
                link: String(data.link || '').trim(),
                category: String(data.category || '').trim(),
                iconKey: String(data.iconKey || '').trim()
            };

            const changes = calculateChanges(oldValues, newValues, fieldsToTrack);
            const summary = generateChangeSummary('CREATE', 'QUICK_LINK', changes, { identifier: newValues.name });
            logAuditEvent('CREATE', 'QUICK_LINK', newValues.name, changes, summary);
        } catch (auditError) {
            Logger.log('Failed to log audit for createQuickLink: ' + auditError.message);
        }

        updateDataVersion();

        return { success: true };

    } catch (error) {
        Logger.log('Error in createQuickLink: ' + error.message);
        return { success: false, error: 'Failed to create Quick Link: ' + error.message };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Update an existing Quick Link
 * @param {Object} data - { rowIndex, name, link, category, iconKey }
 * @returns {Object} { success: boolean, error: string }
 */
function updateQuickLink(data) {
    if (!isCurrentUserAdmin()) {
        return { success: false, error: 'Access denied. Admin privileges required.' };
    }

    const lock = LockService.getScriptLock();

    try {
        lock.waitLock(LOCK_TIMEOUT_MS);

        const spreadsheetId = getMainDbId();
        if (!spreadsheetId) {
            return { success: false, error: 'Main database not configured' };
        }

        const ss = SpreadsheetApp.openById(spreadsheetId);
        const sheet = ss.getSheetByName('QUICK LINKS');

        if (!sheet) {
            return { success: false, error: 'Sheet "QUICK LINKS" not found' };
        }

        // Validate required fields
        if (!data.rowIndex || data.rowIndex < 2) {
            return { success: false, error: 'Invalid row index' };
        }
        if (!data.name || String(data.name).trim() === '') {
            return { success: false, error: 'Name is required' };
        }
        if (!data.link || String(data.link).trim() === '') {
            return { success: false, error: 'Link is required' };
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const headerMap = headers.map(h => String(h).trim().toUpperCase());

        const nameIndex = headerMap.indexOf('NAME');
        const linkIndex = findFirstHeaderIndex(headerMap, ['LINK', 'LINKS', 'URL', 'LINK URL', 'LINK_URL']);
        const categoryIndex = headerMap.indexOf('CATEGORY');
        const iconIndex = headerMap.indexOf('ICON');

        if (nameIndex === -1 || linkIndex === -1 || categoryIndex === -1) {
            return { success: false, error: 'Required columns not found in sheet' };
        }

        // Update the row (capture old values for audit before writing)
        const rowRange = sheet.getRange(data.rowIndex, 1, 1, headers.length);
        const existingRow = rowRange.getValues()[0];

        const fieldsToTrack = ['name', 'link', 'category'];
        if (iconIndex !== -1) fieldsToTrack.push('iconKey');

        const oldValues = {
            name: String(existingRow[nameIndex] || '').trim(),
            link: String(existingRow[linkIndex] || '').trim(),
            category: String(existingRow[categoryIndex] || '').trim(),
            iconKey: iconIndex !== -1 ? String(existingRow[iconIndex] || '').trim() : ''
        };

        const newValues = {
            name: String(data.name || '').trim(),
            link: String(data.link || '').trim(),
            category: String(data.category || '').trim(),
            iconKey: oldValues.iconKey
        };

        // Only update icon if client provided iconKey (preserves older clients)
        if (iconIndex !== -1 && Object.prototype.hasOwnProperty.call(data, 'iconKey')) {
            newValues.iconKey = data.iconKey ? String(data.iconKey).trim() : '';
        }

        const rowData = existingRow.slice();
        rowData[nameIndex] = newValues.name;
        rowData[linkIndex] = newValues.link;
        rowData[categoryIndex] = newValues.category;
        if (iconIndex !== -1 && Object.prototype.hasOwnProperty.call(data, 'iconKey')) {
            rowData[iconIndex] = newValues.iconKey;
        }

        rowRange.setValues([rowData]);

        // Log to audit
        try {
            const changes = calculateChanges(oldValues, newValues, fieldsToTrack);
            const summary = generateChangeSummary('UPDATE', 'QUICK_LINK', changes, { identifier: newValues.name });
            logAuditEvent('UPDATE', 'QUICK_LINK', newValues.name, changes, summary);
        } catch (auditError) {
            Logger.log('Failed to log audit for updateQuickLink: ' + auditError.message);
        }

        updateDataVersion();

        return { success: true };

    } catch (error) {
        Logger.log('Error in updateQuickLink: ' + error.message);
        return { success: false, error: 'Failed to update Quick Link: ' + error.message };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Delete a Quick Link
 * @param {number} rowIndex - 1-based row number
 * @returns {Object} { success: boolean, error: string }
 */
function deleteQuickLink(rowIndex) {
    if (!isCurrentUserAdmin()) {
        return { success: false, error: 'Access denied. Admin privileges required.' };
    }

    const lock = LockService.getScriptLock();

    try {
        lock.waitLock(LOCK_TIMEOUT_MS);

        const spreadsheetId = getMainDbId();
        if (!spreadsheetId) {
            return { success: false, error: 'Main database not configured' };
        }

        const ss = SpreadsheetApp.openById(spreadsheetId);
        const sheet = ss.getSheetByName('QUICK LINKS');

        if (!sheet) {
            return { success: false, error: 'Sheet "QUICK LINKS" not found' };
        }

        if (!rowIndex || rowIndex < 2) {
            return { success: false, error: 'Invalid row index' };
        }

        // Get the row data for audit log before deleting
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const headerMap = headers.map(h => String(h).trim().toUpperCase());
        const nameIndex = headerMap.indexOf('NAME');
        const linkIndex = findFirstHeaderIndex(headerMap, ['LINK', 'LINKS', 'URL', 'LINK URL', 'LINK_URL']);
        const categoryIndex = headerMap.indexOf('CATEGORY');
        const iconIndex = headerMap.indexOf('ICON');

        const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        const itemName = nameIndex !== -1 ? String(rowData[nameIndex] || '').trim() : 'Unknown';

        // Delete the row
        sheet.deleteRow(rowIndex);

        // Log to audit
        try {
            const fieldsToTrack = ['name', 'link', 'category'];
            if (iconIndex !== -1) fieldsToTrack.push('iconKey');

            const oldValues = {
                name: nameIndex !== -1 ? String(rowData[nameIndex] || '').trim() : itemName,
                link: linkIndex !== -1 ? String(rowData[linkIndex] || '').trim() : '',
                category: categoryIndex !== -1 ? String(rowData[categoryIndex] || '').trim() : '',
                iconKey: iconIndex !== -1 ? String(rowData[iconIndex] || '').trim() : ''
            };

            const newValues = {
                name: '',
                link: '',
                category: '',
                iconKey: ''
            };

            const changes = calculateChanges(oldValues, newValues, fieldsToTrack);
            const summary = generateChangeSummary('DELETE', 'QUICK_LINK', changes, { identifier: itemName });
            logAuditEvent('DELETE', 'QUICK_LINK', itemName, changes, summary);
        } catch (auditError) {
            Logger.log('Failed to log audit for deleteQuickLink: ' + auditError.message);
        }

        updateDataVersion();

        return { success: true };

    } catch (error) {
        Logger.log('Error in deleteQuickLink: ' + error.message);
        return { success: false, error: 'Failed to delete Quick Link: ' + error.message };
    } finally {
        lock.releaseLock();
    }
}
