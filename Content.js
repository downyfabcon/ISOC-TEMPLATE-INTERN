const ANNOUNCEMENTS_SHEET_NAME = 'ANNOUNCEMENTS';
const KNOWLEDGE_SHEET_NAME = 'KNOWLEDGE';

function getContentSheet_(name, headers) {
    const databaseId = getMainDbId();
    if (!databaseId) throw new Error('Main database not configured');

    const spreadsheet = SpreadsheetApp.openById(databaseId);
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(name);
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
        sheet.setFrozenRows(1);
    }
    return sheet;
}

function getAnnouncements() {
    if (!isCurrentUserAuthorized()) return { success: false, error: 'Access denied.' };

    try {
        const sheet = getContentSheet_(ANNOUNCEMENTS_SHEET_NAME, ['Title', 'Message', 'Audience', 'Type', 'Send At', 'Created At', 'Created By', 'Email Sent']);
        const values = sheet.getDataRange().getValues();
        const email = getCurrentUserEmail().toLowerCase();
        const now = new Date();
        const items = values.slice(1).map((row, index) => ({
            rowIndex: index + 2,
            title: String(row[0] || ''),
            message: String(row[1] || ''),
            audience: String(row[2] || 'ALL'),
            type: String(row[3] || 'announcement'),
            sendAt: row[4] ? new Date(row[4]).toISOString() : '',
            createdAt: row[5] ? new Date(row[5]).toISOString() : '',
            createdBy: String(row[6] || '')
        })).filter(item => item.title && (item.audience === 'ALL' || item.audience.toLowerCase() === email))
          .filter(item => item.type !== 'reminder' || !item.sendAt || new Date(item.sendAt) <= now)
          .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
        return { success: true, items: items };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

function createAnnouncement(data) {
    if (!isCurrentUserAdmin()) return { success: false, error: 'Admin privileges required.' };

    const title = String(data && data.title || '').trim();
    const message = String(data && data.message || '').trim();
    const audience = String(data && data.audience || 'ALL').trim();
    const type = data && data.type === 'reminder' ? 'reminder' : 'announcement';
    const sendAt = data && data.sendAt ? new Date(data.sendAt) : null;

    if (!title || !message) return { success: false, error: 'Title and message are required.' };
    if (audience !== 'ALL' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(audience)) return { success: false, error: 'Enter a valid recipient email.' };
    if (sendAt && isNaN(sendAt.getTime())) return { success: false, error: 'Enter a valid schedule date.' };

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(LOCK_TIMEOUT_MS);
        const sheet = getContentSheet_(ANNOUNCEMENTS_SHEET_NAME, ['Title', 'Message', 'Audience', 'Type', 'Send At', 'Created At', 'Created By', 'Email Sent']);
        const now = new Date();
        sheet.appendRow([title, message, audience, type, sendAt || '', now, getCurrentUserEmail(), '']);
        if (type === 'reminder' && sendAt && sendAt > now) ensureReminderTrigger_();
        if ((!sendAt || sendAt <= now) && data.sendEmail) sendAnnouncementEmail_(title, message, audience, type);
        logAuditEvent('CREATE', type === 'reminder' ? 'REMINDER' : 'ANNOUNCEMENT', title, { title: title, audience: audience }, 'Created ' + type + ': ' + title);
        updateDataVersion();
        return { success: true, message: type === 'reminder' && sendAt > now ? 'Reminder scheduled.' : 'Announcement published.' };
    } catch (error) {
        return { success: false, error: error.message };
    } finally {
        lock.releaseLock();
    }
}

function deleteAnnouncement(rowIndex) {
    if (!isCurrentUserAdmin()) return { success: false, error: 'Admin privileges required.' };
    try {
        const sheet = getContentSheet_(ANNOUNCEMENTS_SHEET_NAME, ['Title', 'Message', 'Audience', 'Type', 'Send At', 'Created At', 'Created By', 'Email Sent']);
        const row = Number(rowIndex);
        if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) return { success: false, error: 'Item not found.' };
        const title = String(sheet.getRange(row, 1).getValue() || '');
        sheet.deleteRow(row);
        logAuditEvent('DELETE', 'ANNOUNCEMENT', title, null, 'Deleted announcement: ' + title);
        updateDataVersion();
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

function processDueReminders() {
    const sheet = getContentSheet_(ANNOUNCEMENTS_SHEET_NAME, ['Title', 'Message', 'Audience', 'Type', 'Send At', 'Created At', 'Created By', 'Email Sent']);
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    values.slice(1).forEach((row, index) => {
        const scheduledAt = row[4] ? new Date(row[4]) : null;
        if (String(row[3]) === 'reminder' && scheduledAt && scheduledAt <= now && !row[7]) {
            sendAnnouncementEmail_(String(row[0]), String(row[1]), String(row[2]), 'reminder');
            sheet.getRange(index + 2, 8).setValue(new Date());
        }
    });
}

function ensureReminderTrigger_() {
    const handler = 'processDueReminders';
    const exists = ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === handler);
    if (!exists) ScriptApp.newTrigger(handler).timeBased().everyMinutes(15).create();
}

function sendAnnouncementEmail_(title, message, audience, type) {
    const recipients = audience === 'ALL'
        ? getAllUserEmails_()
        : [audience];
    if (recipients.length) MailApp.sendEmail(recipients.join(','), (type === 'reminder' ? 'Reminder: ' : 'Announcement: ') + title, message);
}

function getAllUserEmails_() {
    const sheet = SpreadsheetApp.openById(getMainDbId()).getSheetByName('USERS');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
        .map(row => String(row[0] || '').trim()).filter(Boolean);
}

function getKnowledgeItems() {
    if (!isCurrentUserAuthorized()) return { success: false, error: 'Access denied.' };
    try {
        const sheet = getContentSheet_(KNOWLEDGE_SHEET_NAME, ['Title', 'Category', 'Description', 'URL', 'Created At', 'Created By']);
        const items = sheet.getDataRange().getValues().slice(1).map((row, index) => ({
            rowIndex: index + 2, title: String(row[0] || ''), category: String(row[1] || ''),
            description: String(row[2] || ''), url: String(row[3] || ''), createdAt: row[4] ? new Date(row[4]).toISOString() : ''
        })).filter(item => item.title).sort((a, b) => a.title.localeCompare(b.title));
        return { success: true, items: items };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

function createKnowledgeItem(data) {
    if (!isCurrentUserAdmin()) return { success: false, error: 'Admin privileges required.' };
    const title = String(data && data.title || '').trim();
    let url = String(data && data.url || '').trim();
    const upload = data && data.upload;
    if (!title) return { success: false, error: 'Title is required.' };
    if (upload && upload.base64 && upload.name) {
        if (upload.base64.length > 35 * 1024 * 1024) return { success: false, error: 'Please upload a file smaller than 25 MB.' };
        const blob = Utilities.newBlob(Utilities.base64Decode(upload.base64), String(upload.mimeType || 'application/octet-stream'), String(upload.name));
        const file = DriveApp.createFile(blob);
        getAllUserEmails_().forEach(email => file.addViewer(email));
        url = file.getUrl();
    } else if (!/^https?:\/\//i.test(url)) {
        return { success: false, error: 'Provide a document URL or upload a file.' };
    }
    try {
        const sheet = getContentSheet_(KNOWLEDGE_SHEET_NAME, ['Title', 'Category', 'Description', 'URL', 'Created At', 'Created By']);
        sheet.appendRow([title, String(data.category || '').trim(), String(data.description || '').trim(), url, new Date(), getCurrentUserEmail()]);
        logAuditEvent('CREATE', 'KNOWLEDGE', title, { title: title }, 'Added knowledge resource: ' + title);
        updateDataVersion();
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

function deleteKnowledgeItem(rowIndex) {
    if (!isCurrentUserAdmin()) return { success: false, error: 'Admin privileges required.' };
    try {
        const sheet = getContentSheet_(KNOWLEDGE_SHEET_NAME, ['Title', 'Category', 'Description', 'URL', 'Created At', 'Created By']);
        const row = Number(rowIndex);
        if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) return { success: false, error: 'Resource not found.' };
        const title = String(sheet.getRange(row, 1).getValue() || '');
        sheet.deleteRow(row);
        logAuditEvent('DELETE', 'KNOWLEDGE', title, null, 'Deleted knowledge resource: ' + title);
        updateDataVersion();
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
}