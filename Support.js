/**
 * Sends a support request to all administrators listed in the USERS sheet.
 * @param {Object} request - Support form data from the authenticated user.
 * @returns {Object} Result object with success status and message.
 */
function submitSupportRequest(request) {
    const allowedTypes = {
        broken_link: 'Report broken link',
        new_resource: 'Request new resource',
        other: 'Other support'
    };

    try {
        const currentUser = getCurrentUserInfo();
        if (!currentUser) {
            return { success: false, error: 'You must be authenticated to submit a support request.' };
        }

        const type = String(request && request.type || '').trim();
        const subject = String(request && request.subject || '').trim();
        const description = String(request && request.description || '').trim();
        const link = String(request && request.link || '').trim();

        if (!allowedTypes[type]) return { success: false, error: 'Please select a valid request type.' };
        if (!subject) return { success: false, error: 'Subject is required.' };
        if (subject.length > 120) return { success: false, error: 'Subject must be 120 characters or fewer.' };
        if (!description) return { success: false, error: 'Description is required.' };
        if (description.length > 2000) return { success: false, error: 'Description must be 2,000 characters or fewer.' };
        if (link && !/^https?:\/\/\S+$/i.test(link)) return { success: false, error: 'Link URL must start with http:// or https://.' };

        const mainDbId = getMainDbId();
        if (!mainDbId) return { success: false, error: 'Main database is not configured.' };

        const sheet = SpreadsheetApp.openById(mainDbId).getSheetByName('USERS');
        if (!sheet) return { success: false, error: 'Users database is not configured.' };

        const rows = sheet.getDataRange().getValues();
        const adminEmails = [];
        for (let i = 1; i < rows.length; i++) {
            const email = String(rows[i][0] || '').trim();
            const role = String(rows[i][2] || '').trim().toLowerCase();
            if (email && role === 'admin' && adminEmails.indexOf(email) === -1) adminEmails.push(email);
        }

        const body = [
            'A new support request was submitted in ' + (getAppConfig().name || 'the web app') + '.',
            '',
            'Type: ' + allowedTypes[type],
            'Subject: ' + subject,
            'Submitted by: ' + currentUser.name + ' (' + currentUser.email + ')',
            link ? 'Link URL: ' + link : '',
            '',
            'Description:',
            description
        ].filter(Boolean).join('\n');

        const lock = LockService.getScriptLock();
        lock.waitLock(30000);
        try {
            let supportSheet = SpreadsheetApp.openById(mainDbId).getSheetByName('SUPPORT REQUESTS');
            if (!supportSheet) {
                supportSheet = SpreadsheetApp.openById(mainDbId).insertSheet('SUPPORT REQUESTS');
                supportSheet.getRange(1, 1, 1, 8).setValues([[
                    'Timestamp', 'User Email', 'User Name', 'Request Type',
                    'Subject', 'Link URL', 'Description', 'Status'
                ]]);
                supportSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
                supportSheet.setFrozenRows(1);
            }

            supportSheet.appendRow([
                new Date(),
                currentUser.email,
                currentUser.name,
                allowedTypes[type],
                subject,
                link,
                description,
                'Submitted'
            ]);
        } finally {
            lock.releaseLock();
        }

        if (adminEmails.length === 0) {
            return {
                success: true,
                message: 'Your request was saved in the SUPPORT REQUESTS sheet. No administrator email is configured.'
            };
        }

        try {
            MailApp.sendEmail({
                to: adminEmails.join(','),
                subject: '[Support] ' + subject,
                body: body,
                name: getAppConfig().name || 'Web App Support'
            });
            return { success: true, message: 'Your request was saved and sent to the administrators.' };
        } catch (emailError) {
            Logger.log('Support request saved, but email failed: ' + emailError.message);
            return { success: true, message: 'Your request was saved in the SUPPORT REQUESTS sheet, but email delivery failed.' };
        }
    } catch (error) {
        Logger.log('Error in submitSupportRequest: ' + error.message);
        return { success: false, error: 'Unable to send support request: ' + error.message };
    }
}

/**
 * Gets the number of support requests recorded in the support sheet.
 * @returns {Object} Support request totals for dashboard analytics.
 */
function getSupportRequestStats() {
    try {
        const mainDbId = getMainDbId();
        if (!mainDbId) return { total: 0 };

        const sheet = SpreadsheetApp.openById(mainDbId).getSheetByName('SUPPORT REQUESTS');
        if (!sheet || sheet.getLastRow() < 2) return { total: 0 };
        return { total: sheet.getLastRow() - 1 };
    } catch (error) {
        Logger.log('Error loading support request stats: ' + error.message);
        return { total: 0 };
    }
}
