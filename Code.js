/**
 * Web App Template - Main Entry Point
 * Google Apps Script Web App Template
 */

/**
 * Serves the main HTML page when the web app is accessed
 * @param {Object} e - Event object containing request parameters
 * @returns {HtmlOutput} The rendered HTML page
 */
function doGet(e) {
    const template = HtmlService.createTemplateFromFile('Index');

    let cfg = {};
    try {
        cfg = (typeof getAppConfig === 'function' ? getAppConfig() : {}) || {};
    } catch (err) {
        cfg = {};
    }

    // Expose for server-side templating (Index.html/AuthOverlay/Home/etc.)
    template.appConfig = cfg;

    const pageTitle = cfg.pageTitle || cfg.name || 'Web App';

    const out = template.evaluate()
        .setTitle(pageTitle)
        // Use Apps Script's native favicon support (preferred over DOM <link> manipulation)
        .setFaviconUrl((cfg && cfg.faviconUrl) ? String(cfg.faviconUrl) : '')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

    return out;
}

/**
 * Includes an HTML file's content into another HTML file
 * Used with scriptlets: <?!= include('filename'); ?>
 * @param {string} filename - The name of the HTML file to include (without .html extension)
 * @returns {string} The content of the HTML file
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
