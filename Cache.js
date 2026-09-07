/**
 * Web App Template - Cache Management
 * Implements version tracking with CacheService for smart polling
 */

// Cache Keys
const CACHE_KEYS = {
    DATA_VERSION: 'DB_VERSION'
};

// Cache expiration (6 hours = 21600 seconds, max allowed)
const CACHE_EXPIRATION = 21600;

/**
 * Gets the current data version from cache
 * Used for smart polling - the clients will check this to see if data has changed
 * @returns {string} The current version timestamp or generates new one if not exists 
 */
function getDataVersion() {
    const cache = CacheService.getScriptCache();
    let version = cache.get(CACHE_KEYS.DATA_VERSION);

    // If no version exists, create one
    if (!version) {
        version = Date.now().toString();
        cache.put(CACHE_KEYS.DATA_VERSION, version, CACHE_EXPIRATION);
    }

    return version;
}

/**
 * Updates the data version in cache
 * Call this whenever data is modified (user CRUD, config changes)
 * @returns {string} The new version timestamp
 */
function updateDataVersion() {
    const cache = CacheService.getScriptCache();
    const newVersion = Date.now().toString();
    cache.put(CACHE_KEYS.DATA_VERSION, newVersion, CACHE_EXPIRATION);
    return newVersion;
}

/**
 * Checks if client version matches server version
 * @param {string} clientVersion - The client's current version
 * @returns {Object} Object with hasChanges boolean and current version
 */
function checkVersion(clientVersion) {
    const serverVersion = getDataVersion();

    return {
        hasChanges: clientVersion !== serverVersion,
        currentVersion: serverVersion
    };
}

/**
 * Clears all cache data (use for debugging/reset)
 */
function clearCache() {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.DATA_VERSION);
}
