/**
 * Util_Cache.js
 * Multi-level Hybrid Caching System (L1: Memory, L2: CacheService)
 * 
 * DESIGN:
 * - L1: Simple Object in memory (Resets every execution)
 * - L2: Google Apps Script CacheService (Persists up to 6 hours)
 */

const ScriptCache = {
    _l1: {}, // Level 1: In-Memory

    /**
     * Get a value from cache
     * @param {string} key
     * @param {boolean} bypass (Optional) Force fetch from origin
     */
    get: function (key, bypass = false) {
        if (bypass) return null;

        // 1. Try L1 (Memory)
        if (this._l1[key] !== undefined) {
            return this._l1[key];
        }

        // 2. Try L2 (CacheService)
        const cached = CacheService.getScriptCache().get(key);
        if (cached) {
            try {
                const data = JSON.parse(cached);
                this._l1[key] = data; // Promotion to L1
                return data;
            } catch (e) {
                return cached;
            }
        }

        return null;
    },

    /**
     * Put a value into cache
     * @param {string} key
     * @param {any} value
     * @param {number} ttlSeconds (Optional) Default 900 (15 mins)
     */
    put: function (key, value, ttlSeconds = 900) {
        if (value === null || value === undefined) return;

        // Store in L1
        this._l1[key] = value;

        // Store in L2
        const dataToStore = (typeof value === 'object') ? JSON.stringify(value) : String(value);

        // GAS Limit check: CacheService values must be < 100KB
        if (dataToStore.length < 100 * 1024) {
            CacheService.getScriptCache().put(key, dataToStore, ttlSeconds);
        }
    },

    /**
     * Remove a value from cache
     * @param {string} key
     */
    remove: function (key) {
        delete this._l1[key];
        CacheService.getScriptCache().remove(key);
    }
};
