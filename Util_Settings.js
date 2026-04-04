/**
 * Util_Settings.js
 * Centralized Settings & Properties Manager for SAP.
 * Prevents direct PropertiesService calls in business logic.
 */

const Settings = {
    /**
     * Get a script property by key
     * @param {string} key 
     * @param {*} defaultValue 
     * @returns {*}
     */
    _cache: null,

    /**
     * Get a script property by key
     * @param {string} key 
     * @param {*} defaultValue 
     * @returns {*}
     */
    get: function (key, defaultValue = null) {
        // Build cache on first call within this execution.
        // Do not use cross-execution ScriptCache here, otherwise manual
        // Script Properties edits in GAS UI may stay stale for up to 30 min.
        if (this._cache === null) {
            this._cache = PropertiesService.getScriptProperties().getProperties();
        }

        const value = this._cache[key];
        if (value === null || value === undefined) return defaultValue;
        return value;
    },

    /**
     * Get a numeric script property
     * @param {string} key 
     * @param {number} defaultValue 
     * @returns {number}
     */
    getNumber: function (key, defaultValue = 0) {
        const value = this.get(key);
        return value !== null ? parseFloat(value) : defaultValue;
    },

    /**
     * Set a script property
     * @param {string} key 
     * @param {string} value 
     */
    set: function (key, value) {
        if (value === null || value === undefined) {
            PropertiesService.getScriptProperties().deleteProperty(key);
        } else {
            PropertiesService.getScriptProperties().setProperty(key, value.toString());
        }

        // Invalidate execution-level cache
        this._cache = null;
    },

    /**
     * Bulk get essential system settings
     */
    getSystemSet: function () {
        return {
            adminEmail: this.get('ADMIN_EMAIL'),
            treasuryReserve: this.getNumber('TREASURY_RESERVE_TWD', 100000),
            monthlyDebt: this.getNumber('MONTHLY_DEBT_COST', 12967)
        };
    }
};

