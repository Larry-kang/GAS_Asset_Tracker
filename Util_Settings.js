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
    get: function (key, defaultValue = null) {
        // Always read Script Properties fresh.
        // Apps Script V8 executes files in global scope, and warm runtimes can
        // preserve global objects between executions. Caching settings in a
        // global object can therefore leak stale values such as old tunnel URLs.
        const value = PropertiesService.getScriptProperties().getProperty(key);
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

