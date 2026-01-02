/**
 * Util_Credentials.js
 * Centralized credential management for exchange APIs.
 * Wrapper around PropertiesService to prevent code duplication.
 */

const Credentials = {
    /**
     * Get API keys for specific exchange
     * @param {string} exchange - e.g. 'BINANCE', 'OKX', 'BITOPRO', 'PIONEX'
     * @returns {Object} Keys object { apiKey, apiSecret, ...others }
     */
    get: function (exchange) {
        const prefix = exchange.toUpperCase();

        const creds = {
            apiKey: Settings.get(`${prefix}_API_KEY`),
            apiSecret: Settings.get(`${prefix}_API_SECRET`)
        };

        // Exchange specific extras
        if (exchange === 'OKX') {
            creds.apiPassphrase = Settings.get('OKX_API_PASSPHRASE');
        }

        if (exchange === 'BINANCE') {
            creds.tunnelUrl = Settings.get('TUNNEL_URL');
            creds.proxyPassword = Settings.get('PROXY_PASSWORD');
        }

        return creds;
    },

    /**
     * Validate if essential keys exist
     * @param {Object} creds 
     * @returns {boolean}
     */
    isValid: function (creds) {
        if (!creds.apiKey || !creds.apiSecret) return false;
        return true;
    }
};

