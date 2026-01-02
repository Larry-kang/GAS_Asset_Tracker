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
        const props = PropertiesService.getScriptProperties();
        const prefix = exchange.toUpperCase();

        // Base keys
        const creds = {
            apiKey: props.getProperty(`${prefix}_API_KEY`),
            apiSecret: props.getProperty(`${prefix}_API_SECRET`)
        };

        // Exchange specific extras
        if (exchange === 'OKX') {
            creds.apiPassphrase = props.getProperty('OKX_API_PASSPHRASE');
        }

        if (exchange === 'BINANCE') {
            creds.tunnelUrl = props.getProperty('TUNNEL_URL');
            creds.proxyPassword = props.getProperty('PROXY_PASSWORD');
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
