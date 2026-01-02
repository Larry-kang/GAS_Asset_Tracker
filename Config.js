/**
 * Config.js
 * Centralized configuration constants for the Sovereign Asset Protocol.
 */
const Config = {
    SYSTEM_NAME: "SAP - Sovereign Asset Protocol",
    VERSION: "v24.6",

    // Feature Flags
    FEATURES: {
        ENABLE_LOGGING: true,
        ENABLE_NOTIFICATIONS: true
    },

    // Thresholds & Constants
    THRESHOLDS: {
        PRICE_CACHE_MINUTES: 1, // High-frequency update
        LOG_RETENTION_DAYS: 7
    }
};

