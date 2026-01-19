/**
 * Config.js
 * Centralized configuration constants for the Sovereign Asset Protocol.
 */
const Config = {
    SYSTEM_NAME: "SAP - Sovereign Asset Protocol",
    VERSION: "v24.11", // Added Segmented LTV & Refined Cashflow Monitoring

    // Feature Flags
    FEATURES: {
        ENABLE_LOGGING: true,
        ENABLE_NOTIFICATIONS: true
    },

    // Thresholds & Constants
    THRESHOLDS: {
        PRICE_CACHE_MINUTES: 1, // High-frequency update
        LOG_RETENTION_DAYS: 7
    },

    // Sheet Names
    SHEET_NAMES: {
        BALANCE_SHEET: "Balance Sheet",
        INDICATORS: "Key Market Indicators",
        DASHBOARD: "Dashboard"
    },

    // Strategic Thresholds
    STRATEGIC: {
        REBALANCE_ABS: 0.03,
        PLEDGE_RATIO_SAFE: 2.5,
        PLEDGE_RATIO_ALERT: 2.1,
        PLEDGE_RATIO_CRITICAL: 1.8,
        CRYPTO_LOAN_RATIO_SAFE: 2.0,
        CRYPTO_LOAN_RATIO_ALERT: 1.5,
        CRYPTO_LOAN_RATIO_CRITICAL: 1.3
    },

    // Martingale Strategy
    BTC_MARTINGALE: {
        ENABLED: true,
        BASE_AMOUNT: 20000,
        LEVELS: [
            { drop: -0.40, multiplier: 1, name: "Level 1 (Sniper Zone)" },
            { drop: -0.50, multiplier: 2, name: "Level 2 (Deep Value)" },
            { drop: -0.60, multiplier: 3, name: "Level 3 (Abyss)" },
            { drop: -0.70, multiplier: 4, name: "Level 4 (Capitulation)" }
        ]
    },

    // Asset Layers (2026-01-18 Calibration: MM=0.91 Accumulation Zone)
    // Reference: Multi_Target_Analysis_20260117.md
    ASSET_GROUPS: [
        { id: "L1", name: "Layer 1: Digital Reserve (Attack)", defaultTarget: 0.70, tickers: ["IBIT", "BTC_Spot", "BTC"] },
        { id: "L2", name: "Layer 2: Credit Base (Defend)", defaultTarget: 0.20, tickers: ["00713", "00662", "QQQ"] },
        { id: "L3", name: "Layer 3: Tactical Liquidity", defaultTarget: 0.10, tickers: ["BOXX", "CASH_TWD", "CASH_FC"] }
    ],

    NOISE_ASSETS: ["ETH", "BNB", "TQQQ"]
};

