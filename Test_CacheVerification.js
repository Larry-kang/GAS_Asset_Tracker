/**
 * Test_MultiLevelCache.js
 * Verifies that ScriptCache works across executions (L2) and within execution (L1).
 */

function testCacheFlow() {
    const testKey = "CACHE_TEST_" + new Date().getTime();
    const testValue = { status: "OK", timestamp: new Date().toISOString() };

    console.time("First Put");
    ScriptCache.put(testKey, testValue, 60);
    console.timeEnd("First Put");

    console.time("L1 Hit (Memory)");
    const v1 = ScriptCache.get(testKey);
    console.timeEnd("L1 Hit (Memory)");

    if (JSON.stringify(v1) === JSON.stringify(testValue)) {
        console.log("? L1 Cache Match");
    } else {
        console.error("? L1 Cache Mismatch");
    }
}

/**
 * Run this multiple times to test L2 (Persistent)
 */
function testL2Persistency() {
    const key = "PERSISTENT_TEST_KEY";
    const val = ScriptCache.get(key);

    if (val) {
        console.log("? L2 Cache HIT: " + JSON.stringify(val));
        // Verify bypass
        const bypassed = ScriptCache.get(key, true);
        if (bypassed === null) {
            console.log("? Bypass Flag OK");
        }
    } else {
        console.log("?? L2 Cache MISS. Seeding...");
        ScriptCache.put(key, { message: "Hello from previous execution", time: new Date().toISOString() }, 300);
    }
}

function verifyPriceCaching() {
    const ticker = "BTC";
    console.log("Checking Price Cache for " + ticker);

    console.time("Fetch 1 (Likely Origin or L2)");
    const p1 = fetchCryptoPrice(ticker);
    console.timeEnd("Fetch 1 (Likely Origin or L2)");

    console.time("Fetch 2 (Should be L1)");
    const p2 = fetchCryptoPrice(ticker);
    console.timeEnd("Fetch 2 (Should be L1)");

    console.log(`P1: ${p1}, P2: ${p2}`);
}
