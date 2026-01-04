/**
 * Test_Benchmarking.js
 * Measures the execution time of buildContext to establish a performance baseline.
 * Run this function manually in the GAS Editor.
 */
function benchmarkPerformance() {
    const iterations = 5;
    let totalTime = 0;

    console.log("=== Starting Performance Benchmark ===");
    console.log("This will run buildContext() " + iterations + " times to calculate average execution speed.");

    for (let i = 0; i < iterations; i++) {
        const start = new Date().getTime();
        try {
            buildContext();
            const end = new Date().getTime();
            const duration = end - start;
            totalTime += duration;
            console.log("Iteration " + (i + 1) + ": " + duration + "ms");
        } catch (e) {
            console.error("Iteration " + (i + 1) + " failed: " + e.toString());
        }
    }

    const average = totalTime / iterations;
    console.log("-----------------------------------------");
    console.log("Average Execution Time: " + average.toFixed(2) + "ms");
    console.log("=== Benchmark Complete ===");

    return average;
}
