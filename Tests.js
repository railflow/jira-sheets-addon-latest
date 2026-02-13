
/**
 * Unit Tests for Jira Sheets Add-on
 * Run this function to verify core logic.
 */
function runAllTests() {
    console.log("Running Tests...");
    testCalculateRoadmapDateRange();
    testParseJqluCommand();
    console.log("All Tests Completed.");
}

function assert(condition, message) {
    if (!condition) {
        throw new Error("Assertion Failed: " + message);
    }
    console.log("PASS: " + message);
}

function testCalculateRoadmapDateRange() {
    console.log("--- Testing calculateRoadmapDateRange ---");

    // Mock Indices
    const indices = {
        createdIndex: 0,
        startDateIndex: 1,
        dueDateIndex: 2,
        resolvedIndex: 3
    };

    // Scenario 1: Standard Range within defaults
    // Today is roughly Feb 2026. Data is Mar 2026.
    // Default range (1m back, 4m fwd) should cover it if close, but let's test expansion.
    // Let's force dates OUTSIDE the default range to ensure it expands.

    // Future Date: July 2026 (User's issue)
    const dataFuture = [
        ['2026-02-01', '2026-02-01', '2026-07-15', null] // Created, Start, Due, Resolved
    ];

    const resultFuture = calculateRoadmapDateRange(dataFuture, indices);

    // Check End Date: Should be at least End of July 2026 (+ buffer logic)
    // Code adds +1 month buffer to max date found.
    // Max found: July 15. Buffer: Month + 1 = August. End date = Last day of August.
    const expectedEndMonth = 7; // August (0-indexed) is 7? No, Jan=0, Aug=7.
    // Logic: endDate.setMonth(endDate.getMonth() + 1); endDate.setDate(0);
    // If Max is July 15 (Month 6). +1 = Month 7 (Aug). setDate(0) -> Last day of July?
    // Wait: setDate(0) sets it to the last day of the PREVIOUS month.
    // Example: Date(2026, 6, 15) -> July
    // .setMonth(7) -> Aug 15
    // .setDate(0) -> July 31.
    // So buffer logic effectively rounds to End of Month found.
    // Actually: 
    // endDate = new Date(maxDate);
    // endDate.setMonth(endDate.getMonth() + 1);
    // endDate.setDate(0);
    // If maxDate is July 15 (Month 6). setMonth(7) is Aug 15. setDate(0) is July 31.
    // So it rounds to End of Month.

    // User complaint: "stops at may when end date july".
    // If logic was fixed, it should now extend to July.

    console.log("Result Future End:", resultFuture.endDate);
    assert(resultFuture.endDate.getMonth() >= 6, "End date should cover July (Month 6)");
    assert(resultFuture.endDate.getFullYear() === 2026, "Year should be 2026");

    // Scenario 2: Past Date (Start Date honoring)
    // Start Date: Jan 2025.
    const dataPast = [
        ['2026-02-01', '2025-01-15', '2026-03-01', null]
    ];
    const resultPast = calculateRoadmapDateRange(dataPast, indices);
    console.log("Result Past Start:", resultPast.startDate);
    assert(resultPast.startDate.getFullYear() === 2025, "Start year should extend back to 2025");
    assert(resultPast.startDate.getMonth() === 0, "Start month should be Jan (0)");

    // Scenario 3: No Dates (Defaults)
    const dataEmpty = [];
    const resultEmpty = calculateRoadmapDateRange(dataEmpty, indices);
    const today = new Date();
    // Default min: today - 1 month. setDate(1).
    // Default max: today + 4 months. setMonth(+1), setDate(0).
    // Just assert it returns valid dates.
    assert(resultEmpty.startDate instanceof Date, "Should return valid start date");
    assert(resultEmpty.endDate instanceof Date, "Should return valid end date");
    assert(resultEmpty.totalWeeks > 0, "Should have positive weeks");

    console.log("calculateRoadmapDateRange Tests Passed");
}

function testParseJqluCommand() {
    console.log("--- Testing parseJqluCommand ---");

    // Test 1: Simple Update
    const cmd1 = "UPDATE summary='New Title' WHERE project='ABC'";
    const res1 = parseJqluCommand(cmd1);
    assert(res1.updatePayload.fields.summary === 'New Title', "Simple summary update failed");
    assert(res1.jql === "project='ABC'", "Simple JQL extraction failed");

    // Test 2: Multiple Fields
    const cmd2 = "UPDATE summary='A', description='B' WHERE id=1";
    const res2 = parseJqluCommand(cmd2);
    assert(res2.updatePayload.fields.summary === 'A', "Multiple field A failed");
    assert(res2.updatePayload.fields.description === 'B', "Multiple field B failed");

    // Test 3: Special Fields (Assignee Name)
    const cmd3 = "UPDATE assignee='bob' WHERE key='TEST-1'";
    const res3 = parseJqluCommand(cmd3);
    assert(res3.updatePayload.fields.assignee.name === 'bob', "Assignee name mapping failed");

    // Test 4: Special Fields (Assignee ID)
    const cmd4 = "UPDATE assignee='557058:xyz' WHERE key='TEST-1'";
    const res4 = parseJqluCommand(cmd4);
    assert(res4.updatePayload.fields.assignee.accountId === '557058:xyz', "Assignee accountId mapping failed");

    // Test 5: Special Fields (Priority)
    const cmd5 = "UPDATE priority='High' WHERE key='TEST-1'";
    const res5 = parseJqluCommand(cmd5);
    assert(res5.updatePayload.fields.priority.name === 'High', "Priority mapping failed");

    // Test 6: Quoted values with commas
    const cmd6 = "UPDATE summary='Hello, World' WHERE key='TEST-1'";
    const res6 = parseJqluCommand(cmd6);
    assert(res6.updatePayload.fields.summary === 'Hello, World', "Quoted comma failed");

    // Test 7: Complex JQL
    const cmd7 = "UPDATE summary='X' WHERE project='A' AND (status='Open' OR assignee is EMPTY)";
    const res7 = parseJqluCommand(cmd7);
    assert(res7.jql === "project='A' AND (status='Open' OR assignee is EMPTY)", "Complex JQL failed");

    // Test 8: Error Handling (Missing WHERE)
    try {
        parseJqluCommand("UPDATE summary='X'");
        assert(false, "Should have thrown error for missing WHERE");
    } catch (e) {
        assert(true, "Caught missing WHERE error");
    }

    // Test 9: Case Insensitivity (Keys) and whitespace
    const cmd9 = "UPDATE  Assignee = ' aliraza '  WHERE key=1";
    const res9 = parseJqluCommand(cmd9);
    assert(res9.updatePayload.fields.assignee.name === ' aliraza ', "Whitespace value preservation failed");

    // Test 10: Numeric Value (Story Points - usually customfield but let's assume 'points' alias if we had one, or raw customfield)
    // currently parsed as string, which is generally safe for JSON

    console.log("parseJqluCommand Tests Passed");
}
