import { google } from 'googleapis';
import { GoogleGenerativeAI } from '@google/generative-ai';
import { getAuthenticatedSheetsClient } from './oauthController.js';
import {
    analyzePromptIntent,
    applyIntelligentArrangement,
    updateSpreadsheetWithArrangement,
    createCompactTestCasesContext,
    createCompactTestScenariosContext,
    parseGeminiJSON,
    parseTestScenariosJSON,
    validateAndCleanTestCases,
    callGeminiWithRetry,
    appendTestCasesToExistingSheet,
    appendTestScenariosToExistingSheet,
    addTestCasesSheetData,
    addTestScenariosSheetData,
    // ADD THESE NEW IMPORTS:
    updateGenerateTestCasesPrompt,
    parseGeminiJSONEnhanced
} from '../lib/sheetsHelpers.js';
import dotenv from "dotenv";
dotenv.config();

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

export const getTestCases = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.query;
        const userId = req.user._id.toString();

        console.log("📊 Getting test cases for user:", userId);

        const sheets = await getAuthenticatedSheetsClient(userId);

        // FIXED: Use A:J range for 10 columns
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ testCases: [], message: 'No test cases found' });
        }

        // FIXED: Map to new column structure
        const testCases = rows.slice(1).map((row, index) => ({
            rowIndex: index + 2,
            id: row[0] || '',
            module: row[1] || '',
            submodule: row[2] || '',
            summary: row[3] || '',
            testSteps: row[4] || '',
            expectedResults: row[5] || '',
            actualResult: row[6] || '',
            testCaseType: row[7] || 'Positive',
            environment: row[8] || 'Test',
            status: row[9] || 'Not Tested'
        })).filter(tc => tc.id);

        res.json({
            testCases,
            totalCount: testCases.length,
            sheetName
        });

    } catch (error) {
        console.error('Error getting test cases:', error);
        res.status(500).json({
            message: 'Failed to get test cases',
            error: error.message
        });
    }
};

export const analyzeTestCases = async (req, res) => {
    try {
        const { spreadsheetId, sheetName, analysisType = 'general' } = req.body;
        const userId = req.user._id.toString();

        console.log("🔍 Analyzing test cases for user:", userId);

        const sheets = await getAuthenticatedSheetsClient(userId);

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ analysis: 'No test cases found to analyze' });
        }

        const testCases = rows.slice(1).map((row, index) => ({
            id: row[0] || '',
            module: row[1] || '',
            submodule: row[2] || '',
            summary: row[3] || '',
            testSteps: row[4] || '',
            expectedResults: row[5] || '',
            actualResult: row[6] || '',
            testCaseType: row[7] || 'Positive',
            environment: row[8] || 'Test',
            status: row[9] || 'Not Tested'
        })).filter(tc => tc.id);

        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        let analysisPrompt = '';

        // FIXED: Move the switch statement INSIDE the function
        switch (analysisType) {
            case 'coverage':
                // Instead of sending full JSON, send summary statistics
                const moduleStats = {};
                const typeStats = {};
                testCases.forEach(tc => {
                    moduleStats[tc.module] = (moduleStats[tc.module] || 0) + 1;
                    typeStats[tc.testCaseType] = (typeStats[tc.testCaseType] || 0) + 1;
                });

                analysisPrompt = `
                Analyze test coverage for ${testCases.length} test cases:

                Module distribution: ${Object.entries(moduleStats).map(([k, v]) => `${k}: ${v}`).join(', ')}
                Type distribution: ${Object.entries(typeStats).map(([k, v]) => `${k}: ${v}`).join(', ')}
                
                Sample test cases (first 5):
                ${testCases.slice(0, 5).map(tc => `- ${tc.id}: ${tc.summary} (${tc.testCaseType})`).join('\n')}

                Provide analysis on:
                1. Coverage gaps in modules/submodules
                2. Test type balance recommendations
                3. Missing edge cases
                4. Improvement suggestions
                `;
                break;

            case 'quality':
                // Send only essential quality indicators
                const qualityStats = testCases.slice(0, 10).map(tc => ({
                    id: tc.id,
                    summary: tc.summary,
                    stepsLength: tc.testSteps.length,
                    expectedLength: tc.expectedResults.length,
                    hasDetailedSteps: tc.testSteps.includes('Step 1') || tc.testSteps.includes('1.')
                }));

                analysisPrompt = `
                Analyze quality of ${testCases.length} test cases.

                Sample quality indicators:
                ${qualityStats.map(q => `- ${q.id}: Summary length: ${q.summary.length}, Steps: ${q.stepsLength} chars, Expected: ${q.expectedLength} chars, Structured: ${q.hasDetailedSteps}`).join('\n')}

                Evaluate and provide:
                1. Test case clarity assessment
                2. Completeness scoring
                3. Independence verification
                4. Organization recommendations
                5. Automation potential
                `;
                break;

            case 'duplicates':
                // Send only summaries and key identifiers for duplicate detection
                const summariesOnly = testCases.map(tc => ({
                    id: tc.id,
                    summary: tc.summary.toLowerCase().trim(),
                    submodule: tc.submodule,
                    keyWords: tc.testSteps.toLowerCase().split(' ').filter(w => w.length > 4).slice(0, 5).join(' ')
                }));

                analysisPrompt = `
                Find duplicates in ${testCases.length} test cases:

                Test case identifiers:
                ${summariesOnly.map(s => `- ${s.id}: "${s.summary}" | ${s.submodule} | Keywords: ${s.keyWords}`).join('\n')}

                Identify:
                1. Exact duplicate summaries
                2. Similar test cases (>80% similarity)
                3. Overlapping scenarios
                4. Consolidation recommendations
                `;
                break;

            default:
                // General analysis with minimal data
                const generalStats = {
                    total: testCases.length,
                    modules: [...new Set(testCases.map(tc => tc.module))].length,
                    submodules: [...new Set(testCases.map(tc => tc.submodule))].length,
                    positive: testCases.filter(tc => tc.testCaseType === 'Positive').length,
                    negative: testCases.filter(tc => tc.testCaseType === 'Negative').length
                };

                analysisPrompt = `
                Analyze test suite with ${generalStats.total} test cases:
                
                Statistics:
                - ${generalStats.modules} modules, ${generalStats.submodules} submodules
                - ${generalStats.positive} positive, ${generalStats.negative} negative cases
                
                Sample cases:
                ${testCases.slice(0, 8).map(tc => `- ${tc.id}: ${tc.summary} (${tc.testCaseType})`).join('\n')}

                Provide comprehensive analysis with actionable recommendations.
                `;
        }

        console.log("🤖 Calling Gemini AI for analysis...");
        const result = await model.generateContent(analysisPrompt);
        const analysisResponse = await result.response;
        const analysis = analysisResponse.text();

        res.json({
            success: true,
            analysis,
            testCaseCount: testCases.length,
            analysisType,
            sheetName
        });

    } catch (error) {
        console.error('Error analyzing test cases:', error);
        res.status(500).json({
            message: 'Failed to analyze test cases',
            error: error.message
        });
    }
};

export const modifyTestCases = async (req, res) => {
    try {
        const { spreadsheetId, sheetName, modificationPrompt } = req.body;
        const userId = req.user._id.toString();

        console.log("✏️ Modifying test cases for user:", userId);

        const sheets = await getAuthenticatedSheetsClient(userId);

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`, // FIXED: Changed from A:I to A:J
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ message: 'No test cases found to modify' });
        }

        // FIXED: Map to correct column structure
        const testCases = rows.slice(1).map((row, index) => ({
            rowIndex: index + 2,
            id: row[0] || '',
            module: row[1] || '',
            submodule: row[2] || '',
            summary: row[3] || '',
            testSteps: row[4] || '',
            expectedResults: row[5] || '',
            actualResult: row[6] || '',
            testCaseType: row[7] || 'Positive',
            environment: row[8] || 'Test',
            status: row[9] || 'Not Tested'
        })).filter(tc => tc.id);

        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

        // FIXED: Use optimized prompt
        const modificationPromptForAI = `
Modify test cases based on user request.

Test Cases Summary (${testCases.length} total):
${testCases.map(tc => `- ${tc.id}: ${tc.summary} (${tc.testCaseType})`).join('\n')}

User Request: "${modificationPrompt}"

Return JSON with modifications array containing testCaseId, action, changes, and reason.
Rules: Only include changed fields, be precise with IDs, explain reasons clearly.

Return ONLY JSON, no formatting.
`;

        console.log("🤖 Processing modification request with AI...");
        const result = await model.generateContent(modificationPromptForAI);
        const aiResponse = await result.response;
        let modificationPlan = aiResponse.text();

        modificationPlan = modificationPlan.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
        const modifications = JSON.parse(modificationPlan);

        console.log("📝 Modification plan:", modifications);

        const updates = [];
        const addedTestCases = [];

        for (const mod of modifications.modifications) {
            const testCase = testCases.find(tc => tc.id === mod.testCaseId);

            if (mod.action === 'update' && testCase) {
                // FIXED: Update to use correct field names
                const updatedRow = [
                    testCase.id,
                    mod.changes.module || testCase.module,
                    mod.changes.submodule || testCase.submodule,
                    mod.changes.summary || testCase.summary,
                    mod.changes.testSteps || testCase.testSteps,
                    mod.changes.expectedResults || testCase.expectedResults,
                    mod.changes.actualResult || testCase.actualResult,
                    mod.changes.testCaseType || testCase.testCaseType,
                    mod.changes.environment || testCase.environment,
                    mod.changes.status || testCase.status
                ];

                updates.push({
                    range: `'${sheetName}'!A${testCase.rowIndex}:J${testCase.rowIndex}`, // FIXED: Changed to J
                    values: [updatedRow]
                });

            } else if (mod.action === 'delete' && testCase) {
                updates.push({
                    range: `'${sheetName}'!A${testCase.rowIndex}:J${testCase.rowIndex}`, // FIXED: Changed to J
                    values: [['', '', '', '', '', '', '', '', '', '']] // FIXED: Added more empty strings
                });

            } else if (mod.action === 'add') {
                const newRow = [
                    mod.testCaseId,
                    mod.changes.module || '',
                    mod.changes.submodule || '',
                    mod.changes.summary || '',
                    mod.changes.testSteps || '',
                    mod.changes.expectedResults || '',
                    mod.changes.actualResult || '',
                    mod.changes.testCaseType || 'Positive',
                    mod.changes.environment || 'Test',
                    mod.changes.status || 'Not Tested'
                ];
                addedTestCases.push(newRow);
            }
        }

        if (updates.length > 0) {
            await sheets.spreadsheets.values.batchUpdate({
                spreadsheetId,
                requestBody: {
                    valueInputOption: 'USER_ENTERED',
                    data: updates
                }
            });
        }

        if (addedTestCases.length > 0) {
            const lastRow = rows.length + 1;
            await sheets.spreadsheets.values.update({
                spreadsheetId,
                range: `'${sheetName}'!A${lastRow}`,
                valueInputOption: 'USER_ENTERED',
                requestBody: {
                    values: addedTestCases
                }
            });
        }

        console.log("✅ Test cases modified successfully");

        res.json({
            success: true,
            modifications: modifications.modifications,
            summary: modifications.summary,
            updatedCount: updates.length,
            addedCount: addedTestCases.length,
            message: 'Test cases modified successfully'
        });

    } catch (error) {
        console.error('Error modifying test cases:', error);
        res.status(500).json({
            message: 'Failed to modify test cases',
            error: error.message
        });
    }
};

// NEW: Custom Prompt System for Workflow Arrangement
export const processCustomPrompt = async (req, res) => {
    try {
        const { spreadsheetId, sheetName, customPrompt } = req.body;
        const userId = req.user._id.toString();

        console.log("🧠 Processing custom prompt for user:", userId);
        console.log("📝 Custom prompt:", customPrompt);

        const sheets = await getAuthenticatedSheetsClient(userId);

        // FIXED: Get current test cases with A:J range
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ message: 'No test cases found to arrange' });
        }

        // FIXED: Map to new column structure
        const testCases = rows.slice(1).map((row, index) => ({
            rowIndex: index + 2,
            id: row[0] || '',
            module: row[1] || '',
            submodule: row[2] || '',
            summary: row[3] || '',
            testSteps: row[4] || '',
            expectedResults: row[5] || '',
            actualResult: row[6] || '',
            testCaseType: row[7] || 'Positive',
            environment: row[8] || 'Test',
            status: row[9] || 'Not Tested'
        })).filter(tc => tc.id);

        // Rest of the function remains the same...
        const intentAnalysis = await analyzePromptIntent(customPrompt, testCases);
        console.log("🎯 Intent analysis:", intentAnalysis);

        const arrangedTestCases = await applyIntelligentArrangement(
            testCases,
            customPrompt,
            intentAnalysis
        );

        const success = await updateSpreadsheetWithArrangement(
            sheets,
            spreadsheetId,
            sheetName,
            arrangedTestCases,
            rows[0] // header row
        );

        if (success) {
            res.json({
                success: true,
                originalCount: testCases.length,
                arrangedCount: arrangedTestCases.length,
                intent: intentAnalysis.intent,
                arrangementStrategy: intentAnalysis.strategy,
                summary: intentAnalysis.summary,
                changes: intentAnalysis.changes,
                message: 'Test cases arranged successfully according to your prompt'
            });
        } else {
            throw new Error('Failed to update spreadsheet');
        }

    } catch (error) {
        console.error('Error processing custom prompt:', error);
        res.status(500).json({
            message: 'Failed to process custom prompt',
            error: error.message
        });
    }
};

export const generateTestCasesWithOptions = async (req, res) => {
    try {
        const {
            module,
            summary,
            acceptanceCriteria,
            spreadsheetId,
            generateTestCases = true,
            generateTestScenarios = true,
            testCasesCount = 20,
            testScenariosCount = 10,
            testCasesSheetName = null,
            testScenariosSheetName = null
        } = req.body;

        const userId = req.user._id.toString();
        console.log("🔧 Generate with options for user:", userId);
        console.log(`📊 Requested: ${testCasesCount} test cases, ${testScenariosCount} scenarios`);

        // Validate input
        if (!module || !summary || !acceptanceCriteria) {
            return res.status(400).json({
                success: false,
                message: 'Missing required fields: module, summary, or acceptanceCriteria'
            });
        }

        // Check if at least one generation option is selected
        if (!generateTestCases && !generateTestScenarios) {
            return res.status(400).json({
                success: false,
                message: 'At least one generation option (testCases or testScenarios) must be selected'
            });
        }

        const sheets = await getAuthenticatedSheetsClient(userId);

        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');

        let testCases = [];
        let testScenarios = [];
        let createdSheets = [];

        // OPTIMIZED: Use compact context generation
        let existingTestCasesContext = "";
        let existingTestScenariosContext = "";

        // OPTIMIZED: Get existing test cases context (lightweight)
        if (generateTestCases && testCasesSheetName && testCasesSheetName.trim()) {
            try {
                const existingResponse = await sheets.spreadsheets.values.get({
                    spreadsheetId,
                    range: `'${testCasesSheetName.trim()}'!A:J`,
                });

                const existingRows = existingResponse.data.values || [];
                if (existingRows.length > 1) {
                    // OPTIMIZED: Extract only essential fields
                    const existingTestCases = existingRows.slice(1).map((row, index) => ({
                        id: row[0] || '',
                        module: row[1] || '',
                        submodule: row[2] || '',
                        summary: row[3] || '',
                        testSteps: row[4] || '', // ADD testSteps for enhanced validation
                        testCaseType: row[7] || 'Positive'
                    })).filter(tc => tc.id);

                    if (existingTestCases.length > 0) {
                        existingTestCasesContext = createCompactTestCasesContext(existingTestCases, testCasesSheetName);
                    }
                }
            } catch (error) {
                console.log("⚠️ Could not fetch existing test cases context:", error.message);
            }
        }

        // OPTIMIZED: Get existing test scenarios context (lightweight)
        if (generateTestScenarios && testScenariosSheetName && testScenariosSheetName.trim()) {
            try {
                const existingResponse = await sheets.spreadsheets.values.get({
                    spreadsheetId,
                    range: `'${testScenariosSheetName.trim()}'!A:D`,
                });

                const existingRows = existingResponse.data.values || [];
                if (existingRows.length > 1) {
                    // OPTIMIZED: Extract only essential fields
                    const existingTestScenarios = existingRows.slice(1).map((row, index) => ({
                        condition: row[1] || '',
                        testScenarios: row[2] || ''
                    })).filter(ts => ts.testScenarios);

                    if (existingTestScenarios.length > 0) {
                        existingTestScenariosContext = createCompactTestScenariosContext(existingTestScenarios, testScenariosSheetName);
                    }
                }
            } catch (error) {
                console.log("⚠️ Could not fetch existing test scenarios context:", error.message);
            }
        }

        // ENHANCED: Generate Test Cases with enhanced duplicate prevention
        if (generateTestCases) {
            let nextIdNumber = 1;
            let existingTestCasesForValidation = [];

            if (existingTestCasesContext) {
                const lastIdMatch = existingTestCasesContext.match(/Last ID: PC_(\d+)/);
                if (lastIdMatch) {
                    nextIdNumber = parseInt(lastIdMatch[1]) + 1;
                }

                // Get existing test cases for validation
                try {
                    const existingResponse = await sheets.spreadsheets.values.get({
                        spreadsheetId,
                        range: `'${testCasesSheetName.trim()}'!A:J`,
                    });

                    const existingRows = existingResponse.data.values || [];
                    if (existingRows.length > 1) {
                        existingTestCasesForValidation = existingRows.slice(1).map((row, index) => ({
                            id: row[0] || '',
                            module: row[1] || '',
                            submodule: row[2] || '',
                            summary: row[3] || '',
                            testSteps: row[4] || '',
                            expectedResults: row[5] || '',
                            testCaseType: row[7] || 'Positive'
                        })).filter(tc => tc.id);
                    }
                } catch (error) {
                    console.log("⚠️ Could not fetch existing test cases for validation:", error.message);
                }
            }

            // ENHANCED: Use the improved prompt
            const testCasesPrompt = updateGenerateTestCasesPrompt(
                module,
                summary,
                acceptanceCriteria,
                testCasesCount,
                existingTestCasesContext,
                nextIdNumber
            );

            try {
                console.log("🤖 Generating test cases with enhanced duplicate prevention...");
                const testCasesText = await callGeminiWithRetry(model, testCasesPrompt, 3, 2000);

                // ENHANCED: Use enhanced parsing and validation
                testCases = parseGeminiJSONEnhanced(testCasesText, existingTestCasesForValidation);

                console.log(`📊 Final unique test cases after enhanced validation: ${testCases.length}`);

                if (testCases.length > 0) {
                    let finalTestCasesSheetName;

                    if (testCasesSheetName && testCasesSheetName.trim()) {
                        finalTestCasesSheetName = testCasesSheetName.trim();
                        console.log("📝 Appending test cases to existing sheet:", finalTestCasesSheetName);
                        await appendTestCasesToExistingSheet(sheets, spreadsheetId, finalTestCasesSheetName, testCases, module);
                    } else {
                        finalTestCasesSheetName = `Test Cases - ${module} - ${timestamp}`;
                        console.log("📝 Creating new test cases sheet:", finalTestCasesSheetName);

                        await sheets.spreadsheets.batchUpdate({
                            spreadsheetId,
                            requestBody: {
                                requests: [{
                                    addSheet: {
                                        properties: {
                                            title: finalTestCasesSheetName,
                                            gridProperties: {
                                                rowCount: 100,
                                                columnCount: 10
                                            }
                                        }
                                    }
                                }]
                            }
                        });

                        await addTestCasesSheetData(sheets, spreadsheetId, finalTestCasesSheetName, testCases, module);
                    }

                    createdSheets.push({
                        type: 'testCases',
                        name: finalTestCasesSheetName,
                        count: testCases.length,
                        action: testCasesSheetName ? 'appended' : 'created'
                    });
                }
            } catch (error) {
                console.error('❌ Error generating test cases:', error);

                // Handle specific API errors
                if (error.message.includes('503') || error.message.includes('overloaded')) {
                    return res.status(503).json({
                        success: false,
                        message: 'AI service is currently overloaded. Please try again in a few minutes.',
                        errorType: 'SERVICE_OVERLOADED',
                        retryAfter: 60000
                    });
                }

                if (error.message.includes('429') || error.message.includes('rate limit')) {
                    return res.status(429).json({
                        success: false,
                        message: 'Rate limit exceeded. Please wait before making another request.',
                        errorType: 'RATE_LIMITED',
                        retryAfter: 30000
                    });
                }

                throw error;
            }
        }

        // OPTIMIZED: Generate Test Scenarios with lightweight prompt  
        if (generateTestScenarios) {
            // OPTIMIZED: Much shorter prompt
            const testScenariosPrompt = `
Generate ${testScenariosCount} test scenarios for: ${module}

Summary: ${summary}
Acceptance Criteria: ${acceptanceCriteria}
${existingTestScenariosContext ? `\nExisting Context:\n${existingTestScenariosContext}` : ''}

Requirements:
- Create different conditions/workflows than existing
- Cover various user states and system conditions
- Include valid/invalid inputs, error scenarios

JSON Format:
[
  {
    "id": "TS_1",
    "module": "${module}",
    "condition": "Specific condition/state",
    "testScenarios": "Complete workflow description",
    "status": "Not Tested"
  }
]

Return ONLY JSON array with ${testScenariosCount} scenarios.
`;

            try {
                console.log("🤖 Generating test scenarios...");
                const testScenariosText = await callGeminiWithRetry(model, testScenariosPrompt, 3, 2000);
                testScenarios = parseTestScenariosJSON(testScenariosText);

                if (testScenarios.length > 0) {
                    let finalTestScenariosSheetName;

                    if (testScenariosSheetName && testScenariosSheetName.trim()) {
                        finalTestScenariosSheetName = testScenariosSheetName.trim();
                        console.log("📝 Appending test scenarios to existing sheet:", finalTestScenariosSheetName);
                        await appendTestScenariosToExistingSheet(sheets, spreadsheetId, finalTestScenariosSheetName, testScenarios, module);
                    } else {
                        finalTestScenariosSheetName = `Test Scenarios - ${module} - ${timestamp}`;
                        console.log("📝 Creating new test scenarios sheet:", finalTestScenariosSheetName);

                        await sheets.spreadsheets.batchUpdate({
                            spreadsheetId,
                            requestBody: {
                                requests: [{
                                    addSheet: {
                                        properties: {
                                            title: finalTestScenariosSheetName,
                                            gridProperties: {
                                                rowCount: 50,
                                                columnCount: 4
                                            }
                                        }
                                    }
                                }]
                            }
                        });

                        await addTestScenariosSheetData(sheets, spreadsheetId, finalTestScenariosSheetName, testScenarios, module);
                    }

                    createdSheets.push({
                        type: 'testScenarios',
                        name: finalTestScenariosSheetName,
                        count: testScenarios.length,
                        action: testScenariosSheetName ? 'appended' : 'created'
                    });
                }
            } catch (error) {
                console.error('❌ Error generating test scenarios:', error);

                // Handle specific API errors
                if (error.message.includes('503') || error.message.includes('overloaded')) {
                    return res.status(503).json({
                        success: false,
                        message: 'AI service is currently overloaded. Please try again in a few minutes.',
                        errorType: 'SERVICE_OVERLOADED',
                        retryAfter: 60000
                    });
                }

                if (error.message.includes('429') || error.message.includes('rate limit')) {
                    return res.status(429).json({
                        success: false,
                        message: 'Rate limit exceeded. Please wait before making another request.',
                        errorType: 'RATE_LIMITED',
                        retryAfter: 30000
                    });
                }

                throw error;
            }
        }

        console.log("✅ Generation completed successfully");

        res.json({
            success: true,
            testCases: generateTestCases ? testCases : [],
            testScenarios: generateTestScenarios ? testScenarios : [],
            createdSheets,
            message: `Successfully ${createdSheets.map(s => `${s.action} ${s.count} ${s.type === 'testCases' ? 'test cases' : 'test scenarios'} ${s.action === 'appended' ? 'to' : 'in'} "${s.name}"`).join(' and ')}`
        });

    } catch (error) {
        console.error('❌ Error in generateTestCasesWithOptions:', error);

        // Handle different types of errors
        if (error.message.includes('Authentication')) {
            return res.status(401).json({
                success: false,
                message: 'Authentication failed. Please reconnect your Google Sheets.',
                errorType: 'AUTH_ERROR'
            });
        }

        if (error.message.includes('Permission')) {
            return res.status(403).json({
                success: false,
                message: 'Permission denied. Please check your Google Sheets permissions.',
                errorType: 'PERMISSION_ERROR'
            });
        }

        if (error.message.includes('not found')) {
            return res.status(404).json({
                success: false,
                message: 'Spreadsheet or sheet not found. Please check your selection.',
                errorType: 'NOT_FOUND'
            });
        }

        // Generic error response
        res.status(500).json({
            success: false,
            message: 'An unexpected error occurred. Please try again.',
            errorType: 'INTERNAL_ERROR',
            error: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
};

export const updateSheetFormatting = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.body;
        const userId = req.user._id.toString();

        const sheets = await getAuthenticatedSheetsClient(userId);

        // Get sheet ID
        const spreadsheetInfo = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const sheet = spreadsheetInfo.data.sheets.find(s =>
            s.properties.title === sheetName
        );

        if (!sheet) {
            return res.status(404).json({ message: 'Sheet not found' });
        }

        const sheetId = sheet.properties.sheetId;

        // Update header row to match image format
        const newHeaders = [
            'Test Case ID',
            'Module',
            'Submodule',
            'Summary',
            'Test Steps',
            'Expected Results',
            'Actual Result',
            'Test Case Type',
            'Environment',
            'Status'
        ];

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${sheetName}'!A1:J1`,
            valueInputOption: 'USER_ENTERED',
            requestBody: {
                values: [newHeaders]
            }
        });

        // Apply the formatting
        const formatRequests = [
            // Header formatting
            {
                repeatCell: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: 0,
                        endRowIndex: 1,
                        startColumnIndex: 0,
                        endColumnIndex: 10
                    },
                    cell: {
                        userEnteredFormat: {
                            backgroundColor: {
                                red: 0.2,
                                green: 0.4,
                                blue: 0.7
                            },
                            textFormat: {
                                foregroundColor: {
                                    red: 1.0,
                                    green: 1.0,
                                    blue: 1.0
                                },
                                bold: true,
                                fontSize: 10
                            },
                            horizontalAlignment: 'CENTER',
                            verticalAlignment: 'MIDDLE'
                        }
                    },
                    fields: 'userEnteredFormat'
                }
            }
        ];

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: formatRequests
            }
        });

        res.json({
            success: true,
            message: 'Sheet formatting updated successfully'
        });

    } catch (error) {
        console.error('Error updating sheet formatting:', error);
        res.status(500).json({
            message: 'Failed to update sheet formatting',
            error: error.message
        });
    }
};