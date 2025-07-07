import { google } from 'googleapis';
import { GoogleGenerativeAI } from '@google/generative-ai';
import UserToken from '../models/UserToken.js';
import dotenv from "dotenv"
dotenv.config();

const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
);

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Workflow patterns and domain knowledge
const WORKFLOW_PATTERNS = {
    login: {
        sequence: ['setup', 'authentication', 'authorization', 'session_management', 'logout', 'cleanup'],
        dependencies: ['setup_before_auth', 'auth_before_session', 'session_before_main_features']
    },
    ecommerce: {
        sequence: ['browse', 'search', 'product_view', 'cart', 'checkout', 'payment', 'confirmation', 'order_tracking'],
        criticalPath: ['payment', 'checkout', 'inventory']
    },
    api: {
        sequence: ['authentication', 'crud_operations', 'error_handling', 'performance', 'security'],
        dependencies: ['auth_required_before_data_operations']
    },
    user_management: {
        sequence: ['registration', 'profile_setup', 'permissions', 'user_operations', 'deactivation'],
        criticalPath: ['registration', 'authentication', 'permissions']
    }
};

const TESTING_BEST_PRACTICES = {
    execution_order: ['setup', 'positive_cases', 'negative_cases', 'edge_cases', 'cleanup'],
    priority_mapping: { 'High': 1, 'Medium': 2, 'Low': 3 },
    type_order: ['Functional', 'Integration', 'Edge Case', 'Security', 'Performance']
};

export const getAuthUrl = async (req, res) => {
    try {
        const scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ];

        const authUrl = oauth2Client.generateAuthUrl({
            access_type: 'offline',
            scope: scopes,
            state: req.user._id.toString(),
            prompt: 'consent'
        });

        res.json({ authUrl });
    } catch (error) {
        console.error('Error generating auth URL:', error);
        res.status(500).json({ message: 'Failed to generate authorization URL' });
    }
};

export const handleCallback = async (req, res) => {
    try {
        const { code, state } = req.query;
        const userId = state;

        console.log("🔄 OAuth Callback for user:", userId);

        const { tokens } = await oauth2Client.getToken(code);

        console.log("✅ Received tokens:", {
            hasAccessToken: !!tokens.access_token,
            hasRefreshToken: !!tokens.refresh_token,
            expiryDate: tokens.expiry_date
        });

        await UserToken.findOneAndUpdate(
            { userId: userId },
            {
                userId: userId,
                accessToken: tokens.access_token,
                refreshToken: tokens.refresh_token,
                expiryDate: new Date(tokens.expiry_date),
                scope: tokens.scope?.split(' ') || []
            },
            { upsert: true, new: true }
        );

        console.log("💾 Tokens saved to database for user:", userId);

        res.send(`
            <html>
                <body>
                    <script>
                        window.close();
                    </script>
                    <p>Authorization successful! You can close this window.</p>
                </body>
            </html>
        `);
    } catch (error) {
        console.error('Error handling OAuth callback:', error);
        res.status(500).send('Authorization failed');
    }
};

export const getConnectionStatus = async (req, res) => {
    try {
        const userId = req.user._id.toString();
        const userToken = await UserToken.findOne({ userId: userId });

        if (!userToken) {
            console.log("❌ No tokens found in database for user:", userId);
            return res.json({ connected: false });
        }

        console.log("✅ Tokens found in database for user:", userId);

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const drive = google.drive({ version: 'v3', auth: oauth2Client });

        const response = await drive.files.list({
            q: "mimeType='application/vnd.google-apps.spreadsheet'",
            pageSize: 10,
            fields: 'files(id, name)'
        });

        res.json({
            connected: true,
            spreadsheets: response.data.files
        });
    } catch (error) {
        console.error('Error checking connection status:', error);

        if (error.message.includes('invalid_grant') || error.message.includes('Token has been expired')) {
            console.log("🔄 Token expired, user needs to re-authenticate");
            await UserToken.findOneAndDelete({ userId: req.user._id.toString() });
        }

        res.json({ connected: false });
    }
};

export const getTestCases = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.query;
        const userId = req.user._id.toString();

        console.log("📊 Getting test cases for user:", userId);

        const userToken = await UserToken.findOne({ userId: userId });

        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

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

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

        // FIXED: Use A:J range for 10 columns
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ analysis: 'No test cases found to analyze' });
        }

        // FIXED: Map to new column structure
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

        switch (analysisType) {
            case 'coverage':
                analysisPrompt = `
                Analyze the test coverage of these test cases:

                ${JSON.stringify(testCases, null, 2)}

                Provide analysis on:
                1. Functional coverage gaps
                2. Test case type distribution (Positive vs Negative)
                3. Module and submodule coverage
                4. Missing edge cases
                5. Recommendations for improvement

                Return the analysis in a structured format.
                `;
                break;

            case 'quality':
                analysisPrompt = `
                Analyze the quality of these test cases:

                ${JSON.stringify(testCases, null, 2)}

                Evaluate:
                1. Clarity and completeness of test steps
                2. Adequacy of expected results
                3. Test case independence
                4. Proper module/submodule organization
                5. Potential for automation
                6. Quality score for each test case

                Provide detailed feedback and suggestions.
                `;
                break;

            case 'duplicates':
                analysisPrompt = `
                Find duplicate or similar test cases:

                ${JSON.stringify(testCases, null, 2)}

                Identify:
                1. Exact duplicates
                2. Similar test cases that could be merged
                3. Overlapping test scenarios
                4. Recommendations for consolidation

                List specific test case IDs that are duplicates or similar.
                `;
                break;

            default:
                analysisPrompt = `
                Provide a comprehensive analysis of these test cases:

                ${JSON.stringify(testCases, null, 2)}

                Include:
                1. Overall test suite summary
                2. Strengths and weaknesses
                3. Coverage analysis by module and submodule
                4. Quality assessment
                5. Test case type distribution
                6. Specific recommendations
                7. Test metrics and statistics

                Make the analysis actionable and detailed.
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

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:I`,
        });

        const rows = response.data.values;
        if (!rows || rows.length <= 1) {
            return res.json({ message: 'No test cases found to modify' });
        }

        const testCases = rows.slice(1).map((row, index) => ({
            rowIndex: index + 2,
            id: row[0] || '',
            title: row[1] || '',
            description: row[2] || '',
            preconditions: row[3] || '',
            steps: row[4] || '',
            expectedResult: row[5] || '',
            priority: row[6] || 'Medium',
            type: row[7] || 'Functional',
            status: row[8] || 'Not Executed'
        })).filter(tc => tc.id);

        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

        const modificationPromptForAI = `
        You are helping to modify test cases based on user instructions. 

        Current test cases:
        ${JSON.stringify(testCases, null, 2)}

        User modification request: "${modificationPrompt}"

        Analyze the user's request and return a JSON response with the following structure:
        {
            "modifications": [
                {
                    "testCaseId": "TC001",
                    "action": "update|delete|add",
                    "changes": {
                        "title": "new title if changed",
                        "description": "new description if changed",
                        "steps": "new steps if changed",
                        "expectedResult": "new expected result if changed",
                        "priority": "new priority if changed",
                        "type": "new type if changed",
                        "preconditions": "new preconditions if changed"
                    },
                    "reason": "explanation of why this change was made"
                }
            ],
            "summary": "Summary of all modifications made"
        }

        Rules:
        - Only include fields that are being changed in the "changes" object
        - For "delete" action, changes object can be empty
        - For "add" action, include all required fields for the new test case
        - Be precise about which test cases to modify based on the user's request
        - If the request is unclear, suggest reasonable interpretations

        Return ONLY the JSON response, no additional text or formatting.
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
                const updatedRow = [
                    testCase.id,
                    mod.changes.title || testCase.title,
                    mod.changes.description || testCase.description,
                    mod.changes.preconditions || testCase.preconditions,
                    mod.changes.steps || testCase.steps,
                    mod.changes.expectedResult || testCase.expectedResult,
                    mod.changes.priority || testCase.priority,
                    mod.changes.type || testCase.type,
                    testCase.status
                ];

                updates.push({
                    range: `'${sheetName}'!A${testCase.rowIndex}:I${testCase.rowIndex}`,
                    values: [updatedRow]
                });

            } else if (mod.action === 'delete' && testCase) {
                updates.push({
                    range: `'${sheetName}'!A${testCase.rowIndex}:I${testCase.rowIndex}`,
                    values: [['', '', '', '', '', '', '', '', '']]
                });

            } else if (mod.action === 'add') {
                const newRow = [
                    mod.testCaseId,
                    mod.changes.title || '',
                    mod.changes.description || '',
                    mod.changes.preconditions || '',
                    mod.changes.steps || '',
                    mod.changes.expectedResult || '',
                    mod.changes.priority || 'Medium',
                    mod.changes.type || 'Functional',
                    'Not Executed'
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

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

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

// Analyze the intent behind the custom prompt
async function analyzePromptIntent(prompt, testCases) {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const intentPrompt = `
    Analyze this user prompt for test case arrangement and determine the intent:
    
    User Prompt: "${prompt}"
    
    Current Test Cases Summary:
    - Total: ${testCases.length}
    - Priorities: ${getDistribution(testCases, 'priority')}
    - Types: ${getDistribution(testCases, 'type')}
    - Test Case IDs: ${testCases.map(tc => tc.id).join(', ')}
    
    Available Workflow Patterns: ${Object.keys(WORKFLOW_PATTERNS).join(', ')}
    Testing Best Practices: ${JSON.stringify(TESTING_BEST_PRACTICES)}
    
    Return a JSON response with this structure:
    {
        "intent": "workflow|priority|dependency|risk|module|execution|custom",
        "strategy": "specific arrangement strategy to apply",
        "confidence": 0.95,
        "detectedPattern": "login|ecommerce|api|user_management|custom",
        "arrangementCriteria": {
            "primarySort": "field to sort by first",
            "secondarySort": "field to sort by second", 
            "groupBy": "field to group by if applicable",
            "direction": "ascending|descending"
        },
        "summary": "Clear explanation of what will be done",
        "changes": "Description of expected changes"
    }
    
    Consider these intent categories:
    - "workflow": Arrange in logical business/user flow order
    - "priority": Sort by test priority or business importance  
    - "dependency": Order by technical dependencies
    - "risk": Arrange by risk level or business impact
    - "module": Group by functionality/module
    - "execution": Optimize for test execution efficiency
    - "custom": Other specific arrangements
    
    Return ONLY valid JSON.
    `;

    try {
        const result = await model.generateContent(intentPrompt);
        const response = await result.response;
        const intentText = response.text();

        // Clean and parse the response
        const cleanedText = intentText.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
        return JSON.parse(cleanedText);
    } catch (error) {
        console.error('Error analyzing intent:', error);
        // Return default intent analysis
        return {
            intent: 'workflow',
            strategy: 'basic_workflow_arrangement',
            confidence: 0.7,
            detectedPattern: 'custom',
            arrangementCriteria: {
                primarySort: 'priority',
                secondarySort: 'type',
                groupBy: null,
                direction: 'ascending'
            },
            summary: 'Arranging test cases in a logical workflow order',
            changes: 'Test cases will be reordered based on testing best practices'
        };
    }
}

// Apply intelligent arrangement based on intent
async function applyIntelligentArrangement(testCases, prompt, intentAnalysis) {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const arrangementPrompt = `
    You are an expert test case arrangement system. Arrange these test cases according to the user's request.
    
    User Prompt: "${prompt}"
    Intent Analysis: ${JSON.stringify(intentAnalysis)}
    
    Current Test Cases:
    ${JSON.stringify(testCases, null, 2)}
    
    Workflow Patterns Available:
    ${JSON.stringify(WORKFLOW_PATTERNS, null, 2)}
    
    Testing Best Practices:
    ${JSON.stringify(TESTING_BEST_PRACTICES, null, 2)}
    
    Instructions:
    1. Analyze the test cases and understand their relationships
    2. Apply the detected intent and strategy: ${intentAnalysis.strategy}
    3. Consider dependencies, business logic flow, and testing best practices
    4. Maintain test case integrity (don't modify content, only order)
    5. Return the complete reordered array
    
    Return a JSON response with this structure:
    {
        "arrangedTestCases": [
            // Complete array of test cases in new order
            // Keep all original fields: rowIndex, id, title, description, etc.
        ],
        "arrangementLogic": "Explanation of the arrangement logic applied",
        "groupings": [
            {
                "name": "Setup & Prerequisites", 
                "testCases": ["TC001", "TC002"],
                "reason": "Why this grouping"
            }
        ],
        "dependencies": [
            {
                "before": "TC001",
                "after": "TC005", 
                "reason": "TC001 must run before TC005 because..."
            }
        ]
    }
    
    Key Principles:
    - Respect logical flow (setup → positive → negative → edge cases → cleanup)
    - Group related functionality together
    - Consider technical dependencies
    - Prioritize by business impact when requested
    - Maintain testing efficiency
    
    Return ONLY valid JSON with the complete reordered test case array.
    `;

    try {
        const result = await model.generateContent(arrangementPrompt);
        const response = await result.response;
        const arrangementText = response.text();
        const cleanedText = arrangementText.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
        const arrangementResult = JSON.parse(cleanedText);

        // Validate that we have all original test cases
        if (arrangementResult.arrangedTestCases.length !== testCases.length) {
            console.warn('⚠️ Arrangement result has different count, falling back to priority sort');
            return sortByPriority(testCases);
        }

        return arrangementResult.arrangedTestCases;

    } catch (error) {
        console.error('Error in arrangement:', error);
        // Fallback to simple priority-based sorting
        return sortByPriority(testCases);
    }
}

// Update spreadsheet with new arrangement
async function updateSpreadsheetWithArrangement(sheets, spreadsheetId, sheetName, arrangedTestCases, headerRow) {
    try {
        // Prepare data for update
        const updatedData = [headerRow]; // Keep original header

        arrangedTestCases.forEach(testCase => {
            updatedData.push([
                testCase.id,
                testCase.module,
                testCase.submodule,
                testCase.summary,
                testCase.testSteps,
                testCase.expectedResults,
                testCase.actualResult || '',
                testCase.testCaseType,
                testCase.environment,
                testCase.status
            ]);
        });

        // FIXED: Clear the existing data for 10 columns
        await sheets.spreadsheets.values.clear({
            spreadsheetId,
            range: `'${sheetName}'!A:J`
        });

        // Write the rearranged data
        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${sheetName}'!A1`,
            valueInputOption: 'USER_ENTERED',
            requestBody: {
                values: updatedData
            }
        });

        console.log("✅ Spreadsheet updated with new arrangement");
        return true;
    } catch (error) {
        console.error('Error updating spreadsheet:', error);
        return false;
    }
}

// Helper functions
function getDistribution(testCases, field) {
    const distribution = {};
    testCases.forEach(tc => {
        distribution[tc[field]] = (distribution[tc[field]] || 0) + 1;
    });
    return Object.entries(distribution).map(([key, value]) => `${key}:${value}`).join(', ');
}

function sortByPriority(testCases) {
    return testCases.sort((a, b) => {
        const priorityOrder = { 'High': 1, 'Medium': 2, 'Low': 3 };
        return priorityOrder[a.priority] - priorityOrder[b.priority];
    });
}

export const getAvailableSheets = async (req, res) => {
    try {
        const { spreadsheetId } = req.query;
        const userId = req.user._id.toString();

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

        const response = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const availableSheets = response.data.sheets.map(sheet => ({
            name: sheet.properties.title,
            id: sheet.properties.sheetId,
            index: sheet.properties.index
        }));

        res.json({
            sheets: availableSheets,
            spreadsheetId
        });

    } catch (error) {
        console.error('Error getting available sheets:', error);
        res.status(500).json({
            message: 'Failed to get available sheets',
            error: error.message
        });
    }
};

// Enhanced function to generate test cases with proper formatting

function parseTestScenariosJSON(text) {
    console.log("🔍 Starting test scenarios JSON parsing...");

    let cleanedText = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();

    const startIndex = cleanedText.indexOf('[');
    const endIndex = cleanedText.lastIndexOf(']');

    if (startIndex === -1 || endIndex === -1 || endIndex <= startIndex) {
        console.log("❌ No valid JSON array boundaries found for scenarios");
        return getFallbackTestScenarios();
    }

    const jsonString = cleanedText.substring(startIndex, endIndex + 1);

    try {
        const parsed = JSON.parse(jsonString);
        if (Array.isArray(parsed) && parsed.length > 0) {
            console.log("✅ Test scenarios JSON parsing successful");
            return validateAndCleanTestScenarios(parsed);
        }
    } catch (error) {
        console.log("❌ Test scenarios parsing failed:", error.message);
    }

    console.log("❌ Test scenarios parsing failed, using fallback");
    return getFallbackTestScenarios();
}


function validateAndCleanTestScenarios(scenarios) {
    return scenarios.map((scenario, index) => ({
        id: scenario.id || `TS_${(index + 1)}`,
        module: scenario.module || '',
        condition: scenario.condition || '',
        testScenarios: scenario.testScenarios || scenario.description || `Test Scenario ${index + 1} Description`,
        status: scenario.status || 'Not Tested'
    })).filter(scenario => scenario.testScenarios);
}

function getFallbackTestScenarios() {
    const fallbackScenarios = [];
    const conditions = ['Using Valid Credentials for Login', 'Using Invalid Credentials for Login', 'Forgot Password', 'Language Change', 'Attempt with empty fields'];

    for (let i = 1; i <= 10; i++) {
        fallbackScenarios.push({
            id: `TS_${i}`,
            module: 'Test Module',
            condition: conditions[(i - 1) % conditions.length] || 'Default Condition',
            testScenarios: `Test Scenario ${i}: End-to-end workflow testing for the module functionality`,
            status: 'Not Tested'
        });
    }
    return fallbackScenarios;
}

// Enhanced function to update existing sheet columns to match the image format
export const updateSheetFormatting = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.body;
        const userId = req.user._id.toString();

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

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

        // Apply the formatting as defined in the previous function
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

// Robust JSON parser for Gemini responses
function parseGeminiJSON(text) {
    console.log("🔍 Starting JSON parsing...");
    console.log("📄 Raw text length:", text.length);
    console.log("📄 First 500 chars:", text.substring(0, 500));

    let cleanedText = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();

    const startIndex = cleanedText.indexOf('[');
    const endIndex = cleanedText.lastIndexOf(']');

    if (startIndex === -1 || endIndex === -1 || endIndex <= startIndex) {
        console.log("❌ No valid JSON array boundaries found");
        console.log("🔧 Using fallback test cases");
        return getFallbackTestCases();
    }

    const jsonString = cleanedText.substring(startIndex, endIndex + 1);
    console.log("📝 Extracted JSON string length:", jsonString.length);
    console.log("📝 First 200 chars of JSON:", jsonString.substring(0, 200));

    // Strategy 1: Direct parsing
    try {
        const parsed = JSON.parse(jsonString);
        if (Array.isArray(parsed) && parsed.length > 0) {
            console.log("✅ Direct JSON parsing successful");
            console.log("📊 Parsed", parsed.length, "test cases");
            const cleaned = validateAndCleanTestCases(parsed);
            console.log("📊 After validation:", cleaned.length, "test cases");
            return cleaned;
        }
    } catch (error) {
        console.log("❌ Direct parsing failed:", error.message);
    }

    // Strategy 2: Fix common JSON issues
    try {
        let fixedJson = jsonString
            .replace(/,(\s*[}\]])/g, '$1')  // Remove trailing commas
            .replace(/\n/g, '\\n')          // Escape newlines
            .replace(/\r/g, '\\r')          // Escape carriage returns
            .replace(/\t/g, '\\t');         // Escape tabs

        console.log("🔧 Attempting to parse fixed JSON...");
        const parsed = JSON.parse(fixedJson);
        if (Array.isArray(parsed) && parsed.length > 0) {
            console.log("✅ Fixed JSON parsing successful");
            return validateAndCleanTestCases(parsed);
        }
    } catch (error) {
        console.log("❌ Fixed JSON parsing failed:", error.message);
    }

    // Strategy 3: Manual parsing
    console.log("🔧 Attempting manual parsing...");
    const manualParsed = manuallyParseTestCases(text);
    if (manualParsed.length > 0) {
        console.log("✅ Manual parsing successful");
        return manualParsed;
    }

    console.log("❌ All parsing strategies failed, using fallback");
    return getFallbackTestCases();
}

function validateAndCleanTestCases(testCases) {
    return testCases.map((tc, index) => {
        let steps = tc.testSteps || tc.steps;
        if (Array.isArray(steps)) {
            steps = steps.join('\n');
        }

        return {
            id: tc.id || `PC_${(index + 1)}`,
            module: tc.module || '',
            submodule: tc.submodule || '',
            summary: tc.summary || tc.title || `Test Case ${index + 1}`,
            testSteps: steps || '',
            expectedResults: tc.expectedResults || tc.expectedResult || '',
            testCaseType: tc.testCaseType || 'Positive',
            environment: tc.environment || 'Test',
            status: tc.status || 'Not Tested'
        };
    }).filter(tc => tc.summary);
}

function manuallyParseTestCases(text) {
    const testCases = [];
    const tcRegex = /"id":\s*"(PC_\d+)"/g;  // Changed from TC to PC_
    const matches = [...text.matchAll(tcRegex)];

    console.log(`🔍 Found ${matches.length} test case IDs manually`);

    for (let i = 0; i < matches.length; i++) {
        const tcId = matches[i][1];
        const startPos = matches[i].index;
        const endPos = i < matches.length - 1 ? matches[i + 1].index : text.length;
        const tcText = text.substring(startPos, endPos);

        const testCase = {
            id: tcId,
            module: extractField(tcText, 'module'),
            submodule: extractField(tcText, 'submodule'),
            summary: extractField(tcText, 'summary'),
            testSteps: extractField(tcText, 'testSteps'),
            expectedResults: extractField(tcText, 'expectedResults'),
            testCaseType: extractField(tcText, 'testCaseType') || 'Positive',
            environment: extractField(tcText, 'environment') || 'Test',
            status: extractField(tcText, 'status') || 'Not Tested'
        };

        if (testCase.summary) {
            testCases.push(testCase);
        }
    }

    return testCases;
}

function extractField(text, fieldName) {
    const regex = new RegExp(`"${fieldName}":\\s*"([^"]*)"`, 'i');
    const match = text.match(regex);
    return match ? match[1].replace(/\\n/g, '\n') : '';
}

function getFallbackTestCases() {
    const fallbackCases = [];
    for (let i = 1; i <= 20; i++) {
        fallbackCases.push({
            id: `PC_${i}`,
            module: 'Test Module',
            submodule: 'Test Submodule',
            summary: `Test Case ${i} Summary`,
            testSteps: `Step 1. Perform action for test case ${i}\nStep 2. Verify result\nStep 3. Validate outcome`,
            expectedResults: `Expected result for test case ${i}`,
            testCaseType: i % 4 === 0 ? 'Negative' : 'Positive',
            environment: 'Test',
            status: 'Not Tested'
        });
    }
    return fallbackCases;
}


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
            // NEW: Sheet selection options
            testCasesSheetName = null,  // If provided, append to existing sheet
            testScenariosSheetName = null  // If provided, append to existing sheet
        } = req.body;

        const userId = req.user._id.toString();

        console.log("🔧 Generate with options for user:", userId);
        console.log("📋 Options:", {
            generateTestCases,
            generateTestScenarios,
            testCasesCount,
            testScenariosCount,
            testCasesSheetName,
            testScenariosSheetName
        });

        const userToken = await UserToken.findOne({ userId: userId });
        if (!userToken) {
            return res.status(401).json({ message: 'Google Sheets not connected' });
        }

        const tokens = {
            access_token: userToken.accessToken,
            refresh_token: userToken.refreshToken,
            expiry_date: userToken.expiryDate.getTime()
        };

        oauth2Client.setCredentials(tokens);
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');

        let testCases = [];
        let testScenarios = [];
        let createdSheets = [];

        // Generate Test Cases if requested
        if (generateTestCases) {
            const testCasesPrompt = `
Generate comprehensive test cases for the following software feature:

Module: ${module}
Summary: ${summary}
Acceptance Criteria: ${acceptanceCriteria}

Create EXACTLY ${testCasesCount} detailed test cases in valid JSON format. Each test case must follow this exact structure:

[
  {
    "id": "PC_1",
    "module": "${module}",
    "submodule": "Specific submodule/component within ${module}",
    "summary": "Clear, concise test case description (what is being tested)",
    "testSteps": "Step 1. Detailed action to perform\\nStep 2. Specific verification step\\nStep 3. Expected result validation",
    "expectedResults": "Detailed description of expected outcome",
    "testCaseType": "Positive",
    "environment": "Test",
    "status": "Not Tested"
  }
]

IMPORTANT REQUIREMENTS:
1. Use EXACTLY the ID format: PC_1, PC_2, PC_3, ..., PC_${testCasesCount}
2. Module field should always be: "${module}"
3. Create meaningful submodules that represent different components/areas within ${module}
4. Include both Positive and Negative test cases (70% Positive, 30% Negative)
5. Use \\n for line breaks in testSteps field (not actual line breaks)
6. Make testSteps detailed with numbered steps
7. Environment should be "Test" for all test cases
8. Status should be "Not Tested" for all test cases
9. Summary should be concise but descriptive
10. ExpectedResults should be detailed and specific

Test Case Distribution:
- 70% Positive test cases
- 30% Negative test cases
- Cover all aspects mentioned in the acceptance criteria
- Include edge cases and boundary conditions
- Create logical submodules based on the functionality

Return ONLY the JSON array with exactly ${testCasesCount} test cases, no markdown formatting, no code blocks, no additional text.
`;

            console.log("🤖 Generating test cases...");
            const testCasesResult = await model.generateContent(testCasesPrompt);
            const testCasesResponse = await testCasesResult.response;
            let testCasesText = testCasesResponse.text();
            testCases = parseGeminiJSON(testCasesText);

            if (testCases.length > 0) {
                let finalTestCasesSheetName;

                // Check if user wants to append to existing sheet
                if (testCasesSheetName && testCasesSheetName.trim()) {
                    finalTestCasesSheetName = testCasesSheetName.trim();
                    console.log("📝 Appending test cases to existing sheet:", finalTestCasesSheetName);

                    // Append to existing sheet
                    await appendTestCasesToExistingSheet(sheets, spreadsheetId, finalTestCasesSheetName, testCases, module);
                } else {
                    // Create new sheet
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

                    // Add test cases data and formatting to new sheet
                    await addTestCasesSheetData(sheets, spreadsheetId, finalTestCasesSheetName, testCases, module);
                }

                createdSheets.push({
                    type: 'testCases',
                    name: finalTestCasesSheetName,
                    count: testCases.length,
                    action: testCasesSheetName ? 'appended' : 'created'
                });
            }
        }

        // Generate Test Scenarios if requested
        if (generateTestScenarios) {
            const testScenariosPrompt = `
Generate test scenarios for the following software feature:

Module: ${module}
Summary: ${summary}
Acceptance Criteria: ${acceptanceCriteria}

Create EXACTLY ${testScenariosCount} test scenarios in valid JSON format. Each test scenario must follow this exact structure:

[
  {
    "id": "TS_1",
    "module": "${module}",
    "condition": "Specific condition or state for this test scenario",
    "testScenarios": "Detailed description of the test scenario covering a complete user workflow or business process",
    "status": "Not Tested"
  }
]

IMPORTANT REQUIREMENTS:
1. Use EXACTLY the ID format: TS_1, TS_2, TS_3, ..., TS_${testScenariosCount}
2. Module field should always be: "${module}"
3. Condition should describe the specific state, input, or situation being tested
4. testScenarios should be a detailed description of the complete workflow
5. Status should always be "Not Tested" for all scenarios
6. Cover different conditions like:
   - Valid/Invalid inputs
   - Different user states (logged in/guest)
   - Various system conditions
   - Error scenarios
   - Edge cases

Examples of good test scenarios:
- Condition: "Using Valid Credentials for Login"
  Test Scenario: "Successful login with valid email and password, followed by navigation to the main dashboard."
  
- Condition: "Using Invalid Credentials for Login" 
  Test Scenario: "Login failure due to incorrect password, verifying error message display and retry functionality."

- Condition: "Forgot Password"
  Test Scenario: "Successful password reset starting from the Forgot Password link to successful login with new password."

Return ONLY the JSON array with exactly ${testScenariosCount} test scenarios, no markdown formatting, no code blocks, no additional text.
`;

            console.log("🤖 Generating test scenarios...");
            const testScenariosResult = await model.generateContent(testScenariosPrompt);
            const testScenariosResponse = await testScenariosResult.response;
            let testScenariosText = testScenariosResponse.text();
            testScenarios = parseTestScenariosJSON(testScenariosText);

            if (testScenarios.length > 0) {
                let finalTestScenariosSheetName;

                // Check if user wants to append to existing sheet
                if (testScenariosSheetName && testScenariosSheetName.trim()) {
                    finalTestScenariosSheetName = testScenariosSheetName.trim();
                    console.log("📝 Appending test scenarios to existing sheet:", finalTestScenariosSheetName);

                    // Append to existing sheet
                    await appendTestScenariosToExistingSheet(sheets, spreadsheetId, finalTestScenariosSheetName, testScenarios, module);
                } else {
                    // Create new sheet
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
                                            columnCount: 3
                                        }
                                    }
                                }
                            }]
                        }
                    });

                    // Add test scenarios data and formatting to new sheet
                    await addTestScenariosSheetData(sheets, spreadsheetId, finalTestScenariosSheetName, testScenarios, module);
                }

                createdSheets.push({
                    type: 'testScenarios',
                    name: finalTestScenariosSheetName,
                    count: testScenarios.length,
                    action: testScenariosSheetName ? 'appended' : 'created'
                });
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
        console.error('Error generating with options:', error);
        res.status(500).json({
            message: 'Failed to generate content',
            error: error.message
        });
    }
};

// NEW: Function to append test cases to existing sheet
async function appendTestCasesToExistingSheet(sheets, spreadsheetId, sheetName, testCases, module) {
    try {
        // Get existing data to find the next available row
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const existingRows = existingData.data.values || [];
        const nextRow = existingRows.length + 1;

        console.log(`📍 Appending ${testCases.length} test cases starting at row ${nextRow}`);

        // Prepare new test case data (without header)
        const newTestCasesData = testCases.map(testCase => [
            testCase.id || '',
            testCase.module || module,
            testCase.submodule || '',
            testCase.summary || '',
            testCase.testSteps || '',
            testCase.expectedResults || '',
            '', // Actual Result - empty initially
            testCase.testCaseType || 'Positive',
            testCase.environment || 'Test',
            testCase.status || 'Not Tested'
        ]);

        // Append the new data
        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${sheetName}'!A${nextRow}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: newTestCasesData }
        });

        console.log("✅ Test cases appended successfully");
        return true;
    } catch (error) {
        console.error('Error appending test cases:', error);
        throw error;
    }
}

// NEW: Function to append test scenarios to existing sheet
async function appendTestScenariosToExistingSheet(sheets, spreadsheetId, sheetName, testScenarios, module) {
    try {
        // Get existing data to find the next available row
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:D`,  // Updated to 4 columns
        });

        const existingRows = existingData.data.values || [];
        const nextRow = existingRows.length + 1;

        console.log(`📍 Appending ${testScenarios.length} test scenarios starting at row ${nextRow}`);

        // Prepare new test scenario data (without header) - Updated format
        const newTestScenariosData = testScenarios.map(scenario => [
            scenario.module || module,
            scenario.condition || '',
            scenario.testScenarios || '',
            scenario.status || 'Not Tested'
        ]);

        // Append the new data
        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `'${sheetName}'!A${nextRow}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: newTestScenariosData }
        });

        console.log("✅ Test scenarios appended successfully");
        return true;
    } catch (error) {
        console.error('Error appending test scenarios:', error);
        throw error;
    }
}


// Helper function to add test cases sheet data
async function addTestCasesSheetData(sheets, spreadsheetId, sheetName, testCases, module) {
    // Get sheet ID for formatting
    const spreadsheetInfo = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties'
    });

    const sheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === sheetName);
    const sheetId = sheet.properties.sheetId;

    // Prepare Test Cases Sheet Data
    const testCasesHeaderRow = [
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

    const testCasesDataRows = testCases.map(testCase => [
        testCase.id || '',
        testCase.module || module,
        testCase.submodule || '',
        testCase.summary || '',
        testCase.testSteps || '',
        testCase.expectedResults || '',
        '', // Actual Result - empty initially
        testCase.testCaseType || 'Positive',
        testCase.environment || 'Test',
        testCase.status || 'Not Tested'
    ]);

    const testCasesAllData = [testCasesHeaderRow, ...testCasesDataRows];

    // Insert data into test cases sheet
    await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${sheetName}'!A1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: testCasesAllData }
    });

    // FIXED: Apply formatting to Test Cases Sheet with corrected conditional formatting
    const testCasesFormatRequests = [
        // Format header row - keep your colors
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
                            red: 0.047,
                            green: 0.204,
                            blue: 0.239
                        },
                        textFormat: {
                            foregroundColor: {
                                red: 1.0,
                                green: 1.0,
                                blue: 1.0
                            },
                            bold: true,
                            fontSize: 11
                        },
                        horizontalAlignment: 'CENTER',
                        verticalAlignment: 'MIDDLE',
                        padding: {
                            top: 8,
                            bottom: 8,
                            left: 4,
                            right: 4
                        }
                    }
                },
                fields: 'userEnteredFormat'
            }
        },
        // Set better column widths for test cases
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 0,
                    endIndex: 1
                },
                properties: { pixelSize: 120 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 1,
                    endIndex: 2
                },
                properties: { pixelSize: 100 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 2,
                    endIndex: 3
                },
                properties: { pixelSize: 130 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 3,
                    endIndex: 4
                },
                properties: { pixelSize: 220 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 4,
                    endIndex: 5
                },
                properties: { pixelSize: 300 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 5,
                    endIndex: 6
                },
                properties: { pixelSize: 250 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 6,
                    endIndex: 7
                },
                properties: { pixelSize: 200 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 7,
                    endIndex: 8
                },
                properties: { pixelSize: 120 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 8,
                    endIndex: 9
                },
                properties: { pixelSize: 100 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 9,
                    endIndex: 10
                },
                properties: { pixelSize: 100 },
                fields: 'pixelSize'
            }
        },
        // Set row heights for better spacing
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: 0,
                    endIndex: 1
                },
                properties: { pixelSize: 35 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: 1,
                    endIndex: testCases.length + 1
                },
                properties: { pixelSize: 80 },
                fields: 'pixelSize'
            }
        },
        // Format data rows with better spacing
        {
            repeatCell: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: 1,
                    endRowIndex: testCases.length + 1,
                    startColumnIndex: 0,
                    endColumnIndex: 10
                },
                cell: {
                    userEnteredFormat: {
                        textFormat: {
                            fontSize: 10
                        },
                        verticalAlignment: 'TOP',
                        wrapStrategy: 'WRAP',
                        padding: {
                            top: 6,
                            bottom: 6,
                            left: 6,
                            right: 6
                        }
                    }
                },
                fields: 'userEnteredFormat'
            }
        },
        // FIXED: Test Case Type conditional formatting - index moved outside rule
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testCases.length + 1,
                        startColumnIndex: 7,
                        endColumnIndex: 8
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Positive' }] },
                        format: { backgroundColor: { red: 0.85, green: 0.95, blue: 0.85 } }
                    }
                },
                index: 0
            }
        },
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testCases.length + 1,
                        startColumnIndex: 7,
                        endColumnIndex: 8
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Negative' }] },
                        format: { backgroundColor: { red: 1.0, green: 0.85, blue: 0.85 } }
                    }
                },
                index: 1
            }
        },
        // FIXED: Status conditional formatting - index moved outside rule
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testCases.length + 1,
                        startColumnIndex: 9,
                        endColumnIndex: 10
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Pass' }] },
                        format: {
                            backgroundColor: { red: 0.0, green: 0.5, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 } }
                        }
                    }
                },
                index: 2
            }
        },
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testCases.length + 1,
                        startColumnIndex: 9,
                        endColumnIndex: 10
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Fail' }] },
                        format: {
                            backgroundColor: { red: 0.6, green: 0.0, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 } }
                        }
                    }
                },
                index: 3
            }
        },
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testCases.length + 1,
                        startColumnIndex: 9,
                        endColumnIndex: 10
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Blocked' }] },
                        format: {
                            backgroundColor: { red: 1.0, green: 1.0, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 0.0, green: 0.0, blue: 0.0 } }
                        }
                    }
                },
                index: 4
            }
        }
    ];

    // Apply all formatting
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: testCasesFormatRequests }
    });

    // Add data validation for Test Cases sheet
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests: [
                {
                    setDataValidation: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: 1,
                            endRowIndex: testCases.length + 1,
                            startColumnIndex: 7,
                            endColumnIndex: 8
                        },
                        rule: {
                            condition: {
                                type: 'ONE_OF_LIST',
                                values: [
                                    { userEnteredValue: 'Positive' },
                                    { userEnteredValue: 'Negative' }
                                ]
                            },
                            showCustomUi: true,
                            strict: true
                        }
                    }
                },
                {
                    setDataValidation: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: 1,
                            endRowIndex: testCases.length + 1,
                            startColumnIndex: 9,
                            endColumnIndex: 10
                        },
                        rule: {
                            condition: {
                                type: 'ONE_OF_LIST',
                                values: [
                                    { userEnteredValue: 'Not Tested' },
                                    { userEnteredValue: 'Pass' },
                                    { userEnteredValue: 'Fail' },
                                    { userEnteredValue: 'Blocked' }
                                ]
                            },
                            showCustomUi: true,
                            strict: true
                        }
                    }
                }
            ]
        }
    });
}

// Helper function to add test scenarios sheet data
async function addTestScenariosSheetData(sheets, spreadsheetId, sheetName, testScenarios, module) {
    // Get sheet ID for formatting
    const spreadsheetInfo = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties'
    });

    const sheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === sheetName);
    const sheetId = sheet.properties.sheetId;

    // Updated Test Scenarios Sheet Data with new format
    const testScenariosHeaderRow = [
        'Module',
        'Condition', 
        'Test Scenarios',
        'Status'
    ];

    const testScenariosDataRows = testScenarios.map(scenario => [
        scenario.module || module,
        scenario.condition || '',
        scenario.testScenarios || '',
        scenario.status || 'Not Tested'
    ]);

    const testScenariosAllData = [testScenariosHeaderRow, ...testScenariosDataRows];

    // Insert data into test scenarios sheet
    await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${sheetName}'!A1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: testScenariosAllData }
    });

    // Updated formatting for Test Scenarios Sheet with 4 columns
    const testScenariosFormatRequests = [
        // Format header row - same style as test cases
        {
            repeatCell: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: 4  // Updated to 4 columns
                },
                cell: {
                    userEnteredFormat: {
                        backgroundColor: {
                            red: 0.047,
                            green: 0.204,
                            blue: 0.239
                        },
                        textFormat: {
                            foregroundColor: {
                                red: 1.0,
                                green: 1.0,
                                blue: 1.0
                            },
                            bold: true,
                            fontSize: 11
                        },
                        horizontalAlignment: 'CENTER',
                        verticalAlignment: 'MIDDLE',
                        padding: {
                            top: 8,
                            bottom: 8,
                            left: 4,
                            right: 4
                        }
                    }
                },
                fields: 'userEnteredFormat'
            }
        },
        // Updated column widths for 4 columns
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 0,
                    endIndex: 1
                },
                properties: { pixelSize: 150 },  // Module
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 1,
                    endIndex: 2
                },
                properties: { pixelSize: 250 },  // Condition
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 2,
                    endIndex: 3
                },
                properties: { pixelSize: 400 },  // Test Scenarios
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 3,
                    endIndex: 4
                },
                properties: { pixelSize: 100 },  // Status
                fields: 'pixelSize'
            }
        },
        // Set row heights for scenarios
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: 0,
                    endIndex: 1
                },
                properties: { pixelSize: 35 },
                fields: 'pixelSize'
            }
        },
        {
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: 1,
                    endIndex: testScenarios.length + 1
                },
                properties: { pixelSize: 60 },
                fields: 'pixelSize'
            }
        },
        // Format data rows for scenarios with spacing
        {
            repeatCell: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: 1,
                    endRowIndex: testScenarios.length + 1,
                    startColumnIndex: 0,
                    endColumnIndex: 4  // Updated to 4 columns
                },
                cell: {
                    userEnteredFormat: {
                        textFormat: {
                            fontSize: 10
                        },
                        verticalAlignment: 'TOP',
                        wrapStrategy: 'WRAP',
                        padding: {
                            top: 6,
                            bottom: 6,
                            left: 6,
                            right: 6
                        }
                    }
                },
                fields: 'userEnteredFormat'
            }
        },
        // Add conditional formatting for Status column
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testScenarios.length + 1,
                        startColumnIndex: 3,
                        endColumnIndex: 4
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Pass' }] },
                        format: {
                            backgroundColor: { red: 0.0, green: 0.5, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 } }
                        }
                    }
                },
                index: 0
            }
        },
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testScenarios.length + 1,
                        startColumnIndex: 3,
                        endColumnIndex: 4
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Fail' }] },
                        format: {
                            backgroundColor: { red: 0.6, green: 0.0, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 } }
                        }
                    }
                },
                index: 1
            }
        },
        {
            addConditionalFormatRule: {
                rule: {
                    ranges: [{
                        sheetId: sheetId,
                        startRowIndex: 1,
                        endRowIndex: testScenarios.length + 1,
                        startColumnIndex: 3,
                        endColumnIndex: 4
                    }],
                    booleanRule: {
                        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Blocked' }] },
                        format: {
                            backgroundColor: { red: 1.0, green: 1.0, blue: 0.0 },
                            textFormat: { foregroundColor: { red: 0.0, green: 0.0, blue: 0.0 } }
                        }
                    }
                },
                index: 2
            }
        }
    ];

    // Apply all formatting for scenarios
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: testScenariosFormatRequests }
    });

    // Add data validation for Status column
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests: [
                {
                    setDataValidation: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: 1,
                            endRowIndex: testScenarios.length + 1,
                            startColumnIndex: 3,
                            endColumnIndex: 4
                        },
                        rule: {
                            condition: {
                                type: 'ONE_OF_LIST',
                                values: [
                                    { userEnteredValue: 'Not Tested' },
                                    { userEnteredValue: 'Pass' },
                                    { userEnteredValue: 'Fail' },
                                    { userEnteredValue: 'Blocked' }
                                ]
                            },
                            showCustomUi: true,
                            strict: true
                        }
                    }
                }
            ]
        }
    });
}