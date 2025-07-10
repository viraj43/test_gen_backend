import { GoogleGenerativeAI } from '@google/generative-ai';
import dotenv from "dotenv";
dotenv.config();

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Analyze the intent behind the custom prompt
export async function analyzePromptIntent(prompt, testCases) {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    // FIXED: Use optimized prompt instead of the original heavy one
    const intentPrompt = `
Analyze test case arrangement request:

User Prompt: "${prompt}"
Test Cases: ${testCases.length} total
Sample IDs: ${testCases.slice(0, 10).map(tc => tc.id).join(', ')}

Return JSON with intent analysis including:
- intent: workflow|priority|dependency|risk|module|execution|custom
- strategy: arrangement approach
- confidence: 0-1 score
- arrangementCriteria: sorting/grouping details
- summary: what will be done

Return ONLY valid JSON.
`;

    try {
        const result = await model.generateContent(intentPrompt);
        const response = await result.response;
        const intentText = response.text();

        const cleanedText = intentText.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
        return JSON.parse(cleanedText);
    } catch (error) {
        console.error('Error analyzing intent:', error);
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
export async function applyIntelligentArrangement(testCases, prompt, intentAnalysis) {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const arrangementPrompt = `
Arrange ${testCases.length} test cases per user request.

User Request: "${prompt}"
Strategy: ${intentAnalysis.strategy}

Test Cases (IDs and summaries only):
${testCases.map(tc => `${tc.id}: ${tc.summary}`).join('\n')}

Return JSON with:
- arrangedTestCases: reordered array (keep ALL original fields)
- arrangementLogic: explanation
- groupings: logical groups formed

Apply strategy: ${intentAnalysis.strategy}
Keep ALL original test case data intact, only change order.

Return ONLY valid JSON.
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
export async function updateSpreadsheetWithArrangement(sheets, spreadsheetId, sheetName, arrangedTestCases, headerRow) {
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

// ENHANCED: More comprehensive existing context extraction
export function createEnhancedTestCasesContext(existingTestCases, sheetName) {
    if (!existingTestCases || existingTestCases.length === 0) {
        return "";
    }

    const lastTestCase = existingTestCases[existingTestCases.length - 1];
    const modules = [...new Set(existingTestCases.map(tc => tc.module))];
    const submodules = [...new Set(existingTestCases.map(tc => tc.submodule))];
    const typeDistribution = {};

    existingTestCases.forEach(tc => {
        typeDistribution[tc.testCaseType] = (typeDistribution[tc.testCaseType] || 0) + 1;
    });

    // ENHANCED: Extract key phrases from existing summaries to avoid similar content
    const existingSummaryKeywords = extractKeyPhrasesFromSummaries(existingTestCases);

    // ENHANCED: Extract test patterns to avoid duplication
    const testPatterns = extractTestPatterns(existingTestCases);

    return `
Existing: ${existingTestCases.length} test cases in "${sheetName}"
Last ID: ${lastTestCase.id}
Modules: ${modules.join(', ')}
Submodules: ${submodules.join(', ')}
Types: ${Object.entries(typeDistribution).map(([k, v]) => `${k}: ${v}`).join(', ')}

EXISTING SUMMARY KEYWORDS TO AVOID:
${existingSummaryKeywords}

EXISTING TEST PATTERNS TO AVOID:
${testPatterns}

ALL EXISTING SUMMARIES (for reference):
${existingTestCases.map(tc => `- ${tc.id}: "${tc.summary}"`).join('\n')}

CRITICAL RULES:
1. NO duplicates with existing summaries
2. NO similar test purposes or approaches
3. Each test must have unique objective
4. Different validation approaches required
5. Avoid repeating existing keywords/phrases`;
}

// HELPER: Extract key phrases from existing test case summaries
function extractKeyPhrasesFromSummaries(testCases) {
    const keyPhrases = new Set();

    testCases.forEach(tc => {
        const summary = tc.summary.toLowerCase();

        // Extract meaningful phrases (2-4 words)
        const words = summary.split(/\s+/).filter(word => word.length > 2);

        // Extract 2-word phrases
        for (let i = 0; i < words.length - 1; i++) {
            keyPhrases.add(`${words[i]} ${words[i + 1]}`);
        }

        // Extract 3-word phrases
        for (let i = 0; i < words.length - 2; i++) {
            keyPhrases.add(`${words[i]} ${words[i + 1]} ${words[i + 2]}`);
        }
    });

    return Array.from(keyPhrases).slice(0, 20).join(', ');
}

// HELPER: Extract common test patterns
function extractTestPatterns(testCases) {
    const patterns = new Set();

    testCases.forEach(tc => {
        const steps = tc.testSteps.toLowerCase();

        // Extract action patterns
        if (steps.includes('login') || steps.includes('sign in')) patterns.add('login_flow');
        if (steps.includes('validate') || steps.includes('verify')) patterns.add('validation_pattern');
        if (steps.includes('create') || steps.includes('add')) patterns.add('creation_pattern');
        if (steps.includes('delete') || steps.includes('remove')) patterns.add('deletion_pattern');
        if (steps.includes('update') || steps.includes('edit')) patterns.add('update_pattern');
        if (steps.includes('search') || steps.includes('filter')) patterns.add('search_pattern');
        if (steps.includes('upload') || steps.includes('file')) patterns.add('file_upload_pattern');
        if (steps.includes('download') || steps.includes('export')) patterns.add('download_pattern');
        if (steps.includes('permission') || steps.includes('access')) patterns.add('permission_pattern');
        if (steps.includes('notification') || steps.includes('alert')) patterns.add('notification_pattern');
    });

    return Array.from(patterns).join(', ');
}

// UPDATED: Keep existing function name but use enhanced version
export function createCompactTestCasesContext(existingTestCases, sheetName) {
    return createEnhancedTestCasesContext(existingTestCases, sheetName);
}

// ENHANCED: Improved prompt with stronger duplicate prevention
export function createAntiDuplicatePrompt(module, summary, acceptanceCriteria, testCasesCount, nextIdNumber, existingContext) {
    return `
Generate ${testCasesCount} COMPLETELY UNIQUE test cases for: ${module}

Summary: ${summary}
Acceptance Criteria: ${acceptanceCriteria}

${existingContext}

ANTI-DUPLICATE REQUIREMENTS:
1. Each test case must have a UNIQUE purpose and approach
2. NO overlapping test objectives with existing cases
3. Different validation techniques for each test
4. Vary the user personas, data sets, and scenarios
5. Use different edge cases and boundary conditions
6. Different error scenarios and recovery paths

DIVERSITY GUIDELINES:
- Vary input data types (valid, invalid, boundary, special characters)
- Different user roles/permissions for each test
- Various system states (empty, full, partial data)
- Different browsers/devices if applicable
- Different time scenarios (peak hours, off-hours)
- Various network conditions if applicable

COVERAGE DISTRIBUTION (${testCasesCount} unique tests):
- Happy Path Variations: ${Math.ceil(testCasesCount * 0.35)} tests
- Input Validation (different types): ${Math.ceil(testCasesCount * 0.25)} tests  
- Error Handling (various scenarios): ${Math.ceil(testCasesCount * 0.20)} tests
- Edge Cases & Boundaries: ${Math.ceil(testCasesCount * 0.15)} tests
- Integration & Workflow: ${Math.ceil(testCasesCount * 0.05)} tests

UNIQUENESS STRATEGIES:
- Use different submodules for similar functionality
- Vary the complexity levels (simple vs complex scenarios)
- Different combinations of features
- Various user journey stages
- Different data volumes and types

JSON Format (STRICTLY follow):
[
  {
    "id": "PC_${nextIdNumber}",
    "module": "${module}",
    "submodule": "UNIQUE_COMPONENT_NAME",
    "summary": "UNIQUE test description with specific objective",
    "testSteps": "Step 1. Specific action\\nStep 2. Unique verification\\nStep 3. Distinct validation",
    "expectedResults": "Specific expected outcome",
    "testCaseType": "Positive|Negative",
    "environment": "Test",
    "status": "Not Tested"
  }
]

CRITICAL: Every test case must be completely different in approach, data, and validation method.
Return ONLY the JSON array with ${testCasesCount} absolutely unique test cases.
`;
}

// ENHANCED: Update the main generation function to use these improvements
export function updateGenerateTestCasesPrompt(module, summary, acceptanceCriteria, testCasesCount, existingTestCasesContext, nextIdNumber) {
    // Use the enhanced context and anti-duplicate prompt
    return createAntiDuplicatePrompt(
        module, 
        summary, 
        acceptanceCriteria, 
        testCasesCount, 
        nextIdNumber, 
        existingTestCasesContext
    );
}

// ENHANCED: Improved validation with stronger duplicate detection
export function validateAndCleanTestCasesEnhanced(testCases, existingTestCases = []) {
    console.log(`🔍 Starting enhanced validation for ${testCases.length} generated test cases`);
    
    // Combine new and existing test cases for comprehensive duplicate checking
    const allExistingCases = existingTestCases || [];
    
    const cleaned = testCases.map((tc, index) => {
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

    // ENHANCED: Multi-level duplicate detection
    const uniqueCases = [];
    const seenSummaries = new Set();
    const seenStepsSignatures = new Set();
    const seenPurposes = new Set();
    let duplicateCount = 0;

    // Add existing test cases to seen sets
    allExistingCases.forEach(existingTc => {
        seenSummaries.add(existingTc.summary.toLowerCase().trim());
        const stepsSignature = createStepsSignature(existingTc.testSteps);
        seenStepsSignatures.add(stepsSignature);
        const purpose = extractTestPurpose(existingTc.summary);
        seenPurposes.add(purpose);
    });

    cleaned.forEach((tc, index) => {
        const summaryKey = tc.summary.toLowerCase().trim();
        const stepsSignature = createStepsSignature(tc.testSteps);
        const testPurpose = extractTestPurpose(tc.summary);

        // Check 1: Exact summary duplicate
        if (seenSummaries.has(summaryKey)) {
            console.warn(`🚨 Duplicate summary removed: ${tc.summary}`);
            duplicateCount++;
            return;
        }

        // Check 2: Similar test steps
        let hasSimilarSteps = false;
        for (const existingSteps of seenStepsSignatures) {
            if (calculateAdvancedSimilarity(stepsSignature, existingSteps) > 0.75) {
                console.warn(`🚨 Similar test steps removed: ${tc.summary}`);
                duplicateCount++;
                hasSimilarSteps = true;
                break;
            }
        }
        if (hasSimilarSteps) return;

        // Check 3: Similar test purpose/objective
        let hasSimilarPurpose = false;
        for (const existingPurpose of seenPurposes) {
            if (calculatePurposeSimilarity(testPurpose, existingPurpose) > 0.70) {
                console.warn(`🚨 Similar test purpose removed: ${tc.summary}`);
                duplicateCount++;
                hasSimilarPurpose = true;
                break;
            }
        }
        if (hasSimilarPurpose) return;

        // Check 4: Validate uniqueness within the new batch
        let isDuplicateInBatch = false;
        for (const uniqueCase of uniqueCases) {
            if (calculateAdvancedSimilarity(stepsSignature, createStepsSignature(uniqueCase.testSteps)) > 0.70 ||
                calculatePurposeSimilarity(testPurpose, extractTestPurpose(uniqueCase.summary)) > 0.70) {
                console.warn(`🚨 Duplicate within batch removed: ${tc.summary}`);
                duplicateCount++;
                isDuplicateInBatch = true;
                break;
            }
        }
        if (isDuplicateInBatch) return;

        // If passes all checks, add to unique cases
        seenSummaries.add(summaryKey);
        seenStepsSignatures.add(stepsSignature);
        seenPurposes.add(testPurpose);
        uniqueCases.push(tc);
    });

    if (duplicateCount > 0) {
        console.log(`✅ Removed ${duplicateCount} duplicate/similar test cases`);
        console.log(`📊 Final unique test cases: ${uniqueCases.length}`);
    }

    return uniqueCases;
}

// HELPER: Create a signature for test steps
function createStepsSignature(testSteps) {
    return testSteps
        .toLowerCase()
        .replace(/step \d+\./g, '') // Remove step numbers
        .replace(/\d+\./g, '') // Remove numbered lists
        .replace(/[^\w\s]/g, ' ') // Remove special characters
        .replace(/\s+/g, ' ') // Normalize whitespace
        .trim();
}

// HELPER: Extract the main purpose/objective of a test case
function extractTestPurpose(summary) {
    const purposeKeywords = [
        'verify', 'validate', 'test', 'check', 'ensure', 'confirm',
        'login', 'create', 'update', 'delete', 'search', 'upload',
        'download', 'submit', 'cancel', 'save', 'edit', 'view'
    ];
    
    const words = summary.toLowerCase().split(/\s+/);
    const purposeWords = words.filter(word => 
        purposeKeywords.includes(word) || word.length > 4
    );
    
    return purposeWords.slice(0, 5).join(' ');
}

// HELPER: Advanced similarity calculation
function calculateAdvancedSimilarity(str1, str2) {
    const words1 = str1.split(' ').filter(word => word.length > 3);
    const words2 = str2.split(' ').filter(word => word.length > 3);

    if (words1.length === 0 || words2.length === 0) return 0;

    const set1 = new Set(words1);
    const set2 = new Set(words2);

    const intersection = new Set([...set1].filter(x => set2.has(x)));
    const union = new Set([...set1, ...set2]);

    // Jaccard similarity
    const jaccardSimilarity = intersection.size / union.size;
    
    // Add penalty for common action words appearing in same order
    const actionWords = ['login', 'create', 'update', 'delete', 'verify', 'validate'];
    let actionPenalty = 0;
    
    actionWords.forEach(action => {
        if (str1.includes(action) && str2.includes(action)) {
            actionPenalty += 0.1;
        }
    });

    return Math.min(jaccardSimilarity + actionPenalty, 1);
}

// HELPER: Calculate purpose similarity
function calculatePurposeSimilarity(purpose1, purpose2) {
    if (!purpose1 || !purpose2) return 0;
    
    const words1 = purpose1.split(' ');
    const words2 = purpose2.split(' ');
    
    const commonWords = words1.filter(word => words2.includes(word));
    const totalWords = new Set([...words1, ...words2]).size;
    
    return commonWords.length / totalWords;
}

// Replace parseGeminiJSON validation call with enhanced version
export function parseGeminiJSONEnhanced(text, existingTestCases = []) {
    console.log("🔍 Starting enhanced JSON parsing...");
    
    // Use existing parsing logic first
    const parsed = parseGeminiJSON(text);
    
    // Then apply enhanced validation
    return performPostGenerationValidation(parsed, existingTestCases);
}

// ENHANCED: Post-generation validation
export function performPostGenerationValidation(generatedTestCases, existingTestCases = []) {
    console.log("🔍 Performing post-generation validation...");
    
    // Use enhanced validation
    const validatedCases = validateAndCleanTestCasesEnhanced(generatedTestCases, existingTestCases);
    
    // Additional checks
    if (validatedCases.length < generatedTestCases.length * 0.7) {
        console.warn(`⚠️ High duplicate rate detected. Only ${validatedCases.length} unique cases from ${generatedTestCases.length} generated.`);
    }
    
    // Check for variety in submodules
    const submodules = [...new Set(validatedCases.map(tc => tc.submodule))];
    if (submodules.length < 3) {
        console.warn(`⚠️ Low submodule variety. Consider more diverse test scenarios.`);
    }
    
    // Check test type distribution
    const positiveCount = validatedCases.filter(tc => tc.testCaseType === 'Positive').length;
    const negativeCount = validatedCases.filter(tc => tc.testCaseType === 'Negative').length;
    
    if (positiveCount === 0 || negativeCount === 0) {
        console.warn(`⚠️ Unbalanced test types: ${positiveCount} positive, ${negativeCount} negative`);
    }
    
    return validatedCases;
}

export function createCompactTestScenariosContext(existingTestScenarios, sheetName) {
    if (!existingTestScenarios || existingTestScenarios.length === 0) {
        return "";
    }

    const conditions = [...new Set(existingTestScenarios.map(ts => ts.condition))];
    const recentScenarios = existingTestScenarios.slice(-5).map(ts =>
        `- ${ts.condition}: ${ts.testScenarios.substring(0, 80)}...`
    ).join('\n');

    return `
Existing: ${existingTestScenarios.length} scenarios in "${sheetName}"
Conditions: ${conditions.slice(0, 10).join(', ')}${conditions.length > 10 ? '...' : ''}

Recent scenarios:
${recentScenarios}

CRITICAL: Create NEW scenarios with different conditions!`;
}

// Robust JSON parser for Gemini responses
export function parseGeminiJSON(text) {
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

export function parseTestScenariosJSON(text) {
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

export function validateAndCleanTestCases(testCases) {
    const cleaned = testCases.map((tc, index) => {
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

    // Enhanced duplicate detection
    const uniqueCases = [];
    const seenSummaries = new Set();
    const seenStepsSignatures = new Set();
    let duplicateCount = 0;

    cleaned.forEach(tc => {
        const summaryKey = tc.summary.toLowerCase().trim();
        const stepsSignature = tc.testSteps.toLowerCase().replace(/\s+/g, ' ').trim();

        // Check for duplicate summary
        if (seenSummaries.has(summaryKey)) {
            console.warn(`🚨 Duplicate summary removed: ${tc.summary}`);
            duplicateCount++;
            return;
        }

        // Check for very similar test steps (basic similarity check)
        let isSimilar = false;
        for (const existingSteps of seenStepsSignatures) {
            if (calculateSimpleSimilarity(stepsSignature, existingSteps) > 0.85) {
                console.warn(`🚨 Similar test steps removed: ${tc.summary}`);
                duplicateCount++;
                isSimilar = true;
                break;
            }
        }

        if (!isSimilar) {
            seenSummaries.add(summaryKey);
            seenStepsSignatures.add(stepsSignature);
            uniqueCases.push(tc);
        }
    });

    if (duplicateCount > 0) {
        console.log(`✅ Removed ${duplicateCount} duplicate/similar test cases`);
    }

    return uniqueCases;
}

export function validateAndCleanTestScenarios(scenarios) {
    return scenarios.map((scenario, index) => ({
        id: scenario.id || `TS_${(index + 1)}`,
        module: scenario.module || '',
        condition: scenario.condition || '',
        testScenarios: scenario.testScenarios || scenario.description || `Test Scenario ${index + 1} Description`,
        status: scenario.status || 'Not Tested'
    })).filter(scenario => scenario.testScenarios);
}

// Helper function for simple similarity calculation
function calculateSimpleSimilarity(str1, str2) {
    const words1 = str1.split(' ').filter(word => word.length > 3); // Only significant words
    const words2 = str2.split(' ').filter(word => word.length > 3);

    const set1 = new Set(words1);
    const set2 = new Set(words2);

    const intersection = new Set([...set1].filter(x => set2.has(x)));
    const union = new Set([...set1, ...set2]);

    return intersection.size / union.size;
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

export async function callGeminiWithRetry(model, prompt, maxRetries = 3, baseDelay = 1000) {
    let lastError;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`🤖 Gemini API call attempt ${attempt}/${maxRetries}`);

            const result = await model.generateContent(prompt);
            const response = await result.response;
            const text = response.text();

            console.log(`✅ Gemini API call successful on attempt ${attempt}`);
            return text;

        } catch (error) {
            lastError = error;
            console.error(`❌ Gemini API call failed on attempt ${attempt}:`, error.message);

            // Check if it's a rate limit or overload error
            if (error.message.includes('503') ||
                error.message.includes('overloaded') ||
                error.message.includes('rate limit') ||
                error.message.includes('429')) {

                if (attempt < maxRetries) {
                    // Exponential backoff with jitter
                    const delay = baseDelay * Math.pow(2, attempt - 1) + Math.random() * 1000;
                    console.log(`⏳ Waiting ${Math.round(delay)}ms before retry...`);
                    await new Promise(resolve => setTimeout(resolve, delay));
                    continue;
                }
            }

            // For other errors, don't retry
            if (!error.message.includes('503') &&
                !error.message.includes('overloaded') &&
                !error.message.includes('rate limit') &&
                !error.message.includes('429')) {
                throw error;
            }
        }
    }

    // If all retries failed, throw the last error
    throw lastError;
}

// NEW: Function to append test cases to existing sheet
export async function appendTestCasesToExistingSheet(sheets, spreadsheetId, sheetName, testCases, module) {
    try {
        // Get existing data to find the next available row
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:J`,
        });

        const existingRows = existingData.data.values || [];
        const nextRow = existingRows.length + 1;

        console.log(`📍 Appending ${testCases.length} test cases starting at row ${nextRow}`);

        // Get sheet ID for formatting
        const spreadsheetInfo = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const sheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        const sheetId = sheet.properties.sheetId;

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

        // NOW ADD FORMATTING FOR THE NEW ROWS
        const startRowIndex = nextRow - 1; // Convert to 0-based index
        const endRowIndex = nextRow - 1 + testCases.length;

        console.log(`🎨 Applying formatting to rows ${startRowIndex + 1} to ${endRowIndex}`);

        const formatRequests = [
            // Format the new data rows with proper spacing and alignment
            {
                repeatCell: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: startRowIndex,
                        endRowIndex: endRowIndex,
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
            // Set row heights for better spacing
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'ROWS',
                        startIndex: startRowIndex,
                        endIndex: endRowIndex
                    },
                    properties: { pixelSize: 80 },
                    fields: 'pixelSize'
                }
            }
        ];

        // Apply basic formatting first
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: formatRequests }
        });

        // Add conditional formatting for Test Case Type column (column H = index 7)
        const conditionalFormatRequests = [
            {
                addConditionalFormatRule: {
                    rule: {
                        ranges: [{
                            sheetId: sheetId,
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
            // Status conditional formatting (column J = index 9)
            {
                addConditionalFormatRule: {
                    rule: {
                        ranges: [{
                            sheetId: sheetId,
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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

        // Apply conditional formatting
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: conditionalFormatRequests }
        });

        // Add data validation for the new rows
        const dataValidationRequests = [
            // Test Case Type validation (column H = index 7)
            {
                setDataValidation: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: startRowIndex,
                        endRowIndex: endRowIndex,
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
            // Status validation (column J = index 9)
            {
                setDataValidation: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: startRowIndex,
                        endRowIndex: endRowIndex,
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
        ];

        // Apply data validation
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: dataValidationRequests }
        });

        console.log("✅ Test cases appended successfully with full formatting");
        return true;
    } catch (error) {
        console.error('Error appending test cases:', error);
        throw error;
    }
}

// NEW: Function to append test scenarios to existing sheet
export async function appendTestScenariosToExistingSheet(sheets, spreadsheetId, sheetName, testScenarios, module) {
    try {
        // Get existing data to find the next available row
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:D`,
        });

        const existingRows = existingData.data.values || [];
        const nextRow = existingRows.length + 1;

        console.log(`📍 Appending ${testScenarios.length} test scenarios starting at row ${nextRow}`);

        // Get sheet ID for formatting
        const spreadsheetInfo = await sheets.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const sheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }
        const sheetId = sheet.properties.sheetId;

        // Prepare new test scenario data (without header)
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

        // ADD FORMATTING FOR THE NEW ROWS
        const startRowIndex = nextRow - 1; // Convert to 0-based index
        const endRowIndex = nextRow - 1 + testScenarios.length;

        console.log(`🎨 Applying formatting to scenario rows ${startRowIndex + 1} to ${endRowIndex}`);

        const formatRequests = [
            // Format the new data rows with proper spacing
            {
                repeatCell: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: startRowIndex,
                        endRowIndex: endRowIndex,
                        startColumnIndex: 0,
                        endColumnIndex: 4
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
            // Set row heights for scenarios
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'ROWS',
                        startIndex: startRowIndex,
                        endIndex: endRowIndex
                    },
                    properties: { pixelSize: 60 },
                    fields: 'pixelSize'
                }
            }
        ];

        // Apply basic formatting
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: formatRequests }
        });

        // Add conditional formatting for Status column (column D = index 3)
        const conditionalFormatRequests = [
            {
                addConditionalFormatRule: {
                    rule: {
                        ranges: [{
                            sheetId: sheetId,
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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
                            startRowIndex: startRowIndex,
                            endRowIndex: endRowIndex,
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

        // Apply conditional formatting
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: conditionalFormatRequests }
        });

        // Add data validation for Status column
        const dataValidationRequests = [
            {
                setDataValidation: {
                    range: {
                        sheetId: sheetId,
                        startRowIndex: startRowIndex,
                        endRowIndex: endRowIndex,
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
        ];

        // Apply data validation
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: dataValidationRequests }
        });

        console.log("✅ Test scenarios appended successfully with full formatting");
        return true;
    } catch (error) {
        console.error('Error appending test scenarios:', error);
        throw error;
    }
}


// Helper function to add test cases sheet data
export async function addTestCasesSheetData(sheets, spreadsheetId, sheetName, testCases, module) {
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

    // Apply formatting to Test Cases Sheet with corrected conditional formatting
    const testCasesFormatRequests = [
        // Format header row
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
        // Test Case Type conditional formatting
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
        // Status conditional formatting
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
export async function addTestScenariosSheetData(sheets, spreadsheetId, sheetName, testScenarios, module) {
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