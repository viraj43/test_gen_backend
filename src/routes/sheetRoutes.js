import express from 'express';
import {  getAuthUrl, 
    handleCallback, 
    getConnectionStatus, 
    generateTestCasesWithOptions,
    getTestCases,
    analyzeTestCases,
    modifyTestCases,
    getAvailableSheets,
    processCustomPrompt} from '../controller/sheetsController.js';
import { ProtectRoute } from '../middlewares/authMiddleware.js';

const router = express.Router();

router.get('/auth-url', ProtectRoute, getAuthUrl);
router.get('/callback', handleCallback);
router.get('/status', ProtectRoute, getConnectionStatus);
router.post('/generate', ProtectRoute, generateTestCasesWithOptions);

router.get('/sheets', ProtectRoute, getAvailableSheets);
router.get('/test-cases', ProtectRoute, getTestCases);
router.post('/analyze', ProtectRoute, analyzeTestCases);
router.post('/modify', ProtectRoute, modifyTestCases);

// NEW: Custom Prompt Route for Workflow Arrangement
router.post('/custom-prompt', ProtectRoute, processCustomPrompt);

export default router;

// Add to your main app.js 