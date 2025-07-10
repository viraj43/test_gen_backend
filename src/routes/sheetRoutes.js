import express from 'express';
import {     
    generateTestCasesWithOptions,
    getTestCases,
    analyzeTestCases,
    modifyTestCases,
    processCustomPrompt} from '../controller/sheetsController.js';
import { ProtectRoute } from '../middlewares/authMiddleware.js';
import {
    getAuthUrl,
    handleCallback,
    getConnectionStatus,
    getAvailableSheets,
    removeCredentials
} from '../controller/oauthController.js';

const router = express.Router();

router.get('/auth-url', ProtectRoute, getAuthUrl);
router.get('/callback', handleCallback);
router.get('/status', ProtectRoute, getConnectionStatus);
router.delete('/disconnect', ProtectRoute, removeCredentials);

router.post('/generate', ProtectRoute, generateTestCasesWithOptions);
router.get('/list', ProtectRoute, getAvailableSheets);
router.get('/test-cases', ProtectRoute, getTestCases);
router.post('/analyze', ProtectRoute, analyzeTestCases);
router.post('/modify', ProtectRoute, modifyTestCases);

// NEW: Custom Prompt Route for Workflow Arrangement
router.post('/custom-prompt', ProtectRoute, processCustomPrompt);

export default router;

// Add to your main app.js 