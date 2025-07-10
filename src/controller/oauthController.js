import { google } from 'googleapis';
import UserToken from '../models/UserToken.js';
import dotenv from "dotenv";
dotenv.config();

const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
);

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

// Helper function to get authenticated sheets client
export const getAuthenticatedSheetsClient = async (userId) => {
    const userToken = await UserToken.findOne({ userId: userId });
    if (!userToken) {
        throw new Error('Google Sheets not connected');
    }

    const tokens = {
        access_token: userToken.accessToken,
        refresh_token: userToken.refreshToken,
        expiry_date: userToken.expiryDate.getTime()
    };

    oauth2Client.setCredentials(tokens);
    return google.sheets({ version: 'v4', auth: oauth2Client });
};

export const removeCredentials = async (req, res) => {
    try {
        const userId = req.user._id.toString();

        console.log("🗑️ Removing credentials for user:", userId);

        // Check if user has stored tokens
        const userToken = await UserToken.findOne({ userId: userId });

        if (!userToken) {
            console.log("❌ No tokens found for user:", userId);
            return res.json({
                success: false,
                message: 'No Google Sheets connection found to remove',
                connected: false
            });
        }

        // Optional: Revoke the token with Google (recommended for security)
        try {
            const tokens = {
                access_token: userToken.accessToken,
                refresh_token: userToken.refreshToken,
                expiry_date: userToken.expiryDate.getTime()
            };

            oauth2Client.setCredentials(tokens);

            // Revoke the refresh token to fully disconnect
            if (userToken.refreshToken) {
                await oauth2Client.revokeToken(userToken.refreshToken);
                console.log("🔐 Token revoked with Google");
            }
        } catch (revokeError) {
            // Don't fail the whole operation if revocation fails
            console.warn("⚠️ Failed to revoke token with Google:", revokeError.message);
        }

        // Remove tokens from database
        await UserToken.findOneAndDelete({ userId: userId });

        console.log("✅ Credentials removed successfully for user:", userId);

        res.json({
            success: true,
            message: 'Google Sheets connection removed successfully',
            connected: false
        });

    } catch (error) {
        console.error('Error removing credentials:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to remove Google Sheets connection',
            error: error.message
        });
    }
};

// Export oauth2Client for use in other controllers
export { oauth2Client };