
// debug-server.js - Create this file to test routes one by one
import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import cookieParser from 'cookie-parser';

dotenv.config();
const app = express();
const PORT = process.env.PORT || 3000;

const corsOptions = {
    origin: function (origin, callback) {
        if (!origin) return callback(null, true);
        
        const allowedOrigins = [
            'https://test-gen-frontend.onrender.com',
            'http://localhost:3000',
            'http://localhost:3001',
            'http://localhost:5173',
            'http://localhost:5174'
        ];
        
        if (allowedOrigins.includes(origin)) {
            callback(null, true);
        } else {
            callback(new Error('Not allowed by CORS'));
        }
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
    exposedHeaders: ['Set-Cookie'],
    optionsSuccessStatus: 200,
};

app.use(cors(corsOptions));
app.use(cookieParser());
app.use(express.json());

// Test basic route first
app.get('/test', (req, res) => {
    res.json({ message: 'Server is working' });
});

console.log('Basic server setup complete, testing auth routes...');

// Test 1: Try loading auth routes
try {
    const authRoutes = await import('./src/routes/authRoutes.js');
    console.log('✅ Auth routes imported successfully');
    app.use('/auth', authRoutes.default);
    console.log('✅ Auth routes mounted successfully');
} catch (error) {
    console.error('❌ Error with auth routes:', error.message);
}

// Test 2: Try loading sheet routes
try {
    const sheetRoutes = await import('./src/routes/sheetRoutes.js');
    console.log('✅ Sheet routes imported successfully');
    app.use('/api/sheets', sheetRoutes.default);
    console.log('✅ Sheet routes mounted successfully');
} catch (error) {
    console.error('❌ Error with sheet routes:', error.message);
    console.error('This is likely where the problem is!');
}

app.listen(PORT, () => {
    console.log("Debug server running on PORT:" + PORT);
});