import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import mongoose from 'mongoose';
import cookieParser from 'cookie-parser';
import authRoutes from './src/routes/authRoutes.js';
import sheetRoutes from "./src/routes/sheetRoutes.js";
import { connectDBWithRetry } from './src/config/db.js';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

const corsOptions = {
    origin: function (origin, callback) {
        if (!origin) return callback(null, true);
        
        const allowedOrigins = [
            'https://test-gen-frontend.onrender.com',
            'http://localhost:3000',
            'http://192.168.0.193:4173',
            'http://localhost:3001',
            'http://localhost:5173',
            'http://localhost:4173',
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

// Health check endpoint
app.get('/health', (req, res) => {
    const dbState = mongoose.connection.readyState;
    const dbStatus = {
        0: 'disconnected',
        1: 'connected',
        2: 'connecting',
        3: 'disconnecting'
    };
    
    res.json({ 
        status: 'OK', 
        database: dbStatus[dbState],
        timestamp: new Date().toISOString() 
    });
});

// Debug endpoint
app.get('/debug/cookies', (req, res) => {
    res.cookie('test-cookie', 'test-value', {
        httpOnly: false,
        secure: false,
        sameSite: 'lax',
        maxAge: 60000
    });
    
    res.json({
        message: 'Debug endpoint',
        cookies: req.cookies,
        origin: req.get('origin')
    });
});

// Routes
app.use('/auth', authRoutes);
app.use('/api/sheets', sheetRoutes);

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Global error:', error);
    
    if (error.name === 'MongooseError' || error.name === 'MongoError') {
        return res.status(503).json({ 
            message: 'Database temporarily unavailable',
            code: 'DB_ERROR'
        });
    }
    
    res.status(500).json({ 
        message: 'Internal server error',
        code: 'SERVER_ERROR'
    });
});

// Start server and connect to database
const startServer = async () => {
    try {
        // Connect to MongoDB first
        await connectDBWithRetry();
        
        // Start server after successful DB connection
        app.listen(PORT, () => {
            console.log(`🚀 Server running on PORT: ${PORT}`);
            console.log(`🌍 Environment: ${process.env.NODE_ENV || 'development'}`);
        });
        
    } catch (error) {
        console.error('💥 Failed to start server:', error);
        process.exit(1);
    }
};

// Graceful shutdown
process.on('SIGINT', async () => {
    console.log('📴 Shutting down gracefully...');
    try {
        await mongoose.connection.close();
        console.log('✅ MongoDB connection closed');
        process.exit(0);
    } catch (error) {
        console.error('❌ Error during shutdown:', error);
        process.exit(1);
    }
});

startServer();
