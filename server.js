import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import mongoose from 'mongoose';
import cookieParser from 'cookie-parser';
import authRoutes from './src/routes/authRoutes.js'
import sheetRoutes from "./src/routes/sheetRoutes.js"
import { connectDB } from './src/config/db.js';

dotenv.config();
const app = express();
const PORT = process.env.PORT || 3000;

const corsOptions = {
    origin: function (origin, callback) {
        // Allow requests with no origin (like mobile apps or curl requests)
        if (!origin) return callback(null, true);
        
        const allowedOrigins = [
            'https://test-gen-frontend.onrender.com',
            'http://localhost:3000',
            'http://localhost:3001',
            'http://localhost:5173', // Common Vite dev server port
            'http://localhost:5174'
        ];
        
        if (allowedOrigins.includes(origin)) {
            callback(null, true);
        } else {
            callback(new Error('Not allowed by CORS'));
        }
    },
    credentials: true, // This is crucial for cookies
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
    exposedHeaders: ['Set-Cookie'], // Expose Set-Cookie header
    optionsSuccessStatus: 200, // Some legacy browsers choke on 204
};

// Apply CORS before other middleware
app.use(cors(corsOptions));

// Handle preflight requests explicitly
app.options('*', cors(corsOptions));

// Middleware Configuration
app.use(cookieParser()); // Ensure cookie parsing middleware is set
app.use(express.json());  // Parse JSON request bodies

 // Use CORS with the provided 

app.use("/auth/",authRoutes)
app.use('/api/sheets', sheetRoutes);
app.use('/api', sheetRoutes);
 // Connect to MongoDB
app.listen(PORT, () => {
    console.log("server is running on PORT:" + PORT);
    connectDB();
});