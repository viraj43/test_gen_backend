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
  origin: 'http://localhost:5173', // The URL of your frontend app
  credentials: true,  // Ensure cookies are allowed to be sent with requests
};

app.use(cors(corsOptions));
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