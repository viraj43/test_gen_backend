// authController.js - Enhanced with debugging
import User from "../models/User.js";
import bcrypt from "bcryptjs";
import { generateToken } from "../lib/utils.js";
import jwt from "jsonwebtoken";

export const signup = async (req, res) => {
    const { email, username, password } = req.body;
    
    console.log('📝 Signup attempt:', { email, username });
    console.log('🌐 Request origin:', req.get('origin'));
    console.log('📋 Request headers:', req.headers);
    
    if (!email || !username || !password) {
        return res.status(400).json({ message: "All fields are required" });
    }
    
    try {
        if (password.length < 8) {
            return res.status(400).json({ message: "Password must be at least 8 characters long" });
        }
        
        const user = await User.findOne({ email });
        if (user) {
            return res.status(400).json({ message: "User already exists" });
        }
        
        const salt = await bcrypt.genSalt(10);
        const hashedPassword = await bcrypt.hash(password, salt);

        const newUser = new User({
            email,
            username,
            password: hashedPassword
        });
        
        if (newUser) {
            await newUser.save();
            console.log('✅ User created successfully:', newUser._id);
            
            // Generate token and set cookie
            generateToken(newUser._id, res);
            
            console.log('🍪 Cookie should be set now');
            console.log('📤 Response headers:', res.getHeaders());
            
            return res.status(201).json({
                _id: newUser._id,
                email: newUser.email,
                username: newUser.username,
            });
        } else {
            return res.status(400).json({ message: "User could not be created" });
        }

    } catch (error) {
        console.error('❌ Signup error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const login = async (req, res) => {
    const { email, password } = req.body;
    
    console.log('🔐 Login attempt:', { email });
    console.log('🌐 Request origin:', req.get('origin'));
    console.log('📋 Request headers:', req.headers);
    console.log('🍪 Incoming cookies:', req.cookies);
    
    if (!email || !password) {
        return res.status(400).json({ message: "All fields are required" });
    }

    try {
        const user = await User.findOne({ email });
        if (!user) {
            console.log('❌ User not found:', email);
            return res.status(400).json({ message: "Invalid email or password" });
        }
        
        const isValidPassword = await bcrypt.compare(password, user.password);
        if (!isValidPassword) {
            console.log('❌ Invalid password for:', email);
            return res.status(400).json({ message: "Invalid email or password" });
        }
        
        console.log('✅ Login successful for:', email);
        
        // Generate token and set cookie
        generateToken(user._id, res);
        
        console.log('🍪 Cookie should be set now for user:', user._id);
        console.log('📤 Response headers being sent:', res.getHeaders());
        
        return res.status(200).json({
            _id: user._id,
            email: user.email,
            username: user.username,
        });
    } catch (error) {
        console.error('❌ Login error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const logout = async (req, res) => {
    console.log('🚪 Logout attempt');
    console.log('🌐 Request origin:', req.get('origin'));
    console.log('🍪 Incoming cookies:', req.cookies);
    
    try {
        res.cookie("jwt", "", { 
            maxAge: 0,
            httpOnly: true,
            sameSite: "strict",
            secure: process.env.NODE_ENV !== 'development'
        });
        
        console.log('✅ Logout cookie cleared');
        console.log('📤 Response headers:', res.getHeaders());
        
        return res.status(200).json({ message: "Logged out successfully" });
    } catch (error) {
        console.error('❌ Logout error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const getMe = async (req, res) => {
    console.log('👤 GetMe request');
    console.log('🌐 Request origin:', req.get('origin'));
    console.log('🍪 Incoming cookies:', req.cookies);
    console.log('📋 All headers:', req.headers);
    
    try {
        const token = req.cookies.jwt;
        
        if (!token) {
            console.log('❌ No JWT token found in cookies');
            return res.status(401).json({ message: "Not authenticated" });
        }

        console.log('🔑 JWT token found, verifying...');
        const decoded = jwt.verify(token, process.env.JWT_SECRET);
        console.log('✅ JWT decoded successfully:', decoded.userId);
        
        const user = await User.findById(decoded.userId).select('-password');
    
        if (!user) {
            console.log('❌ User not found in database:', decoded.userId);
            return res.status(401).json({ message: "User not found" });
        }

        console.log('✅ User found:', user.email);
        
        return res.status(200).json({
            _id: user._id,
            email: user.email,
            username: user.username,
        });
    } catch (error) {
        console.error('❌ GetMe error:', error);
        if (error.name === 'JsonWebTokenError') {
            console.log('🔑 Invalid JWT token');
        } else if (error.name === 'TokenExpiredError') {
            console.log('⏰ JWT token expired');
        }
        return res.status(401).json({ message: "Invalid token" });
    }
};