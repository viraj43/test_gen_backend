import User from "../models/User.js";
import bcrypt from "bcryptjs";
import { generateToken } from "../lib/utils.js";
import jwt from "jsonwebtoken";

export const signup = async (req, res) => {
    const { email, username, password } = req.body;
    
    console.log('Signup attempt for:', email); // Debug log
    
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
            generateToken(newUser._id, res); // FIXED: _id instead of *id
            
            console.log('User created successfully:', newUser._id); // Debug log
            
            return res.status(201).json({
                _id: newUser._id, // FIXED: _id instead of *id
                email: newUser.email,
                username: newUser.username,
            });
        } else {
            return res.status(400).json({ message: "User could not be created" });
        }
    } catch (error) {
        console.log('Signup error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const login = async (req, res) => {
    const { email, password } = req.body;
    
    console.log('Login attempt for:', email); // Debug log
    console.log('Request origin:', req.get('origin')); // Debug log
    
    if (!email || !password) {
        return res.status(400).json({ message: "All fields are required" });
    }
    
    try {
        const user = await User.findOne({ email });
        if (!user) {
            return res.status(400).json({ message: "Invalid email or password" });
        }
        
        const isValidPassword = await bcrypt.compare(password, user.password);
        if (!isValidPassword) {
            return res.status(400).json({ message: "Invalid email or password" });
        }
        
        generateToken(user._id, res);
        
        console.log('Login successful for user:', user._id); // Debug log
        console.log('Setting cookie for origin:', req.get('origin')); // Debug log
        
        return res.status(200).json({
            _id: user._id, // FIXED: _id instead of *id
            email: user.email,
            username: user.username,
        });
    } catch (error) {
        console.log('Login error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const logout = async (req, res) => {
    try {
        res.cookie("jwt", "", { 
            maxAge: 0,
            httpOnly: true,
            sameSite: "strict",
            secure: process.env.NODE_ENV !== 'development'
        });
        
        console.log('User logged out successfully'); // Debug log
        
        return res.status(200).json({ message: "Logged out successfully" });
    } catch (error) {
        console.log('Logout error:', error);
        return res.status(500).json({ message: "Internal Server Error" });
    }
};

export const getMe = async (req, res) => {
    try {
        const token = req.cookies.jwt;
        
        console.log('GetMe request - token present:', !!token); // Debug log
        console.log('All cookies:', req.cookies); // Debug log
        
        if (!token) {
            return res.status(401).json({ message: "Not authenticated" });
        }
        
        const decoded = jwt.verify(token, process.env.JWT_SECRET);
        const user = await User.findById(decoded.userId).select('-password');
    
        if (!user) {
            return res.status(401).json({ message: "User not found" });
        }
        
        console.log('GetMe successful for user:', user._id); // Debug log
        
        return res.status(200).json({
            _id: user._id, // FIXED: _id instead of *id
            email: user.email,
            username: user.username,
        });
    } catch (error) {
        console.log("Get me error:", error);
        return res.status(401).json({ message: "Invalid token" });
    }
};
