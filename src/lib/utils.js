// utils.js - Enhanced token generation with debugging
import jwt from 'jsonwebtoken';

export const generateToken = (userId, res) => {
    console.log('🏭 Generating token for user:', userId);
    
    if (!process.env.JWT_SECRET) {
        console.error('❌ JWT_SECRET not found in environment variables');
        throw new Error('JWT_SECRET is required');
    }
    
    const token = jwt.sign({ userId }, process.env.JWT_SECRET, {
        expiresIn: '7d'
    });

    const isProduction = process.env.NODE_ENV === 'production';
    
    const cookieOptions = {
        maxAge: 7 * 24 * 60 * 60 * 1000, // 7 days
        httpOnly: true,
        secure: isProduction, // Only secure in production
        sameSite: isProduction ? 'none' : 'lax', // 'none' for cross-site in production
        path: '/'
    };

    console.log('🍪 Setting cookie with options:', cookieOptions);
    console.log('🔧 Environment:', process.env.NODE_ENV);
    console.log('🔒 Is production:', isProduction);
    
    res.cookie('jwt', token, cookieOptions);
    
    console.log('✅ Cookie set successfully');
    
    return token;
};