import mongoose from 'mongoose';

export const connectDB = async () => {
    try {
        // Check if MONGODB_URI exists
        if (!process.env.MONGODB_URI) {
            console.error('❌ MONGODB_URI environment variable is not set');
            process.exit(1);
        }

        console.log('🔄 Attempting to connect to MongoDB...');
        console.log('🔗 Connection string format:', process.env.MONGODB_URI.replace(/\/\/.*@/, '//***:***@'));

        // Fixed connection options - removed deprecated options
        const conn = await mongoose.connect(process.env.MONGODB_URI, {
            // Connection pool options
            maxPoolSize: 10, // Maintain up to 10 socket connections

            // Timeout options
            serverSelectionTimeoutMS: 10000, // Keep trying to send operations for 10 seconds
            socketTimeoutMS: 45000, // Close sockets after 45 seconds of inactivity
            connectTimeoutMS: 10000, // Give up initial connection after 10 seconds

            // Idle connection options
            maxIdleTimeMS: 30000, // Close connections after 30 seconds of inactivity

            // Network options
            family: 4, // Use IPv4, skip trying IPv6

            // Write concern
            w: 'majority',

            // Read preference
            readPreference: 'primary'
        });

        console.log(`✅ MongoDB Connected: ${conn.connection.host}`);
        console.log(`📊 Database Name: ${conn.connection.name}`);

        // Handle connection events
        mongoose.connection.on('error', (err) => {
            console.error('❌ MongoDB connection error:', err);
        });

        mongoose.connection.on('disconnected', () => {
            console.warn('⚠️ MongoDB disconnected');
        });

        mongoose.connection.on('reconnected', () => {
            console.log('🔄 MongoDB reconnected');
        });

    } catch (error) {
        console.error('❌ MongoDB connection failed:', error);

        // Log specific error types
        if (error.name === 'MongoNetworkError') {
            console.error('🌐 Network error - check your internet connection and MongoDB URI');
        } else if (error.name === 'MongoServerSelectionError') {
            console.error('🖥️ Server selection error - MongoDB server might be down or unreachable');
        } else if (error.name === 'MongoParseError') {
            console.error('📝 Connection string parse error - check your MONGODB_URI format');
        }

        throw error; // Re-throw to be caught by retry logic
    }
};

export const connectDBWithRetry = async (retries = 5) => {
    for (let i = 0; i < retries; i++) {
        try {
            await connectDB();
            return; // Success, exit retry loop
        } catch (error) {
            console.log(`❌ Connection attempt ${i + 1} failed:`, error.message);

            if (i === retries - 1) {
                console.error('💥 All connection attempts failed');
                process.exit(1);
            }

            // Wait before retrying (exponential backoff)
            const delay = Math.pow(2, i) * 1000; // 1s, 2s, 4s, 8s, 16s
            console.log(`⏳ Retrying in ${delay / 1000} seconds...`);
            await new Promise(resolve => setTimeout(resolve, delay));
        }
    }
};

// Alternative minimal connection for quick testing
export const connectDBMinimal = async () => {
    try {
        if (!process.env.MONGODB_URI) {
            console.error('❌ MONGODB_URI environment variable is not set');
            process.exit(1);
        }

        console.log('🔄 Attempting minimal MongoDB connection...');

        // Minimal connection options - only the essentials
        const conn = await mongoose.connect(process.env.MONGODB_URI);

        console.log(`✅ MongoDB Connected (minimal): ${conn.connection.host}`);
        console.log(`📊 Database Name: ${conn.connection.name}`);

    } catch (error) {
        console.error('❌ Minimal MongoDB connection failed:', error);
        throw error;
    }
};

