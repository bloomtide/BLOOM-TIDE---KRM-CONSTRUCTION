import dotenv from 'dotenv';
import connectDB from '../config/db.js';
import User from '../models/User.js';

// Load environment variables
dotenv.config();

const createAdmin = async () => {
    try {
        // Connect to database
        await connectDB();

        // Admin credentials
        const adminData = {
            name: 'KRM Admin',
            email: 'admin@krm.com',
            password: 'admin123456', // Change this password!
            role: 'admin',
            isActive: true,
        };

        // Check if admin already exists
        let admin = await User.findOne({ email: adminData.email });

        if (admin) {
            console.log('⚠️  Admin user already exists! Resetting password...');
            admin.password = adminData.password;
            await admin.save();
            console.log('✅ Admin password reset successfully!');
            console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
            console.log('📧 Email:', adminData.email);
            console.log('🔑 Password:', adminData.password);
            console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
            process.exit(0);
        }

        // Create admin user
        const newAdmin = await User.create(adminData);

        console.log('✅ Admin user created successfully!');
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        console.log('📧 Email:', adminData.email);
        console.log('🔑 Password:', adminData.password);
        console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        console.log('⚠️  IMPORTANT: Change this password after first login!');

        process.exit(0);
    } catch (error) {
        console.error('❌ Error creating admin:', error.message);
        process.exit(1);
    }
};

createAdmin();
