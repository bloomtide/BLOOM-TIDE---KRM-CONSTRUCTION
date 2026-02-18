import dotenv from 'dotenv';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import connectDB from '../config/db.js';
import User from '../models/User.js';
import Proposal from '../models/Proposal.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const DUMP_PATH = join(__dirname, '../data/proposals-dump.json');
const ADMIN_EMAIL = 'admin@krm.com';

const seedProposals = async () => {
    try {
        await connectDB();

        // Find admin user (must exist; run npm run seed:admin first)
        const admin = await User.findOne({ email: ADMIN_EMAIL });
        if (!admin) {
            console.error('❌ No admin user found. Run: npm run seed:admin');
            process.exit(1);
        }

        const raw = readFileSync(DUMP_PATH, 'utf8');
        const items = JSON.parse(raw);
        if (!Array.isArray(items) || items.length === 0) {
            console.log('⚠️ No proposals in dump file.');
            process.exit(0);
        }

        const created = [];
        for (const item of items) {
            const { name, client, project, template, rawExcelData } = item;
            if (!name || !client || !project || !template || !rawExcelData) {
                console.warn('⚠️ Skipping invalid item (missing required fields):', item.name || item);
                continue;
            }
            const proposal = await Proposal.create({
                name,
                client,
                project,
                template,
                rawExcelData: {
                    fileName: rawExcelData.fileName || 'dump.xlsx',
                    sheetName: rawExcelData.sheetName || 'Sheet1',
                    headers: Array.isArray(rawExcelData.headers) ? rawExcelData.headers : [],
                    rows: Array.isArray(rawExcelData.rows) ? rawExcelData.rows : [],
                },
                createdBy: admin._id,
            });
            created.push(proposal.name);
        }

        console.log('✅ Proposals seeded successfully:', created.length);
        created.forEach((name, i) => console.log(`   ${i + 1}. ${name}`));
        process.exit(0);
    } catch (error) {
        console.error('❌ Seed failed:', error.message);
        process.exit(1);
    }
};

seedProposals();
