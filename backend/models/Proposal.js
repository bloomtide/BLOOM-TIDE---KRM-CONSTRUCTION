import mongoose from 'mongoose';

const proposalSchema = new mongoose.Schema(
    {
        name: {
            type: String,
            required: [true, 'Please provide a proposal name'],
            trim: true,
        },
        client: {
            type: String,
            required: [true, 'Please provide a client name'],
            trim: true,
        },
        project: {
            type: String,
            required: [true, 'Please provide a project name'],
            trim: true,
        },
        template: {
            type: String,
            required: [true, 'Please select a template'],
            enum: ['capstone', 'dbi', 'pd_steel', 'sperrin_tony', 'tristate_martin'],
            default: 'capstone',
        },
        rawExcelData: {
            fileName: {
                type: String,
                required: true,
            },
            sheetName: {
                type: String,
            },
            headers: [String],
            rows: [[mongoose.Schema.Types.Mixed]], // 2D array of cell values
        },
        spreadsheetJson: {
            type: mongoose.Schema.Types.Mixed, // Syncfusion saveAsJson() output
            default: null,
        },
        // Store unused raw data rows (rows not processed by any processor)
        unusedRawDataRows: [{
            rowIndex: {
                type: Number,
                required: true,
            },
            rowData: [mongoose.Schema.Types.Mixed], // Array of cell values for this row
            isUsed: {
                type: Boolean,
                default: false, // Client can mark as used after manual processing
            },
        }],
        // Store images separately since Syncfusion saveAsJson() doesn't include them
        images: [{
            sheetIndex: {
                type: Number,
                required: true,
            },
            imageId: {
                type: String,
                required: true,
            },
            src: {
                type: String, // Base64 data URL
                required: true,
            },
            top: {
                type: Number,
                default: 0,
            },
            left: {
                type: Number,
                default: 0,
            },
            width: {
                type: Number,
            },
            height: {
                type: Number,
            },
        }],
        createdBy: {
            type: mongoose.Schema.Types.ObjectId,
            ref: 'User',
            required: true,
        },
    },
    {
        timestamps: true,
    }
);

// Index for faster queries by user
proposalSchema.index({ createdBy: 1, createdAt: -1 });

const Proposal = mongoose.model('Proposal', proposalSchema);
export default Proposal;
