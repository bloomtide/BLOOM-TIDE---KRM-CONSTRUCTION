import multer from 'multer';
import path from 'path';

// Configure multer for memory storage (we'll process the file immediately)
const storage = multer.memoryStorage();

// File filter to only accept Excel files
const fileFilter = (req, file, cb) => {
    const allowedExtensions = ['.xlsx', '.xls'];
    const ext = path.extname(file.originalname).toLowerCase();
    
    if (allowedExtensions.includes(ext)) {
        cb(null, true);
    } else {
        cb(new Error('Only Excel files (.xlsx, .xls) are allowed'), false);
    }
};

// Configure upload with limits
const upload = multer({
    storage: storage,
    fileFilter: fileFilter,
    limits: {
        fileSize: 20 * 1024 * 1024, // 20MB limit
    },
});

export default upload;
