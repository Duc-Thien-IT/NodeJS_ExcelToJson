import express from 'express';
import multer from 'multer';
import { 
          convertExcelToJson
        } from '../controller/excelController';

const router = express.Router();

// Cấu hình Multer để xử lý file upload
const upload = multer({ dest: 'uploads/' });

/**
 * @swagger
 * /api/excel/convert:
 *   post:
 *     summary: Upload an Excel file and convert the first sheet to JSON
 *     tags:
 *       - Excel
 *     requestBody:
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               file:
 *                 type: string
 *                 format: binary
 *                 description: The Excel file to upload
 *     responses:
 *       200:
 *         description: File processed successfully
 *       400:
 *         description: No file uploaded
 *       500:
 *         description: Error processing file
 */
router.post('/convert', upload.single('file'), convertExcelToJson);

export default router;