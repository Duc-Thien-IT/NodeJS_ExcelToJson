import express from 'express';
import multer from 'multer';
import { handleFirstSheet, handleAllSheets } from '../controller/excelController';

const router = express.Router();

// Cấu hình Multer để xử lý file upload
const upload = multer({ dest: 'uploads/' });

/**
 * @swagger
 * /api/excel/upload-first-sheet:
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
router.post('/upload-first-sheet', upload.single('file'), handleFirstSheet);

/**
 * @swagger
 * /api/excel/upload-all-sheets:
 *   post:
 *     summary: Upload an Excel file and convert all sheets to JSON
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
router.post('/upload-all-sheets', upload.single('file'), handleAllSheets);

export default router;
