import express from 'express';
import { convertJsonToExcel } from '../controller/jsonController';

const router = express.Router();

/**
 * @swagger
 * /api/excel/convert-json-to-excel:
 *   post:
 *     summary: Convert JSON data to Excel file
 *     tags:
 *       - Json
 *     requestBody:
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               jsonData:
 *                 type: object
 *                 description: JSON data to convert to Excel
 *               outputFile:
 *                 type: string
 *                 description: Name of the output Excel file (optional)
 *     responses:
 *       200:
 *         description: Excel file processed and ready for download
 *       400:
 *         description: Invalid or missing JSON data
 *       500:
 *         description: Error processing file
 */
router.post('/convert-json-to-excel', convertJsonToExcel);

export default router;
