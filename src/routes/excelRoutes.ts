import express from 'express';
import multer from 'multer';
import { handleFirstSheet, 
          handleAllSheets, 
          handleSheetByName, 
          handleFirstSheetWithDataOnly, 
          handleJsonToExcel ,
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

/**
 * @swagger
 * /api/excel/convert-file-by-name:
 *   post:
 *      summary: Convert file sheet to Json by name
 *      tags:
 *         - Excel
 *      requestBody:
 *          content:
 *              multipart/form-data:
 *                  schema:
 *                      type: object
 *                      properties:
 *                          file:
 *                              type: string
 *                              format: binary
 *                              description: The Excel to upload
 *                          name:
 *                              type: string
 *                              description: Fill name of sheet you want convert to Json
 *      responses:
 *          200:
 *              description: File processed successfully
 *          400:
 *              description: No file upload
 *          500:
 *              description: Error processing file
 */
router.post('/convert-file-by-name', upload.single('file'), handleSheetByName);

/**
 * @swagger
 * /api/excel/convert-file-with-data-only:
 *   post:
 *      summary: Convert first sheet to JSON with data only
 *      tags:
 *         - Excel
 *      requestBody:
 *          content:
 *              multipart/form-data:
 *                  schema:
 *                      type: object
 *                      properties:
 *                          file:
 *                              type: string
 *                              format: binary
 *                              description: The Excel file to upload
 *      responses:
 *          200:
 *              description: File processed successfully with data only
 *          400:
 *              description: No file uploaded
 *          500:
 *              description: Error processing file
 */
router.post('/convert-file-with-data-only', upload.single('file'), handleFirstSheetWithDataOnly);

/**
 * @swagger
 * /api/excel/convert-json-to-excel:
 *   post:
 *     summary: Convert JSON data to Excel file
 *     tags:
 *       - Excel
 *     requestBody:
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               jsonData:
 *                 type: object
 *                 description: JSON data to convert to Excel
 *     responses:
 *       200:
 *         description: Excel file processed and ready for download
 *       400:
 *         description: Invalid or missing JSON data
 *       500:
 *         description: Error processing file
 */
router.post('/convert-json-to-excel', handleJsonToExcel);

// /**
//  * @swagger
//  * /api/excel/json-to-excel-and-print:
//  *   post:
//  *     summary: Convert JSON data to an Excel file and print it
//  *     tags:
//  *       - Excel
//  *     requestBody:
//  *       content:
//  *         application/json:
//  *           schema:
//  *             type: object
//  *             properties:
//  *               jsonData:
//  *                 type: object
//  *                 description: JSON data to convert
//  *               printerName:
//  *                 type: string
//  *                 description: Name of the printer
//  *     responses:
//  *       200:
//  *         description: Excel file created and sent to printer successfully
//  *       400:
//  *         description: Invalid JSON data or printer name
//  *       500:
//  *         description: Error creating or printing Excel file
//  */
// router.post('/json-to-excel-and-print', handleJsonToExcelAndPrint);

export default router;
