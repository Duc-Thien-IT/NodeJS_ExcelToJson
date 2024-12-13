"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const multer_1 = __importDefault(require("multer"));
const excelController_1 = require("../controller/excelController");
//===================================================
const router = express_1.default.Router();
// Cấu hình Multer để xử lý file upload
const upload = (0, multer_1.default)({ dest: 'uploads/' });
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
router.post('/upload-first-sheet', upload.single('file'), excelController_1.handleFirstSheet);
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
router.post('/upload-all-sheets', upload.single('file'), excelController_1.handleAllSheets);
exports.default = router;
