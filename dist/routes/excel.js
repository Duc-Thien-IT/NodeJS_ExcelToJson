"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const multer_1 = __importDefault(require("multer"));
const xlsx_1 = __importDefault(require("xlsx"));
const router = express_1.default.Router();
// Cấu hình Multer để xử lý file upload
const upload = (0, multer_1.default)({ dest: 'uploads/' });
// Hàm xử lý logic chính
const processExcelFile = (req, res, next) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        if (!req.file) {
            res.status(400).json({ message: 'No file uploaded' });
            return;
        }
        // Đọc file Excel
        const filePath = req.file.path;
        const workbook = xlsx_1.default.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Lấy sheet đầu tiên
        const sheet = workbook.Sheets[sheetName];
        // Chuyển đổi dữ liệu từ sheet thành JSON
        const jsonData = xlsx_1.default.utils.sheet_to_json(sheet);
        res.json({
            message: 'File processed successfully',
            data: jsonData,
        });
    }
    catch (error) {
        next(error); // Truyền lỗi cho middleware xử lý lỗi
    }
});
/**
 * @swagger
 * /api/excel/upload:
 *   post:
 *     summary: Upload an Excel file and convert it to JSON
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
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *                 data:
 *                   type: array
 *                   items:
 *                     type: object
 *       400:
 *         description: No file uploaded
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *       500:
 *         description: Error processing file
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *                 error:
 *                   type: object
 */
router.post('/upload', upload.single('file'), processExcelFile);
exports.default = router;
