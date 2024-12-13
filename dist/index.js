"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const swagger_1 = __importDefault(require("./swagger"));
const swagger_ui_express_1 = __importDefault(require("swagger-ui-express"));
const excelRoutes_1 = __importDefault(require("./routes/excelRoutes"));
//=====================================================================
const app = (0, express_1.default)();
const port = 3000;
app.use(express_1.default.json());
// Cấu hình Swagger
app.use('/api-docs', swagger_ui_express_1.default.serve, swagger_ui_express_1.default.setup(swagger_1.default));
// Đăng ký routes
app.use('/api/excel', excelRoutes_1.default);
//======== Test frontend and notification port running================
app.get('/', (req, res) => {
    res.send('Hello World from Express and TypeScript!');
});
app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
    console.log(`Swagger docs available at http://localhost:${port}/api-docs`);
});
