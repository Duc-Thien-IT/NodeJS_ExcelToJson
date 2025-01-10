import express from 'express';
import swaggerDocs from './swagger';
import SwaggerUI from 'swagger-ui-express';

import excelRoutes from './routes/excelRoutes';
import jsonRoutes from './routes/jsonRoutes';
//=====================================================================
const app = express();
const port = 3000;

app.use(express.json());

// Cấu hình Swagger
app.use('/api-docs', SwaggerUI.serve, SwaggerUI.setup(swaggerDocs));

// Đăng ký routes
app.use('/api/excel', excelRoutes);
app.use('/api/json', jsonRoutes);

//======== Test frontend and notification port running================
app.get('/', (req, res) => {
  res.send('Hello World from Express and TypeScript!');
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
  console.log(`Swagger docs available at http://localhost:${port}/api-docs`);
});
