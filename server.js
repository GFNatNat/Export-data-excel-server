import express from 'express';
import dotenv from 'dotenv';
import cors from 'cors';
import companyRoutes from './routes/companyRoutes.js';
import connectDB from './config/db.js';

dotenv.config();
const app = express();

app.use(cors());
app.use(express.json());
app.use('/api/companies', companyRoutes);

connectDB();

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));