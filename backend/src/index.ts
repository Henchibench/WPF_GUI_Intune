import express from 'express';
import cors from 'cors';
import authRoutes from './routes/auth';
import intuneRoutes from './routes/intune';
import { graphClientMiddleware } from './middleware/graphClient';

const app = express();
const port = process.env.PORT || 4000;

// Middleware
app.use(cors());
app.use(express.json());

// Routes
app.use('/api/auth', authRoutes);
app.use('/api/intune', graphClientMiddleware, intuneRoutes);

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
}); 