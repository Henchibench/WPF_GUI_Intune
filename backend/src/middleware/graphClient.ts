import { Request, Response, NextFunction } from 'express';
import { Client } from '@microsoft/microsoft-graph-client';

declare global {
  namespace Express {
    interface Request {
      graphClient?: Client;
    }
  }
}

// Store the graph client globally (not ideal but works for demo)
let globalGraphClient: Client | null = null;

export const setGlobalGraphClient = (client: Client) => {
  globalGraphClient = client;
};

export const graphClientMiddleware = (req: Request, res: Response, next: NextFunction) => {
  if (!globalGraphClient) {
    return res.status(401).json({ error: 'Not authenticated with Microsoft Graph' });
  }
  req.graphClient = globalGraphClient;
  next();
}; 