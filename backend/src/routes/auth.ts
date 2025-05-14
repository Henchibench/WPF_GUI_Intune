import { Router, Request, Response } from 'express';
import { DeviceCodeCredential, DeviceCodeInfo as AzureDeviceCodeInfo } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { setGlobalGraphClient } from '../middleware/graphClient';

const router = Router();

// Store credential globally (not ideal but works for demo)
let globalCredential: DeviceCodeCredential | null = null;

// Extend Request type to include graphClient
interface AuthenticatedRequest extends Request {
  graphClient?: Client;
}

interface DeviceCodeInfo {
  userCode: string;
  verificationUri: string;
  expiresOn?: Date;
  message?: string;
}

router.post('/device-code', async (req: Request, res: Response) => {
  try {
    console.log('Initiating device code flow...');
    
    // Create a promise that will resolve when we get the device code info
    const deviceCodePromise = new Promise<DeviceCodeInfo>((resolve, reject) => {
      let isResolved = false;

      const credential = new DeviceCodeCredential({
        clientId: '14d82eec-204b-4c2f-b7e8-296a70dab67e', // Default PowerShell client ID
        tenantId: 'common',
        userPromptCallback: (info: AzureDeviceCodeInfo) => {
          console.log('Received device code info:', info);
          if (!isResolved) {
            isResolved = true;
            resolve(info as unknown as DeviceCodeInfo);
          }
        },
      });

      globalCredential = credential;

      // Initialize Graph client
      const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default']
      });

      const graphClient = Client.initWithMiddleware({
        authProvider: authProvider
      });

      // Set the global graph client
      setGlobalGraphClient(graphClient);

      // Trigger the device code flow by attempting to get a token
      credential.getToken(['https://graph.microsoft.com/.default']).catch((error) => {
        if (!isResolved) {
          reject(error);
        }
      });

      // Set a timeout in case we don't get the callback
      setTimeout(() => {
        if (!isResolved) {
          reject(new Error('Timeout waiting for device code. Please try again.'));
        }
      }, 60000); // 60 seconds timeout
    });

    // Wait for the device code info
    const deviceCodeInfo = await deviceCodePromise;

    if (!deviceCodeInfo || !deviceCodeInfo.userCode || !deviceCodeInfo.verificationUri) {
      throw new Error('Invalid device code information received');
    }

    res.json({
      message: 'Device code flow initiated',
      userCode: deviceCodeInfo.userCode,
      verificationUri: deviceCodeInfo.verificationUri
    });
  } catch (error) {
    console.error('Error in device code flow:', error);
    if (error instanceof Error) {
      res.status(500).json({ 
        error: 'Failed to initiate device code flow',
        details: error.message,
        stack: error.stack
      });
    } else {
      res.status(500).json({ 
        error: 'Failed to initiate device code flow',
        details: 'Unknown error occurred'
      });
    }
  }
});

router.get('/status', async (req: AuthenticatedRequest, res: Response) => {
  try {
    if (!globalCredential) {
      return res.json({ authenticated: false });
    }

    // Try to get a token silently to verify authentication
    try {
      const token = await globalCredential.getToken(['https://graph.microsoft.com/.default']);
      if (token) {
        // Test the connection by making a simple Graph API call
        const graphClient = req.graphClient;
        if (graphClient) {
          await graphClient.api('/me').get();
        }
        return res.json({ authenticated: true });
      }
    } catch (error) {
      console.log('Token acquisition failed:', error);
      return res.json({ authenticated: false });
    }

    res.json({ authenticated: false });
  } catch (error) {
    console.error('Error checking auth status:', error);
    res.json({ authenticated: false });
  }
});

export default router; 