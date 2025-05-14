import React, { useState, useEffect, useCallback } from 'react';
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Typography,
  CircularProgress,
  Link,
  Box
} from '@mui/material';

const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:4000';

interface AuthDialogProps {
  open: boolean;
  onClose: () => void;
  onAuthenticated: () => void;
}

const AuthDialog: React.FC<AuthDialogProps> = ({ open, onClose, onAuthenticated }) => {
  const [deviceCode, setDeviceCode] = useState<string | null>(null);
  const [verificationUri, setVerificationUri] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [pollInterval, setPollInterval] = useState<NodeJS.Timeout | null>(null);
  const [isInitiating, setIsInitiating] = useState(false);

  const clearPolling = useCallback(() => {
    if (pollInterval) {
      clearInterval(pollInterval);
      setPollInterval(null);
    }
  }, [pollInterval]);

  useEffect(() => {
    return () => {
      clearPolling();
      setIsInitiating(false);
    };
  }, [clearPolling]);

  useEffect(() => {
    if (open && !deviceCode && !isInitiating && !error) {
      initiateDeviceCode();
    } else if (!open) {
      clearPolling();
      setDeviceCode(null);
      setVerificationUri(null);
      setError(null);
      setIsInitiating(false);
    }
  }, [open, deviceCode, isInitiating, error, clearPolling]);

  const initiateDeviceCode = async () => {
    if (isInitiating) return;

    setLoading(true);
    setError(null);
    clearPolling();
    setIsInitiating(true);

    try {
      console.log('Initiating device code flow...');
      const response = await fetch(`${API_URL}/api/auth/device-code`, {
        method: 'POST'
      });
      
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.details || 'Failed to initiate authentication');
      }

      const data = await response.json();
      console.log('Received device code response:', data);
      
      if (!data.userCode || !data.verificationUri) {
        throw new Error('Invalid response from server: missing device code information');
      }

      setDeviceCode(data.userCode);
      setVerificationUri(data.verificationUri);
      startPolling();
    } catch (error) {
      console.error('Error:', error);
      setError(error instanceof Error ? error.message : 'Failed to initiate authentication');
    } finally {
      setLoading(false);
      setIsInitiating(false);
    }
  };

  const startPolling = () => {
    const interval = setInterval(async () => {
      try {
        const response = await fetch(`${API_URL}/api/auth/status`);
        const data = await response.json();
        
        if (data.authenticated) {
          clearPolling();
          onAuthenticated();
        }
      } catch (error) {
        console.error('Error polling auth status:', error);
      }
    }, 5000);

    setPollInterval(interval);

    // Clean up interval after 15 minutes
    setTimeout(() => {
      clearPolling();
      setError('Authentication timeout. Please try again.');
    }, 15 * 60 * 1000);
  };

  const handleRetry = () => {
    setDeviceCode(null);
    setVerificationUri(null);
    setError(null);
    setIsInitiating(false);
    initiateDeviceCode();
  };

  return (
    <Dialog open={open} onClose={onClose} maxWidth="sm" fullWidth>
      <DialogTitle>Connect to Microsoft Graph</DialogTitle>
      <DialogContent>
        {loading ? (
          <Box display="flex" justifyContent="center" p={3}>
            <CircularProgress />
          </Box>
        ) : error ? (
          <Typography color="error">{error}</Typography>
        ) : deviceCode && verificationUri ? (
          <>
            <Typography variant="body1" paragraph>
              To sign in, use a web browser to open the page{' '}
              <Link href={verificationUri} target="_blank" rel="noopener">
                {verificationUri}
              </Link>{' '}
              and enter the code <strong>{deviceCode}</strong> to authenticate.
            </Typography>
            <Typography variant="body2" color="textSecondary">
              Waiting for authentication...
            </Typography>
          </>
        ) : null}
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Cancel</Button>
        {error && (
          <Button onClick={handleRetry} color="primary">
            Retry
          </Button>
        )}
      </DialogActions>
    </Dialog>
  );
};

export default AuthDialog; 