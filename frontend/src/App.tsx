import React, { useState, useEffect } from 'react';
import { Container, AppBar, Toolbar, Typography, Button, Box, Paper, useMediaQuery, useTheme } from '@mui/material';
import DeviceList from './components/DeviceList';
import GraphVisualization from './components/GraphVisualization';
import AuthDialog from './components/AuthDialog';

const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:4000';

interface Group {
  color?: string;
  shape?: string;
}

interface GraphData {
  nodes: any[];
  edges: any[];
  groups?: Record<string, Group>;
}

function App() {
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down('md'));
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [showAuthDialog, setShowAuthDialog] = useState(false);
  const [selectedDevice, setSelectedDevice] = useState<string | null>(null);
  const [graphData, setGraphData] = useState<GraphData>({ nodes: [], edges: [] });

  useEffect(() => {
    checkAuthStatus();
  }, []);

  const checkAuthStatus = async () => {
    try {
      const response = await fetch(`${API_URL}/api/auth/status`);
      const data = await response.json();
      setIsAuthenticated(data.authenticated);
    } catch (error) {
      console.error('Error checking auth status:', error);
    }
  };

  const handleAuth = () => {
    if (isAuthenticated) {
      // Handle logout
      setIsAuthenticated(false);
      setSelectedDevice(null);
    } else {
      setShowAuthDialog(true);
    }
  };

  const handleDeviceSelect = async (deviceId: string) => {
    try {
      const response = await fetch(`${API_URL}/api/intune/device/${deviceId}/graph`);
      const data = await response.json();
      setGraphData(data);
      setSelectedDevice(deviceId);
    } catch (error) {
      console.error('Error fetching device graph:', error);
    }
  };

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      <AppBar position="static">
        <Toolbar>
          <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
            Intune Explorer
          </Typography>
          <Button color="inherit" onClick={handleAuth}>
            {isAuthenticated ? 'Disconnect' : 'Connect to Microsoft Graph'}
          </Button>
        </Toolbar>
      </AppBar>

      {isAuthenticated ? (
        <Box 
          sx={{ 
            flexGrow: 1, 
            display: 'flex', 
            flexDirection: isMobile ? 'column' : 'row', 
            p: 2,
            gap: 2,
            height: 'calc(100vh - 64px)',
            overflow: 'hidden'
          }}
        >
          <Paper 
            elevation={3} 
            sx={{ 
              width: isMobile ? '100%' : '25%', 
              height: isMobile ? '40%' : '100%',
              overflow: 'auto'
            }}
          >
            <DeviceList onDeviceSelect={handleDeviceSelect} />
          </Paper>
          
          <Paper 
            elevation={3} 
            sx={{ 
              width: isMobile ? '100%' : '75%', 
              height: isMobile ? '60%' : '100%',
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden'
            }}
          >
            {selectedDevice ? (
              <GraphVisualization
                nodes={graphData.nodes}
                edges={graphData.edges}
                groups={graphData.groups}
              />
            ) : (
              <Box 
                display="flex" 
                justifyContent="center" 
                alignItems="center" 
                height="100%" 
                p={3}
              >
                <Typography>
                  Select a device from the list to view its relationships
                </Typography>
              </Box>
            )}
          </Paper>
        </Box>
      ) : (
        <Box 
          sx={{ 
            flexGrow: 1,
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center'
          }}
        >
          <Paper elevation={3} sx={{ p: 4, maxWidth: 500, width: '100%' }}>
            <Typography variant="h5" sx={{ mb: 2, textAlign: 'center' }}>
              Welcome to Intune Explorer
            </Typography>
            <Typography variant="body1" sx={{ mb: 3, textAlign: 'center' }}>
              Connect to Microsoft Graph to view and analyze your Intune devices
            </Typography>
            <Box display="flex" justifyContent="center">
              <Button 
                variant="contained" 
                color="primary" 
                onClick={() => setShowAuthDialog(true)}
              >
                Connect
              </Button>
            </Box>
          </Paper>
        </Box>
      )}

      <AuthDialog
        open={showAuthDialog}
        onClose={() => setShowAuthDialog(false)}
        onAuthenticated={() => {
          setIsAuthenticated(true);
          setShowAuthDialog(false);
        }}
      />
    </Box>
  );
}

export default App; 