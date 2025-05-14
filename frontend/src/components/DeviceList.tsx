import React, { useEffect, useState } from 'react';
import {
  List,
  ListItem,
  ListItemText,
  ListItemButton,
  Paper,
  Typography,
  CircularProgress,
  Box
} from '@mui/material';

const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:4000';

interface Device {
  id: string;
  deviceName: string;
  operatingSystem: string;
  complianceState: string;
}

interface DeviceListProps {
  onDeviceSelect: (deviceId: string) => void;
}

const DeviceList: React.FC<DeviceListProps> = ({ onDeviceSelect }) => {
  const [devices, setDevices] = useState<Device[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchDevices();
  }, []);

  const fetchDevices = async () => {
    try {
      const response = await fetch(`${API_URL}/api/intune/devices`);
      const data = await response.json();
      setDevices(data.value || []);
      setError(null);
    } catch (error) {
      console.error('Error fetching devices:', error);
      setError('Failed to fetch devices');
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return (
      <Box display="flex" justifyContent="center" p={3}>
        <CircularProgress />
      </Box>
    );
  }

  if (error) {
    return (
      <Typography color="error" p={2}>
        {error}
      </Typography>
    );
  }

  return (
    <Paper elevation={2}>
      <List>
        {devices.map((device) => (
          <ListItem key={device.id} disablePadding>
            <ListItemButton onClick={() => onDeviceSelect(device.id)}>
              <ListItemText
                primary={device.deviceName}
                secondary={
                  <>
                    {device.operatingSystem}
                    <br />
                    Status: {device.complianceState}
                  </>
                }
              />
            </ListItemButton>
          </ListItem>
        ))}
        {devices.length === 0 && (
          <ListItem>
            <ListItemText primary="No devices found" />
          </ListItem>
        )}
      </List>
    </Paper>
  );
};

export default DeviceList; 