import { Router, Request, Response } from 'express';

const router = Router();

router.get('/devices', async (req: Request, res: Response) => {
  try {
    const graphClient = req.graphClient;
    if (!graphClient) {
      return res.status(401).json({ error: 'Not authenticated' });
    }

    const devices = await graphClient
      .api('/deviceManagement/managedDevices')
      .select('id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName,model,manufacturer,serialNumber')
      .get();
    res.json(devices);
  } catch (error) {
    console.error('Error fetching devices:', error);
    res.status(500).json({ error: 'Failed to fetch devices' });
  }
});

router.get('/device/:id', async (req: Request, res: Response) => {
  try {
    const graphClient = req.graphClient;
    if (!graphClient) {
      return res.status(401).json({ error: 'Not authenticated' });
    }

    const deviceId = req.params.id;
    const device = await graphClient
      .api(`/deviceManagement/managedDevices/${deviceId}`)
      .select('id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName,model,manufacturer,serialNumber,userId,userDisplayName')
      .get();
    res.json(device);
  } catch (error) {
    console.error('Error fetching device:', error);
    res.status(500).json({ error: 'Failed to fetch device' });
  }
});

router.get('/configurations', async (req: Request, res: Response) => {
  try {
    const graphClient = req.graphClient;
    if (!graphClient) {
      return res.status(401).json({ error: 'Not authenticated' });
    }

    const configs = await graphClient
      .api('/deviceManagement/deviceConfigurations')
      .select('id,displayName,description,lastModifiedDateTime,createdDateTime')
      .get();
    res.json(configs);
  } catch (error) {
    console.error('Error fetching configurations:', error);
    res.status(500).json({ error: 'Failed to fetch configurations' });
  }
});

router.get('/device/:id/graph', async (req: Request, res: Response) => {
  try {
    const graphClient = req.graphClient;
    if (!graphClient) {
      return res.status(401).json({ error: 'Not authenticated' });
    }

    const deviceId = req.params.id;
    
    // Fetch device and its relationships
    const [device, configurationProfiles, compliancePolicies] = await Promise.all([
      graphClient.api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,userPrincipalName,model,manufacturer,serialNumber,userId,userDisplayName')
        .get(),
      graphClient.api(`/deviceManagement/managedDevices/${deviceId}/deviceConfigurationStates`)
        .get(),
      graphClient.api(`/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates`)
        .get()
    ]);

    // Create a map to track used IDs
    const usedIds = new Set<string>();
    
    // Transform data into graph format
    const nodes = [
      { 
        id: 'device', 
        label: device.deviceName,
        group: 'device',
        title: `OS: ${device.operatingSystem} ${device.osVersion}\nStatus: ${device.complianceState}\nModel: ${device.model}\nSerial: ${device.serialNumber}`,
        color: '#ff7f0e' 
      }
    ];
    usedIds.add('device');

    const edges = [];

    // Add configuration profiles with unique IDs
    if (configurationProfiles.value) {
      configurationProfiles.value.forEach((config: any, index: number) => {
        // Create a unique ID for this config
        let configId = `config_${config.id || index}`;
        
        // If this ID is already used, make it unique by adding an index
        if (usedIds.has(configId)) {
          let counter = 1;
          while (usedIds.has(`${configId}_${counter}`)) {
            counter++;
          }
          configId = `${configId}_${counter}`;
        }
        
        // Track this ID as used
        usedIds.add(configId);
        
        nodes.push({
          id: configId,
          label: config.displayName || 'Configuration Profile',
          group: 'config',
          title: `State: ${config.state}\nLastReported: ${config.lastReportedDateTime}`,
          color: '#2ca02c'
        });
        edges.push({
          from: 'device',
          to: configId,
          label: config.state || 'Applied'
        });
      });
    }

    // Add compliance policies with unique IDs
    if (compliancePolicies.value) {
      compliancePolicies.value.forEach((policy: any, index: number) => {
        // Create a unique ID for this policy
        let policyId = `policy_${policy.id || index}`;
        
        // If this ID is already used, make it unique by adding an index
        if (usedIds.has(policyId)) {
          let counter = 1;
          while (usedIds.has(`${policyId}_${counter}`)) {
            counter++;
          }
          policyId = `${policyId}_${counter}`;
        }
        
        // Track this ID as used
        usedIds.add(policyId);
        
        nodes.push({
          id: policyId,
          label: policy.displayName || 'Compliance Policy',
          group: 'policy',
          title: `State: ${policy.state}\nLastReported: ${policy.lastReportedDateTime}`,
          color: '#1f77b4'
        });
        edges.push({
          from: 'device',
          to: policyId,
          label: policy.state || 'Applied'
        });
      });
    }

    // Add user information if available
    if (device.userId) {
      const userId = 'user';
      
      if (!usedIds.has(userId)) {
        usedIds.add(userId);
        
        nodes.push({
          id: userId,
          label: device.userPrincipalName || 'User',
          group: 'user',
          title: device.userDisplayName || 'Device User',
          color: '#d62728'
        });
        edges.push({
          from: 'device',
          to: userId,
          label: 'Primary User'
        });
      }
    }

    res.json({ 
      nodes, 
      edges,
      groups: {
        device: { color: '#ff7f0e', shape: 'box' },
        config: { color: '#2ca02c', shape: 'dot' },
        policy: { color: '#1f77b4', shape: 'diamond' },
        user: { color: '#d62728', shape: 'triangle' }
      }
    });
  } catch (error) {
    console.error('Error fetching device graph:', error);
    res.status(500).json({ error: 'Failed to fetch device graph' });
  }
});

export default router; 