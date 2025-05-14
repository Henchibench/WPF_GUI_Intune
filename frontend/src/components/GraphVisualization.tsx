// If you see type errors, run: npm install --save-dev @types/vis-network
import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Network, Node as VisNode, Edge as VisEdge, Options } from 'vis-network/standalone';
import { Box, Typography, Paper, IconButton, Divider, Tooltip, Switch, FormControlLabel, Chip, List, ListItem, ListItemText, Button } from '@mui/material';

interface Node {
  id: string;
  label: string;
  color?: string;
  group?: string;
  title?: string;
  level?: number;
  widthConstraint?: number;
  font?: any;
}

interface Edge {
  from: string;
  to: string;
  label?: string;
}

interface Group {
  color?: string;
  shape?: string;
}

interface GraphData {
  nodes: Node[];
  edges: Edge[];
  groups?: Record<string, Group>;
}

interface GraphProps extends GraphData {}

const GraphVisualization: React.FC<GraphProps> = ({ nodes, edges, groups }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const networkRef = useRef<Network | null>(null);
  const [hierarchicalLayout, setHierarchicalLayout] = useState(true);
  const [selectedNode, setSelectedNode] = useState<Node | null>(null);
  
  // Organize nodes by compliance state
  const organizeNodesByState = useCallback((originalNodes: Node[], originalEdges: Edge[]): [Node[], Edge[]] => {
    // Find the device node
    const deviceNode = originalNodes.find(node => node.id === 'device');
    if (!deviceNode) return [originalNodes, originalEdges];
    
    const newNodes: Node[] = [];
    const newEdges: Edge[] = [];
    
    // Add device node
    newNodes.push({ ...deviceNode, level: 0 });
    
    // Add user nodes directly connected to device
    const userNodes = originalNodes.filter(node => node.group === 'user');
    userNodes.forEach(userNode => {
      newNodes.push({ ...userNode, level: 1 });
      
      // Find edges connecting user to device
      const userEdges = originalEdges.filter(edge => 
        (edge.from === deviceNode.id && edge.to === userNode.id) ||
        (edge.to === deviceNode.id && edge.from === userNode.id)
      );
      
      newEdges.push(...userEdges);
    });
    
    // Group configuration nodes by compliance state
    const nodesByState = new Map<string, Node[]>();
    const stateColors = {
      'compliant': '#4CAF50',
      'noncompliant': '#F44336',
      'error': '#FF9800',
      'unknown': '#9E9E9E'
    };
    
    // Identify config and policy nodes
    const configNodes = originalNodes.filter(node => 
      node.group === 'config' || node.group === 'policy'
    );
    
    configNodes.forEach(node => {
      // Find the edge that connects this node to the device
      const edgeToDevice = originalEdges.find(edge => 
        (edge.from === deviceNode.id && edge.to === node.id) ||
        (edge.to === deviceNode.id && edge.from === node.id)
      );
      
      // Determine state from edge label or node properties
      let state = 'unknown';
      if (edgeToDevice && edgeToDevice.label) {
        state = edgeToDevice.label.toLowerCase();
      } else if (node.title && node.title.toLowerCase().includes('state:')) {
        const stateMatch = node.title.match(/state:\s*(\w+)/i);
        if (stateMatch) {
          state = stateMatch[1].toLowerCase();
        }
      }
      
      // Normalize states
      if (state.includes('compliant')) {
        state = state.includes('non') ? 'noncompliant' : 'compliant';
      }
      
      // Add node to appropriate state group
      if (!nodesByState.has(state)) {
        nodesByState.set(state, []);
      }
      nodesByState.get(state)?.push(node);
    });
    
    // Create state parent nodes and connect children
    nodesByState.forEach((stateNodes, state) => {
      if (stateNodes.length > 0) {
        // Create state parent node
        const stateNodeId = `state_${state}`;
        const stateNode: Node = {
          id: stateNodeId,
          label: `${state.charAt(0).toUpperCase() + state.slice(1)} (${stateNodes.length})`,
          group: 'state',
          level: 1,
          color: stateColors[state as keyof typeof stateColors] || '#9E9E9E',
          widthConstraint: 200
        };
        
        // Add state node
        newNodes.push(stateNode);
        
        // Connect device to state node
        newEdges.push({
          from: deviceNode.id,
          to: stateNodeId,
          label: state
        });
        
        // Add child nodes and connect to state parent
        stateNodes.forEach(configNode => {
          // Calculate width based on label length
          const labelLength = configNode.label.length;
          const width = Math.max(250, labelLength * 10); // 10px per character, minimum 250px width
          
          // Add config node with level 2 and ensure text doesn't get too small
          newNodes.push({ 
            ...configNode, 
            level: 2,
            // Add specific font settings to ensure readability
            font: {
              size: 14,
              face: 'Tahoma',
              multi: 'html'
            },
            // Add width constraint based on label length
            widthConstraint: width
          });
          
          // Connect state parent to config node
          newEdges.push({
            from: stateNodeId,
            to: configNode.id
          });
        });
      }
    });
    
    return [newNodes, newEdges];
  }, []);
  
  // Create network options based on current layout mode
  const createNetworkOptions = useCallback((hierarchical: boolean): Options => {
    return {
      nodes: { 
        shape: hierarchical ? 'box' : 'dot',
        size: hierarchical ? 25 : 15,
        font: {
          size: 14,
          face: 'Tahoma'
        },
        margin: { top: 20, right: 20, bottom: 20, left: 20 },
        widthConstraint: hierarchical ? { minimum: 200, maximum: 450 } : undefined
      },
      edges: { 
        arrows: 'to',
        smooth: {
          enabled: true,
          type: hierarchical ? 'cubicBezier' : 'dynamic',
          roundness: 0.5
        },
        font: {
          size: 12,
          align: 'middle',
          background: 'rgba(255, 255, 255, 0.7)'
        },
        color: '#666666',
        length: 300 // Set fixed edge length
      },
      physics: {
        enabled: !hierarchical,
        stabilization: {
          enabled: true,
          iterations: 1000
        }
      },
      layout: {
        hierarchical: {
          enabled: hierarchical,
          direction: 'UD',
          sortMethod: 'directed',
          nodeSpacing: 400, // Increased from 300 to 400
          levelSeparation: 350, // Increased from 250 to 350
          treeSpacing: 500, // Increased from 400 to 500
          blockShifting: true,
          edgeMinimization: true,
          parentCentralization: true
        }
      },
      groups: {
        ...groups,
        state: {
          shape: 'box',
          borderWidth: 2,
          shadow: true,
          font: {
            size: 16,
            color: '#000000',
            face: 'Tahoma',
            bold: true
          }
        }
      },
      interaction: {
        navigationButtons: true,
        keyboard: true,
        hover: true,
        selectable: true,
        multiselect: false,
        tooltipDelay: 300,
        zoomView: true
      }
    };
  }, [groups]);

  // Initialize or recreate the network
  const initializeNetwork = useCallback(() => {
    if (containerRef.current && nodes.length > 0) {
      try {
        // Destroy previous network if exists
        if (networkRef.current) {
          networkRef.current.destroy();
        }

        // Organize nodes by compliance state
        const [organizedNodes, organizedEdges] = organizeNodesByState(nodes, edges);

        // Create options based on current layout mode
        const options = createNetworkOptions(hierarchicalLayout);

        // Create network with organized data
        networkRef.current = new Network(
          containerRef.current, 
          { 
            nodes: organizedNodes as VisNode[], 
            edges: organizedEdges as VisEdge[] 
          }, 
          options
        );

        // Add event listeners
        networkRef.current.on('click', function(params) {
          if (params.nodes.length > 0) {
            const nodeId = params.nodes[0];
            const node = organizedNodes.find(n => n.id === nodeId);
            if (node) {
              setSelectedNode(node);
            }
          } else {
            setSelectedNode(null);
          }
        });

        // After network is stabilized, perform final layout adjustments
        networkRef.current.on("stabilizationIterationsDone", function() {
          // Fit the network to the container
          if (networkRef.current) {
            networkRef.current.fit({
              animation: true
            });
          }
        });

        // Initial fit 
        setTimeout(() => {
          if (networkRef.current) {
            networkRef.current.fit({
              animation: true
            });
          }
        }, 500);
      } catch (error) {
        console.error('Error creating network visualization:', error);
      }
    }
  }, [nodes, edges, hierarchicalLayout, createNetworkOptions, organizeNodesByState]);

  // Toggle layout mode
  const toggleLayout = () => {
    setHierarchicalLayout(!hierarchicalLayout);
  };

  // Effect for initializing network
  useEffect(() => {
    initializeNetwork();

    // Cleanup function
    return () => {
      if (networkRef.current) {
        networkRef.current.destroy();
        networkRef.current = null;
      }
    };
  }, [initializeNetwork]);

  // Format node details for display
  const formatNodeDetails = (node: Node) => {
    if (!node.title) return null;
    
    // Parse the title which is newline-separated key-value pairs
    const lines = node.title.split('\n');
    const details = lines.map(line => {
      const parts = line.split(':');
      const key = parts.shift()?.trim() || '';
      const value = parts.join(':').trim();
      return { key, value };
    });
    
    return (
      <List dense>
        {details.map((detail, index) => (
          <ListItem key={index} disablePadding>
            <ListItemText 
              primary={
                <Typography variant="body2">
                  <strong>{detail.key}:</strong> {detail.value}
                </Typography>
              } 
            />
          </ListItem>
        ))}
      </List>
    );
  };

  // Render empty state
  if (nodes.length === 0) {
    return (
      <Box display="flex" justifyContent="center" alignItems="center" height="600px">
        <Typography variant="body1">Select a device to view its relationships</Typography>
      </Box>
    );
  }

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', p: 1, borderBottom: '1px solid #eee' }}>
        <Typography variant="subtitle1">Device Configuration Visualization</Typography>
        <FormControlLabel
          control={<Switch checked={hierarchicalLayout} onChange={toggleLayout} />}
          label="Hierarchical View"
        />
      </Box>
      
      <Box sx={{ display: 'flex', flexGrow: 1, height: 'calc(100% - 48px)' }}>
        <Box sx={{ flexGrow: 1, height: '100%' }}>
          <div ref={containerRef} style={{ height: '100%', width: '100%' }} />
        </Box>
        
        {selectedNode && (
          <Paper 
            sx={{ 
              width: 300, 
              height: '100%', 
              p: 2, 
              overflowY: 'auto',
              borderLeft: '1px solid #eee'
            }}
            elevation={0}
          >
            <Box sx={{ mb: 2 }}>
              <Typography variant="h6">{selectedNode.label}</Typography>
              <Chip 
                label={selectedNode.group || 'Unknown'} 
                size="small" 
                sx={{ 
                  backgroundColor: selectedNode.color || '#ccc',
                  color: '#fff',
                  mt: 1
                }} 
              />
            </Box>
            <Divider sx={{ my: 1 }} />
            
            {formatNodeDetails(selectedNode)}
          </Paper>
        )}
      </Box>
    </Box>
  );
};

export default GraphVisualization; 