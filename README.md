# Intune Explorer Web

A modern, cross-platform web application for exploring Microsoft Intune data, featuring interactive graph visualizations (Vis.js or Cytoscape.js), built with React (TypeScript) and Node.js (TypeScript), and fully containerized with Docker.

## Features
- Device code authentication (no Azure app registration required)
- Browse Intune devices, users, apps, groups, and more
- Interactive graph/network visualizations of configurations and relationships
- Clean, modern UI (Material-UI or Tailwind CSS)
- Easy deployment with Docker Compose

## Project Structure

```
intune-explorer-web/
├── docker-compose.yml
├── frontend/
│   ├── Dockerfile
│   ├── package.json
│   ├── src/
│   │   ├── components/
│   │   ├── pages/
│   │   ├── services/
│   │   └── App.tsx
│   └── tsconfig.json
├── backend/
│   ├── Dockerfile
│   ├── package.json
│   ├── src/
│   │   ├── routes/
│   │   ├── services/
│   │   └── index.ts
│   └── tsconfig.json
└── README.md
```

## Setup & Usage

### Prerequisites
- Docker & Docker Compose

### Quick Start
1. Clone the repository
2. Run `docker-compose up --build`
3. Access the app at `http://localhost:3000`

### Authentication Flow
- Click "Connect to Microsoft Graph" in the web UI
- A device code and verification URL will be displayed
- Authenticate in your browser
- Once authenticated, browse and visualize your Intune data

### Customization
- To use Cytoscape.js instead of Vis.js, swap the visualization component in `frontend/src/components/GraphVisualization.tsx`

## Development
- Frontend: React + TypeScript (see `frontend/`)
- Backend: Node.js + TypeScript (see `backend/`)
- Both can be run independently for development (see respective READMEs)

## License
MIT 