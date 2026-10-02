import React from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { HubPage } from './pages/HubPage';
import { IsoDrawingManagerPage } from './pages/IsoDrawingManagerPage';
import App from './App';

/**
 * AppRouter defines the multi-tool routing architecture.
 * - '/' : Landing Page / Engineering Tool Hub
 * - '/tools/armature-manager' : Complete Armature List Manager application
 * - '/tools/iso-drawing-manager' : ISO Drawing Manager tool
 * - '*' : Fallback redirecting back to Hub
 */
export const AppRouter: React.FC = () => {
  return (
    <BrowserRouter>
      <Routes>
        {/* Main Tool Hub Landing Page */}
        <Route path="/" element={<HubPage />} />

        {/* Existing Armature Import & Master Manager Tool */}
        <Route path="/tools/armature-manager" element={<App />} />

        {/* ISO Drawing Manager Tool */}
        <Route path="/tools/iso-drawing-manager" element={<IsoDrawingManagerPage />} />

        {/* Catch-all route returns to Hub */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  );
};

export default AppRouter;
