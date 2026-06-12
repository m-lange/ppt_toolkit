import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import './index.css';

Office.onReady((info: Office.IsReadyInfo) => {
  const isOffice = info.host === Office.HostType.PowerPoint;

  const container = document.getElementById('root') as HTMLElement;
  const root = createRoot(container);

  root.render(<App isOffice={isOffice} />);
});