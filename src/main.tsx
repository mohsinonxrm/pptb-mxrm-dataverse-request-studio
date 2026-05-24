import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';

const splash = document.getElementById('boot-splash');
if (splash) {
  splash.classList.add('hide');
  setTimeout(() => splash.remove(), 250);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
