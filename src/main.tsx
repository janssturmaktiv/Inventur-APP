import { StrictMode, lazy, Suspense } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

// „…/Inventur-APP/#lager“ öffnet die Lagerinventur, sonst die Konsilager-Inventur
const Lager = lazy(() => import('./Lager'));
const isLager = () => /^#lager/i.test(window.location.hash) || new URLSearchParams(window.location.search).has('lager');

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    {isLager() ? (
      <Suspense fallback={null}>
        <Lager />
      </Suspense>
    ) : (
      <App />
    )}
  </StrictMode>,
);
