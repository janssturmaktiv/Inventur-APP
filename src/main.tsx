import { StrictMode, lazy, Suspense } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

// „…/Inventur-APP/#lager“ öffnet die Lagerinventur, sonst die Konsilager-Inventur
const Lager = lazy(() => import('./Lager'));
// Nach der Microsoft-Anmeldung kommt die Antwort ohne „#lager“ zurück: dann merkt sich die Seite das vorher
const isLager = () =>
  /^#lager/i.test(window.location.hash) ||
  new URLSearchParams(window.location.search).has('lager') ||
  (sessionStorage.getItem('afterLogin') === 'lager' && /(code|error)=/.test(window.location.hash + window.location.search));

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
