import { useEffect, useRef, useState } from 'react';
import { BarcodeDetector, prepareZXingModule } from 'barcode-detector/ponyfill';
import wasmUrl from 'zxing-wasm/reader/zxing_reader.wasm?url';

// Die Erkennungs-Engine wird mit der App ausgeliefert (kein externer Server nötig).
prepareZXingModule({
  overrides: {
    locateFile: (path: string, prefix: string) => (path.endsWith('.wasm') ? wasmUrl : prefix + path),
  },
});

let detector: BarcodeDetector | null = null;
function getDetector() {
  detector ??= new BarcodeDetector({
    formats: ['qr_code', 'data_matrix', 'code_128', 'code_39', 'ean_13', 'ean_8', 'upc_a', 'upc_e', 'itf', 'pdf417', 'aztec'],
  });
  return detector;
}

const FORMAT_NAMES: Record<string, string> = { qr_code: 'qr', data_matrix: 'datamatrix', code_128: 'code128', code_39: 'code39' };

// Im Dauermodus wird derselbe Code erst wieder gezählt, wenn er kurz aus dem Bild war
// (mind. 1 s nicht gesehen) und seit dem letzten Treffer mind. 2 s vergangen sind.
const SAME_CODE_GAP = 2000;
const SAME_CODE_ABSENT = 1000;

type Props = {
  /** true zurückgeben = gespeichert, Kamera bleibt offen (Dauermodus) */
  onResult: (text: string, format: string) => boolean | void;
  onClose: () => void;
  continuous?: boolean;
  feedback?: string;
};

export function Scanner({ onResult, onClose, continuous, feedback }: Props) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const [error, setError] = useState('');
  const [torch, setTorch] = useState<boolean | null>(null); // null = nicht verfügbar
  const trackRef = useRef<MediaStreamTrack | null>(null);
  const onResultRef = useRef(onResult);
  onResultRef.current = onResult;
  const continuousRef = useRef(continuous);
  continuousRef.current = continuous;

  useEffect(() => {
    let stream: MediaStream | null = null;
    let stopped = false;
    let timer = 0;

    (async () => {
      try {
        stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } },
        });
        if (stopped) return stream.getTracks().forEach((t) => t.stop());
        const video = videoRef.current!;
        video.srcObject = stream;
        await video.play();
        const track = stream.getVideoTracks()[0];
        trackRef.current = track;
        const caps = (track.getCapabilities?.() ?? {}) as MediaTrackCapabilities & { torch?: boolean; focusMode?: string[] };
        if (caps.torch) setTorch(false);
        if (caps.focusMode?.includes('continuous')) {
          track.applyConstraints({ advanced: [{ focusMode: 'continuous' } as MediaTrackConstraintSet] }).catch(() => {});
        }

        const det = getDetector();
        let last = '';
        let lastHit = 0;
        let lastSeen = 0;
        const tick = async () => {
          if (stopped) return;
          try {
            if (video.readyState >= 2) {
              const codes = await det.detect(video);
              const now = Date.now();
              if (codes.length && !stopped) {
                const text = codes[0].rawValue;
                const repeat = text === last && (now - lastHit < SAME_CODE_GAP || now - lastSeen < SAME_CODE_ABSENT);
                if (text === last) lastSeen = now;
                if (!repeat) {
                  navigator.vibrate?.(80);
                  const keepOpen = continuousRef.current && onResultRef.current(text, FORMAT_NAMES[codes[0].format] ?? codes[0].format) === true;
                  if (!continuousRef.current) {
                    stopped = true;
                    onResultRef.current(text, FORMAT_NAMES[codes[0].format] ?? codes[0].format);
                    return;
                  }
                  if (!keepOpen) return; // Kontrolle wird angezeigt, Scanner schließt
                  last = text;
                  lastHit = lastSeen = now;
                }
              }
            }
          } catch (e) {
            console.warn(e);
          }
          timer = window.setTimeout(tick, 120);
        };
        tick();
      } catch (e) {
        const name = (e as DOMException)?.name;
        setError(
          name === 'NotAllowedError'
            ? 'Kamerazugriff wurde verweigert. Bitte in den Browser-Einstellungen die Kamera für diese Seite erlauben.'
            : name === 'NotFoundError'
              ? 'Keine Kamera gefunden.'
              : `Kamera konnte nicht gestartet werden (${name ?? e}).`,
        );
      }
    })();

    return () => {
      stopped = true;
      clearTimeout(timer);
      stream?.getTracks().forEach((t) => t.stop());
    };
  }, []);

  async function toggleTorch() {
    const track = trackRef.current;
    if (!track || torch === null) return;
    try {
      await track.applyConstraints({ advanced: [{ torch: !torch } as MediaTrackConstraintSet] });
      setTorch(!torch);
    } catch {
      setTorch(null);
    }
  }

  return (
    <div className="scanner">
      <video ref={videoRef} playsInline muted />
      <div className="scanner-frame" />
      <div className="scanner-bar">
        {error ? (
          <p className="scanner-error">{error}</p>
        ) : feedback ? (
          <p className="scanner-feedback">{feedback}</p>
        ) : (
          <p>UDI-Code in den Rahmen halten</p>
        )}
        {!error && continuous && <p className="scanner-hint">Schnell-Scannen an: nächsten Code einfach ins Bild halten</p>}
        <div className="row-buttons">
          {torch !== null && (
            <button className="btn" onClick={toggleTorch}>
              {torch ? 'Licht aus' : 'Licht an'}
            </button>
          )}
          <button className="btn" onClick={onClose}>
            {continuous ? 'Fertig' : 'Abbrechen'}
          </button>
        </div>
      </div>
    </div>
  );
}
