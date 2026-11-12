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

export function Scanner({ onResult, onClose }: { onResult: (text: string, format: string) => void; onClose: () => void }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const [error, setError] = useState('');
  const [torch, setTorch] = useState<boolean | null>(null); // null = nicht verfügbar
  const trackRef = useRef<MediaStreamTrack | null>(null);

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
        const tick = async () => {
          if (stopped) return;
          try {
            if (video.readyState >= 2) {
              const codes = await det.detect(video);
              if (codes.length && !stopped) {
                stopped = true;
                navigator.vibrate?.(80);
                onResult(codes[0].rawValue, FORMAT_NAMES[codes[0].format] ?? codes[0].format);
                return;
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
  }, [onResult]);

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
        {error ? <p className="scanner-error">{error}</p> : <p>UDI-Code in den Rahmen halten</p>}
        <div className="row-buttons">
          {torch !== null && (
            <button className="btn" onClick={toggleTorch}>
              {torch ? 'Licht aus' : 'Licht an'}
            </button>
          )}
          <button className="btn" onClick={onClose}>
            Abbrechen
          </button>
        </div>
      </div>
    </div>
  );
}
