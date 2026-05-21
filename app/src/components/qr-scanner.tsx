import * as DialogPrimitive from "@radix-ui/react-dialog";
import { Html5Qrcode } from "html5-qrcode";
import { Camera, Upload, X } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { Button } from "@/components/ui/button";

interface Props {
  open: boolean;
  onClose: () => void;
  onScan: (code: string) => void;
}

const CONTAINER_ID = "qr-reader-fullscreen";

export function QrScanner({ open, onClose, onScan }: Props) {
  const [facing, setFacing] = useState<"environment" | "user">("environment");
  const [err, setErr] = useState<string | null>(null);
  const scannerRef = useRef<Html5Qrcode | null>(null);

  async function stopScanner() {
    try {
      if (scannerRef.current?.isScanning) {
        await scannerRef.current.stop();
      }
    } catch (_) {
      // ignore stop errors
    }
  }

  async function startScanner(f: "environment" | "user") {
    await stopScanner();
    setErr(null);
    const el = document.getElementById(CONTAINER_ID);
    if (!el) return;
    // Use explicit dimensions so html5-qrcode can read container size
    const box = Math.round(Math.min(el.clientWidth || 260, el.clientHeight || 300) * 0.68) || 220;
    const scanner = new Html5Qrcode(CONTAINER_ID);
    scannerRef.current = scanner;
    try {
      await scanner.start(
        { facingMode: f },
        { fps: 12, qrbox: { width: box, height: box } },
        (text) => {
          onScan(text);
          stopScanner();
          onClose();
        },
        () => {},
      );
    } catch (_) {
      // Fallback: try without facingMode constraint (any camera)
      try {
        const scanner2 = new Html5Qrcode(CONTAINER_ID);
        scannerRef.current = scanner2;
        await scanner2.start(
          { deviceId: { ideal: "environment" } } as unknown as string,
          { fps: 12, qrbox: { width: box, height: box } },
          (text) => {
            onScan(text);
            stopScanner();
            onClose();
          },
          () => {},
        );
      } catch (_2) {
        setErr("Câmera não disponível. Verifique a permissão no navegador e tente novamente — ou use o upload de imagem.");
      }
    }
  }

  useEffect(() => {
    if (!open) {
      stopScanner();
      return;
    }
    let alive = true;
    // 200ms delay: ensures the dialog is fully painted and the container has dimensions
    const t = window.setTimeout(() => {
      if (alive) startScanner(facing);
    }, 200);
    return () => {
      alive = false;
      clearTimeout(t);
      stopScanner();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  async function flipCamera() {
    const next = facing === "environment" ? "user" : "environment";
    setFacing(next);
    startScanner(next);
  }

  async function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setErr(null);
    try {
      if (!scannerRef.current) {
        scannerRef.current = new Html5Qrcode(CONTAINER_ID);
      }
      const result = await scannerRef.current.scanFileV2(file, false);
      onScan(result.decodedText);
      onClose();
    } catch (_) {
      setErr("Nenhum código encontrado na imagem. Tente outra foto.");
    }
    e.target.value = "";
  }

  return (
    <DialogPrimitive.Root open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogPrimitive.Portal>
        <DialogPrimitive.Overlay className="fixed inset-0 z-50 bg-black/80 data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0" />
        <DialogPrimitive.Content
          className="fixed inset-0 z-50 flex flex-col bg-black data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0"
          aria-describedby={undefined}
        >
          {/* Header */}
          <div className="flex items-center justify-between px-4 py-3 bg-black/70 backdrop-blur-sm shrink-0">
            <div>
              <DialogPrimitive.Title className="text-white font-bold text-sm leading-tight">
                Scanner QR / Código de Barras
              </DialogPrimitive.Title>
              <p className="text-white/60 text-xs mt-0.5">
                Aponte para o código — EAN-13 · Code-128 · QR Code
              </p>
            </div>
            <Button
              variant="ghost"
              size="icon"
              onClick={onClose}
              className="text-white hover:bg-white/10 shrink-0"
            >
              <X className="size-5" />
            </Button>
          </div>

          {/* Scanner viewport */}
          <div className="flex-1 relative overflow-hidden bg-black">
            <div id={CONTAINER_ID} style={{ width: "100%", height: "100%" }} />

            {/* Targeting overlay */}
            <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
              <div className="relative size-60">
                {/* Corner marks */}
                <div className="absolute top-0 left-0 w-7 h-7 border-t-[3px] border-l-[3px] border-primary rounded-tl-xl" />
                <div className="absolute top-0 right-0 w-7 h-7 border-t-[3px] border-r-[3px] border-primary rounded-tr-xl" />
                <div className="absolute bottom-0 left-0 w-7 h-7 border-b-[3px] border-l-[3px] border-primary rounded-bl-xl" />
                <div className="absolute bottom-0 right-0 w-7 h-7 border-b-[3px] border-r-[3px] border-primary rounded-br-xl" />
                {/* Scan line */}
                <div className="absolute inset-x-2 top-1/2 -translate-y-px h-px bg-gradient-to-r from-transparent via-primary to-transparent animate-pulse" />
              </div>
            </div>
          </div>

          {/* Error banner with retry */}
          {err && (
            <div className="px-4 py-2 bg-destructive/30 border-t border-destructive/40 text-white/90 text-xs text-center shrink-0 space-y-2">
              <p>{err}</p>
              <button
                onClick={() => startScanner(facing)}
                className="px-3 py-1 rounded-md bg-white/20 hover:bg-white/30 text-white text-xs font-semibold transition-colors"
              >
                ↺ Tentar novamente
              </button>
            </div>
          )}

          {/* Footer controls */}
          <div className="flex gap-3 px-4 py-4 bg-black/70 backdrop-blur-sm shrink-0">
            <Button
              variant="outline"
              size="sm"
              onClick={flipCamera}
              className="flex-1 bg-white/10 border-white/20 text-white hover:bg-white/20 hover:text-white"
            >
              <Camera className="size-3.5 mr-1.5" />
              Virar câmera
            </Button>
            <label className="flex-1 cursor-pointer">
              <div className="flex items-center justify-center gap-1.5 h-9 px-3 rounded-md border border-white/20 bg-white/10 text-white text-sm font-medium hover:bg-white/20 transition-colors">
                <Upload className="size-3.5" />
                Upload imagem
              </div>
              <input
                type="file"
                accept="image/*"
                className="sr-only"
                onChange={onFileChange}
              />
            </label>
          </div>
        </DialogPrimitive.Content>
      </DialogPrimitive.Portal>
    </DialogPrimitive.Root>
  );
}
