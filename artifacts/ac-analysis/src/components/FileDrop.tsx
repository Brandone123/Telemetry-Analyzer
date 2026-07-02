import { useRef, useState, type DragEvent } from "react";
import { Upload, FileSpreadsheet, X } from "lucide-react";
import { cn } from "@/lib/utils";

interface FileDropProps {
  label: string;
  hint?: string;
  required?: boolean;
  accept?: string;
  file: File | null;
  onChange: (f: File | null) => void;
}

export function FileDrop({ label, hint, required, accept = ".xlsx,.xls,.csv", file, onChange }: FileDropProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);

  function handleDrop(e: DragEvent) {
    e.preventDefault();
    setDrag(false);
    const f = e.dataTransfer.files?.[0];
    if (f) onChange(f);
  }

  return (
    <div className="space-y-2">
      <label className="text-sm font-medium flex items-center gap-2">
        {label}
        {required && <span className="text-destructive">*</span>}
        {hint && <span className="text-xs font-normal text-muted-foreground">— {hint}</span>}
      </label>
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current?.click()}
        className={cn(
          "relative cursor-pointer rounded-lg border-2 border-dashed p-4 transition-colors",
          drag ? "border-primary bg-primary/5" : "border-border bg-muted/30 hover:bg-muted/50",
          file && "border-accent bg-accent/5",
        )}
      >
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          className="hidden"
          onChange={(e) => onChange(e.target.files?.[0] ?? null)}
        />
        {file ? (
          <div className="flex items-center gap-3">
            <FileSpreadsheet className="h-8 w-8 text-accent shrink-0" />
            <div className="flex-1 min-w-0">
              <div className="font-medium truncate">{file.name}</div>
              <div className="text-xs text-muted-foreground">{(file.size / 1024).toFixed(1)} KB</div>
            </div>
            <button
              type="button"
              onClick={(e) => { e.stopPropagation(); onChange(null); }}
              className="rounded p-1 hover:bg-destructive/10 text-destructive"
              aria-label="Remove file"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
        ) : (
          <div className="flex items-center gap-3 text-muted-foreground">
            <Upload className="h-8 w-8 shrink-0" />
            <div className="text-sm">
              <div className="font-medium text-foreground">Click or drop a file</div>
              <div className="text-xs">Excel (.xlsx, .xls) or CSV</div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
