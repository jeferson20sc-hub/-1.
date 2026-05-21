import * as DialogPrimitive from "@radix-ui/react-dialog";
import { X } from "lucide-react";
import type { ComponentPropsWithoutRef, HTMLAttributes } from "react";

import { cn } from "@/lib/utils";

const Sheet = DialogPrimitive.Root;
const SheetTrigger = DialogPrimitive.Trigger;
const SheetClose = DialogPrimitive.Close;
const SheetPortal = DialogPrimitive.Portal;

function SheetOverlay({
  className,
  ...props
}: ComponentPropsWithoutRef<typeof DialogPrimitive.Overlay>) {
  return (
    <DialogPrimitive.Overlay
      className={cn(
        "fixed inset-0 z-50 bg-black/50 backdrop-blur-sm",
        "data-[state=open]:animate-in data-[state=closed]:animate-out",
        "data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0",
        className,
      )}
      {...props}
    />
  );
}

interface SheetContentProps
  extends ComponentPropsWithoutRef<typeof DialogPrimitive.Content> {
  side?: "left" | "right" | "bottom";
}

function SheetContent({
  side = "left",
  className,
  children,
  ...props
}: SheetContentProps) {
  return (
    <SheetPortal>
      <SheetOverlay />
      <DialogPrimitive.Content
        className={cn(
          "fixed z-50 flex flex-col bg-card shadow-2xl transition ease-in-out",
          "data-[state=open]:animate-in data-[state=closed]:animate-out",
          "data-[state=open]:duration-300 data-[state=closed]:duration-200",
          side === "left" && [
            "inset-y-0 left-0 w-72 border-r border-border",
            "data-[state=open]:slide-in-from-left",
            "data-[state=closed]:slide-out-to-left",
          ],
          side === "right" && [
            "inset-y-0 right-0 w-80 border-l border-border",
            "data-[state=open]:slide-in-from-right",
            "data-[state=closed]:slide-out-to-right",
          ],
          side === "bottom" && [
            "inset-x-0 bottom-0 border-t border-border rounded-t-2xl max-h-[85vh]",
            "data-[state=open]:slide-in-from-bottom",
            "data-[state=closed]:slide-out-to-bottom",
          ],
          className,
        )}
        {...props}
      >
        {children}
        <DialogPrimitive.Close className="absolute right-4 top-4 rounded-lg p-1 opacity-70 hover:opacity-100 hover:bg-accent transition-all focus:outline-none focus:ring-2 focus:ring-ring">
          <X className="size-4" />
          <span className="sr-only">Fechar</span>
        </DialogPrimitive.Close>
      </DialogPrimitive.Content>
    </SheetPortal>
  );
}

function SheetHeader({ className, ...props }: HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn("px-5 pt-5 pb-3 border-b border-border/60", className)}
      {...props}
    />
  );
}

function SheetTitle({
  className,
  ...props
}: ComponentPropsWithoutRef<typeof DialogPrimitive.Title>) {
  return (
    <DialogPrimitive.Title
      className={cn("text-sm font-bold text-foreground", className)}
      {...props}
    />
  );
}

export {
  Sheet,
  SheetClose,
  SheetContent,
  SheetHeader,
  SheetPortal,
  SheetTitle,
  SheetTrigger,
};
