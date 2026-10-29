import { useId, useRef, useState, type ReactNode } from 'react';

interface HeaderDropdownProps {
  label: string;
  triggerClassName: string;
  panelClassName: string;
  children: ReactNode;
}

export function HeaderDropdown({ label, triggerClassName, panelClassName, children }: HeaderDropdownProps) {
  const panelId = useId();
  const panelRef = useRef<HTMLDivElement>(null);
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      <button
        type="button"
        className={triggerClassName}
        popoverTarget={panelId}
        aria-expanded={isOpen}
        aria-controls={panelId}
        onClick={(event) => {
          const panel = panelRef.current;
          if (!panel) return;
          const bounds = event.currentTarget.getBoundingClientRect();
          const width = Math.min(320, window.innerWidth - 16);
          panel.style.left = `${Math.max(8, Math.min(bounds.right - width, window.innerWidth - width - 8))}px`;
          const below = window.innerHeight - bounds.bottom - 14;
          if (below < 160 && bounds.top > 160) {
            panel.style.top = 'auto';
            panel.style.bottom = `${window.innerHeight - bounds.top + 6}px`;
            panel.style.maxHeight = `${bounds.top - 14}px`;
          } else {
            panel.style.bottom = 'auto';
            panel.style.top = `${bounds.bottom + 6}px`;
            panel.style.maxHeight = `${Math.max(80, below)}px`;
          }
        }}
      >
        {label}<span aria-hidden="true">&#9662;</span>
      </button>
      <div
        ref={panelRef}
        id={panelId}
        popover="auto"
        className={panelClassName}
        role="group"
        aria-label={`${label} actions`}
        onToggle={(event) => setIsOpen(event.newState === 'open')}
        onClick={(event) => {
          if ((event.target as HTMLElement).closest('button')) event.currentTarget.hidePopover();
        }}
      >
        {children}
      </div>
    </>
  );
}