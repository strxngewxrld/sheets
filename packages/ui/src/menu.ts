export interface MenuItem {
  label: string;
  action: () => void;
  disabled?: boolean;
}

export type MenuEntry = MenuItem | 'separator';

let openMenu: HTMLElement | null = null;
let cleanupMenu: (() => void) | null = null;

export function closeMenu(): void {
  cleanupMenu?.();
  cleanupMenu = null;
  openMenu?.remove();
  openMenu = null;
}

/** Shows a context menu at viewport coordinates. */
export function showMenu(root: HTMLElement, x: number, y: number, entries: MenuEntry[]): void {
  closeMenu();
  const previousFocus = document.activeElement;
  const restoreFocus = () => {
    if (previousFocus instanceof HTMLElement && previousFocus.isConnected) previousFocus.focus({ preventScroll: true });
  };
  const menu = document.createElement('div');
  menu.className = 'sx-menu';
  menu.setAttribute('role', 'menu');
  for (const e of entries) {
    if (e === 'separator') {
      menu.appendChild(document.createElement('hr'));
      continue;
    }
    const b = document.createElement('button');
    b.type = 'button';
    b.setAttribute('role', 'menuitem');
    b.textContent = e.label;
    b.disabled = !!e.disabled;
    b.addEventListener('click', () => {
      closeMenu();
      restoreFocus();
      e.action();
    });
    menu.appendChild(b);
  }
  root.appendChild(menu);
  const bounds = root.getBoundingClientRect();
  menu.style.maxHeight = `${Math.max(0, bounds.height - 8)}px`;
  menu.style.maxWidth = `${Math.max(0, bounds.width - 8)}px`;
  const rect = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(4, Math.min(x - bounds.left, bounds.width - rect.width - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(y - bounds.top, bounds.height - rect.height - 4))}px`;
  openMenu = menu;
  (menu.querySelector('button:not(:disabled)') as HTMLButtonElement | null)?.focus();
  menu.addEventListener('keydown', (ev) => {
    const buttons = [...menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
    let next: number;
    switch (ev.key) {
      case 'ArrowDown': next = (index + 1) % buttons.length; break;
      case 'ArrowUp': next = (index - 1 + buttons.length) % buttons.length; break;
      case 'Home': next = 0; break;
      case 'End': next = buttons.length - 1; break;
      default: return;
    }
    ev.preventDefault();
    ev.stopPropagation();
    buttons[next]?.focus();
  });
  const dismiss = (ev: Event) => {
    if (ev instanceof KeyboardEvent && ev.key !== 'Escape' && ev.key !== 'Tab') return;
    if (ev.type === 'mousedown' && menu.contains(ev.target as Node)) return;
    closeMenu();
    if (ev instanceof KeyboardEvent) {
      if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); }
      restoreFocus();
    }
  };
  const timer = setTimeout(() => {
    window.addEventListener('mousedown', dismiss, true);
    window.addEventListener('keydown', dismiss, true);
    window.addEventListener('blur', dismiss);
    window.addEventListener('resize', dismiss);
  });
  // Editor groups can resize without resizing the browser window.
  const observer = new ResizeObserver(() => {
    const current = root.getBoundingClientRect();
    if (current.width !== bounds.width || current.height !== bounds.height) closeMenu();
  });
  observer.observe(root);
  cleanupMenu = () => {
    clearTimeout(timer);
    observer.disconnect();
    window.removeEventListener('mousedown', dismiss, true);
    window.removeEventListener('keydown', dismiss, true);
    window.removeEventListener('blur', dismiss);
    window.removeEventListener('resize', dismiss);
  };
}
