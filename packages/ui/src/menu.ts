export interface MenuItem {
  label: string;
  action: () => void;
  disabled?: boolean;
}

export type MenuEntry = MenuItem | 'separator';

let openMenu: HTMLElement | null = null;

export function closeMenu(): void {
  openMenu?.remove();
  openMenu = null;
}

/** Shows a context menu at viewport coordinates. */
export function showMenu(root: HTMLElement, x: number, y: number, entries: MenuEntry[]): void {
  closeMenu();
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
      e.action();
    });
    menu.appendChild(b);
  }
  root.appendChild(menu);
  const rect = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(4, Math.min(x, window.innerWidth - rect.width - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(y, window.innerHeight - rect.height - 4))}px`;
  openMenu = menu;
  (menu.querySelector('button:not(:disabled)') as HTMLButtonElement | null)?.focus();
  const dismiss = (ev: Event) => {
    if (ev instanceof KeyboardEvent && ev.key !== 'Escape') return;
    if (ev.type === 'mousedown' && menu.contains(ev.target as Node)) return;
    closeMenu();
    window.removeEventListener('mousedown', dismiss, true);
    window.removeEventListener('keydown', dismiss, true);
    window.removeEventListener('blur', dismiss);
  };
  setTimeout(() => {
    window.addEventListener('mousedown', dismiss, true);
    window.addEventListener('keydown', dismiss, true);
    window.addEventListener('blur', dismiss);
  });
}
