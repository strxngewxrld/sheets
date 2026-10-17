const svg = (body: string) =>
  `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;

export const ICONS = {
  undo: svg('<path d="M5 3 2 6l3 3"/><path d="M2 6h7.5a4 4 0 0 1 0 8H7"/>'),
  redo: svg('<path d="m11 3 3 3-3 3"/><path d="M14 6H6.5a4 4 0 0 0 0 8H9"/>'),
  alignLeft: svg('<path d="M2 3h12M2 6.5h8M2 10h12M2 13.5h8"/>'),
  alignCenter: svg('<path d="M2 3h12M4 6.5h8M2 10h12M4 13.5h8"/>'),
  alignRight: svg('<path d="M2 3h12M6 6.5h8M2 10h12M6 13.5h8"/>'),
  wrap: svg('<path d="M2 3h12M2 7h10a2 2 0 0 1 0 4H8"/><path d="m9.5 9.5-1.5 1.5 1.5 1.5"/><path d="M2 11h3"/>'),
  merge: svg('<rect x="1.5" y="3" width="13" height="10" rx="1"/><path d="m4 8 2-2v4zM12 8l-2-2v4z" fill="currentColor"/>'),
  borders: svg('<rect x="2" y="2" width="12" height="12"/><path d="M8 2v12M2 8h12" stroke-dasharray="1.5 1.5"/>'),
  chart: svg('<path d="M2 14h12"/><rect x="3" y="8" width="2.5" height="5" rx=".5"/><rect x="7" y="4" width="2.5" height="9" rx=".5"/><rect x="11" y="6" width="2.5" height="7" rx=".5"/>'),
  sortAsc: svg('<path d="M4 2v12M1.5 11.5 4 14l2.5-2.5"/><path d="M9 3h5M9 7h4M9 11h3"/>'),
  sortDesc: svg('<path d="M4 2v12M1.5 11.5 4 14l2.5-2.5"/><path d="M9 3h3M9 7h4M9 11h5"/>'),
  fill: svg('<path d="m3 7.5 5-5 5 5-5 5z"/><path d="M13 9.5s1.5 1.7 1.5 2.7a1.5 1.5 0 0 1-3 0c0-1 1.5-2.7 1.5-2.7z" fill="currentColor" stroke="none"/>'),
  freeze: svg('<rect x="2" y="2" width="12" height="12" rx="1"/><path d="M2 6h12M6 2v12" stroke-width="2"/>'),
};
