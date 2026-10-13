import type { ReactNode } from 'react';

const svg = (d: ReactNode) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    {d}
  </svg>
);

export const I = {
  search: svg(<><circle cx="11" cy="11" r="7" /><path d="m20 20-3.5-3.5" /></>),
  plus: svg(<path d="M12 5v14M5 12h14" />),
  more: svg(<><circle cx="5" cy="12" r="1" /><circle cx="12" cy="12" r="1" /><circle cx="19" cy="12" r="1" /></>),
  upload: svg(<path d="M12 16V4m0 0-4 4m4-4 4 4M4 20h16" />),
  download: svg(<path d="M12 4v12m0 0-4-4m4 4 4-4M4 20h16" />),
  file: svg(<><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" /><path d="M14 3v5h5" /></>),
  trash: svg(<path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4h6v3" />),
  edit: svg(<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z" />),
  filter: svg(<path d="M4 5h16l-6 8v6l-4-2v-4z" />),
  close: svg(<path d="M6 6l12 12M18 6 6 18" />),
};
