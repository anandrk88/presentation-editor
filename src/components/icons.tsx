import React from "react";

const P: Record<string, React.ReactNode> = {
  undo: <path d="M7 5 3 9l4 4M3 9h9a5 5 0 0 1 0 10h-3" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />,
  redo: <path d="M13 5l4 4-4 4M17 9H8a5 5 0 0 0 0 10h3" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />,
  addSlide: (<>
    <rect x="2.5" y="4.5" width="11" height="8" rx="1" fill="none" stroke="currentColor" strokeWidth="1.4" />
    <path d="M15 12v6M12 15h6" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
  </>),
  duplicate: (<>
    <rect x="6.5" y="6.5" width="10" height="8" rx="1" fill="none" stroke="currentColor" strokeWidth="1.4" />
    <path d="M4 12.5V5.6C4 5 4.5 4.5 5.1 4.5H14" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
  </>),
  trash: (<>
    <path d="M4 6h12M8 6V4.5h4V6M6 6l.7 10h6.6L14 6" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    <path d="M8.5 9v4.5M11.5 9v4.5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </>),
  textbox: (<>
    <rect x="3" y="4.5" width="14" height="11" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" strokeDasharray="2.5 1.8" />
    <path d="M7 8h6M10 8v5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </>),
  image: (<>
    <rect x="3" y="4.5" width="14" height="11" rx="1" fill="none" stroke="currentColor" strokeWidth="1.4" />
    <circle cx="7.2" cy="8.2" r="1.3" fill="currentColor" />
    <path d="M4.5 14.5 9 10l3 3 2-2 1.5 3.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
  </>),
  shapes: (<>
    <rect x="3" y="10" width="7" height="6.5" fill="none" stroke="currentColor" strokeWidth="1.4" />
    <circle cx="13.5" cy="7" r="3.6" fill="none" stroke="currentColor" strokeWidth="1.4" />
  </>),
  bold: <path d="M6.5 4.5h4.2a2.9 2.9 0 0 1 0 5.8H6.5zm0 5.8h5a3 3 0 0 1 0 6h-5z" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />,
  italic: <path d="M11 4.5h4M5.5 16.5h4M12.8 4.5l-5 12" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />,
  underline: <path d="M6 4v6a4 4 0 0 0 8 0V4M5 17.5h10" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />,
  strike: <path d="M6 5h8M6 16h8M4 10.5h12" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />,
  alignL: <path d="M3 5h14M3 8.5h9M3 12h14M3 15.5h9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />,
  alignC: <path d="M3 5h14M5.5 8.5h9M3 12h14M5.5 15.5h9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />,
  alignR: <path d="M3 5h14M8 8.5h9M3 12h14M8 15.5h9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />,
  alignJ: <path d="M3 5h14M3 8.5h14M3 12h14M3 15.5h14" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />,
  bullets: (<>
    <circle cx="4" cy="5.5" r="1.2" fill="currentColor" /><circle cx="4" cy="10.5" r="1.2" fill="currentColor" /><circle cx="4" cy="15.5" r="1.2" fill="currentColor" />
    <path d="M7.5 5.5h9M7.5 10.5h9M7.5 15.5h9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </>),
  numbering: (<>
    <path d="M3.2 4h1.4v3.4M3 7.4h3M3 11.2c0-2 3-2 3-.4 0 1-3 1.6-3 3h3M7.5 5.5h9M7.5 10.5h9M7.5 15.5h9" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
  </>),
  fill: (<>
    <path d="m9 3 6 6-5.2 5.2a1.5 1.5 0 0 1-2.1 0L4.5 11a1.5 1.5 0 0 1 0-2.1z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    <path d="M16.5 13.5s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3z" fill="currentColor" />
  </>),
  outline: <rect x="4" y="4" width="12" height="12" rx="1.5" fill="none" stroke="currentColor" strokeWidth="2.2" />,
  arrange: (<>
    <rect x="3" y="3" width="9" height="9" fill="currentColor" opacity="0.35" />
    <rect x="8" y="8" width="9" height="9" fill="none" stroke="currentColor" strokeWidth="1.4" />
  </>),
  present: <path d="M5 4.5v11l9-5.5z" fill="currentColor" />,
  chevDown: <path d="m5 8 5 5 5-5" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />,
  zoomIn: <path d="M10 5v10M5 10h10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />,
  zoomOut: <path d="M5 10h10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />,
  fit: <rect x="3.5" y="5" width="13" height="10" rx="1" fill="none" stroke="currentColor" strokeWidth="1.4" />,
  fitW: (<>
    <rect x="3" y="5" width="14" height="10" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" />
    <path d="M6 10h8M6 10l1.5-1.5M6 10l1.5 1.5M14 10l-1.5-1.5M14 10l-1.5 1.5" stroke="currentColor" strokeWidth="1.2" fill="none" strokeLinecap="round" />
  </>),
  open: <path d="M3 16V5.5C3 5 3.4 4.5 4 4.5h4l1.5 2H17c.5 0 1 .4 1 1V9M3 16l2.2-6h13.3L16 16z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />,
  save: (<>
    <path d="M4 4h10l2.5 2.5V16a.9.9 0 0 1-1 1H5a.9.9 0 0 1-1-1z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    <path d="M7 4v4h6V4M7 17v-5h6v5" fill="none" stroke="currentColor" strokeWidth="1.3" />
  </>),
  doc: <path d="M5.5 3h6L16 7.5V17a1 1 0 0 1-1 1H5.5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1zM11 3v5h5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />,
  anchorT: (<>
    <rect x="3.5" y="3.5" width="13" height="13" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" />
    <path d="M6 6.5h8M6 9h8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </>),
  anchorM: (<>
    <rect x="3.5" y="3.5" width="13" height="13" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" />
    <path d="M6 9h8M6 11.5h8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </>),
  anchorB: (<>
    <rect x="3.5" y="3.5" width="13" height="13" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" />
    <path d="M6 11.5h8M6 14h8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </>),
  fontColor: <path d="M5.5 14.5 10 4l4.5 10.5M7 11h6" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />,
  changeCase: (<>
    <text x="2" y="14.5" fontFamily="Arial" fontSize="13" fontWeight="bold" fill="currentColor">A</text>
    <text x="11" y="14.5" fontFamily="Arial" fontSize="11" fill="currentColor">a</text>
  </>),
  superscript: (<>
    <path d="M3.5 15.5 8 5.5l4.5 10M5 12.2h6" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    <text x="13" y="9" fontFamily="Arial" fontSize="8.5" fontWeight="bold" fill="currentColor">2</text>
  </>),
  subscript: (<>
    <path d="M3.5 13.5 8 3.5l4.5 10M5 10.2h6" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    <text x="13" y="17.5" fontFamily="Arial" fontSize="8.5" fontWeight="bold" fill="currentColor">2</text>
  </>),
  highlighter: (<>
    <path d="m11.5 3.5 5 5L9 16H6.5L4 13.5z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    <path d="M3 17.5h14" stroke="currentColor" strokeWidth="2.4" />
  </>),
  eraser: (<>
    <path d="m11 4 5 5-6.5 6.5H6L3.5 13z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    <path d="M7.5 8.5 12 13M10 17.5h7" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
  </>),
  indentInc: <path d="M9 5h8M9 8.3h8M9 11.6h8M3 16h14M3 6.5v6l3.4-3z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" fill="currentColor" />,
  indentDec: <path d="M9 5h8M9 8.3h8M9 11.6h8M3 16h14M6.4 6.5v6L3 9.5z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" fill="currentColor" />,
  columns: (<>
    <path d="M3.5 4.5h5.4v11H3.5zM11.1 4.5h5.4v11h-5.4z" fill="none" stroke="currentColor" strokeWidth="1.3" />
    <path d="M5 7.5h2.4M5 10h2.4M5 12.5h2.4M12.6 7.5H15M12.6 10H15M12.6 12.5H15" stroke="currentColor" strokeWidth="1" strokeLinecap="round" />
  </>),
  ruler: (<>
    <rect x="2.5" y="7" width="15" height="6" rx="1" fill="none" stroke="currentColor" strokeWidth="1.4" />
    <path d="M5.5 7v2.5M8.5 7v3.5M11.5 7v2.5M14.5 7v3.5" stroke="currentColor" strokeWidth="1.1" />
  </>),
  fontUp: <path d="M3 15 7.5 4.5 12 15M4.8 11.5h5.4M14 11V4M14 4l-2 2M14 4l2 2" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />,
  fontDown: <path d="M3 15 7.5 4.5 12 15M4.8 11.5h5.4M14 4v7M14 11l-2-2M14 11l2-2" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />,
};

export function Icon({ name, size = 20 }: { name: keyof typeof P | string; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true" focusable="false">
      {P[name] ?? <circle cx="10" cy="10" r="6" fill="none" stroke="currentColor" />}
    </svg>
  );
}
