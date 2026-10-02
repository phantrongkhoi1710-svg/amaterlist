import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  FileSpreadsheet,
  Calculator,
  Gauge,
  Activity,
  Tag,
  Waves,
  Sparkles,
  FolderSync,
  type LucideIcon,
} from 'lucide-react';
import { ENGINEERING_TOOLS, WebToolItem } from '../data/toolsConfig';

/*
  Palette
  hull   #0b2f3a  ink / lines
  steel  #9db2b9  pipe body
  mist   #e6edf0  page
  flow   #ff6a2b  fluid + active
  brass  #c99a3b  valve hubs
*/

const ICONS: Partial<Record<WebToolItem['iconType'], LucideIcon>> = {
  iso: FolderSync,
  armature: FileSpreadsheet,
  calculator: Calculator,
  flange: Gauge,
  valve: Activity,
  sfi: Tag,
  hydrotest: Waves,
};

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,500..900&display=swap');
.wtr { font-family: 'Archivo', 'Helvetica Neue', Arial, sans-serif; }
.wtr-grid {
  background-image:
    linear-gradient(rgba(11,47,58,.07) 1px, transparent 1px),
    linear-gradient(90deg, rgba(11,47,58,.07) 1px, transparent 1px);
  background-size: 32px 32px;
}
.wtr-pipe-v { background: linear-gradient(90deg,#6f878f 0%,#c9d8dc 35%,#9db2b9 60%,#5d747c 100%); }
.wtr-pipe-h { background: linear-gradient(180deg,#6f878f 0%,#c9d8dc 35%,#9db2b9 60%,#5d747c 100%); }
.wtr-flow-v {
  background: repeating-linear-gradient(to bottom,#ff6a2b 0 16px,transparent 16px 34px);
  animation: wtr-flow-v 1.1s linear infinite;
}
.wtr-flow-h {
  background: repeating-linear-gradient(to right,#ff6a2b 0 12px,transparent 12px 26px);
  animation: wtr-flow-h .9s linear infinite;
}
@keyframes wtr-flow-v { to { background-position-y: 34px; } }
@keyframes wtr-flow-h { to { background-position-x: 26px; } }
.wtr-needle {
  transform-origin: 120px 120px;
  transform: rotate(46deg);
  animation: wtr-swing 2.4s cubic-bezier(.2,1.4,.4,1) both;
}
@keyframes wtr-swing { from { transform: rotate(-120deg); } to { transform: rotate(46deg); } }
@media (prefers-reduced-motion: reduce) {
  .wtr-flow-v, .wtr-flow-h, .wtr-needle { animation: none; }
}
`;

const PressureGauge: React.FC = () => {
  const ticks = Array.from({ length: 13 }, (_, i) => -120 + i * 20);
  return (
    <figure className="relative mx-auto w-64 sm:w-80" aria-label="Pressure gauge showing test pressure at 1.5 times design">
      <svg viewBox="0 0 240 240" className="w-full drop-shadow-[0_22px_30px_rgba(11,47,58,.35)]">
        <circle cx="120" cy="120" r="116" fill="#0b2f3a" />
        <circle cx="120" cy="120" r="108" fill="#c99a3b" />
        <circle cx="120" cy="120" r="102" fill="#f4f7f6" />
        <path d="M206.6 70 A100 100 0 0 1 206.6 170" fill="none" stroke="#ff6a2b" strokeWidth="7" strokeLinecap="round" transform="translate(-2 0) scale(.96) translate(5 5)" />
        {ticks.map((a, i) => (
          <line
            key={a}
            x1="120" y1="28" x2="120" y2={i % 2 === 0 ? 46 : 40}
            stroke="#0b2f3a" strokeWidth={i % 2 === 0 ? 3 : 1.5}
            transform={`rotate(${a} 120 120)`}
          />
        ))}
        <text x="120" y="158" textAnchor="middle" fontSize="13" fontWeight="700" fill="#0b2f3a">bar</text>
        <text x="120" y="184" textAnchor="middle" fontSize="11" fontWeight="600" fill="#5d747c">test = 1.5 × design</text>
        <g className="wtr-needle">
          <line x1="120" y1="132" x2="120" y2="42" stroke="#ff6a2b" strokeWidth="4" strokeLinecap="round" />
        </g>
        <circle cx="120" cy="120" r="10" fill="#0b2f3a" />
        <circle cx="120" cy="120" r="4" fill="#c99a3b" />
      </svg>
    </figure>
  );
};

const ValveRow: React.FC<{ tool: WebToolItem; onOpen: (t: WebToolItem) => void }> = ({ tool, onOpen }) => {
  const Icon = ICONS[tool.iconType] ?? Gauge;
  const live = tool.path !== '#';
  const label = tool.btnText.replace(/\s*→\s*$/, '');

  return (
    <li className="group relative pl-[92px] pb-7 last:pb-0">
      {/* valve hub sitting on the main line */}
      <div className="absolute left-0 top-1/2 -translate-y-1/2 h-14 w-14 -mt-3.5">
        <div className="absolute inset-0 rounded-full border-[3px] border-dashed border-[#0b2f3a] transition-transform duration-500 group-hover:rotate-90" />
        <div className="absolute inset-[7px] flex items-center justify-center rounded-full bg-[#c99a3b] text-[#0b2f3a] ring-2 ring-[#0b2f3a] transition-colors group-hover:bg-[#ff6a2b]">
          <Icon className="h-5 w-5" strokeWidth={2.2} />
        </div>
      </div>

      {/* branch pipe */}
      <div className="absolute left-14 top-1/2 -mt-3.5 h-3.5 w-9 -translate-y-1/2 wtr-pipe-h overflow-hidden rounded-sm">
        <div className="absolute inset-x-0 top-[5px] h-1 wtr-flow-h opacity-0 transition-opacity group-hover:opacity-100" />
      </div>

      <button
        type="button"
        onClick={() => onOpen(tool)}
        className="flex w-full flex-col gap-4 rounded-xl border-2 border-[#0b2f3a] bg-white p-5 text-left shadow-[5px_5px_0_#0b2f3a] transition-all duration-150 hover:-translate-x-0.5 hover:-translate-y-0.5 hover:shadow-[8px_8px_0_#ff6a2b] focus-visible:outline focus-visible:outline-[3px] focus-visible:outline-offset-4 focus-visible:outline-[#ff6a2b] sm:flex-row sm:items-center sm:gap-6 cursor-pointer"
      >
        <span className="flex-1">
          <span className="flex items-center gap-2.5">
            <span className="text-lg font-extrabold leading-tight text-[#0b2f3a]">{tool.title}</span>
            <span
              className={`inline-flex items-center gap-1.5 rounded-full px-2 py-0.5 text-[11px] font-bold ${
                live ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
              }`}
            >
              <span className={`h-1.5 w-1.5 rounded-full ${live ? 'bg-emerald-500' : 'bg-slate-400'}`} />
              {live ? 'Live' : 'Soon'}
            </span>
          </span>
          <span className="mt-1.5 block max-w-xl text-sm leading-relaxed text-slate-600">{tool.description}</span>
        </span>
        <span className="inline-flex shrink-0 items-center justify-center rounded-lg bg-[#0b2f3a] px-4 py-2.5 text-sm font-bold text-white transition-colors group-hover:bg-[#ff6a2b] group-hover:text-[#0b2f3a]">
          {label}
        </span>
      </button>
    </li>
  );
};

export const HubPage: React.FC = () => {
  const navigate = useNavigate();
  const [toastMsg, setToastMsg] = useState<string | null>(null);

  const handleOpen = (tool: WebToolItem) => {
    if (tool.isMainTool || tool.path.startsWith('/tools/')) {
      navigate(tool.path);
      return;
    }
    setToastMsg(`${tool.title} is coming in the next update.`);
    setTimeout(() => setToastMsg(null), 3500);
  };

  return (
    <div className="wtr wtr-grid min-h-screen bg-[#e6edf0] text-[#0b2f3a] selection:bg-[#ff6a2b] selection:text-[#0b2f3a]">
      <style>{CSS}</style>

      {/* header */}
      <header className="mx-auto flex max-w-6xl items-center justify-between px-5 pt-6 sm:px-8">
        <div className="flex items-center gap-2.5">
          <span className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#0b2f3a]">
            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="#ff6a2b" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
              <path d="M4 6h9a4 4 0 0 1 0 8H8" />
              <path d="M8 14l-3 4" />
            </svg>
          </span>
          <span className="text-xl font-black tracking-tight">WebToolRush</span>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => navigate('/tools/iso-drawing-manager')}
            className="cursor-pointer rounded-lg border-2 border-[#0b2f3a] bg-[#ff6a2b] px-3.5 py-1.5 text-sm font-bold text-[#0b2f3a] transition-colors hover:bg-[#0b2f3a] hover:text-white focus-visible:outline focus-visible:outline-[3px] focus-visible:outline-offset-2 focus-visible:outline-[#ff6a2b]"
          >
            ISO Manager
          </button>
          <button
            type="button"
            onClick={() => navigate('/tools/armature-manager')}
            className="cursor-pointer rounded-lg border-2 border-[#0b2f3a] px-3.5 py-1.5 text-sm font-bold transition-colors hover:bg-[#0b2f3a] hover:text-white focus-visible:outline focus-visible:outline-[3px] focus-visible:outline-offset-2 focus-visible:outline-[#ff6a2b]"
          >
            Armature App
          </button>
        </div>
      </header>

      {/* hero */}
      <section className="mx-auto grid max-w-6xl items-center gap-10 px-5 pb-16 pt-14 sm:px-8 lg:grid-cols-[1.25fr_1fr]">
        <div>
          <h1 className="text-[2.75rem] font-black leading-[0.98] tracking-tight sm:text-6xl lg:text-[4.5rem]" style={{ fontStretch: '88%' }}>
            Messy vendor sheets in. A clean valve schedule out.
          </h1>
          <p className="mt-6 max-w-xl text-base leading-relaxed text-slate-700">
            Import Excel from any vendor, standardize it to 26 columns, check wall thickness, look up flange ratings and set the hydrotest pressure. All in your browser, nothing to install.
          </p>
          <div className="mt-8 flex flex-wrap items-center gap-4">
            <button
              type="button"
              onClick={() => navigate('/tools/armature-manager')}
              className="cursor-pointer rounded-xl bg-[#ff6a2b] px-6 py-3.5 text-base font-extrabold text-[#0b2f3a] shadow-[5px_5px_0_#0b2f3a] transition-all hover:-translate-x-0.5 hover:-translate-y-0.5 hover:shadow-[8px_8px_0_#0b2f3a] focus-visible:outline focus-visible:outline-[3px] focus-visible:outline-offset-4 focus-visible:outline-[#0b2f3a]"
            >
              Import a valve schedule
            </button>
            <a href="#tools" className="text-sm font-bold underline decoration-[#ff6a2b] decoration-[3px] underline-offset-4">
              Browse all six tools
            </a>
          </div>
        </div>
        <PressureGauge />
      </section>

      {/* pipeline */}
      <section id="tools" className="mx-auto max-w-4xl px-5 pb-20 sm:px-8">
        <p className="mb-3 pl-[92px] text-sm font-bold text-slate-600">Raw vendor sheets</p>

        <div className="relative">
          {/* main line */}
          <div className="absolute left-[18px] top-[-12px] bottom-[-12px] w-5 wtr-pipe-v overflow-hidden rounded-full shadow-[0_0_0_2px_#0b2f3a]">
            <div className="absolute inset-y-0 left-[7px] w-[6px] wtr-flow-v" />
          </div>

          <ul className="relative">
            {ENGINEERING_TOOLS.map((tool) => (
              <ValveRow key={tool.id} tool={tool} onOpen={handleOpen} />
            ))}
          </ul>
        </div>

        <p className="mt-3 pl-[92px] text-sm font-bold text-slate-600">Standardized schedule, ready for review</p>
      </section>

      {toastMsg && (
        <div role="status" className="fixed bottom-6 left-1/2 z-50 flex -translate-x-1/2 items-center gap-2 rounded-xl border-2 border-[#0b2f3a] bg-[#ff6a2b] px-4 py-3 text-sm font-bold shadow-[5px_5px_0_#0b2f3a]">
          <Sparkles className="h-4 w-4" />
          <span>{toastMsg}</span>
        </div>
      )}

      <footer className="border-t-2 border-[#0b2f3a] bg-[#0b2f3a] py-6 text-sm text-slate-300">
        <div className="mx-auto flex max-w-6xl flex-col items-center justify-between gap-3 px-5 sm:flex-row sm:px-8">
          <p>&copy; 2026 WebToolRush. Files are processed in the browser and deleted afterwards.</p>
          <div className="flex gap-5">
            <span>Privacy Policy</span>
            <span>Terms of Service</span>
            <span>Contact Support</span>
          </div>
        </div>
      </footer>
    </div>
  );
};
