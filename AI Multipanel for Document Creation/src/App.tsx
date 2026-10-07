import { useState, useRef, useEffect } from 'react'
import {
  createPresentation, refinePresentation, presentationExportUrl, uploadPresentationFile,
  createExcelWorkbook, excelDownloadUrl, excelChartUrl, refineExcelChart, uploadExcelFile,
  createVerifiedResearch,
  createResearchPlan, refineResearchPlan,
  uploadCleanerFile, sendCleanerChat, cleanerDownloadUrl,
  uploadExtractorImage, extractorDownloadUrl,
  checkApiHealth,
  type AIProvider,
  type PresentationInfo, type ExcelResult, type VerifiedResearchReport,
  type ResearchPlan, type CleanerUploadResult, type ExtractorResult,
} from './api'

type Mode = 'presentation' | 'excel' | 'research' | 'cleaner' | 'planner' | 'extractor'

interface Message { role: 'user' | 'assistant'; text: string }

// ─── Icons ────────────────────────────────────────────────────────────────────

const Icon = {
  Slides: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <rect x="2" y="4" width="20" height="14" rx="2"/><path d="M8 20h8M12 18v2"/>
    </svg>
  ),
  Grid: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
      <rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>
    </svg>
  ),
  Search: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
    </svg>
  ),
  Clean: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/>
    </svg>
  ),
  Plan: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/>
      <line x1="3" y1="10" x2="21" y2="10"/>
    </svg>
  ),
  Extract: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/>
    </svg>
  ),
  Sparkle: ({ s = 14, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill={c}><path d="M12 2l2.4 7.6L22 12l-7.6 2.4L12 22l-2.4-7.6L2 12l7.6-2.4L12 2z"/></svg>
  ),
  Send: ({ s = 14, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
    </svg>
  ),
  Plus: ({ s = 14, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
    </svg>
  ),
  Download: ({ s = 13, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
  ),
  Check: ({ s = 12, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12"/>
    </svg>
  ),
  Link: ({ s = 12, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
      <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
    </svg>
  ),
  Warning: ({ s = 12, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
      <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
    </svg>
  ),
  Upload: ({ s = 16, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
    </svg>
  ),
  ChevronRight: ({ s = 12, c = 'currentColor' }) => (
    <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="9 18 15 12 9 6"/>
    </svg>
  ),
}

// ─── Mode config ──────────────────────────────────────────────────────────────

const MODES: { id: Mode; verb: string; label: string; desc: string; sub: string; color: string; raw: string; ModeIcon: React.FC<{s?:number;c?:string}> }[] = [
  { id: 'presentation', verb: 'Create', label: 'Presentation Studio', desc: 'Draft and refine a polished PowerPoint narrative.', sub: 'Start with a topic, audience, and tone.', color: 'var(--c-ppt)', raw: '#c9d8e4', ModeIcon: Icon.Slides },
  { id: 'excel',        verb: 'Build',  label: 'Excel Research',       desc: 'Create sourced tables, charts, and workbook sheets.', sub: 'Export an analysis-ready workbook.', color: 'var(--c-xls)', raw: '#94c8a8', ModeIcon: Icon.Grid },
  { id: 'research',     verb: 'Research', label: 'Verified Research',  desc: 'Answer current questions with claim-level evidence.', sub: 'Use sources, citations, and visible uncertainty.', color: 'var(--c-res)', raw: '#83b7d1', ModeIcon: Icon.Search },
  { id: 'cleaner',      verb: 'Clean',  label: 'Data Cleaner',         desc: 'Upload data and change it with plain language.', sub: 'Every operation is whitelisted and previewed.', color: 'var(--c-cln)', raw: '#a9a2d0', ModeIcon: Icon.Clean },
  { id: 'planner',      verb: 'Plan',   label: 'Research Planner',     desc: 'Turn a thesis into a timeline, method, and deliverables.', sub: 'Refine the plan conversationally.', color: 'var(--c-pln)', raw: '#d5a5b8', ModeIcon: Icon.Plan },
  { id: 'extractor',    verb: 'Extract','label': 'Document Reader',    desc: 'Turn image tables into structured data.', sub: '', color: 'var(--c-ext)', raw: '#e6bd7d', ModeIcon: Icon.Extract },
]

// ─── Presentation Panel ───────────────────────────────────────────────────────

const presentationThemes = {
  Editorial: { background: 'linear-gradient(135deg,#f4f0e8,#e9e3d8)', panel: 'rgba(255,255,255,.58)', title: '#20252b', body: '#41464a', muted: '#74736d', accent: '#a95036', accent2: '#d7b68a', font: 'Georgia,serif' },
  Executive: { background: 'linear-gradient(135deg,#18273d,#101a2b)', panel: 'rgba(255,255,255,.07)', title: '#f7f9fc', body: '#d9e2f0', muted: '#aab8cc', accent: '#68a8f5', accent2: '#50d5c2', font: 'var(--font-sans)' },
  Modern: { background: 'linear-gradient(135deg,#fbfcfd,#eaf0f4)', panel: 'rgba(255,255,255,.7)', title: '#132a34', body: '#354b54', muted: '#687b83', accent: '#147d80', accent2: '#a5d5cf', font: 'var(--font-sans)' },
  Midnight: { background: 'linear-gradient(135deg,#171f2b,#0b1018)', panel: 'rgba(255,255,255,.055)', title: '#f5f3ee', body: '#d9dee5', muted: '#9aa8b5', accent: '#a7c7d8', accent2: '#738c9c', font: 'var(--font-sans)' },
  Warm: { background: 'linear-gradient(135deg,#fff8ef,#f1e2d1)', panel: 'rgba(255,255,255,.6)', title: '#39281f', body: '#59453b', muted: '#8a7264', accent: '#bd5a36', accent2: '#e1ae70', font: 'Georgia,serif' },
}

type PreviewSlide =
  | { id: number; type: 'title'; title: string; badge?: string; subtitle?: string }
  | { id: number; type: 'metrics'; title: string; metrics?: { l: string; v: string; d: string }[] }
  | { id: number; type: 'bullets'; title: string; bullets: string[]; visual?: PresentationInfo['slides'][number]['visual']; image?: PresentationInfo['slides'][number]['image'] }
  | { id: number; type: 'quote'; title: string; quote: string; author?: string }

function DiagramPreview({ visual, theme }: { visual: NonNullable<PresentationInfo['slides'][number]['visual']>; theme: typeof presentationThemes[keyof typeof presentationThemes] }) {
  const isComparison = visual.type === 'comparison'
  const isCycle = visual.type === 'cycle'
  const items = visual.items.slice(0, isCycle ? 4 : 5)
  const tile = (item: typeof items[number], index: number, className = '') => <div key={index} className={`relative flex min-w-0 flex-col justify-center rounded-lg border px-2 py-1.5 ${className}`}
    style={{ background:theme.panel, borderColor:index % 2 ? theme.accent2 : theme.accent }}>
    <span className="text-[9px] font-semibold leading-tight" style={{ color:theme.title }}>{item.label}</span>
    {item.detail && <span className="mt-1 line-clamp-2 text-[7px] leading-tight" style={{ color:theme.body }}>{item.detail}</span>}
  </div>
  if (!isComparison && !isCycle) return <div className="flex min-h-0 flex-1 items-stretch gap-1.5">
    {items.map((item, index) => <div key={index} className="relative flex min-w-0 flex-1">{tile(item, index, 'w-full')}{index < items.length - 1 && <span className="absolute -right-2 top-1/2 z-10 -translate-y-1/2 text-xs" style={{ color:theme.accent }}>›</span>}</div>)}
  </div>
  return <div className="relative grid min-h-0 flex-1 grid-cols-2 grid-rows-2 gap-1.5">
    {items.map((item, index) => tile(item, index))}
    {isCycle && <span className="pointer-events-none absolute inset-0 flex items-center justify-center text-xl" style={{ color:theme.accent }}>↻</span>}
  </div>
}

function PresentationPanel({ data, onUpload, styleName, onStyleChange, includeWebImages, onToggleWebImages }: { data: PresentationInfo | null; onUpload: (file: File) => void; styleName: keyof typeof presentationThemes; onStyleChange: (name: keyof typeof presentationThemes) => void; includeWebImages: boolean; onToggleWebImages: () => void }) {
  const [active, setActive] = useState(0)
  const [animateSlides, setAnimateSlides] = useState(true)
  const [exporting, setExporting] = useState(false)
  const [uploading, setUploading] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const liveSlides = data?.slides
  const displaySlides: PreviewSlide[] = liveSlides?.map((sl, i) => ({ id: i + 1, type: 'bullets', title: sl.title, bullets: sl.bullets, visual: sl.visual, image: sl.image })) ?? []
  const s: PreviewSlide = displaySlides[Math.min(active, displaySlides.length - 1)] ?? { id: 0, type: 'bullets', title: 'Your slides will appear here.', bullets: [] }
  const theme = presentationThemes[styleName]

  useEffect(() => { setActive(0) }, [data])

  const handleUpload = async (file: File | undefined) => {
    if (!file) return
    setUploading(true)
    try { await onUpload(file) } finally { setUploading(false) }
  }

  const exportPptx = async () => {
    if (!data) return
    setExporting(true)
    try {
      const res = await fetch(presentationExportUrl(), {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...data, style: styleName }),
      })
      if (!res.ok) {
        const body = await res.json().catch(() => null)
        throw new Error(body?.detail || `Export failed (${res.status})`)
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url; a.download = 'presentation.pptx'; a.click()
      URL.revokeObjectURL(url)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Could not export the presentation.')
    } finally {
      setExporting(false)
    }
  }

  return (
    <div className="flex h-full gap-3 p-3 overflow-hidden">
      {/* Rail */}
      <div className="flex flex-col gap-2 w-28 shrink-0 overflow-y-auto">
        {displaySlides.map((sl, i) => (
          <button key={sl.id} onClick={() => setActive(i)} aria-current={i === active ? 'page' : undefined} className="cursor-pointer text-left">
            <div className="text-[9px] mb-1 font-mono" style={{ color: 'var(--muted-fg)' }}>{i + 1}</div>
            <div className="aspect-video rounded overflow-hidden flex items-center justify-center p-2 transition-all"
              style={{ background: theme.background, border: i === active ? `2px solid ${theme.accent}` : '1px solid var(--border)', transform: i === active ? 'scale(1.03)' : 'scale(1)' }}>
              <div className="text-[5px] font-semibold text-center leading-tight" style={{ color: theme.accent }}>{sl.title}</div>
            </div>
          </button>
        ))}
      </div>

      {/* Canvas */}
      <div className="flex-1 flex flex-col gap-2 min-w-0">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <label className="flex items-center gap-1.5 text-xs" style={{ color:'var(--muted-fg)' }}>
              Style
              <select aria-label="Presentation style" value={styleName} onChange={e=>onStyleChange(e.target.value as keyof typeof presentationThemes)} className="rounded px-2 py-1 outline-none" style={{ color:'var(--fg)', background:'var(--panel)', border:'1px solid var(--border)' }}>
                {Object.keys(presentationThemes).map(name => <option key={name}>{name}</option>)}
              </select>
            </label>
            <button onClick={onToggleWebImages} aria-pressed={includeWebImages} className="px-2.5 py-1 rounded text-xs transition-colors" style={{ color:includeWebImages ? 'var(--c-ppt)' : 'var(--muted-fg)', background:'var(--secondary)', border:'1px solid var(--border)' }}>
              Web images: {includeWebImages ? 'On' : 'Off'}
            </button>
            <button onClick={()=>setAnimateSlides(value=>!value)} aria-pressed={animateSlides} className="px-2.5 py-1 rounded text-xs transition-colors" style={{ color: animateSlides ? 'var(--c-ppt)' : 'var(--muted-fg)', background:'var(--secondary)', border:'1px solid var(--border)' }}>
              Animate: {animateSlides ? 'On' : 'Off'}
            </button>
          </div>
          <div className="flex items-center gap-2">
            <input ref={fileInputRef} type="file" accept=".pptx" className="hidden" onChange={e=>handleUpload(e.target.files?.[0])}/>
            <button onClick={()=>fileInputRef.current?.click()} disabled={uploading} className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold hover:opacity-90 transition-opacity disabled:opacity-40"
              style={{ background:'var(--secondary)', color:'var(--fg)', border:'1px solid var(--border)' }}>
              <Icon.Upload s={12}/> {uploading ? 'Uploading…' : 'Upload .pptx'}
            </button>
            <button onClick={exportPptx} disabled={!data || exporting} className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold hover:opacity-90 transition-opacity disabled:opacity-40"
              style={{ background: 'var(--c-ppt)', color: '#fff' }}>
              <Icon.Download /> {exporting ? 'Exporting…' : 'Export .pptx'}
            </button>
          </div>
        </div>

        <div className="flex-1 flex items-center justify-center">
          <div className="w-full max-w-xl rounded-xl overflow-hidden relative"
            style={{ aspectRatio: '16/9', background: theme.background, boxShadow: '0 12px 48px rgba(0,0,0,0.25),0 0 0 1px rgba(255,255,255,0.05)' }}>
            <div style={{ position:'absolute', top:0, left:0, width:'3px', height:'100%', background:theme.accent }}/>
            <div style={{ position:'absolute', bottom:0, right:0, width:'28%', height:'2px', background:`linear-gradient(90deg,transparent,${theme.accent})` }}/>
            <div key={active} className="absolute inset-0 flex flex-col justify-center px-10 py-7" style={{ animation: animateSlides ? 'fade-up 350ms ease both' : undefined }}>
              {s.type === 'title' && (
                <>
                  {s.badge && <span className="text-[9px] font-mono font-semibold tracking-widest mb-4 px-2 py-0.5 rounded-full self-start"
                    style={{ background:`${theme.accent}22`, color:theme.accent, border:`1px solid ${theme.accent}55` }}>{s.badge}</span>}
                  <h1 className="text-2xl font-bold mb-2 leading-tight" style={{ fontFamily:'var(--font-serif)', color:theme.title }}>{s.title}</h1>
                  <p className="text-xs" style={{ color:theme.muted }}>{s.subtitle}</p>
                </>
              )}
              {s.type === 'metrics' && (
                <>
                  <h2 className="text-lg font-semibold mb-5" style={{ fontFamily:'var(--font-serif)', color:theme.title }}>{s.title}</h2>
                  <div className="grid grid-cols-3 gap-3">
                    {s.metrics?.map(m => (
                      <div key={m.l} className="rounded-lg p-3" style={{ background:'rgba(255,255,255,0.04)', border:'1px solid var(--border)' }}>
                        <div className="text-xl font-bold font-mono mb-0.5" style={{ color:theme.title }}>{m.v}</div>
                        <div className="text-[10px] mb-0.5" style={{ color:theme.muted }}>{m.l}</div>
                        <div className="text-[10px] font-semibold" style={{ color:'var(--c-res)' }}>{m.d}</div>
                      </div>
                    ))}
                  </div>
                </>
              )}
              {s.type === 'bullets' && (
                <div className="absolute inset-0 flex flex-col px-10 py-7">
                  <h2 className="text-lg font-semibold mb-3 shrink-0" style={{ fontFamily:'var(--font-serif)', color:theme.title }}>{s.title}</h2>
                  <div className="flex min-h-0 flex-1 gap-4">
                    <ul className={`space-y-2.5 min-h-0 overflow-hidden ${s.visual || s.image ? 'w-[45%]' : 'w-full'}`}>
                      {s.bullets?.map((b, i) => (
                        <li key={i} className="flex items-start gap-2.5 text-xs leading-relaxed" style={{ color:theme.body }}>
                          <span className="mt-1 w-1.5 h-1.5 rounded-full shrink-0" style={{ background:theme.accent }}/>
                          {b}
                        </li>
                      ))}
                    </ul>
                    {(s.visual || s.image) && <div className="flex min-w-0 flex-1 flex-col gap-2">
                      {s.visual && <DiagramPreview visual={s.visual} theme={theme} />}
                      {s.image && <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border" style={{ borderColor:theme.accent2 }}>
                        <img src={s.image.url} alt={s.title} className="min-h-0 flex-1 object-cover" />
                        <a href={s.image.file_page_url} target="_blank" rel="noreferrer" className="block truncate px-1.5 py-1 text-[7px]" style={{ color:theme.muted }}>
                          Photo: {s.image.artist} · {s.image.license} · Wikimedia Commons
                        </a>
                      </div>}
                    </div>}
                  </div>
                </div>
              )}
              {s.type === 'quote' && (
                <div className="flex flex-col items-center text-center px-4">
                  <div className="text-4xl mb-3" style={{ color:theme.accent, fontFamily:'var(--font-serif)' }}>"</div>
                  <p className="text-sm leading-relaxed mb-3" style={{ fontFamily:'var(--font-serif)', fontStyle:'italic', color:theme.title }}>{s.quote}</p>
                  <p className="text-[10px]" style={{ color:theme.muted }}>{s.author}</p>
                </div>
              )}
            </div>
            {data && <div className="absolute bottom-2 right-3 text-[9px] font-mono" style={{ color:theme.muted }}>{active+1}/{displaySlides.length}</div>}
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Excel Panel ──────────────────────────────────────────────────────────────

const COLS = ['A','B','C','D','E','F']


function ExcelPanel({ data, onUpload }: { data: ExcelResult | null; onUpload: (file: File) => void }) {
  const [uploading, setUploading] = useState(false)
  const [chartVersion, setChartVersion] = useState(() => Date.now())
  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => { setChartVersion(Date.now()) }, [data])

  const handleUpload = async (file: File | undefined) => {
    if (!file) return
    setUploading(true)
    try { await onUpload(file) } finally { setUploading(false) }
  }

  const uploadButton = (
    <>
      <input ref={fileInputRef} type="file" accept=".xlsx,.csv" className="hidden" onChange={e=>handleUpload(e.target.files?.[0])}/>
      <button onClick={()=>fileInputRef.current?.click()} disabled={uploading} className="flex items-center gap-1 px-3 py-1 rounded text-xs font-semibold hover:opacity-90 disabled:opacity-40"
        style={{ background:'var(--secondary)', color:'var(--fg)', border:'1px solid var(--border)' }}>
        <Icon.Upload s={12}/> {uploading ? 'Uploading…' : 'Upload'}
      </button>
    </>
  )

  if (data) {
    return (
      <div className="flex flex-col h-full">
        <div className="flex items-center gap-2 px-3 py-2 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
          <span className="text-xs font-semibold" style={{ color:'var(--fg)' }}>Research workbook</span>
          <div className="flex items-center gap-2 ml-auto">
            {uploadButton}
            <a href={excelDownloadUrl()} className="flex items-center gap-1 px-3 py-1 rounded text-xs font-semibold hover:opacity-90" style={{ background:'var(--c-xls)', color:'#fff' }}>
              <Icon.Download /> .xlsx
            </a>
          </div>
        </div>
        <div className="flex-1 overflow-auto p-3">
          <div className="mb-3 rounded-xl p-3" style={{ background:'var(--secondary)', border: '1px solid var(--border)' }}>
            <div className="text-[10px] uppercase tracking-[0.18em]" style={{ color:'var(--muted-fg)' }}>Executive overview</div>
            <div className="mt-2 text-sm" style={{ color:'var(--fg)' }}>{data.summary || 'No executive summary was returned.'}</div>
          </div>
          {data.chart_available !== false && <div className="mb-3 overflow-hidden rounded-xl border" style={{ background:'#fff', borderColor:'var(--border)' }}>
            <div className="px-3 pt-2 text-[10px] font-semibold uppercase tracking-wider" style={{ color:'var(--muted-fg)' }}>Chart preview · editable in the downloaded workbook</div>
            <img src={`${excelChartUrl()}?v=${chartVersion}`} alt="Generated chart preview" className="block h-64 w-full object-contain" />
          </div>}
          <table className="w-full text-xs border-collapse" style={{ fontFamily:'var(--font-mono)' }}>
            <thead>
              <tr>
                {data.columns.map(c => (
                  <th key={c} className="sticky top-0 py-1.5 px-3 text-left font-medium"
                    style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color:'var(--muted-fg)' }}>{c}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, ri) => (
                <tr key={ri}>
                  {data.columns.map(col => (
                    <td key={col} className="px-3 py-1.5" style={{ borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color:'var(--fg)' }}>
                      {String(row[col] ?? '')}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          {data.total_rows > data.rows.length && <div className="mt-2 text-[10px]" style={{ color:'var(--muted-fg)' }}>Previewing {data.rows.length} of {data.total_rows} rows. Chart actions use the complete workbook data.</div>}
        </div>
        <div className="px-3 py-1.5 border-t flex items-center gap-2" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
          <span className="text-[10px] font-mono" style={{ color:'var(--muted-fg)' }}>{data.total_rows} rows</span>
          <span className="ml-auto text-[10px]" style={{ color:'var(--muted-fg)' }}>{data.sources.length} sources</span>
        </div>
      </div>
    )
  }

  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center gap-2 px-3 py-2 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
        <span className="text-xs font-mono px-2 py-0.5 rounded" style={{ background:'var(--secondary)', color:'var(--muted-fg)' }}>—</span>
        <div className="flex-1 px-2 py-0.5 rounded text-xs font-mono" style={{ background:'var(--secondary)', color:'var(--muted-fg)' }}>Upload a workbook to preview cell values</div>
        <div className="flex gap-1 ml-2">
          {['Sheet 1','Chart View','Pivot'].map((t,i) => (
            <span key={t} className="px-2 py-0.5 rounded text-[10px]" style={{ background: i===0 ? 'var(--c-xls)' : 'var(--secondary)', color: i===0 ? '#fff' : 'var(--muted-fg)' }}>{t}</span>
          ))}
          {uploadButton}
          <button disabled className="flex items-center gap-1 ml-2 px-3 py-1 rounded text-xs font-semibold opacity-40" style={{ background:'var(--c-xls)', color:'#fff' }}>
            <Icon.Download /> .xlsx
          </button>
        </div>
      </div>
      <div className="flex-1 overflow-auto p-3">
        <div className="mb-3 rounded-xl p-3" style={{ background: 'var(--secondary)', border: '1px solid var(--border)' }}>
            <div className="text-[10px] uppercase tracking-[0.18em]" style={{ color:'var(--muted-fg)' }}>No workbook yet — ask the assistant a question or upload data</div>
        </div>
        <div className="h-24 mb-3 rounded-xl p-3 flex items-end gap-2" style={{ background:'radial-gradient(70% 100% at 50% 100%,rgba(125,157,182,.18),transparent 70%),linear-gradient(180deg,#101923,#0a0f15)', border:'1px solid var(--border)' }}>
          <div className="m-auto text-[10px]" style={{ color:'var(--muted-fg)' }}>Workbook chart preview appears after data is loaded</div>
        </div>
        <table className="w-full text-xs border-collapse" style={{ fontFamily:'var(--font-mono)' }}>
          <thead>
            <tr>
              <th className="w-8 sticky top-0 py-1.5" style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color:'var(--muted-fg)' }}></th>
              {COLS.map(c => <th key={c} className="sticky top-0 py-1.5 px-3 text-center min-w-[110px] font-medium"
                style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color:'var(--muted-fg)' }}>{c}</th>)}
            </tr>
          </thead>
          <tbody>
            {[1,2,3,4,5,6,7].map(row => (
              <tr key={row}>
                <td className="text-center py-1.5 select-none text-[10px]" style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color:'var(--muted-fg)' }}>{row}</td>
                {COLS.map(col => {
                  const key=`${col}${row}`; const val=''; const isHead=row===1||col==='F'; const isTotal=row===7
                  return (
                    <td key={key} className="px-3 py-1.5"
                      style={{ background:isTotal?'rgba(94,234,212,0.04)':'transparent', borderBottom:'1px solid var(--border)', borderRight:'1px solid var(--border)', color: isTotal?'var(--c-res)':isHead?'var(--muted-fg)':'var(--fg)', fontWeight: isHead||isTotal?'600':'400' }}>
                      {val}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="px-3 py-1.5 border-t flex items-center gap-2" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
        <span className="text-[10px] font-mono" style={{ color:'var(--muted-fg)' }}>No live workbook loaded</span>
        <span className="ml-auto text-[10px]" style={{ color:'var(--muted-fg)' }}>Use the assistant or upload a file</span>
      </div>
    </div>
  )
}

// ─── Research Panel ───────────────────────────────────────────────────────────

const confColor: Record<string,string> = { high:'var(--c-xls)', medium:'var(--c-ppt)', low:'var(--c-pln)' }

function ResearchPanel({ data, onAsk }: { data: VerifiedResearchReport | null; onAsk: (question: string) => Promise<void> }) {
  const [query, setQuery] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const submitQuestion = async () => {
    if (!query.trim() || loading) return
    setLoading(true)
    setError('')
    try { await onAsk(query.trim()) }
    catch (reason) { setError(reason instanceof Error ? reason.message : String(reason)) }
    finally { setLoading(false) }
  }

  if (data) {
    const sourceMap = new Map(data.sources.map(s => [s.id, s]))
    return (
      <div className="flex flex-col h-full">
        <div className="px-4 py-3 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
          <div className="flex gap-2">
            <div className="flex-1 flex items-center gap-2 px-3 py-2 rounded-lg" style={{ background:'var(--secondary)', border:'1px solid var(--border)' }}>
              <Icon.Search s={14} c="var(--muted-fg)" />
              <input maxLength={1500} className="flex-1 bg-transparent text-sm outline-none" style={{ color:'var(--fg)' }} placeholder="Ask another research question…" value={query} onChange={e=>setQuery(e.target.value)} onKeyDown={e=>{if(e.key==='Enter') submitQuestion()}}/>
              <button onClick={submitQuestion} disabled={loading || !query.trim()} className="px-3 py-1 rounded text-xs font-semibold disabled:opacity-50" style={{ background:'var(--c-res)', color:'#fff' }}>{loading ? 'Searching…' : 'Research'}</button>
            </div>
          </div>
          <div className="flex items-center gap-1.5 mt-2 text-[10px]" style={{ color: error ? 'var(--red)' : 'var(--muted-fg)' }}>
            {error || `${data.claims.length} claims · ${data.sources.length} sources cross-referenced · as of ${data.as_of}`}
          </div>
        </div>
        <div className="px-4 py-2 border-b text-[10px]" style={{ borderColor:'var(--border)', color:'var(--muted-fg)' }}>
          <Icon.Check s={10} c="var(--c-xls)"/> Verified research report
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          <div className="rounded-xl p-4" style={{ background:'var(--card)', border:'1px solid var(--border)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Answer</div>
            <p className="text-sm leading-relaxed" style={{ color:'var(--fg)' }}>{data.answer}</p>
          </div>
          <div className="text-xs font-semibold uppercase tracking-wider mb-1 mt-3" style={{ color:'var(--muted-fg)' }}>Claim-level Evidence</div>
          {data.claims.map((cl, i) => (
            <div key={i} className="rounded-xl p-4" style={{ background:'var(--card)', border:'1px solid var(--border)' }}>
              <div className="flex items-start justify-between gap-3 mb-3">
                <p className="text-sm leading-relaxed flex-1" style={{ color:'var(--fg)' }}>{cl.text}</p>
                <span className="text-[10px] font-mono px-1.5 py-0.5 rounded-full shrink-0" style={{ background:`${confColor[cl.confidence]||'var(--c-pln)'}18`, color:confColor[cl.confidence]||'var(--c-pln)' }}>
                  {cl.confidence} confidence
                </span>
              </div>
              <div className="space-y-1">
                {cl.source_ids.map(id => {
                  const src = sourceMap.get(id)
                  return src ? (
                    <div key={id} className="flex items-center gap-1.5 text-[10px]" style={{ color:'var(--muted-fg)' }}>
                      <Icon.Link s={10} c="var(--c-res)"/>
                      <a href={src.url} target="_blank" rel="noreferrer">{src.title}</a> · {src.publisher}
                    </div>
                  ) : null
                })}
              </div>
            </div>
          ))}
          {data.limitations.length > 0 && (
            <div className="rounded-xl p-4" style={{ background:'rgba(251,191,36,0.08)', border:'1px solid rgba(251,191,36,0.25)' }}>
              <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--c-ext)' }}>Limitations</div>
              {data.limitations.map((l, i) => <p key={i} className="text-xs mb-1" style={{ color:'var(--fg)' }}>{l}</p>)}
            </div>
          )}
        </div>
      </div>
    )
  }

  return (
    <div className="flex flex-col h-full">
      {/* Search bar */}
      <div className="px-4 py-3 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
        <div className="flex gap-2">
          <div className="flex-1 flex items-center gap-2 px-3 py-2 rounded-lg" style={{ background:'var(--secondary)', border:'1px solid var(--border)' }}>
            <Icon.Search s={14} c="var(--muted-fg)" />
            <input maxLength={1500} className="flex-1 bg-transparent text-sm outline-none" style={{ color:'var(--fg)' }} placeholder="Ask a research question…" value={query} onChange={e=>setQuery(e.target.value)} onKeyDown={e=>{if(e.key==='Enter') submitQuestion()}}/>
            <button onClick={submitQuestion} disabled={loading || !query.trim()} className="px-3 py-1 rounded text-xs font-semibold disabled:opacity-50" style={{ background:'var(--c-res)', color:'#fff' }}>{loading ? 'Searching…' : 'Research'}</button>
          </div>
        </div>
        <div className="flex items-center gap-1.5 mt-2 text-[10px]" style={{ color:'var(--muted-fg)' }}>
          {error || 'Submit a question here or ask in the assistant to fetch current, cited evidence.'}
        </div>
      </div>

      {/* Claims */}
      <div className="flex-1 overflow-y-auto p-4 space-y-3">
        <div className="rounded-xl p-5 text-sm" style={{ background:'var(--card)', border:'1px dashed var(--border)', color:'var(--muted-fg)' }}>No verified research yet. Ask a question in the assistant to fetch sourced claims.</div>
      </div>
    </div>
  )
}

// ─── Data Cleaner Panel ───────────────────────────────────────────────────────

const ops: { id:number; label:string; color:string }[] = []

function CleanerPanel({ session, onUpload }: { session: CleanerUploadResult | null; onUpload: (file: File) => void }) {
  const [dragging, setDragging] = useState(false)
  const [uploading, setUploading] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleFile = async (file: File | undefined) => {
    if (!file) return
    setUploading(true)
    try { await onUpload(file) } finally { setUploading(false); setDragging(false) }
  }

  const columns = session?.columns ?? []
  const rows = session ? session.rows.map(r => columns.map(c => r[c] ?? '')) : []
  const duplicateRowsRemoved = session ? Math.max(0, session.source_rows - session.total_rows) : 0

  return (
    <div className="flex flex-col h-full">
      {/* Upload area */}
      <div className="px-4 pt-4 pb-3 border-b" style={{ borderColor:'var(--border)' }}>
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx" className="hidden" onChange={e=>handleFile(e.target.files?.[0])}/>
        {dragging ? (
          <div className="rounded-xl flex flex-col items-center justify-center py-6 text-center"
            onDragLeave={()=>setDragging(false)}
            onDrop={e=>{e.preventDefault(); handleFile(e.dataTransfer.files?.[0])}}
            style={{ border:'2px dashed var(--c-cln)', background:'rgba(167,139,250,0.06)' }}>
            <Icon.Upload s={22} c="var(--c-cln)"/>
            <p className="text-xs mt-2" style={{ color:'var(--c-cln)' }}>Drop to upload</p>
          </div>
        ) : (
          <div onClick={()=>fileInputRef.current?.click()} className="rounded-xl flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-white/5 transition-colors"
            onDragOver={e=>{e.preventDefault();setDragging(true)}} style={{ border:'1px dashed var(--border)', background:'var(--secondary)' }}>
            <Icon.Upload s={18} c="var(--muted-fg)"/>
            <div>
                <div className="text-xs font-medium" style={{ color:'var(--fg)' }}>{uploading ? 'Uploading…' : session ? 'Data loaded' : 'Click or drop a CSV or .xlsx file'}</div>
              <div className="text-[10px]" style={{ color:'var(--muted-fg)' }}>
                {session ? `${session.total_rows} rows · ${session.columns.length} columns · ${duplicateRowsRemoved} duplicate rows removed on upload · click to replace` : 'Upload data to start a cleaning session'}
              </div>
            </div>
          </div>
        )}
      </div>

      <div className="flex flex-1 min-h-0">
        {/* Column summary */}
        <div className="w-56 shrink-0 border-r flex flex-col" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
          <div className="px-3 py-2 border-b text-[10px] font-semibold uppercase tracking-wider" style={{ borderColor:'var(--border)', color:'var(--muted-fg)' }}>
            {session ? 'Missing values' : 'Operations'}
          </div>
          <div className="flex-1 overflow-y-auto py-2">
            {session ? (
              Object.entries(session.missing_by_column).map(([col, count]) => (
                <div key={col} className="flex items-center justify-between px-3 py-2">
                  <span className="text-[11px]" style={{ color:'var(--fg)' }}>{col}</span>
                  <span className="text-[10px] font-mono" style={{ color: count > 0 ? 'var(--c-ext)' : 'var(--muted-fg)' }}>{count}</span>
                </div>
              ))
            ) : ops.map(op => (
              <div key={op.id} className="flex items-start gap-2 px-3 py-2">
                <div className="w-3.5 h-3.5 rounded shrink-0 mt-0.5 flex items-center justify-center" style={{ background: op.color, border: `1px solid ${op.color}` }}>
                  <Icon.Check s={8} c="#fff"/>
                </div>
                <span className="text-[11px] leading-snug" style={{ color:'var(--fg)' }}>{op.label}</span>
              </div>
            ))}
          </div>
          <div className="p-3 border-t" style={{ borderColor:'var(--border)' }}>
            {session ? (
              <a href={cleanerDownloadUrl(session.session_id)} className="block text-center w-full py-1.5 rounded-lg text-xs font-semibold hover:opacity-90 transition-opacity" style={{ background:'var(--c-cln)', color:'#fff' }}>
                Export Clean Data
              </a>
            ) : (
              <button disabled className="w-full py-1.5 rounded-lg text-xs font-semibold opacity-40" style={{ background:'var(--c-cln)', color:'#fff' }}>
                Upload data to export
              </button>
            )}
          </div>
        </div>

        {/* Preview table */}
        <div className="flex-1 overflow-auto">
          <table className="w-full text-[11px] border-collapse" style={{ fontFamily:'var(--font-mono)' }}>
            <thead>
              <tr>
                {columns.map(h => (
                  <th key={h} className="sticky top-0 px-3 py-2 text-left font-semibold"
                    style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', color:'var(--muted-fg)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, ri) => (
                <tr key={ri} className="hover:bg-white/3 transition-colors">
                  {row.map((cell, ci) => {
                    const isEmpty = cell === '' || cell === null
                    return (
                      <td key={ci} className="px-3 py-2" style={{ borderBottom:'1px solid var(--border)', color: isEmpty ? 'var(--c-pln)' : 'var(--fg)' }}>
                        {isEmpty ? '⚠ missing' : String(cell)}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ─── Research Planner Panel ───────────────────────────────────────────────────

const phases: { id:number; phase:string; name:string; weeks:string; tasks:string[]; deliverable:string; done:boolean }[] = []

function PlannerPanel({ data }: { data: ResearchPlan | null }) {
  const [expanded, setExpanded] = useState<number|null>(0)

  if (data) {
    return (
      <div className="flex flex-col h-full">
        <div className="px-4 py-3 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
          <div className="text-sm font-semibold" style={{ color:'var(--fg)' }}>{data.title}</div>
          <div className="text-xs mt-0.5" style={{ color:'var(--muted-fg)' }}>{data.field} · {data.thesis_statement}</div>
          <div className="flex gap-3 mt-2">
            <div className="flex flex-col">
              <span className="text-sm font-bold font-mono" style={{ color:'var(--c-pln)' }}>{data.timeline.length}</span>
              <span className="text-[10px]" style={{ color:'var(--muted-fg)' }}>Phases</span>
            </div>
            <div className="flex flex-col">
              <span className="text-sm font-bold font-mono" style={{ color:'var(--c-pln)' }}>{data.timeline.reduce((sum,p)=>sum+(p.duration_weeks||0),0)}</span>
              <span className="text-[10px]" style={{ color:'var(--muted-fg)' }}>Weeks</span>
            </div>
          </div>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-2">
          {data.timeline.map((ph, i) => (
            <div key={i} className="rounded-xl overflow-hidden" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
              <div className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-white/3 transition-colors" onClick={()=>setExpanded(expanded===i?null:i)}>
                <div className="w-5 h-5 rounded-full flex items-center justify-center shrink-0" style={{ background: expanded===i ? 'var(--c-pln)' : 'var(--secondary)', border:'1px solid var(--border)' }}>
                  <span className="text-[8px] font-mono" style={{ color:'var(--muted-fg)' }}>{i+1}</span>
                </div>
                <div className="flex-1 min-w-0">
                  <span className="text-xs font-semibold" style={{ color:'var(--fg)' }}>{ph.phase}</span>
                  <div className="text-[10px]" style={{ color:'var(--muted-fg)' }}>{ph.duration_weeks} weeks</div>
                </div>
                <Icon.ChevronRight s={12} c="var(--muted-fg)"/>
              </div>
              {expanded===i && (
                <div className="px-4 pb-4 border-t" style={{ borderColor:'var(--border)' }}>
                  <p className="pt-3 text-xs leading-relaxed" style={{ color:'var(--text-2)' }}>{ph.description}</p>
                </div>
              )}
            </div>
          ))}
          {data.expected_outcomes.length > 0 && (
            <div className="rounded-xl p-4 mt-2" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
              <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Expected outcomes</div>
              {data.expected_outcomes.map((o,i) => <div key={i} className="text-xs mb-1" style={{ color:'var(--fg)' }}>• {o}</div>)}
            </div>
          )}
          {data.objectives.length > 0 && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Objectives</div>
            {data.objectives.map((objective,i) => <div key={i} className="mb-2"><div className="text-xs font-semibold" style={{ color:'var(--fg)' }}>{objective.title}</div><div className="text-xs" style={{ color:'var(--muted-fg)' }}>{objective.description}</div></div>)}
          </div>}
          {data.methodology && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Methodology</div>
            <div className="text-xs font-semibold" style={{ color:'var(--fg)' }}>{data.methodology.research_design}</div>
            <div className="text-xs mt-1" style={{ color:'var(--muted-fg)' }}>{data.methodology.analysis_approach} · Scope: {data.methodology.sample_size_or_scope}</div>
            <div className="text-xs mt-2" style={{ color:'var(--fg)' }}>{data.methodology.data_collection.join(' · ')}</div>
          </div>}
          {data.resources_needed.length > 0 && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Resources needed</div>
            {data.resources_needed.map((item,i) => <div key={i} className="text-xs mb-1" style={{ color:'var(--fg)' }}><b>{item.resource}</b> — {item.justification}</div>)}
          </div>}
          {data.potential_challenges.length > 0 && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Challenges and mitigations</div>
            {data.potential_challenges.map((item,i) => <div key={i} className="text-xs mb-1" style={{ color:'var(--fg)' }}><b>{item.challenge}</b> — {item.mitigation}</div>)}
          </div>}
          {data.literature_review && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>Literature review</div>
            <div className="text-xs" style={{ color:'var(--fg)' }}>{data.literature_review.scope}</div>
            <div className="text-xs mt-2" style={{ color:'var(--muted-fg)' }}>Themes: {data.literature_review.key_themes.join(' · ')}</div>
            <div className="text-xs mt-1" style={{ color:'var(--muted-fg)' }}>Explore: {data.literature_review.sources_to_explore.join(' · ')}</div>
          </div>}
          {data.references_to_start && data.references_to_start.length > 0 && <div className="rounded-xl p-4" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color:'var(--muted-fg)' }}>References to start</div>
            {data.references_to_start.map((reference,i) => <div key={i} className="text-xs mb-1" style={{ color:'var(--fg)' }}>• {reference}</div>)}
          </div>}
        </div>
      </div>
    )
  }

  return (
    <div className="flex flex-col h-full">
      <div className="px-4 py-3 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
        <div className="text-sm font-semibold" style={{ color:'var(--fg)' }}>No research plan yet</div>
        <div className="text-xs mt-0.5" style={{ color:'var(--muted-fg)' }}>Ask the assistant to create a plan for your thesis or research question.</div>
        <div className="flex gap-3 mt-2">
          {[{ l:'Phases', v:'—' },{ l:'Weeks', v:'—' },{ l:'Deliverables', v:'—' },{ l:'Progress', v:'—' }].map(s => (
            <div key={s.l} className="flex flex-col">
              <span className="text-sm font-bold font-mono" style={{ color:'var(--c-pln)' }}>{s.v}</span>
              <span className="text-[10px]" style={{ color:'var(--muted-fg)' }}>{s.l}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Timeline */}
      <div className="flex-1 overflow-y-auto p-4 space-y-2">
        {!phases.length && <div className="rounded-xl p-5 text-sm" style={{ background:'var(--card)', border:'1px dashed var(--border)', color:'var(--muted-fg)' }}>Your generated phases, methods, resources, and deliverables will appear here.</div>}
        {phases.map(ph => (
          <div key={ph.id} className="rounded-xl overflow-hidden" style={{ border:'1px solid var(--border)', background:'var(--card)' }}>
            <div className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-white/3 transition-colors" onClick={()=>setExpanded(expanded===ph.id?null:ph.id)}>
              <div className="w-5 h-5 rounded-full flex items-center justify-center shrink-0"
                style={{ background: ph.done ? 'var(--c-xls)' : expanded===ph.id ? 'var(--c-pln)' : 'var(--secondary)', border: ph.done ? 'none' : '1px solid var(--border)' }}>
                {ph.done ? <Icon.Check s={9} c="#fff"/> : <span className="text-[8px] font-mono" style={{ color:'var(--muted-fg)' }}>{ph.id}</span>}
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold" style={{ color: ph.done ? 'var(--muted-fg)' : 'var(--fg)' }}>{ph.name}</span>
                  {ph.done && <span className="text-[9px] px-1.5 py-0.5 rounded-full font-mono" style={{ background:'rgba(34,197,94,0.12)', color:'var(--c-xls)' }}>Complete</span>}
                </div>
                <div className="text-[10px]" style={{ color:'var(--muted-fg)' }}>{ph.weeks}</div>
              </div>
              <Icon.ChevronRight s={12} c="var(--muted-fg)"/>
            </div>
            {expanded===ph.id && (
              <div className="px-4 pb-4 border-t" style={{ borderColor:'var(--border)' }}>
                <div className="pt-3 space-y-1.5 mb-3">
                  {ph.tasks.map((t,i) => (
                    <div key={i} className="flex items-start gap-2 text-xs" style={{ color:'var(--card-fg)' }}>
                      <span className="mt-1 w-1 h-1 rounded-full shrink-0" style={{ background:'var(--c-pln)' }}/>
                      {t}
                    </div>
                  ))}
                </div>
                <div className="flex items-center gap-1.5 text-[10px] rounded-lg px-2.5 py-1.5" style={{ background:'rgba(251,113,133,0.08)', border:'1px solid rgba(251,113,133,0.2)', color:'var(--c-pln)' }}>
                  <Icon.Check s={9} c="var(--c-pln)"/> Deliverable: {ph.deliverable}
                </div>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  )
}

// ─── Document Extractor Panel ─────────────────────────────────────────────────

function ExtractorPanel({ data, onUpload }: { data: ExtractorResult | null; onUpload: (file: File) => void }) {
  const [hovering, setHovering] = useState(false)
  const [uploading, setUploading] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleFile = async (file: File | undefined) => {
    if (!file) return
    setUploading(true)
    try { await onUpload(file) } finally { setUploading(false); setHovering(false) }
  }

  const columns = data?.columns ?? []
  const rows = data ? data.rows.map(r => columns.map(c => r[c] ?? '')) : []

  return (
    <div className="flex flex-col h-full">
      {/* Drop zone */}
      <div className="px-4 pt-4 pb-3 border-b" style={{ borderColor:'var(--border)', background:'var(--panel)' }}>
        <input ref={fileInputRef} type="file" accept=".jpg,.jpeg,.png,.webp,image/jpeg,image/png,image/webp" className="hidden" onChange={e=>handleFile(e.target.files?.[0])}/>
        <div onClick={()=>fileInputRef.current?.click()} className="rounded-xl flex items-center gap-4 px-4 py-3 transition-all cursor-pointer"
          onDragOver={e=>{e.preventDefault();setHovering(true)}} onDragLeave={()=>setHovering(false)}
          onDrop={e=>{e.preventDefault(); handleFile(e.dataTransfer.files?.[0])}}
          style={{ border: hovering ? '2px dashed var(--c-ext)' : '1px dashed var(--border)', background: hovering ? 'rgba(251,191,36,0.06)' : 'var(--secondary)' }}>
          <Icon.Upload s={20} c={hovering ? 'var(--c-ext)' : 'var(--muted-fg)'}/>
          <div>
            <div className="text-xs font-medium" style={{ color:'var(--fg)' }}>{uploading ? 'Extracting…' : data ? 'Image processed' : 'Click or drop an image to extract its table'}</div>
            <div className="text-[10px]" style={{ color:'var(--muted-fg)' }}>
              {data ? `${data.total_rows} rows × ${data.columns.length} cols extracted` : 'Printed, cursive, or faint handwriting · PNG, JPG, WebP'}
            </div>
          </div>
          {data && (
            <div className="ml-auto flex items-center gap-1.5 text-[10px] font-semibold px-2 py-1 rounded-full" style={{ background:'rgba(251,191,36,0.12)', color:'var(--c-ext)' }}>
              <Icon.Check s={10} c="var(--c-ext)"/> Extracted
            </div>
          )}
        </div>
      </div>

      {/* Extracted table */}
      <div className="flex-1 flex flex-col overflow-hidden">
        <div className="flex items-center justify-between px-4 py-2 border-b" style={{ borderColor:'var(--border)' }}>
          <div className="text-[10px]" style={{ color:'var(--muted-fg)' }}>{data ? (data.total_rows ? 'Extracted table · review uncertain cells' : 'Transcription · no table was confidently detected') : 'No extraction yet'}</div>
          {data && data.total_rows > 0 && (
            <div className="flex gap-2">
              <a href={extractorDownloadUrl('csv')} className="flex items-center gap-1 px-2.5 py-1 rounded text-[10px] font-semibold hover:opacity-90" style={{ background:'var(--c-xls)', color:'#fff' }}>
                <Icon.Download s={10}/> .csv
              </a>
              <a href={extractorDownloadUrl('xlsx')} className="flex items-center gap-1 px-2.5 py-1 rounded text-[10px] font-semibold hover:opacity-90" style={{ background:'var(--secondary)', color:'var(--fg)', border:'1px solid var(--border)' }}>
                <Icon.Download s={10}/> .xlsx
              </a>
            </div>
          )}
        </div>
        <div className="flex-1 overflow-auto">
          {data && data.total_rows > 0 ? <table className="w-full text-xs border-collapse" style={{ fontFamily:'var(--font-mono)' }}>
            <thead>
              <tr>
                {columns.map(h => (
                  <th key={h} className="sticky top-0 px-3 py-2 text-left font-semibold"
                    style={{ background:'var(--panel)', borderBottom:'1px solid var(--border)', color:'var(--muted-fg)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row,i) => (
                <tr key={i} className="hover:bg-white/3 transition-colors">
                  {row.map((cell, ci) => (
                    <td key={ci} className="px-3 py-2" style={{ borderBottom:'1px solid var(--border)', color:'var(--fg)' }}>{String(cell)}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table> : data ? <div className="p-4 text-xs whitespace-pre-wrap" style={{ color:'var(--fg)' }}>{data.markdown}</div> : <div className="p-4 text-xs" style={{ color:'var(--muted-fg)' }}>The transcription and any uncertain readings will appear here. Please review marked values against the image before using them.</div>}
          {data?.markdown && data.total_rows > 0 && <details className="m-4 rounded-lg p-3" style={{ background:'var(--secondary)', border:'1px solid var(--border)' }}>
            <summary className="cursor-pointer text-xs font-semibold" style={{ color:'var(--fg)' }}>Full transcription and uncertain readings</summary>
            <pre className="mt-3 whitespace-pre-wrap text-xs" style={{ color:'var(--muted-fg)', fontFamily:'var(--font-sans)' }}>{data.markdown}</pre>
          </details>}
        </div>
      </div>
    </div>
  )
}

// ─── AI Chat ──────────────────────────────────────────────────────────────────

const initMsgs: Record<Mode, Message[]> = {
  presentation: [{ role:'assistant', text:'Describe the presentation topic, audience, and tone to generate a deck. You can also upload an existing .pptx file.' }],
  excel:        [{ role:'assistant', text:'Ask a research question to build a sourced workbook, or upload an existing CSV or Excel file.' }],
  research:     [{ role:'assistant', text:'Ask a question to retrieve current claims with citations and confidence levels.' }],
  cleaner:      [{ role:'assistant', text:'Upload a CSV or Excel file first. Then describe the cleaning operation to apply it to your data.' }],
  planner:      [{ role:'assistant', text:'Tell me your thesis topic, deadline, and research method. I\'ll turn it into a phased timeline with milestones and deliverables you can refine conversationally.' }],
  extractor:    [{ role:'assistant', text:'Upload an image containing a table to extract its text and download structured CSV or Excel data.' }],
}

function initialChatMessages(mode: Mode, cleanerSession?: CleanerUploadResult | null): Message[] {
  if (mode === 'cleaner' && cleanerSession) {
    return [{ role:'assistant', text:`Your data is loaded (${cleanerSession.total_rows} rows across ${cleanerSession.columns.length} columns). Tell me what you’d like to clean, such as removing duplicates, filling missing values, or standardizing dates.` }]
  }
  return initMsgs[mode]
}

const suggestions: Record<Mode, string[]> = {
  presentation: ['Draft a slide on competition','Adjust tone to executive','Add a data slide'],
  excel:        ['Research revenue and profit for Microsoft','Research a comparison of recent EV sales'],
  research:     ['Compare current and historical evidence','Research recent government statistics','Find primary sources on this topic'],
  cleaner:      ['Normalize all dates','Remove duplicates','Fix missing values','Standardize currency'],
  planner:      ['Extend Phase 3 timeline','Add risk log','Generate Gantt chart','Add bibliography milestone'],
  extractor:    [],
}

function AIChatPanel({ mode, onSend, apiOnline, aiConfigured, excelData, cleanerSession }: { mode: Mode; onSend: (mode: Mode, text: string) => Promise<string>; apiOnline: boolean | null; aiConfigured: boolean; excelData?: ExcelResult | null; cleanerSession?: CleanerUploadResult | null }) {
  const [msgs, setMsgs] = useState<Message[]>(() => initialChatMessages(mode, cleanerSession))
  const [input, setInput] = useState('')
  const [typing, setTyping] = useState(false)
  const bottomRef = useRef<HTMLDivElement>(null)
  const cfg = MODES.find(m=>m.id===mode)!
  const inputLimit = mode === 'research' ? 1500 : mode === 'cleaner' ? 4000 : 12000
  const metricColumn = excelData?.columns.find(column => ['metric','measure','indicator'].includes(column.trim().toLowerCase()))
  const numericField = excelData?.columns.find(column => {
    if (/^(id|index|year|date|period)$/i.test(column.trim())) return false
    return excelData.rows.some(row => {
      const value = row[column]
      return typeof value === 'number' || /^\s*[$€£¥]?\(?[-+]?\d[\d,.]*(?:\.\d+)?\s*[kmbt]?\s*%?\)?\s*$/i.test(String(value ?? ''))
    })
  })
  const chartMetric = metricColumn ? [...new Set(excelData?.rows.map(row => String(row[metricColumn] ?? '').trim()).filter(Boolean))] : numericField ? [numericField] : []
  const dimension = excelData?.columns.find(column => /^(region|country|company|organization|entity|period|year|date|category|segment|market)$/i.test(column.trim()))
  const quickActions = mode === 'excel' && excelData
    ? chartMetric.slice(0, 4).map(metric => dimension ? `Chart ${metric} by ${dimension}` : `Chart ${metric}`)
    : suggestions[mode]

  useEffect(() => {
    if (mode !== 'cleaner' || !cleanerSession) return
    setMsgs(current => current.length === 1 && current[0].text === initMsgs.cleaner[0].text
      ? initialChatMessages(mode, cleanerSession)
      : current)
  }, [mode, cleanerSession])
  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior:'smooth' }) }, [msgs, typing])

  const send = async (text: string) => {
    if (!text.trim() || typing) return
    setMsgs(m => [...m, { role:'user', text }]); setInput(''); setTyping(true)
    try {
      const reply = await onSend(mode, text.trim())
      setMsgs(m => [...m, { role:'assistant', text: reply }])
    } catch (error) {
      setMsgs(m => [...m, { role:'assistant', text: `Something went wrong: ${error instanceof Error ? error.message : String(error)}` }])
    } finally {
      setTyping(false)
    }
  }


  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center justify-between gap-2 px-4 py-3 border-b shrink-0" style={{ borderColor:'var(--border)', background:'rgba(10,15,21,0.92)' }}>
        <div className="flex items-center gap-2">
          <div className="w-6 h-6 rounded-full flex items-center justify-center" style={{ background:'linear-gradient(135deg,var(--c-cln),var(--c-res))' }}>
            <Icon.Sparkle s={11} c="#fff"/>
          </div>
          <div>
            <div className="text-xs font-semibold" style={{ color:'var(--fg)' }}>AI Assistant</div>
            <div className="text-[10px]" style={{ color:cfg.color }}>{cfg.label}{apiOnline && !aiConfigured ? ' · setup needed' : ''}</div>
          </div>
        </div>
        <div aria-label={apiOnline && aiConfigured ? 'AI ready' : 'AI unavailable'} title={apiOnline ? aiConfigured ? 'AI ready' : 'Add a Gemini or xAI API key to enable generation' : 'API offline'} className="w-2 h-2 rounded-full" style={{ background:apiOnline && aiConfigured ? 'var(--green)' : 'var(--amber)', boxShadow:apiOnline && aiConfigured ? '0 0 0 4px rgba(22,163,74,0.12)' : 'none' }} />
      </div>

      <div className="flex-1 overflow-y-auto px-3 py-3 space-y-3" style={{ background:'linear-gradient(180deg,#0b1017 0%,#0e151d 100%)' }}>
        {msgs.map((m,i) => (
          <div key={i} className={`flex ${m.role==='user'?'justify-end':'justify-start'}`}>
            {m.role==='assistant' && (
              <div className="w-5 h-5 rounded-full flex items-center justify-center shrink-0 mr-2 mt-0.5" style={{ background:'linear-gradient(135deg,var(--c-cln),var(--c-res))' }}>
                <Icon.Sparkle s={9} c="#fff"/>
              </div>
            )}
            <div className="max-w-[88%] px-3 py-2 text-xs leading-relaxed"
              style={{ background: m.role==='user' ? cfg.color : 'var(--secondary)', color: m.role==='user' ? '#fff' : 'var(--fg)', borderRadius: m.role==='user' ? '14px 14px 4px 14px' : '4px 14px 14px 14px', border: m.role==='assistant' ? '1px solid var(--border)' : 'none', boxShadow: m.role==='user' ? `0 8px 18px ${cfg.raw}30` : 'none' }}>
              {m.text}
            </div>
          </div>
        ))}
        {typing && (
          <div className="flex items-start gap-2">
            <div className="w-5 h-5 rounded-full flex items-center justify-center shrink-0" style={{ background:'linear-gradient(135deg,var(--c-cln),var(--c-res))' }}>
              <Icon.Sparkle s={9} c="#fff"/>
            </div>
            <div className="px-3 py-2.5 text-xs ai-typing" style={{ background:'var(--secondary)', border:'1px solid var(--border)', borderRadius:'4px 12px 12px 12px' }}>Thinking…</div>
          </div>
        )}
        <div ref={bottomRef}/>
      </div>

      <div className="px-3 py-2 flex gap-1.5 overflow-x-auto shrink-0 border-t" style={{ borderColor:'var(--border)', background:'rgba(10,15,21,0.92)' }}>
        {quickActions.map(p => (
          <button key={p} onClick={()=>send(p)} className="shrink-0 px-2.5 py-1.5 rounded-full text-[10px] font-medium transition-all hover:-translate-y-0.5"
            style={{ background:'var(--secondary)', color:'var(--muted-fg)', border:'1px solid var(--border)' }}>
            {p}
          </button>
        ))}
      </div>

      <div className="px-3 pb-3 pt-1 shrink-0 border-t" style={{ borderColor:'var(--border)', background:'rgba(10,15,21,0.92)' }}>
        <div className="flex items-end gap-2 rounded-2xl p-2.5" style={{ background:'var(--secondary)', border:'1px solid var(--border)' }}>
          <textarea rows={1} maxLength={inputLimit} className="flex-1 bg-transparent resize-none outline-none text-xs leading-relaxed placeholder:opacity-40"
            style={{ color:'var(--fg)', maxHeight:'80px' }} placeholder={`Ask about your ${cfg.verb.toLowerCase()} task…`}
            value={input} onChange={e=>setInput(e.target.value)}
            onKeyDown={e=>{ if(e.key==='Enter'&&!e.shiftKey){ e.preventDefault(); send(input) } }}/>
          <span className="self-center text-[9px] tabular-nums" style={{ color:'var(--muted-fg)' }}>{input.length}/{inputLimit}</span>
          <button aria-label="Send message" disabled={!input.trim() || typing} onClick={()=>send(input)} className="w-8 h-8 rounded-xl flex items-center justify-center shrink-0 transition-all disabled:cursor-not-allowed"
            style={{ background: input.trim() && !typing ? cfg.color : '#cbd5e1', opacity: input.trim()?1:0.9, boxShadow: input.trim() ? `0 10px 20px ${cfg.raw}30` : 'none' }}>
            <Icon.Send s={12} c="#fff"/>
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Home / Mode Picker ───────────────────────────────────────────────────────

function ModePicker({ onSelect, apiOnline, aiConfigured }: { onSelect: (m: Mode) => void; apiOnline: boolean | null; aiConfigured: boolean }) {
  return (
    <div className="landing-screen">
      <nav className="landing-nav"><div className="landing-wrap landing-nav-inner">
        <a className="landing-logo" href="#top" aria-label="Multipanel home">Multipanel<sup>AI</sup></a>
        <div className="landing-links"><a href="#workspaces">Workspaces</a><a href="#workflow">How it works</a><a href="#about">About</a></div>
        <button className="landing-button landing-button-ghost landing-nav-cta" onClick={()=>onSelect('research')}>Open workspace <span aria-hidden="true">↗</span></button>
      </div></nav>

      <header className="landing-hero" id="top">
        <div className="landing-hero-glow" />
        <div className="landing-grid-lines" />
        <div className="landing-wrap landing-hero-content">
          <div className="landing-status"><span className={`status-dot ${apiOnline && aiConfigured ? 'is-ready' : ''}`} />
            {apiOnline === null ? 'Connecting to your workspace…' : !apiOnline ? 'Local API is offline' : aiConfigured ? 'Your AI workspace is ready' : 'Workspace connected · AI setup needed'}
          </div>
          <p className="landing-kicker">Research · Create · Organize</p>
          <h1>Where ideas <em>take shape.</em></h1>
          <p className="landing-hero-copy">From your first question to the final deliverable, Multipanel gives your research and creative work a place to come together.</p>
          <div className="landing-actions"><a className="landing-button landing-button-ghost" href="#workspaces">Explore workspaces <span aria-hidden="true">↓</span></a><button className="landing-button landing-button-solid" onClick={()=>onSelect('presentation')}>Start creating <span aria-hidden="true">↗</span></button></div>
          {!aiConfigured && <p className="landing-key-note">AI generation needs a configured Gemini or xAI API key. You can still preview and work with uploaded files.</p>}
          <div className="landing-scroll-note"><span /> Scroll to explore</div>
        </div>
      </header>

      <section className="landing-section landing-workspaces" id="workspaces"><div className="landing-wrap">
        <p className="landing-kicker">A considered set of tools</p>
        <h2>Everything you need to move<br className="desktop-break"/> from question to <em>deliverable.</em></h2>
        <div className="landing-workspace-grid">
          {MODES.map((m, i) => <button key={m.id} onClick={()=>onSelect(m.id)} className="landing-workspace-card" style={{'--workspace-accent': m.raw} as React.CSSProperties}>
            <div className="landing-card-top"><span className="landing-card-index">0{i + 1}</span><span className="landing-card-icon"><m.ModeIcon s={18}/></span></div>
            <h3>{m.label}</h3><p>{m.desc}</p>
            <span className="landing-card-link">Open workspace <span aria-hidden="true">↗</span></span>
          </button>)}
        </div>
      </div></section>

      <section className="landing-section landing-workflow" id="workflow"><div className="landing-wrap landing-split">
        <div><p className="landing-kicker">A clearer way to work</p><h2>Good work begins with <em>room to think.</em></h2>
          <p className="landing-lead">Bring your ideas, source material, and next steps into one calm workspace. Draft a presentation, shape a research plan, or turn raw data into something ready to use.</p>
          <ul className="landing-chips"><li>Research with sources</li><li>Build presentations</li><li>Make data useful</li><li>Keep your work together</li></ul>
          <button className="landing-button landing-button-ghost" onClick={()=>onSelect('planner')}>Plan a project <span aria-hidden="true">↗</span></button>
        </div>
        <div className="landing-preview" aria-label="Multipanel workspaces preview">
          <div className="preview-window-bar"><i/><i/><i/><span>YOUR WORKSPACE</span></div>
          <div className="preview-window-content"><div className="preview-eyebrow">MULTIPANEL / STUDIO</div><div className="preview-serif">Make space<br/>for the work.</div>
            <div className="preview-work-list">{MODES.slice(0, 3).map((m, i)=><div className="preview-work-row" key={m.id}><span>0{i + 1}</span><b>{m.label}</b><span>↗</span></div>)}</div>
          </div>
          <span className="preview-caption">One workspace. Many ways forward.</span>
        </div>
      </div></section>

      <section className="landing-cta" id="about"><div className="landing-cta-glow"/><div className="landing-wrap landing-cta-content">
        <p className="landing-kicker">Your next idea starts here</p><h2>Make something<br/><em>worth sharing.</em></h2>
        <p>Choose a workspace and take the next step, wherever you are in the process.</p>
        <button className="landing-button landing-button-solid" onClick={()=>onSelect('research')}>Enter Multipanel <span aria-hidden="true">↗</span></button>
      </div></section>

      <footer className="landing-footer"><div className="landing-wrap landing-footer-inner"><a className="landing-logo" href="#top">Multipanel<sup>AI</sup></a><span>Thoughtful tools for meaningful work.</span><a href="#top">Back to top ↑</a></div></footer>
    </div>
  )
}

// ─── App Shell ────────────────────────────────────────────────────────────────

export default function App() {
  const [mode, setMode] = useState<Mode|null>(null)
  const cfg = mode ? MODES.find(m=>m.id===mode)! : null

  const [presentationResult, setPresentationResult] = useState<PresentationInfo|null>(null)
  const [excelResult, setExcelResult] = useState<ExcelResult|null>(null)
  const [researchReport, setResearchReport] = useState<VerifiedResearchReport|null>(null)
  const [planResult, setPlanResult] = useState<ResearchPlan|null>(null)
  const [cleanerSession, setCleanerSession] = useState<CleanerUploadResult|null>(null)
  const [extractorResult, setExtractorResult] = useState<ExtractorResult|null>(null)
  const [apiOnline, setApiOnline] = useState<boolean|null>(null)
  const [aiConfigured, setAiConfigured] = useState(false)
  const [providers, setProviders] = useState<Record<AIProvider, boolean>>({ gemini: false, grok: false })
  const [provider, setProvider] = useState<AIProvider>(() => window.localStorage.getItem('multipanel-provider') === 'gemini' ? 'gemini' : 'grok')
  const [presentationStyle, setPresentationStyle] = useState<keyof typeof presentationThemes>(() => {
    const saved = window.localStorage.getItem('multipanel-presentation-style')
    return saved && saved in presentationThemes ? saved as keyof typeof presentationThemes : 'Editorial'
  })
  const [includeWebImages, setIncludeWebImages] = useState(false)
  const [uiError, setUiError] = useState('')

  const runUpload = async <T,>(operation: () => Promise<T>, apply: (result: T) => void) => {
    setUiError('')
    try { apply(await operation()) }
    catch (error) { setUiError(error instanceof Error ? error.message : String(error)) }
  }

  useEffect(() => {
    let active = true
    const check = async () => {
      const online = await checkApiHealth()
      if (active) {
        setApiOnline(online.online); setAiConfigured(online.aiConfigured); setProviders(online.providers)
        setProvider(current => online.providers[current] ? current : online.providers.grok ? 'grok' : 'gemini')
      }
    }
    void check()
    const timer = window.setInterval(check, 30000)
    return () => { active = false; window.clearInterval(timer) }
  }, [])

  useEffect(() => { window.localStorage.setItem('multipanel-provider', provider) }, [provider])
  useEffect(() => { window.localStorage.setItem('multipanel-presentation-style', presentationStyle) }, [presentationStyle])

  const handleChatSend = async (currentMode: Mode, text: string): Promise<string> => {
    switch (currentMode) {
      case 'presentation': {
        if (!presentationResult) {
          const result = await createPresentation(text, 6, 'general audience', 'clear and persuasive', provider, presentationStyle, includeWebImages)
          setPresentationResult(result)
          return `Drafted "${result.title}" with ${result.slides.length} slides. Open the canvas to review, then ask me for changes.`
        }
        const revised = await refinePresentation(presentationResult, text, provider, presentationStyle, includeWebImages)
        setPresentationResult(revised)
        return `Updated the deck: "${revised.title}" now has ${revised.slides.length} slides.`
      }
      case 'excel': {
        if (excelResult && /\b(chart|graph|plot|visuali[sz]e)\b/i.test(text)) {
          const result = await refineExcelChart(excelResult, text)
          setExcelResult(result)
          return 'Updated the workbook chart using a comparable measure, explicit units, and a chart type suited to the data.'
        }
        if (excelResult && !/^(new|research|search|build|create)\b/i.test(text.trim())) {
          return 'This workbook can be charted or downloaded. For another sourced dataset, start your request with “Research…”. Formula and pivot-sheet editing are not available in this MVP yet.'
        }
        const result = await createExcelWorkbook(text, provider)
        setExcelResult(result)
        return result.summary || `Built a workbook with ${result.total_rows} rows.`
      }
      case 'research': {
        const report = await createVerifiedResearch(text, 'latest available', 'credible primary sources', provider)
        setResearchReport(report)
        return report.answer
      }
      case 'planner': {
        if (!planResult) {
          const plan = await createResearchPlan(text, 'Research', 'Thesis/Dissertation', 12, provider)
          setPlanResult(plan)
          return `Created a research plan: "${plan.title}" across ${plan.timeline.length} phases.`
        }
        const revised = await refineResearchPlan(planResult, text, provider)
        setPlanResult(revised)
        return `Updated the plan: "${revised.title}".`
      }
      case 'cleaner': {
        if (!cleanerSession) {
          return 'Upload a CSV or Excel file using the upload area in the Data Cleaner panel first.'
        }
        const result = await sendCleanerChat(cleanerSession.session_id, text, provider)
        setCleanerSession(prev => prev ? { ...prev, columns: result.columns, rows: result.rows, total_rows: result.total_rows, missing_by_column: result.missing_by_column, dtypes: result.dtypes } : prev)
        return result.log.join('\n')
      }
      case 'extractor':
        return 'Upload an image in the Document Reader panel to extract its table.'
      default:
        return "I couldn't process that request."
    }
  }

  return (
    <div className="min-h-screen flex flex-col" style={{ background:'var(--background)', fontFamily:'var(--font-sans)' }}>
      {mode && <header className="flex items-center justify-between px-4 py-2.5 border-b shrink-0"
        style={{ background:'rgba(7,9,12,0.92)', borderColor:'var(--border)', backdropFilter:'blur(12px)', zIndex:10 }}>
        <div className="flex items-center gap-3">
          <button onClick={()=>setMode(null)} className="flex items-center gap-2 hover:opacity-80 transition-opacity">
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background:'linear-gradient(135deg,var(--c-cln),var(--c-res))' }}>
              <Icon.Sparkle s={13} c="#fff"/>
            </div>
            <span className="font-semibold text-sm" style={{ color:'var(--fg)' }}>Multipanel</span>
          </button>
          {cfg && (
            <>
              <span style={{ color:'var(--border)' }}>/</span>
              <div className="flex items-center gap-1.5">
                <cfg.ModeIcon s={13} c={cfg.color}/>
                <span className="text-sm font-medium" style={{ color:cfg.color }}>{cfg.label}</span>
              </div>
            </>
          )}
        </div>

        {mode && (
          <div className="flex items-center gap-1 p-1 rounded-xl absolute left-1/2 -translate-x-1/2" style={{ background:'var(--secondary)', border:'1px solid var(--border)', boxShadow:'0 8px 20px rgba(15, 23, 42, 0.04)' }}>
            {MODES.map(m => (
              <button key={m.id} onClick={()=>setMode(m.id)}
                className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-medium transition-all"
                style={{ background: mode===m.id ? m.color : 'transparent', color: mode===m.id ? '#fff' : 'var(--muted-fg)', boxShadow: mode===m.id ? `0 6px 16px ${m.raw}35` : 'none' }}>
                <m.ModeIcon s={12}/>
                {m.verb}
              </button>
            ))}
          </div>
        )}

        <div className="flex items-center gap-2">
        {cfg && <div className="flex items-center gap-1.5 text-[10px] font-mono" style={{ color:'var(--muted-fg)' }}>
            <div className="w-1.5 h-1.5 rounded-full" style={{ background:apiOnline ? 'var(--green)' : apiOnline === false ? 'var(--red)' : 'var(--amber)', boxShadow:apiOnline ? '0 0 5px var(--green)' : 'none' }}/>
            {apiOnline === null ? 'Checking API…' : apiOnline ? 'Backend connected' : 'Backend offline'}
          </div>}
          <div className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold" style={{ background: cfg?.color || 'var(--secondary)', color:'#fff' }}>P</div>
          <label className="flex items-center gap-2 rounded-lg border px-2.5 py-1.5 text-[11px]" style={{ borderColor:'var(--border)', color:'var(--muted-fg)', background:'var(--panel)' }}>
            <span>AI</span>
            <select aria-label="AI provider" value={provider} onChange={event=>setProvider(event.target.value as AIProvider)} className="bg-transparent outline-none" style={{ color:'var(--fg)' }}>
              <option value="gemini" disabled={!providers.gemini}>Gemini{providers.gemini ? '' : ' · not configured'}</option>
              <option value="grok" disabled={!providers.grok}>Grok{providers.grok ? '' : ' · not configured'}</option>
            </select>
          </label>
        </div>
      </header>}

      {uiError && <div role="alert" className="mx-4 mt-3 rounded-lg px-3 py-2 text-xs flex items-center justify-between" style={{ background:'rgba(239,68,68,0.08)', color:'#b91c1c', border:'1px solid rgba(239,68,68,0.2)' }}>
        <span>{uiError}</span><button onClick={()=>setUiError('')} className="ml-3 font-semibold">Dismiss</button>
      </div>}

      {!mode ? (
        <ModePicker onSelect={setMode} apiOnline={apiOnline} aiConfigured={aiConfigured} />
      ) : (
        <main className="app-workspace-main flex-1 flex overflow-hidden p-4">
          <div className="workspace-main-panel flex-1 overflow-hidden rounded-[22px] border" style={{ background:'linear-gradient(160deg,#0d141c 0%,#0a0f15 100%)', borderColor:'var(--border)', minWidth:0, boxShadow:'0 20px 50px rgba(0,0,0,0.28)' }}>
            {mode==='presentation' && <PresentationPanel data={presentationResult} onUpload={file => runUpload(() => uploadPresentationFile(file), setPresentationResult)} styleName={presentationStyle} onStyleChange={setPresentationStyle} includeWebImages={includeWebImages} onToggleWebImages={()=>setIncludeWebImages(value=>!value)}/>}
            {mode==='excel'        && <ExcelPanel data={excelResult} onUpload={file => runUpload(() => uploadExcelFile(file), setExcelResult)}/>}
            {mode==='research'     && <ResearchPanel data={researchReport} onAsk={async question => setResearchReport(await createVerifiedResearch(question, 'latest available', 'credible primary sources', provider))}/>}
            {mode==='cleaner'      && <CleanerPanel session={cleanerSession} onUpload={file => runUpload(() => uploadCleanerFile(file), setCleanerSession)}/>}
            {mode==='planner'      && <PlannerPanel data={planResult}/>}
            {mode==='extractor'    && <ExtractorPanel data={extractorResult} onUpload={file => runUpload(() => uploadExtractorImage(file, provider), setExtractorResult)}/>}
          </div>
          <div className="workspace-assistant-panel w-[350px] shrink-0 flex flex-col overflow-hidden ml-4 rounded-[22px] border" style={{ background:'var(--panel)', borderColor:'var(--border)', boxShadow:'0 20px 50px rgba(0,0,0,0.28)' }}>
            <AIChatPanel key={mode} mode={mode} onSend={handleChatSend} apiOnline={apiOnline} aiConfigured={aiConfigured} excelData={mode==='excel' ? excelResult : null} cleanerSession={mode==='cleaner' ? cleanerSession : null}/>
          </div>
        </main>
      )}

      {mode && (
        <footer className="flex items-center justify-between px-4 py-1.5 border-t shrink-0"
          style={{ background:'rgba(10,15,21,0.9)', borderColor:'var(--border)' }}>
          <span className="text-[10px] font-mono" style={{ color:'var(--muted-fg)' }}>
            {mode==='presentation'&&(presentationResult ? `${presentationResult.slides.length} slides · 16:9 widescreen` : 'No presentation yet')}
            {mode==='excel'&&(excelResult ? `${excelResult.total_rows} rows · ${excelResult.sources.length} sources` : 'No workbook yet')}
            {mode==='research'&&(researchReport ? `${researchReport.claims.length} claims · ${researchReport.sources.length} sources` : 'No research yet')}
            {mode==='cleaner'&&(cleanerSession ? `${cleanerSession.total_rows} rows · ${cleanerSession.columns.length} columns` : 'No file uploaded yet')}
            {mode==='planner'&&(planResult ? `${planResult.timeline.length} phases · ${planResult.timeline.reduce((s,p)=>s+(p.duration_weeks||0),0)} weeks` : 'No plan yet')}
            {mode==='extractor'&&(extractorResult ? `${extractorResult.total_rows} rows extracted` : 'No image uploaded yet')}
          </span>
          <span className="text-[10px]" style={{ color:'var(--muted-fg)' }}>{`AI uses ${provider === 'grok' ? 'Grok' : 'Gemini'}${mode==='research'||mode==='excel' ? ' with live web search' : ''}`} · Results remain in this session</span>
        </footer>
      )}
    </div>
  )
}
