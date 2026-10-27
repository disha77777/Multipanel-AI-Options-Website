// Client for the Multipanel Python backend (see Multipanel/api_server.py).
// Use same-origin API routes by default so production deployments work without
// hard-coding a developer's localhost address. Set VITE_API_BASE_URL when the
// API is hosted on a separate origin.
const API_BASE = ((import.meta as any).env?.VITE_API_BASE_URL || '').replace(/\/+$/, '')
export type AIProvider = 'gemini' | 'grok'

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${API_BASE}${path}`, init)
  if (!res.ok) {
    let detail = res.statusText
    try {
      const body = await res.json()
      if (typeof body.detail === 'string') detail = body.detail
      else if (Array.isArray(body.detail)) {
        detail = body.detail
          .map((item: { msg?: string; loc?: Array<string | number> }) => {
            const field = item.loc?.slice(1).join('.')
            return `${field ? `${field}: ` : ''}${item.msg || 'Invalid input'}`
          })
          .join('; ')
      }
    } catch {
      // ignore non-JSON error bodies
    }
    throw new Error(detail)
  }
  return res.json() as Promise<T>
}

function postJson<T>(path: string, payload: unknown): Promise<T> {
  return request<T>(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

// ─── Presentation Studio ──────────────────────────────────────────────────────

export interface SlideVisual { type: 'process' | 'cycle' | 'comparison'; items: { label: string; detail: string }[]; image_query?: string }
export interface SlideImage {
  url?: string
  file_page_url?: string
  artist?: string
  license: string
  license_url?: string
  file_name?: string
  kind?: 'report_chart'
  data?: string
  chart_title?: string
  report_title?: string
  publisher?: string
  page_number?: number
  source_url?: string
}
export interface Slide { title: string; bullets: string[]; visual?: SlideVisual | null; image?: SlideImage }
export interface PresentationInfo { title: string; slides: Slide[]; chart_notice?: string }

export const createPresentation = (topic: string, slideCount = 6, audience = 'general audience', tone = 'clear and persuasive', provider: AIProvider = 'grok', style = 'Editorial', includeReportCharts = false) =>
  postJson<PresentationInfo>('/api/presentation', { topic, slide_count: slideCount, audience, tone, provider, style, include_report_charts: includeReportCharts })

export const refinePresentation = (presentationInfo: PresentationInfo, revision: string, provider: AIProvider = 'grok', style = 'Editorial', includeReportCharts = false) =>
  postJson<PresentationInfo>('/api/presentation/refine', { presentation: presentationInfo, revision, provider, style, include_report_charts: includeReportCharts })

export const presentationExportUrl = () => `${API_BASE}/api/presentation/export`

export const uploadPresentationFile = (file: File) => {
  const form = new FormData()
  form.append('file', file)
  return request<PresentationInfo>('/api/presentation/upload', { method: 'POST', body: form })
}

// ─── Excel Research ────────────────────────────────────────────────────────────

export interface TablePreview { columns: string[]; rows: Record<string, unknown>[]; total_rows: number }
export interface ExcelResult extends TablePreview { summary: string; sources: string[]; query?: string; session_id?: string; chart_available?: boolean }

export const createExcelWorkbook = (query: string, provider: AIProvider = 'grok') => postJson<ExcelResult>('/api/excel', { query, provider })
export const refineExcelChart = (data: ExcelResult, chart_request: string) => postJson<ExcelResult>('/api/excel/chart', { session_id: data.session_id, ...(!data.session_id ? data : {}), chart_request })
export const excelDownloadUrl = () => `${API_BASE}/api/excel/download`
export const excelChartUrl = () => `${API_BASE}/api/excel/chart`

export const uploadExcelFile = (file: File) => {
  const form = new FormData()
  form.append('file', file)
  return request<ExcelResult>('/api/excel/upload', { method: 'POST', body: form })
}

// ─── Verified Research ──────────────────────────────────────────────────────────

export interface ResearchSource { id: string; title: string; url: string; publisher: string; published_date: string }
export interface ResearchClaim { text: string; kind: string; confidence: string; source_ids: string[] }
export interface VerifiedResearchReport {
  answer: string
  as_of: string
  verified_at: string
  claims: ResearchClaim[]
  sources: ResearchSource[]
  limitations: string[]
}

export const createVerifiedResearch = (question: string, dateRange = 'latest available', sourcePreference = 'credible primary sources', provider: AIProvider = 'grok') =>
  postJson<VerifiedResearchReport>('/api/research', { question, date_range: dateRange, source_preference: sourcePreference, provider })

export async function checkApiHealth(): Promise<{ online: boolean; aiConfigured: boolean; providers: Record<AIProvider, boolean> }> {
  try {
    const result = await request<{ status: string; ai_configured?: boolean; providers?: Partial<Record<AIProvider, boolean>> }>('/api/health')
    return { online: result.status === 'ok', aiConfigured: result.ai_configured === true,
      providers: { gemini: result.providers?.gemini === true, grok: result.providers?.grok === true } }
  } catch {
    return { online: false, aiConfigured: false, providers: { gemini: false, grok: false } }
  }
}

// ─── Research Planner ──────────────────────────────────────────────────────────

export interface ResearchPlan {
  title: string
  thesis_statement: string
  field: string
  objectives: { title: string; description: string }[]
  literature_review: { scope: string; key_themes: string[]; sources_to_explore: string[] }
  methodology: { research_design: string; data_collection: string[]; analysis_approach: string; sample_size_or_scope: string }
  timeline: { phase: string; description: string; duration_weeks: number }[]
  resources_needed: { resource: string; justification: string }[]
  expected_outcomes: string[]
  potential_challenges: { challenge: string; mitigation: string }[]
  references_to_start?: string[]
}

export const createResearchPlan = (thesisStatement: string, fieldOfStudy = 'Research', researchType = 'Thesis/Dissertation', durationMonths = 12, provider: AIProvider = 'grok') =>
  postJson<ResearchPlan>('/api/planner', { thesis_statement: thesisStatement, field_of_study: fieldOfStudy, research_type: researchType, duration_months: durationMonths, provider })

export const refineResearchPlan = (plan: ResearchPlan, revision: string, provider: AIProvider = 'grok') =>
  postJson<ResearchPlan>('/api/planner/refine', { plan, revision, provider })

// ─── Data Cleaner ───────────────────────────────────────────────────────────────

export interface CleanerUploadResult extends TablePreview {
  session_id: string
  source_rows: number
  missing_by_column: Record<string, number>
  dtypes: Record<string, string>
}
export interface CleanerChatResult extends TablePreview {
  log: string[]
  missing_by_column: Record<string, number>
  dtypes: Record<string, string>
}

export const uploadCleanerFile = (file: File) => {
  const form = new FormData()
  form.append('file', file)
  return request<CleanerUploadResult>('/api/cleaner/upload', { method: 'POST', body: form })
}

export const sendCleanerChat = (sessionId: string, message: string, provider: AIProvider = 'grok') =>
  postJson<CleanerChatResult>('/api/cleaner/chat', { session_id: sessionId, message, provider })

export const cleanerDownloadUrl = (sessionId: string) => `${API_BASE}/api/cleaner/download/${sessionId}`

// ─── Document Reader ────────────────────────────────────────────────────────────

export interface ExtractorResult extends TablePreview { markdown: string }

export const uploadExtractorImage = (file: File, provider: AIProvider = 'grok') => {
  const form = new FormData()
  form.append('file', file)
  form.append('provider', provider)
  return request<ExtractorResult>('/api/extractor/upload', { method: 'POST', body: form })
}

export const extractorDownloadUrl = (fmt: 'csv' | 'xlsx') => `${API_BASE}/api/extractor/download/${fmt}`
