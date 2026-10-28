export type FormatId = 'txt' | 'md' | 'docx' | 'pdf' | 'wps' | 'csv' | 'xlsx' | 'pptx'

export interface FormatInfo {
  id: FormatId
  label: string
  kind: 'document' | 'table' | 'presentation'
  icon: string
}

export interface Capability {
  source: FormatId
  target: FormatId
  status: 'supported' | 'beta' | 'disabled'
  quality: 'high' | 'normal' | 'content' | 'beta' | null
  engines: string[]
  modes: string[]
  reason: string | null
}

export interface EngineInfo {
  id: string
  label: string
  available: boolean
  path: string | null
  version: string | null
  busy: boolean
  error: string | null
}

export interface CapabilitiesResponse {
  formats: FormatInfo[]
  matrix: Capability[]
  engines: Record<string, EngineInfo>
}

export interface TaskItem {
  id: string
  task_id: string
  source_name: string
  source_format: FormatId
  target_format: FormatId
  mode: string
  status: TaskStatus
  progress: number
  engine: string | null
  error_code: string | null
  error_message: string | null
  warning: string | null
  outputs: string[]
  output_count: number
  download_url: string
  download_name: string
  published_dir: string | null
  meta: Record<string, unknown>
  created_at: string
  updated_at: string
}

export type TaskStatus =
  | 'queued'
  | 'running'
  | 'succeeded'
  | 'partial'
  | 'failed'
  | 'canceled'
  | 'interrupted'

export interface Task {
  id: string
  status: TaskStatus
  options: {
    target_format?: FormatId
    mode?: string
    style_id?: string | null
    csv_all_text?: boolean
    sheet_name?: string | null
  }
  items: TaskItem[]
  progress: number
  queue_position: number
  created_at: string
  updated_at: string
}

export interface PlaceholderField {
  name: string
  count: number
  sample: string
}

export interface DataTemplate {
  id: string
  name: string
  format: 'docx' | 'xlsx' | 'pptx'
  path: string
  original_name: string
  fields: PlaceholderField[]
  builtin?: boolean
  created_at: string
  updated_at: string
}

export interface StylePreset {
  id: string
  name: string
  applies_to: string[]
  data: Record<string, any>
  builtin: boolean
  created_at: string
  updated_at: string
}

export interface AppSettings {
  output_dir: string
  retention_days: number
  csv_all_text: boolean
}

export interface PreviewResult {
  pages: string[]
  warnings: string[]
  format?: string
  values?: Record<string, string>
}

export interface GenerateResult {
  generation_id: string
  files: { name: string; url: string }[]
  archive_url: string | null
  archive_name: string | null
  generated: number
  rows: number
  warnings: string[]
  headers: string[]
}