import type {
  AppSettings,
  CapabilitiesResponse,
  DataTemplate,
  EngineInfo,
  GenerateResult,
  PreviewResult,
  StylePreset,
  Task,
} from './types'

const TOKEN_KEY = 'format_studio_token'

function initToken(): string {
  const params = new URLSearchParams(window.location.search)
  const fromUrl = params.get('token')
  if (fromUrl) {
    localStorage.setItem(TOKEN_KEY, fromUrl)
    params.delete('token')
    const next = `${window.location.pathname}${params.toString() ? `?${params}` : ''}${window.location.hash}`
    window.history.replaceState({}, '', next)
    return fromUrl
  }
  return localStorage.getItem(TOKEN_KEY) || ''
}

export const token = initToken()

export class ApiError extends Error {
  code: string
  status: number
  constructor(message: string, code = 'REQUEST_FAILED', status = 0) {
    super(message)
    this.code = code
    this.status = status
  }
}

function authHeaders(extra: Record<string, string> = {}): Record<string, string> {
  return token ? { 'X-App-Token': token, ...extra } : extra
}

async function parseError(response: Response): Promise<ApiError> {
  let code = `HTTP_${response.status}`
  let message = `请求失败（${response.status}）`
  try {
    const payload = await response.json()
    const detail = payload?.detail ?? payload
    if (typeof detail === 'string') {
      message = detail
    } else if (detail && typeof detail === 'object') {
      code = detail.code || code
      message = detail.message || message
    } else if (payload?.message) {
      message = payload.message
    }
  } catch {
    // ignore
  }
  return new ApiError(message, code, response.status)
}

export async function api<T>(path: string, options: RequestInit = {}): Promise<T> {
  const headers = authHeaders(options.body instanceof FormData ? {} : { 'Content-Type': 'application/json' })
  const response = await fetch(`/api/v1${path}`, {
    ...options,
    headers: { ...headers, ...(options.headers as Record<string, string> | undefined) },
  })
  if (!response.ok) throw await parseError(response)
  if (response.status === 204) return undefined as T
  const type = response.headers.get('content-type') || ''
  if (type.includes('application/json')) return (await response.json()) as T
  return (await response.text()) as T
}

export function withToken(path: string): string {
  if (!token) return path
  const separator = path.includes('?') ? '&' : '?'
  return `${path}${separator}token=${encodeURIComponent(token)}`
}

declare global {
  interface Window {
    pywebview?: {
      api?: {
        save_file?: (
          url: string,
          filename: string,
        ) => Promise<{ ok: boolean; path?: string; error?: string; cancelled?: boolean }>
      }
    }
  }
}

export async function download(path: string, filename = 'download') {
  const relative = path.includes('token=') ? path : withToken(path)
  const url = new URL(relative, window.location.href).toString()
  try {
    const bridge = window.pywebview?.api?.save_file
    if (bridge) {
      const result = await bridge(url, filename)
      if (result?.ok || result?.cancelled) return
      throw new Error(result?.error || '??????')
    }

    const response = await fetch(url)
    if (!response.ok) throw await parseError(response)
    const blob = await response.blob()
    const objectUrl = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = objectUrl
    anchor.download = filename
    document.body.appendChild(anchor)
    anchor.click()
    anchor.remove()
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 60_000)
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    window.alert(`下载失败：${message}`)
    throw error
  }
}

export const apiClient = {
  capabilities: () => api<CapabilitiesResponse>('/capabilities'),
  engines: () => api<Record<string, EngineInfo>>('/engines'),
  probeEngines: () => api<Record<string, EngineInfo>>('/engines/probe', { method: 'POST' }),
  tasks: () => api<Task[]>('/tasks?limit=100'),
  task: (id: string) => api<Task>(`/tasks/${id}`),
  createTask: (form: FormData) => api<Task>('/tasks', { method: 'POST', body: form }),
  cancelTask: (id: string) => api<Task>(`/tasks/${id}/cancel`, { method: 'POST' }),
  retryTask: (id: string) => api<Task>(`/tasks/${id}/retry`, { method: 'POST' }),
  openFolder: (id: string) => api<{ ok: boolean; path: string }>(`/tasks/${id}/open-folder`, { method: 'POST' }),
  settings: () => api<AppSettings>('/settings'),
  updateSettings: (payload: Partial<AppSettings>) =>
    api<AppSettings>('/settings', { method: 'PUT', body: JSON.stringify(payload) }),
  cleanup: () => api<{ task_dirs_removed: number }>('/system/cleanup', { method: 'POST' }),
  templates: () => api<DataTemplate[]>('/templates'),
  uploadTemplate: (form: FormData) => api<DataTemplate>('/templates', { method: 'POST', body: form }),
  deleteTemplate: (id: string) => api<{ ok: boolean }>(`/templates/${id}`, { method: 'DELETE' }),
  scanTemplate: (id: string) => api<{ fields: DataTemplate['fields'] }>(`/templates/${id}/scan`, { method: 'POST' }),
  previewTemplate: (id: string, form: FormData) =>
    api<PreviewResult>(`/templates/${id}/preview`, { method: 'POST', body: form }),
  generateTemplate: (id: string, form: FormData) =>
    api<GenerateResult>(`/templates/${id}/generate`, { method: 'POST', body: form }),
  styles: () => api<StylePreset[]>('/styles'),
  createStyle: (payload: { name: string; applies_to: string[]; data: Record<string, any> }) =>
    api<StylePreset>('/styles', { method: 'POST', body: JSON.stringify(payload) }),
  updateStyle: (id: string, payload: { name: string; applies_to: string[]; data: Record<string, any> }) =>
    api<StylePreset>(`/styles/${id}`, { method: 'PUT', body: JSON.stringify(payload) }),
  deleteStyle: (id: string) => api<{ ok: boolean }>(`/styles/${id}`, { method: 'DELETE' }),
  importStyle: (form: FormData) => api<StylePreset & { warnings: string[] }>('/styles/import', { method: 'POST', body: form }),
  previewStyle: (id: string, fmt?: string) =>
    api<PreviewResult>(`/styles/${id}/preview${fmt ? `?fmt=${fmt}` : ''}`, { method: 'POST' }),
}

export function sseUrl(taskId: string): string {
  return withToken(`/api/v1/tasks/${taskId}/events`)
}

export function diagnosticsUrl(): string {
  return withToken('/api/v1/system/diagnostics')
}