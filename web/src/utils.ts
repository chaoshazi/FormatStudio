import clsx, { type ClassValue } from 'clsx'
import { twMerge } from 'tailwind-merge'
import type { FormatId, TaskStatus } from './types'

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export const formatLabels: Record<FormatId, string> = {
  txt: 'TXT',
  md: 'Markdown',
  docx: 'Word',
  pdf: 'PDF',
  wps: 'WPS',
  csv: 'CSV',
  xlsx: 'Excel',
  pptx: 'PPT',
}

export const qualityLabels: Record<string, string> = {
  high: '高保真',
  normal: '结构转换',
  content: '仅内容',
  beta: 'Beta',
}

export const statusLabels: Record<TaskStatus, string> = {
  queued: '排队中',
  running: '转换中',
  succeeded: '已完成',
  partial: '部分完成',
  failed: '失败',
  canceled: '已取消',
  interrupted: '已中断',
}

export function statusTone(status: TaskStatus): string {
  switch (status) {
    case 'succeeded':
      return 'bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200'
    case 'partial':
      return 'bg-amber-50 text-amber-700 ring-1 ring-amber-200'
    case 'failed':
      return 'bg-rose-50 text-rose-700 ring-1 ring-rose-200'
    case 'running':
      return 'bg-brand-50 text-brand-700 ring-1 ring-brand-200'
    case 'canceled':
    case 'interrupted':
      return 'bg-slate-100 text-slate-600 ring-1 ring-slate-200'
    default:
      return 'bg-slate-50 text-slate-600 ring-1 ring-slate-200'
  }
}

export function formatBytes(size?: number): string {
  if (!size || size <= 0) return '—'
  const units = ['B', 'KB', 'MB', 'GB']
  let value = size
  let index = 0
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024
    index += 1
  }
  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`
}

export function formatTime(value?: string): string {
  if (!value) return '—'
  try {
    return new Date(value).toLocaleString('zh-CN', { hour12: false })
  } catch {
    return value
  }
}

export function metaSummary(meta: Record<string, unknown>): string[] {
  const labels: Record<string, string> = {
    pages: '页',
    paragraphs: '段落',
    tables: '表格',
    words: '字',
    sheets: '工作表',
    rows: '行',
    slides: '页幻灯片',
    lines: '行',
    chars: '字符',
  }
  return Object.entries(meta)
    .filter(([key, value]) => key in labels && typeof value === 'number' && value > 0)
    .map(([key, value]) => `${value} ${labels[key]}`)
}

export function extOf(name: string): string {
  const index = name.lastIndexOf('.')
  return index >= 0 ? name.slice(index + 1).toLowerCase() : ''
}

export function formatOf(name: string): FormatId | null {
  const ext = extOf(name)
  const mapping: Record<string, FormatId> = {
    txt: 'txt',
    text: 'txt',
    md: 'md',
    markdown: 'md',
    docx: 'docx',
    pdf: 'pdf',
    wps: 'wps',
    csv: 'csv',
    xlsx: 'xlsx',
    pptx: 'pptx',
  }
  return mapping[ext] || null
}