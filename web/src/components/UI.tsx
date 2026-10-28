import type { ReactNode } from 'react'
import { AlertCircle, CheckCircle2, Info, Loader2, X } from 'lucide-react'
import type { TaskStatus } from '../types'
import { cn, statusLabels, statusTone } from '../utils'

export function PageHeader({
  title,
  description,
  actions,
}: {
  title: string
  description?: string
  actions?: ReactNode
}) {
  return (
    <div className="mb-7 flex items-start justify-between gap-6">
      <div>
        <h1 className="text-[28px] font-semibold tracking-tight text-slate-900">{title}</h1>
        {description ? <p className="mt-1.5 max-w-2xl text-sm leading-6 text-slate-500">{description}</p> : null}
      </div>
      {actions ? <div className="flex shrink-0 items-center gap-2">{actions}</div> : null}
    </div>
  )
}

export function Card({ children, className }: { children: ReactNode; className?: string }) {
  return <div className={cn('card p-6', className)}>{children}</div>
}

export function StatusBadge({ status }: { status: TaskStatus }) {
  return <span className={cn('badge', statusTone(status))}>{statusLabels[status]}</span>
}

export function ProgressBar({ value, className }: { value: number; className?: string }) {
  return (
    <div className={cn('h-1.5 w-full overflow-hidden rounded-full bg-slate-100', className)}>
      <div
        className="h-full rounded-full bg-brand-600 transition-all duration-500"
        style={{ width: `${Math.max(0, Math.min(100, value))}%` }}
      />
    </div>
  )
}

export function Spinner({ label }: { label?: string }) {
  return (
    <span className="inline-flex items-center gap-2 text-sm text-slate-500">
      <Loader2 size={15} className="animate-spin" />
      {label}
    </span>
  )
}

export function Notice({
  tone = 'info',
  children,
  className,
}: {
  tone?: 'info' | 'success' | 'error' | 'warning'
  children: ReactNode
  className?: string
}) {
  const styles = {
    info: 'border-brand-100 bg-brand-50/70 text-brand-800',
    success: 'border-emerald-100 bg-emerald-50/70 text-emerald-800',
    error: 'border-rose-100 bg-rose-50/70 text-rose-800',
    warning: 'border-amber-100 bg-amber-50/70 text-amber-800',
  }
  const Icon = tone === 'success' ? CheckCircle2 : tone === 'error' ? AlertCircle : tone === 'warning' ? AlertCircle : Info
  return (
    <div className={cn('flex items-start gap-2.5 rounded-xl border px-3.5 py-3 text-sm leading-6', styles[tone], className)}>
      <Icon size={16} className="mt-0.5 shrink-0" />
      <div className="min-w-0 flex-1">{children}</div>
    </div>
  )
}

export function EmptyState({
  icon,
  title,
  description,
  action,
}: {
  icon?: ReactNode
  title: string
  description?: string
  action?: ReactNode
}) {
  return (
    <div className="flex flex-col items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-white/60 px-6 py-16 text-center">
      {icon ? <div className="mb-3 text-slate-300">{icon}</div> : null}
      <div className="text-sm font-medium text-slate-700">{title}</div>
      {description ? <div className="mt-1 max-w-md text-xs leading-5 text-slate-400">{description}</div> : null}
      {action ? <div className="mt-4">{action}</div> : null}
    </div>
  )
}

export function Modal({
  open,
  title,
  children,
  onClose,
  width = 'max-w-2xl',
}: {
  open: boolean
  title: string
  children: ReactNode
  onClose: () => void
  width?: string
}) {
  if (!open) return null
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/30 p-6 backdrop-blur-sm">
      <div className={cn('max-h-[88vh] w-full overflow-hidden rounded-2xl bg-white shadow-2xl', width)}>
        <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4">
          <h2 className="text-base font-semibold text-slate-900">{title}</h2>
          <button className="btn-ghost !p-2" onClick={onClose} aria-label="关闭">
            <X size={18} />
          </button>
        </div>
        <div className="max-h-[calc(88vh-64px)] overflow-y-auto p-6">{children}</div>
      </div>
    </div>
  )
}