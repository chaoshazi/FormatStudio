import type { Capability, FormatId, FormatInfo } from '../types'
import { cn, formatLabels, qualityLabels } from '../utils'

export default function FormatGrid({
  formats,
  matrix,
  sources,
  value,
  onChange,
}: {
  formats: FormatInfo[]
  matrix: Capability[]
  sources: (FormatId | null)[]
  value: FormatId | null
  onChange: (format: FormatId) => void
}) {
  const activeSources = sources.filter(Boolean) as FormatId[]

  function capabilityFor(target: FormatId): Capability | undefined {
    if (activeSources.length === 0) return undefined
    const caps = activeSources.map((source) => matrix.find((entry) => entry.source === source && entry.target === target))
    if (caps.some((entry) => !entry || entry.status === 'disabled')) return undefined
    return caps[0]
  }

  return (
    <div className="grid grid-cols-4 gap-2.5">
      {formats.map((format) => {
        const cap = capabilityFor(format.id)
        const disabled = activeSources.length > 0 && (!cap || activeSources.includes(format.id))
        const selected = value === format.id
        const reason = activeSources.length > 0 && !cap ? '当前文件组合不支持' : ''
        return (
          <button
            key={format.id}
            type="button"
            disabled={disabled}
            onClick={() => onChange(format.id)}
            title={disabled ? reason || '源格式与目标格式相同' : cap?.reason || ''}
            className={cn(
              'group relative rounded-xl border px-3 py-3 text-left transition',
              selected
                ? 'border-brand-500 bg-brand-50/80 ring-2 ring-brand-100'
                : 'border-slate-200 bg-white hover:border-brand-300 hover:bg-brand-50/30',
              disabled && 'cursor-not-allowed opacity-40 hover:border-slate-200 hover:bg-white',
            )}
          >
            <div className="flex items-center justify-between">
              <span className={cn('text-sm font-semibold', selected ? 'text-brand-700' : 'text-slate-700')}>
                {formatLabels[format.id]}
              </span>
              {cap ? (
                <span className="rounded-md bg-white/80 px-1.5 py-0.5 text-[10px] font-medium text-slate-500 ring-1 ring-slate-200">
                  {qualityLabels[cap.quality || 'normal']}
                </span>
              ) : null}
            </div>
            <div className="mt-0.5 truncate text-[11px] text-slate-400">{format.label}</div>
          </button>
        )
      })}
    </div>
  )
}