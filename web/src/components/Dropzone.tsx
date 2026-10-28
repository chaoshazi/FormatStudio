import { useRef, useState } from 'react'
import { FilePlus2, UploadCloud, X } from 'lucide-react'
import { cn, formatBytes, formatOf } from '../utils'
import { Notice } from './UI'

const ACCEPT = '.txt,.text,.md,.markdown,.docx,.pdf,.wps,.csv,.xlsx,.pptx'

export default function Dropzone({
  files,
  onChange,
  maxFiles = 50,
}: {
  files: File[]
  onChange: (files: File[]) => void
  maxFiles?: number
}) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState('')

  function addFiles(incoming: FileList | File[]) {
    const next = [...files]
    const rejected: string[] = []
    for (const file of Array.from(incoming)) {
      if (!formatOf(file.name)) {
        rejected.push(file.name)
        continue
      }
      if (next.some((entry) => entry.name === file.name && entry.size === file.size)) continue
      if (next.length >= maxFiles) break
      next.push(file)
    }
    setError(rejected.length ? `已跳过不支持的文件：${rejected.join('、')}` : '')
    onChange(next)
  }

  return (
    <div>
      <div
        onDragOver={(event) => {
          event.preventDefault()
          setDragging(true)
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(event) => {
          event.preventDefault()
          setDragging(false)
          addFiles(event.dataTransfer.files)
        }}
        onClick={() => inputRef.current?.click()}
        className={cn(
          'flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed px-6 py-10 text-center transition',
          dragging
            ? 'border-brand-400 bg-brand-50/70'
            : 'border-slate-200 bg-white hover:border-brand-300 hover:bg-brand-50/30',
        )}
      >
        <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-brand-50 text-brand-600">
          <UploadCloud size={22} />
        </div>
        <div className="text-sm font-medium text-slate-700">拖拽文件到这里，或点击选择</div>
        <div className="mt-1.5 text-xs leading-5 text-slate-400">
          支持 TXT / Markdown / Word / PDF / WPS / CSV / Excel / PPT，单文件不超过 200MB
        </div>
        <input
          ref={inputRef}
          type="file"
          multiple
          accept={ACCEPT}
          className="hidden"
          onChange={(event) => {
            if (event.target.files) addFiles(event.target.files)
            event.target.value = ''
          }}
        />
      </div>

      {error ? (
        <Notice tone="warning" className="mt-3">
          {error}
        </Notice>
      ) : null}

      {files.length > 0 ? (
        <div className="mt-4 space-y-2">
          <div className="flex items-center justify-between text-xs text-slate-400">
            <span>已选择 {files.length} 个文件</span>
            <button
              className="inline-flex items-center gap-1 text-slate-500 hover:text-rose-600"
              onClick={() => onChange([])}
            >
              <X size={13} /> 清空
            </button>
          </div>
          {files.map((file, index) => {
            const fmt = formatOf(file.name)
            return (
              <div
                key={`${file.name}-${file.size}-${index}`}
                className="flex items-center gap-3 rounded-xl border border-slate-100 bg-white px-3.5 py-2.5"
              >
                <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-slate-50 text-slate-400">
                  <FilePlus2 size={16} />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm text-slate-700">{file.name}</div>
                  <div className="text-[11px] text-slate-400">
                    {fmt?.toUpperCase()} · {formatBytes(file.size)}
                  </div>
                </div>
                <button
                  className="rounded-lg p-1.5 text-slate-300 transition hover:bg-rose-50 hover:text-rose-500"
                  onClick={(event) => {
                    event.stopPropagation()
                    onChange(files.filter((_, entryIndex) => entryIndex !== index))
                  }}
                  aria-label={`移除 ${file.name}`}
                >
                  <X size={15} />
                </button>
              </div>
            )
          })}
        </div>
      ) : null}
    </div>
  )
}