import { useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import {
  Copy,
  Download,
  FileSpreadsheet,
  FileText,
  Layers,
  Loader2,
  RefreshCw,
  ScanSearch,
  Sparkles,
  Trash2,
  UploadCloud,
} from 'lucide-react'
import { apiClient, ApiError, download, withToken } from '../api'
import { Card, EmptyState, Notice, PageHeader, Spinner } from '../components/UI'
import type { DataTemplate, GenerateResult, PreviewResult, StylePreset } from '../types'
import { cn, formatBytes, formatLabels, formatTime } from '../utils'

type Tab = 'data' | 'style'

const GRADIENTS = [
  ['#1D4ED8', '#60A5FA'],
  ['#047857', '#34D399'],
  ['#C2410C', '#FDBA74'],
  ['#6D28D9', '#A78BFA'],
  ['#BE123C', '#FB7185'],
  ['#0F766E', '#2DD4BF'],
  ['#4338CA', '#818CF8'],
  ['#B45309', '#FBBF24'],
  ['#0F172A', '#64748B'],
  ['#0E7490', '#22D3EE'],
]

function gradientFor(id: string): string {
  let hash = 0
  for (const char of id) hash = (hash * 31 + char.charCodeAt(0)) % 997
  const [start, end] = GRADIENTS[hash % GRADIENTS.length]
  return `linear-gradient(135deg, ${start} 0%, ${end} 100%)`
}

function TemplateBadge({ id, format, large = false }: { id: string; format: string; large?: boolean }) {
  return (
    <div
      className={cn(
        'flex shrink-0 items-center justify-center rounded-xl font-semibold text-white shadow-sm',
        large ? 'h-12 w-12 text-[11px]' : 'h-9 w-9 text-[9px]',
      )}
      style={{ background: gradientFor(id) }}
    >
      {format.toUpperCase()}
    </div>
  )
}

function TemplateIcon({ format }: { format: string }) {
  if (format === 'xlsx') return <FileSpreadsheet size={18} />
  if (format === 'pptx') return <Layers size={18} />
  return <FileText size={18} />
}

function PreviewImages({ preview }: { preview: PreviewResult | null }) {
  if (!preview) return null
  return (
    <div className="space-y-3">
      {preview.warnings.length ? (
        <Notice tone="warning">
          {preview.warnings.map((warning) => (
            <div key={warning}>{warning}</div>
          ))}
        </Notice>
      ) : null}
      {preview.pages.length ? (
        <div className="flex gap-3 overflow-x-auto pb-1">
          {preview.pages.map((page) => (
            <img
              key={page}
              src={withToken(page)}
              alt="预览页"
              className="h-64 rounded-lg border border-slate-200 bg-white object-contain shadow-sm"
            />
          ))}
        </div>
      ) : (
        <div className="rounded-xl border border-dashed border-slate-200 px-4 py-6 text-center text-xs text-slate-400">
          当前无法渲染页面预览，但不影响实际转换与生成。
        </div>
      )}
    </div>
  )
}

function DataTemplates({
  templates,
  styles,
  onChanged,
}: {
  templates: DataTemplate[]
  styles: StylePreset[]
  onChanged: () => void
}) {
  const [selectedId, setSelectedId] = useState<string | null>(templates[0]?.id || null)
  const [formatFilter, setFormatFilter] = useState<'all' | 'docx' | 'xlsx' | 'pptx'>('all')
  const [uploadFile, setUploadFile] = useState<File | null>(null)
  const [uploadName, setUploadName] = useState('')
  const [dataFile, setDataFile] = useState<File | null>(null)
  const [styleId, setStyleId] = useState('')
  const [preview, setPreview] = useState<PreviewResult | null>(null)
  const [generated, setGenerated] = useState<GenerateResult | null>(null)
  const [error, setError] = useState('')

  useEffect(() => {
    if (!selectedId && templates.length) setSelectedId(templates[0].id)
    if (selectedId && !templates.some((template) => template.id === selectedId)) {
      setSelectedId(templates[0]?.id || null)
    }
  }, [templates, selectedId])

  const selected = templates.find((template) => template.id === selectedId) || null
  const filteredTemplates = formatFilter === 'all' ? templates : templates.filter((template) => template.format === formatFilter)

  useEffect(() => {
    if (filteredTemplates.length && !filteredTemplates.some((template) => template.id === selectedId)) {
      setSelectedId(filteredTemplates[0].id)
    }
  }, [formatFilter, templates])
  const applicableStyles = styles.filter((style) => (selected ? style.applies_to.includes(selected.format) : false))

  const upload = useMutation({
    mutationFn: (form: FormData) => apiClient.uploadTemplate(form),
    onSuccess: (template) => {
      setSelectedId(template.id)
      setUploadFile(null)
      setUploadName('')
      setError('')
      onChanged()
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const scan = useMutation({
    mutationFn: (id: string) => apiClient.scanTemplate(id),
    onSuccess: () => onChanged(),
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const previewMutation = useMutation({
    mutationFn: ({ id, form }: { id: string; form: FormData }) => apiClient.previewTemplate(id, form),
    onSuccess: (result) => {
      setPreview(result)
      setError('')
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const generate = useMutation({
    mutationFn: ({ id, form }: { id: string; form: FormData }) => apiClient.generateTemplate(id, form),
    onSuccess: (result) => {
      setGenerated(result)
      setError('')
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const remove = useMutation({
    mutationFn: (id: string) => apiClient.deleteTemplate(id),
    onSuccess: () => {
      setSelectedId(null)
      onChanged()
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  function handleUpload() {
    if (!uploadFile) {
      setError('请选择 docx / xlsx / pptx 模板文件')
      return
    }
    const form = new FormData()
    form.append('file', uploadFile)
    if (uploadName.trim()) form.append('name', uploadName.trim())
    upload.mutate(form)
  }

  function handlePreview() {
    if (!selected) return
    const form = new FormData()
    if (dataFile) form.append('data_file', dataFile)
    previewMutation.mutate({ id: selected.id, form })
  }

  function handleGenerate() {
    if (!selected) return
    if (!dataFile) {
      setError('请先选择 CSV / JSON / XLSX 数据源')
      return
    }
    const form = new FormData()
    form.append('data_file', dataFile)
    if (styleId) form.append('style_id', styleId)
    generate.mutate({ id: selected.id, form })
  }
  return (
    <div className="grid grid-cols-12 gap-5">
      <div className="col-span-4 space-y-5">
        <Card>
          <div className="section-title mb-3">上传数据模板</div>
          <div className="space-y-3">
            <input
              type="file"
              accept=".docx,.xlsx,.pptx"
              className="block w-full text-xs text-slate-500 file:mr-3 file:rounded-lg file:border-0 file:bg-brand-50 file:px-3 file:py-2 file:text-xs file:font-medium file:text-brand-700 hover:file:bg-brand-100"
              onChange={(event) => setUploadFile(event.target.files?.[0] || null)}
            />
            <input
              className="input"
              placeholder="模板名称（可选）"
              value={uploadName}
              onChange={(event) => setUploadName(event.target.value)}
            />
            <button className="btn-primary w-full" disabled={upload.isPending} onClick={handleUpload}>
              <UploadCloud size={15} /> 上传并扫描字段
            </button>
          </div>
          <div className="mt-3 text-[11px] leading-5 text-slate-400">
            在 Word / Excel / PPT 中使用 <span className="font-mono text-slate-600">{'{{字段名}}'}</span> 占位；一行数据生成一份文件。
          </div>
        </Card>

        <Card className="!p-3">
          <div className="flex items-center justify-between px-2 py-1">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-400">{`模板库`}</span>
            <span className="text-[11px] text-slate-400">{templates.length}</span>
          </div>
          <div className="mt-2 flex gap-1 px-1">
            {[{ id: 'all', label: '全部' }, { id: 'docx', label: 'Word' }, { id: 'xlsx', label: 'Excel' }, { id: 'pptx', label: 'PPT' }].map((item) => (
              <button
                key={item.id}
                onClick={() => setFormatFilter(item.id as typeof formatFilter)}
                className={cn(
                  'rounded-lg px-2 py-1 text-[11px] transition',
                  formatFilter === item.id ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-500 hover:bg-slate-200',
                )}
              >
                {item.label}
              </button>
            ))}
          </div>
          <div className="mt-2 max-h-[520px] space-y-1 overflow-y-auto pr-1">
            {filteredTemplates.map((template) => (
              <button
                key={template.id}
                className={cn(
                  'flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left transition',
                  template.id === selectedId ? 'bg-brand-50 text-brand-800' : 'hover:bg-slate-50',
                )}
                onClick={() => {
                  setSelectedId(template.id)
                  setPreview(null)
                  setGenerated(null)
                  setDataFile(null)
                }}
              >
                <TemplateBadge id={template.id} format={template.format} />
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium">{template.name}</div>
                  <div className="flex items-center gap-1 text-[11px] text-slate-400">
                    <span>{template.format.toUpperCase()} {'\u00b7'} {template.fields.length} {'\u4e2a\u5b57\u6bb5'}</span>
                    {template.builtin ? <span className="rounded bg-brand-50 px-1 text-[10px] text-brand-600">{'\u5185\u7f6e'}</span> : null}
                  </div>
                </div>
              </button>
            ))}
            {filteredTemplates.length === 0 ? <div className="px-3 py-6 text-center text-xs text-slate-400">该格式暂无模板</div> : null}
          </div>
        </Card>
      </div>

      <div className="col-span-8">
        {!selected ? (
          <EmptyState
            icon={<Sparkles size={30} />}
            title="选择一个模板开始"
            description="上传带占位符的 Word / Excel / PPT 模板后，可扫描字段、预览首行效果并批量生成。"
          />
        ) : (
          <div className="space-y-5">
            <Card>
              <div className="flex items-start justify-between gap-4">
                <div className="flex items-start gap-3">
                  <TemplateBadge id={selected.id} format={selected.format} large />

                  <div>
                    <div className="flex items-center gap-2">
                      <span className="text-lg font-semibold text-slate-900">{selected.name}</span>
                      {selected.builtin ? <span className="badge bg-brand-50 text-brand-700">{'\u5185\u7f6e\u6a21\u677f'}</span> : null}
                    </div>
                    <div className="mt-0.5 text-xs text-slate-400">
                      {selected.original_name} · {selected.format.toUpperCase()} · 更新于 {formatTime(selected.updated_at)}
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button className="btn-secondary" disabled={scan.isPending} onClick={() => scan.mutate(selected.id)}>
                    {scan.isPending ? <Loader2 size={14} className="animate-spin" /> : <ScanSearch size={14} />}
                    重新扫描
                  </button>
                  {selected.builtin ? (
                    <span className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-400">{'\u53ef\u76f4\u63a5\u4f7f\u7528'}</span>
                  ) : (
                    <button
                      className="btn-danger"
                      onClick={() => {
                        if (window.confirm(`\u786e\u8ba4\u5220\u9664\u6a21\u677f\u300c${selected.name}\u300d\uff1f`)) remove.mutate(selected.id)
                      }}
                    >
                      <Trash2 size={14} /> {'\u5220\u9664'}
                    </button>
                  )}
                </div>
              </div>
              <div className="mt-5">
                <div className="mb-2 text-xs font-medium uppercase tracking-wider text-slate-400">占位字段</div>
                {selected.fields.length ? (
                  <div className="flex flex-wrap gap-2">
                    {selected.fields.map((field) => (
                      <span
                        key={field.name}
                        className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 font-mono text-xs text-slate-600"
                        title={field.sample}
                      >
                        {`{{${field.name}}}`}
                        <span className="rounded bg-white px-1 text-[10px] text-slate-400 ring-1 ring-slate-200">{field.count}</span>
                      </span>
                    ))}
                  </div>
                ) : (
                  <Notice tone="warning">未扫描到占位符。请确认模板中使用 {'{{字段名}}'} 语法。</Notice>
                )}
              </div>
            </Card>

            <Card>
              <div className="mb-4 flex items-center gap-3 rounded-xl border border-brand-100 bg-brand-50/70 px-3.5 py-3">
                <TemplateBadge id={selected.id} format={selected.format} />
                <div className="min-w-0">
                  <div className="text-[11px] text-slate-500">{"本次生成将使用数据模板"}</div>
                  <div className="truncate text-sm font-semibold text-brand-900">{selected.name}</div>
                </div>
              </div>
              <div className="section-title mb-4">数据源与预览</div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <span className="label">数据文件</span>
                  <input
                    type="file"
                    accept=".csv,.json,.xlsx"
                    className="block w-full text-xs text-slate-500 file:mr-3 file:rounded-lg file:border-0 file:bg-slate-100 file:px-3 file:py-2 file:text-xs file:font-medium file:text-slate-700 hover:file:bg-slate-200"
                    onChange={(event) => setDataFile(event.target.files?.[0] || null)}
                  />
                  <div className="mt-1.5 text-[11px] text-slate-400">CSV / JSON / XLSX，首行为字段名，最多 200 行。</div>
                </div>
                <div>
                  <span className="label">{"生成后追加样式（仅字体/配色，不改版式）"}</span>
                  <select className="input" value={styleId} onChange={(event) => setStyleId(event.target.value)}>
                    <option value="">不套用样式</option>
                    {applicableStyles.map((style) => (
                      <option key={style.id} value={style.id}>
                        {style.name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="mt-4 flex flex-wrap items-center gap-2">
                <button className="btn-secondary" disabled={previewMutation.isPending} onClick={handlePreview}>
                  {previewMutation.isPending ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
                  首行预览
                </button>
                <button className="btn-primary" disabled={generate.isPending || !selected.fields.length} onClick={handleGenerate}>
                  {generate.isPending ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                  批量生成
                </button>
                {dataFile ? <span className="text-[11px] text-slate-400">数据：{dataFile.name} · {formatBytes(dataFile.size)}</span> : null}
              </div>

              {error ? (
                <Notice tone="error" className="mt-4">
                  {error}
                </Notice>
              ) : null}

              {preview ? (
                <div className="mt-5">
                  <PreviewImages preview={preview} />
                </div>
              ) : null}

              {generated ? (
                <div className="mt-5 rounded-xl border border-emerald-100 bg-emerald-50/50 p-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <div className="text-sm font-medium text-emerald-900">
                        已生成 {generated.generated} / {generated.rows} 份文件
                      </div>
                      <div className="mt-0.5 text-[11px] text-emerald-700/70">{"基于模板："}{selected.name}</div>
                    </div>
                    {generated.archive_url ? (
                      <button
                        className="btn-primary !px-3 !py-1.5 text-xs"
                        onClick={() => void download(generated.archive_url!, generated.archive_name || 'generated.zip')}
                      >
                        <Download size={13} /> {'\u4e0b\u8f7d\u5168\u90e8 ZIP'}
                      </button>
                    ) : null}
                  </div>
                  {generated.warnings.length ? (
                    <div className="mt-2 space-y-1 text-[11px] leading-5 text-amber-700">
                      {generated.warnings.map((warning) => (
                        <div key={warning}>· {warning}</div>
                      ))}
                    </div>
                  ) : null}
                  <div className="mt-3 grid grid-cols-2 gap-2">
                    {generated.files.slice(0, 20).map((file) => (
                      <button
                        key={file.url}
                        onClick={() => void download(file.url, file.name)}
                        className="flex items-center gap-2 rounded-lg border border-emerald-100 bg-white px-3 py-2 text-left text-xs text-slate-600 hover:border-emerald-300"
                      >
                        <Download size={12} /> <span className="truncate">{file.name}</span>
                      </button>
                    ))}
                  </div>
                </div>
              ) : null}
            </Card>
          </div>
        )}
      </div>
    </div>
  )
}

function NumberField({
  label,
  value,
  onChange,
  step = 1,
}: {
  label: string
  value: number
  onChange: (value: number) => void
  step?: number
}) {
  return (
    <label>
      <span className="label">{label}</span>
      <input
        className="input"
        type="number"
        step={step}
        value={Number.isFinite(value) ? value : 0}
        onChange={(event) => onChange(Number(event.target.value))}
      />
    </label>
  )
}

function TextField({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label>
      <span className="label">{label}</span>
      <input className="input" value={value} onChange={(event) => onChange(event.target.value)} />
    </label>
  )
}
function StyleEditor({ style, onChanged }: { style: StylePreset; onChanged: () => void }) {
  const queryClient = useQueryClient()
  const [name, setName] = useState(style.name)
  const [appliesTo, setAppliesTo] = useState<string[]>(style.applies_to)
  const [data, setData] = useState<Record<string, any>>(JSON.parse(JSON.stringify(style.data)))
  const [fmt, setFmt] = useState(style.applies_to[0] || 'docx')
  const [preview, setPreview] = useState<PreviewResult | null>(null)
  const [error, setError] = useState('')
  const [saved, setSaved] = useState('')

  useEffect(() => {
    setName(style.name)
    setAppliesTo(style.applies_to)
    setData(JSON.parse(JSON.stringify(style.data)))
    setFmt(style.applies_to[0] || 'docx')
    setPreview(null)
    setError('')
    setSaved('')
  }, [style])

  function patch(section: string, key: string, value: any) {
    setData((current) => ({ ...current, [section]: { ...(current[section] || {}), [key]: value } }))
  }

  function patchWatermark(key: string, value: any) {
    setData((current) => ({
      ...current,
      pdf: { ...(current.pdf || {}), watermark: { ...(current.pdf?.watermark || {}), [key]: value } },
    }))
  }

  const save = useMutation({
    mutationFn: async () => {
      if (!appliesTo.length) throw new ApiError('请至少选择一个适用格式')
      const payload = { name: name.trim() || '自定义样式', applies_to: appliesTo, data }
      if (style.builtin) return apiClient.createStyle(payload)
      return apiClient.updateStyle(style.id, payload)
    },
    onSuccess: () => {
      setSaved(style.builtin ? '已复制为新的自定义样式' : '样式已保存')
      setError('')
      onChanged()
      queryClient.invalidateQueries({ queryKey: ['styles'] })
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const remove = useMutation({
    mutationFn: () => apiClient.deleteStyle(style.id),
    onSuccess: onChanged,
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const previewMutation = useMutation({
    mutationFn: async () => {
      if (!style.builtin) {
        const payload = { name: name.trim() || style.name, applies_to: appliesTo, data }
        await apiClient.updateStyle(style.id, payload)
        onChanged()
      }
      return apiClient.previewStyle(style.id, fmt)
    },
    onSuccess: (result) => {
      setPreview(result)
      setError('')
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  const docx = data.docx || {}
  const xlsx = data.xlsx || {}
  const pptx = data.pptx || {}
  const watermark = data.pdf?.watermark || {}
  return (
    <div className="space-y-5">
      <Card>
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-lg font-semibold text-slate-900">{style.builtin ? style.name : name}</span>
              {style.builtin ? <span className="badge bg-slate-100 text-slate-500">内置</span> : null}
            </div>
            <div className="mt-1 text-xs text-slate-400">更新于 {formatTime(style.updated_at)}</div>
          </div>
          <div className="flex items-center gap-2">
            <button className="btn-secondary" disabled={previewMutation.isPending} onClick={() => previewMutation.mutate()}>
              {previewMutation.isPending ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
              预览
            </button>
            <button className="btn-primary" disabled={save.isPending} onClick={() => save.mutate()}>
              {save.isPending ? <Loader2 size={14} className="animate-spin" /> : <Copy size={14} />}
              {style.builtin ? '复制并保存' : '保存'}
            </button>
            {!style.builtin ? (
              <button
                className="btn-danger"
                onClick={() => {
                  if (window.confirm(`确认删除样式「${style.name}」？`)) remove.mutate()
                }}
              >
                <Trash2 size={14} />
              </button>
            ) : null}
          </div>
        </div>

        <div className="mt-5 grid grid-cols-2 gap-4">
          <TextField label="样式名称" value={name} onChange={setName} />
          <div>
            <span className="label">适用格式</span>
            <div className="flex gap-2">
              {['docx', 'xlsx', 'pptx'].map((format) => (
                <button
                  key={format}
                  className={cn(
                    'rounded-lg border px-3 py-2 text-xs transition',
                    appliesTo.includes(format)
                      ? 'border-brand-500 bg-brand-50 text-brand-700'
                      : 'border-slate-200 bg-white text-slate-500',
                  )}
                  onClick={() => {
                    const next = appliesTo.includes(format) ? appliesTo.filter((item) => item !== format) : [...appliesTo, format]
                    setAppliesTo(next)
                    if (!next.includes(fmt) && next.length) setFmt(next[0])
                  }}
                >
                  {formatLabels[format as keyof typeof formatLabels] || format}
                </button>
              ))}
            </div>
          </div>
        </div>

        <div className="mt-4 flex gap-2">
          <button
            className={cn(
              'rounded-lg px-3 py-1.5 text-xs transition',
              fmt === 'pdf' ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-500 hover:bg-slate-200',
            )}
            onClick={() => setFmt('pdf')}
          >
            PDF 水印
          </button>
          {appliesTo.map((format) => (
            <button
              key={format}
              className={cn(
                'rounded-lg px-3 py-1.5 text-xs transition',
                fmt === format ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-500 hover:bg-slate-200',
              )}
              onClick={() => setFmt(format)}
            >
              {formatLabels[format as keyof typeof formatLabels] || format}
            </button>
          ))}
        </div>
      </Card>

      <Card>
        <div className="section-title mb-4">{formatLabels[fmt as keyof typeof formatLabels] || fmt} 样式参数</div>
        {fmt === 'docx' ? (
          <div className="grid grid-cols-3 gap-4">
            <TextField label="中文字体" value={docx.font_cn || ''} onChange={(value) => patch('docx', 'font_cn', value)} />
            <TextField label={"主题色"} value={docx.primary_color || ''} onChange={(value) => patch('docx', 'primary_color', value)} />
            <TextField label={"页面背景"} value={docx.page_background || ''} onChange={(value) => patch('docx', 'page_background', value)} />
            <TextField label={"页眉色带"} value={docx.header_fill || ''} onChange={(value) => patch('docx', 'header_fill', value)} />
            <TextField label="西文字体" value={docx.font_en || ''} onChange={(value) => patch('docx', 'font_en', value)} />
            <NumberField label="正文字号 pt" value={docx.body_size} onChange={(value) => patch('docx', 'body_size', value)} />
            <NumberField label="标题字号 pt" value={docx.title_size} onChange={(value) => patch('docx', 'title_size', value)} />
            <NumberField label="一级标题 pt" value={docx.h1_size} onChange={(value) => patch('docx', 'h1_size', value)} />
            <NumberField label="二级标题 pt" value={docx.h2_size} onChange={(value) => patch('docx', 'h2_size', value)} />
            <NumberField label="行距" value={docx.line_spacing} step={0.1} onChange={(value) => patch('docx', 'line_spacing', value)} />
            <NumberField label="段后 pt" value={docx.space_after} onChange={(value) => patch('docx', 'space_after', value)} />
            <NumberField label="首行缩进 字符" value={docx.first_line_indent} step={0.5} onChange={(value) => patch('docx', 'first_line_indent', value)} />
            <NumberField label="上边距 cm" value={docx.margin_top_cm} step={0.1} onChange={(value) => patch('docx', 'margin_top_cm', value)} />
            <NumberField label="下边距 cm" value={docx.margin_bottom_cm} step={0.1} onChange={(value) => patch('docx', 'margin_bottom_cm', value)} />
            <NumberField label="左边距 cm" value={docx.margin_left_cm} step={0.1} onChange={(value) => patch('docx', 'margin_left_cm', value)} />
            <NumberField label="右边距 cm" value={docx.margin_right_cm} step={0.1} onChange={(value) => patch('docx', 'margin_right_cm', value)} />
            <TextField label="页眉文字" value={docx.header_text || ''} onChange={(value) => patch('docx', 'header_text', value)} />
            <TextField label="页脚文字" value={docx.footer_text || ''} onChange={(value) => patch('docx', 'footer_text', value)} />
            <div className="flex flex-col justify-end gap-2 pb-1">
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input
                  type="checkbox"
                  className="accent-brand-600"
                  checked={Boolean(docx.page_number)}
                  onChange={(event) => patch('docx', 'page_number', event.target.checked)}
                />
                页脚显示页码
              </label>
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input
                  type="checkbox"
                  className="accent-brand-600"
                  checked={Boolean(docx.force_direct_format)}
                  onChange={(event) => patch('docx', 'force_direct_format', event.target.checked)}
                />
                强制覆盖直接格式
              </label>
            </div>
          </div>
        ) : null}

        {fmt === 'xlsx' ? (
          <div className="grid grid-cols-3 gap-4">
            <TextField label="字体" value={xlsx.font_name || ''} onChange={(value) => patch('xlsx', 'font_name', value)} />
            <NumberField label="字号 pt" value={xlsx.font_size} onChange={(value) => patch('xlsx', 'font_size', value)} />
            <TextField label={'\u6b63\u6587\u989c\u8272'} value={xlsx.font_color || ''} onChange={(value) => patch('xlsx', 'font_color', value)} />
            <TextField label={'\u80cc\u666f\u586b\u5145'} value={xlsx.body_fill || ''} onChange={(value) => patch('xlsx', 'body_fill', value)} />
            <TextField label="表头填充色" value={xlsx.header_fill || ''} onChange={(value) => patch('xlsx', 'header_fill', value)} />
            <TextField label={"工作表标签色"} value={xlsx.tab_color || ''} onChange={(value) => patch('xlsx', 'tab_color', value)} />
            <TextField label="表头字色" value={xlsx.header_color || ''} onChange={(value) => patch('xlsx', 'header_color', value)} />
            <TextField label="边框颜色" value={xlsx.border_color || ''} onChange={(value) => patch('xlsx', 'border_color', value)} />
            <TextField label="斑马纹填充" value={xlsx.zebra_fill || ''} onChange={(value) => patch('xlsx', 'zebra_fill', value)} />
            <div className="flex flex-col justify-end gap-2 pb-1">
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" className="accent-brand-600" checked={Boolean(xlsx.freeze_header)} onChange={(event) => patch('xlsx', 'freeze_header', event.target.checked)} />
                冻结表头
              </label>
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" className="accent-brand-600" checked={Boolean(xlsx.landscape)} onChange={(event) => patch('xlsx', 'landscape', event.target.checked)} />
                横向打印
              </label>
            </div>
            <div className="flex flex-col justify-end gap-2 pb-1">
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" className="accent-brand-600" checked={Boolean(xlsx.header_bold)} onChange={(event) => patch('xlsx', 'header_bold', event.target.checked)} />
                表头加粗
              </label>
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" className="accent-brand-600" checked={Boolean(xlsx.auto_width)} onChange={(event) => patch('xlsx', 'auto_width', event.target.checked)} />
                自动列宽
              </label>
            </div>
          </div>
        ) : null}

        {fmt === 'pptx' ? (
          <div className="grid grid-cols-2 gap-4">
            <TextField label="标题字体" value={pptx.title_font || ''} onChange={(value) => patch('pptx', 'title_font', value)} />
            <TextField label="正文字体" value={pptx.body_font || ''} onChange={(value) => patch('pptx', 'body_font', value)} />
            <NumberField label="幻灯片宽度（英寸）" value={pptx.slide_width_inches} step={0.1} onChange={(value) => patch('pptx', 'slide_width_inches', value)} />
            <NumberField label="幻灯片高度（英寸）" value={pptx.slide_height_inches} step={0.1} onChange={(value) => patch('pptx', 'slide_height_inches', value)} />
            <TextField label={"背景色"} value={pptx.background_color || ''} onChange={(value) => patch('pptx', 'background_color', value)} />
            <TextField label={"主题色"} value={pptx.primary_color || ''} onChange={(value) => patch('pptx', 'primary_color', value)} />
            <TextField label={"强调色"} value={pptx.accent_color || ''} onChange={(value) => patch('pptx', 'accent_color', value)} />
            <TextField label={"正文颜色"} value={pptx.text_color || ''} onChange={(value) => patch('pptx', 'text_color', value)} />
          </div>
        ) : null}

        {fmt === 'pdf' ? (
          <div className="grid grid-cols-3 gap-4">
            <div className="col-span-3">
              <label className="flex items-center gap-2 text-xs text-slate-600">
                <input
                  type="checkbox"
                  className="accent-brand-600"
                  checked={Boolean(watermark.enabled)}
                  onChange={(event) => patchWatermark('enabled', event.target.checked)}
                />
                启用 PDF 水印（仅 PDF 输出）
              </label>
            </div>
            <TextField label="水印文字" value={watermark.text || ''} onChange={(value) => patchWatermark('text', value)} />
            <NumberField label="字号" value={watermark.size} onChange={(value) => patchWatermark('size', value)} />
            <TextField label="颜色（HEX）" value={watermark.color || ''} onChange={(value) => patchWatermark('color', value)} />
            <NumberField label="透明度 0-1" value={watermark.opacity} step={0.05} onChange={(value) => patchWatermark('opacity', value)} />
            <NumberField label="角度" value={watermark.angle} onChange={(value) => patchWatermark('angle', value)} />
            <label className="flex items-end gap-2 pb-2 text-xs text-slate-600">
              <input type="checkbox" className="accent-brand-600" checked={Boolean(watermark.tile)} onChange={(event) => patchWatermark('tile', event.target.checked)} />
              平铺
            </label>
          </div>
        ) : null}
      </Card>

      {error ? <Notice tone="error">{error}</Notice> : null}
      {saved ? <Notice tone="success">{saved}</Notice> : null}
      {preview ? <PreviewImages preview={preview} /> : null}
    </div>
  )
}
function StylesPanel({ styles, onChanged }: { styles: StylePreset[]; onChanged: () => void }) {
  const [selectedId, setSelectedId] = useState<string | null>(styles[0]?.id || null)
  const [formatFilter, setFormatFilter] = useState<'all' | 'docx' | 'xlsx' | 'pptx'>('all')
  const [file, setFile] = useState<File | null>(null)
  const filteredStyles = formatFilter === 'all' ? styles : styles.filter((style) => style.applies_to.includes(formatFilter))

  useEffect(() => {
    if (filteredStyles.length && !filteredStyles.some((style) => style.id === selectedId)) {
      setSelectedId(filteredStyles[0].id)
    }
  }, [formatFilter, styles])
  const [name, setName] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    if (!selectedId && styles.length) setSelectedId(styles[0].id)
    if (selectedId && !styles.some((style) => style.id === selectedId)) setSelectedId(styles[0]?.id || null)
  }, [styles, selectedId])

  const selected = styles.find((style) => style.id === selectedId) || null

  const importMutation = useMutation({
    mutationFn: (form: FormData) => apiClient.importStyle(form),
    onSuccess: (style) => {
      setSelectedId(style.id)
      setFile(null)
      setName('')
      setError('')
      onChanged()
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  function handleImport() {
    if (!file) {
      setError('请选择 docx / xlsx / pptx 样板文件')
      return
    }
    const form = new FormData()
    form.append('file', file)
    if (name.trim()) form.append('name', name.trim())
    importMutation.mutate(form)
  }

  return (
    <div className="grid grid-cols-12 gap-5">
      <div className="col-span-4 space-y-5">
        <Card>
          <div className="section-title mb-3">从样板导入</div>
          <div className="space-y-3">
            <input
              type="file"
              accept=".docx,.xlsx,.pptx"
              className="block w-full text-xs text-slate-500 file:mr-3 file:rounded-lg file:border-0 file:bg-slate-100 file:px-3 file:py-2 file:text-xs file:font-medium file:text-slate-700 hover:file:bg-slate-200"
              onChange={(event) => setFile(event.target.files?.[0] || null)}
            />
            <input className="input" placeholder="样式名称（可选）" value={name} onChange={(event) => setName(event.target.value)} />
            <button className="btn-primary w-full" disabled={importMutation.isPending} onClick={handleImport}>
              <UploadCloud size={15} /> 导入样式
            </button>
          </div>
        </Card>

        <Card className="!p-3">
          <div className="flex items-center justify-between px-2 py-1">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-400">{'\u6837\u5f0f\u5e93'}</span>
            <span className="text-[11px] text-slate-400">{filteredStyles.length}</span>
          </div>
          <div className="mt-2 flex gap-1 px-1">
            {[{ id: 'all', label: '全部' }, { id: 'docx', label: 'Word' }, { id: 'xlsx', label: 'Excel' }, { id: 'pptx', label: 'PPT' }].map((item) => (
              <button
                key={item.id}
                onClick={() => setFormatFilter(item.id as typeof formatFilter)}
                className={cn(
                  'rounded-lg px-2 py-1 text-[11px] transition',
                  formatFilter === item.id ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-500 hover:bg-slate-200',
                )}
              >
                {item.label}
              </button>
            ))}
          </div>
          <div className="mt-1 max-h-[560px] space-y-1 overflow-y-auto">
            {filteredStyles.map((style) => (
              <button
                key={style.id}
                className={cn(
                  'flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left transition',
                  style.id === selectedId ? 'bg-brand-50 text-brand-800' : 'hover:bg-slate-50',
                )}
                onClick={() => setSelectedId(style.id)}
              >
                <div className="h-8 w-8 shrink-0 rounded-lg shadow-sm" style={{ background: gradientFor(style.id) }} />
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium">{style.name}</div>
                  <div className="text-[11px] text-slate-400">{style.applies_to.map((item) => item.toUpperCase()).join(' / ')}</div>
                </div>
                {style.builtin ? <span className="text-[10px] text-slate-400">{'\u5185\u7f6e'}</span> : null}
              </button>
            ))}
          </div>
        </Card>
      </div>

      <div className="col-span-8">
        {error ? (
          <Notice tone="error" className="mb-4">
            {error}
          </Notice>
        ) : null}
        {selected ? (
          <StyleEditor style={selected} onChanged={onChanged} />
        ) : (
          <EmptyState title="选择一个样式" description="内置样式可复制后编辑，也可从现有 Word / Excel / PPT 样板导入。" />
        )}
      </div>
    </div>
  )
}

export default function TemplatesPage() {
  const queryClient = useQueryClient()
  const [tab, setTab] = useState<Tab>(() => (new URLSearchParams(window.location.search).get('tab') === 'style' ? 'style' : 'data'))
  const { data: templates = [], isLoading: loadingTemplates } = useQuery({ queryKey: ['templates'], queryFn: apiClient.templates })
  const { data: styles = [], isLoading: loadingStyles } = useQuery({ queryKey: ['styles'], queryFn: apiClient.styles })

  function refresh() {
    queryClient.invalidateQueries({ queryKey: ['templates'] })
    queryClient.invalidateQueries({ queryKey: ['styles'] })
  }

  return (
    <div>
      <PageHeader
        title="模板中心"
        description="数据模板负责批量生成，样式模板负责统一排版。两者可以组合使用：先填充数据，再套用样式。"
      />
      <Notice tone="info" className="mb-5">
        {"要生成带封面、章节和完整版式的文件，请在「数据模板」里选择模板；「样式模板」只调整字体、配色和页边距，不会改变封面与版式。"}
      </Notice>



      <div className="mb-5 inline-flex rounded-xl border border-slate-200 bg-white p-1">
        {[
          { id: 'data' as Tab, label: '数据模板', icon: FileText },
          { id: 'style' as Tab, label: '样式模板', icon: Sparkles },
        ].map((item) => {
          const Icon = item.icon
          return (
            <button
              key={item.id}
              className={cn(
                'flex items-center gap-2 rounded-lg px-4 py-2 text-sm font-medium transition',
                tab === item.id ? 'bg-slate-900 text-white' : 'text-slate-500 hover:bg-slate-50',
              )}
              onClick={() => setTab(item.id)}
            >
              <Icon size={15} />
              {item.label}
            </button>
          )
        })}
      </div>

      {tab === 'data' ? (
        loadingTemplates ? (
          <Card><Spinner label="加载模板…" /></Card>
        ) : (
          <DataTemplates templates={templates} styles={styles} onChanged={refresh} />
        )
      ) : loadingStyles ? (
        <Card><Spinner label="加载样式…" /></Card>
      ) : (
        <StylesPanel styles={styles} onChanged={refresh} />
      )}
    </div>
  )
}
