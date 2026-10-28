import { useEffect, useMemo, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { ArrowRight, CheckCircle2, Download, FolderOpen, Loader2, LockKeyhole, Sparkles } from 'lucide-react'
import { Link } from 'react-router-dom'
import { apiClient, ApiError, download, sseUrl } from '../api'
import Dropzone from '../components/Dropzone'
import FormatGrid from '../components/FormatGrid'
import { Card, Notice, PageHeader, ProgressBar, Spinner, StatusBadge } from '../components/UI'
import type { FormatId, Task } from '../types'
import { cn, formatLabels, formatOf, metaSummary } from '../utils'

function ActiveTask({ initial, onReset }: { initial: Task; onReset: () => void }) {
  const [task, setTask] = useState<Task>(initial)
  const queryClient = useQueryClient()

  useEffect(() => {
    const source = new EventSource(sseUrl(initial.id))
    let fallback: number | undefined
    source.addEventListener('task', (event) => {
      try {
        setTask(JSON.parse((event as MessageEvent).data))
      } catch {
        // ignore
      }
    })
    source.addEventListener('done', () => {
      source.close()
      queryClient.invalidateQueries({ queryKey: ['tasks'] })
    })
    source.onerror = () => {
      source.close()
      fallback = window.setInterval(async () => {
        try {
          const next = await apiClient.task(initial.id)
          setTask(next)
          if (['succeeded', 'failed', 'canceled', 'partial', 'interrupted'].includes(next.status)) {
            window.clearInterval(fallback)
            queryClient.invalidateQueries({ queryKey: ['tasks'] })
          }
        } catch {
          // ignore
        }
      }, 1500)
    }
    return () => {
      source.close()
      if (fallback) window.clearInterval(fallback)
    }
  }, [initial.id, queryClient])

  const done = ['succeeded', 'failed', 'canceled', 'partial', 'interrupted'].includes(task.status)

  return (
    <Card className="mt-6">
      <div className="flex items-start justify-between gap-4">
        <div>
          <div className="flex items-center gap-2.5">
            <span className="section-title">当前任务</span>
            <StatusBadge status={task.status} />
          </div>
          <div className="mt-1 text-xs text-slate-400">任务编号 {task.id.slice(0, 8)}</div>
        </div>
        <div className="flex items-center gap-2">
          <Link to={`/tasks/${task.id}`} className="btn-secondary">
            任务详情
          </Link>
          {done ? (
            <button className="btn-primary" onClick={onReset}>
              继续转换
            </button>
          ) : null}
        </div>
      </div>

      <div className="mt-5">
        <div className="mb-2 flex items-center justify-between text-xs text-slate-500">
          <span>
            {task.status === 'queued'
              ? `排队中${task.queue_position ? ` · 第 ${task.queue_position} 位` : ''}`
              : '转换进度'}
          </span>
          <span>{task.progress}%</span>
        </div>
        <ProgressBar value={task.progress} />
      </div>

      <div className="mt-5 space-y-2">
        {task.items.map((item) => (
          <div key={item.id} className="rounded-xl border border-slate-100 bg-slate-50/40 p-3.5">
            <div className="flex items-center gap-3">
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-white text-slate-400 ring-1 ring-slate-100">
                {item.status === 'succeeded' ? (
                  <CheckCircle2 size={16} className="text-emerald-500" />
                ) : item.status === 'running' ? (
                  <Loader2 size={16} className="animate-spin text-brand-600" />
                ) : (
                  <span className="text-[10px] font-semibold">{formatLabels[item.source_format]}</span>
                )}
              </div>
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2 text-sm">
                  <span className="truncate text-slate-700">{item.source_name}</span>
                  <ArrowRight size={13} className="shrink-0 text-slate-300" />
                  <span className="shrink-0 text-slate-500">{formatLabels[item.target_format]}</span>
                </div>
                <div className="mt-0.5 flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] text-slate-400">
                  <span>{item.progress}%</span>
                  {item.engine ? <span>引擎 {item.engine}</span> : null}
                  {metaSummary(item.meta).length ? <span>{metaSummary(item.meta).join(' · ')}</span> : null}
                </div>
              </div>
              {item.status === 'succeeded' ? (
                <button className="btn-secondary !px-3 !py-1.5 text-xs" onClick={() => void download(item.download_url, item.download_name)}>
                  <Download size={13} /> 下载
                </button>
              ) : null}
            </div>
            {item.error_message ? (
              <Notice tone="error" className="mt-2.5">
                {item.error_message}
              </Notice>
            ) : null}
            {item.warning ? (
              <Notice tone="warning" className="mt-2.5">
                {item.warning}
              </Notice>
            ) : null}
          </div>
        ))}
      </div>
    </Card>
  )
}
export default function Workbench() {
  const queryClient = useQueryClient()
  const { data: capabilities } = useQuery({ queryKey: ['capabilities'], queryFn: apiClient.capabilities })
  const { data: styles = [] } = useQuery({ queryKey: ['styles'], queryFn: apiClient.styles })
  const { data: settings } = useQuery({ queryKey: ['settings'], queryFn: apiClient.settings })

  const [files, setFiles] = useState<File[]>([])
  const [target, setTarget] = useState<FormatId>('docx')
  const [mode, setMode] = useState('default')
  const [styleId, setStyleId] = useState('')
  const [pdfPassword, setPdfPassword] = useState('')
  const [csvAllText, setCsvAllText] = useState(false)
  const [sheetName, setSheetName] = useState('')
  const [error, setError] = useState('')
  const [activeTask, setActiveTask] = useState<Task | null>(null)

  const sources = files.map((file) => formatOf(file.name))
  const activeSources = sources.filter(Boolean) as FormatId[]

  const supportedTargets = useMemo(() => {
    if (!capabilities || activeSources.length === 0) return []
    return capabilities.formats
      .map((format) => format.id)
      .filter((candidate) =>
        activeSources.every((source) => {
          const cap = capabilities.matrix.find((entry) => entry.source === source && entry.target === candidate)
          return cap && cap.status !== 'disabled'
        }),
      )
  }, [capabilities, activeSources])

  useEffect(() => {
    if (supportedTargets.length === 0) return
    if (!supportedTargets.includes(target)) {
      setTarget(supportedTargets.includes('docx') ? 'docx' : supportedTargets[0])
    }
  }, [supportedTargets, target])

  const targetCapability = useMemo(() => {
    if (!capabilities || activeSources.length === 0) return undefined
    const caps = activeSources.map((source) =>
      capabilities.matrix.find((entry) => entry.source === source && entry.target === target),
    )
    const first = caps[0]
    if (!first || caps.some((entry) => !entry || entry.status === 'disabled')) return undefined
    return first
  }, [capabilities, activeSources, target])

  const modeKey = targetCapability?.modes.join(',') || ''

  useEffect(() => {
    if (targetCapability) {
      setMode(targetCapability.modes.includes('editable') ? 'editable' : targetCapability.modes[0] || 'default')
    }
  }, [target, modeKey])

  const applicableStyles = styles.filter(
    (style) => style.applies_to.includes(target) || (target === 'pdf' && style.applies_to.length > 0),
  )

  const mutation = useMutation({
    mutationFn: (form: FormData) => apiClient.createTask(form),
    onSuccess: (task) => {
      setActiveTask(task)
      setError('')
      queryClient.invalidateQueries({ queryKey: ['tasks'] })
    },
    onError: (err) => {
      setError(err instanceof ApiError ? err.message : String(err))
    },
  })

  function start() {
    setError('')
    if (files.length === 0) {
      setError('请先选择要转换的文件')
      return
    }
    if (!targetCapability) {
      setError('当前文件组合不支持所选目标格式')
      return
    }
    const form = new FormData()
    files.forEach((file) => form.append('files', file))
    form.append('target_format', target)
    form.append('mode', mode)
    if (styleId) form.append('style_id', styleId)
    if (pdfPassword) form.append('pdf_password', pdfPassword)
    form.append('csv_all_text', String(csvAllText))
    if (sheetName.trim()) form.append('sheet_name', sheetName.trim())
    mutation.mutate(form)
  }

  const hasPdf = activeSources.includes('pdf')
  const hasCsv = activeSources.includes('csv')
  const hasXlsx = activeSources.includes('xlsx')
  const modes = targetCapability?.modes || []
  const showMode = modes.includes('editable') && modes.includes('layout')
  return (
    <div>
      <PageHeader
        title="转换工作台"
        description="选择文件与目标格式，本地完成转换。批量任务会按顺序执行，结果自动归档到输出目录。"
        actions={
          settings ? (
            <div className="flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-500">
              <FolderOpen size={14} />
              <span className="max-w-[280px] truncate">{settings.output_dir}</span>
            </div>
          ) : null
        }
      />

      <div className="grid grid-cols-12 gap-5">
        <div className="col-span-7 space-y-5">
          <Card>
            <div className="mb-4 flex items-center justify-between">
              <div>
                <div className="section-title">1. 选择文件</div>
                <div className="mt-0.5 text-xs text-slate-400">支持混合选择，但所有文件必须都能转为同一个目标格式</div>
              </div>
              {files.length ? <span className="badge bg-brand-50 text-brand-700">{files.length} 个文件</span> : null}
            </div>
            <Dropzone files={files} onChange={setFiles} />
          </Card>

          <Card>
            <div className="mb-4">
              <div className="section-title">2. 目标格式</div>
              <div className="mt-0.5 text-xs text-slate-400">
                {activeSources.length
                  ? `源格式：${[...new Set(activeSources.map((source) => formatLabels[source]))].join(' / ')}`
                  : '选择文件后会自动过滤不支持的方向'}
              </div>
            </div>
            {capabilities ? (
              <FormatGrid
                formats={capabilities.formats}
                matrix={capabilities.matrix}
                sources={sources}
                value={target}
                onChange={setTarget}
              />
            ) : (
              <Spinner label="加载格式能力…" />
            )}
            {targetCapability ? (
              <div className="mt-3 flex items-center gap-2 text-[11px] text-slate-400">
                <Sparkles size={13} />
                质量：
                {targetCapability.quality === 'high'
                  ? '高保真'
                  : targetCapability.quality === 'beta'
                    ? 'Beta 能力'
                    : '结构/内容转换'}
                {targetCapability.status === 'beta' ? ' · 表格提取为尽力而为' : ''}
              </div>
            ) : null}
          </Card>
        </div>

        <div className="col-span-5 space-y-5">
          <Card>
            <div className="section-title mb-4">3. 转换选项</div>

            {showMode ? (
              <div className="mb-4">
                <span className="label">PDF 转换模式</span>
                <div className="grid grid-cols-2 gap-2">
                  {[
                    { id: 'editable', title: '可编辑', description: '标题 / 段落 / 表格规整' },
                    { id: 'layout', title: '高保真', description: '保留原版面位置' },
                  ].map((option) => (
                    <button
                      key={option.id}
                      className={cn(
                        'rounded-xl border px-3 py-2.5 text-left transition',
                        mode === option.id
                          ? 'border-brand-500 bg-brand-50/80'
                          : 'border-slate-200 bg-white hover:border-brand-300',
                      )}
                      onClick={() => setMode(option.id)}
                    >
                      <div className={cn('text-sm font-medium', mode === option.id ? 'text-brand-700' : 'text-slate-700')}>
                        {option.title}
                      </div>
                      <div className="mt-0.5 text-[11px] leading-4 text-slate-400">{option.description}</div>
                    </button>
                  ))}
                </div>
              </div>
            ) : null}

            <div className="mb-4">
              <span className="label">{"输出样式（仅字体/配色）"}</span>
              <select className="input" value={styleId} onChange={(event) => setStyleId(event.target.value)}>
                <option value="">不使用样式模板</option>
                {applicableStyles.map((style) => (
                  <option key={style.id} value={style.id}>
                    {style.name} · {style.applies_to.map((item) => formatLabels[item as FormatId] || item).join(' / ')}
                  </option>
                ))}
              </select>
              <div className="mt-1.5 text-[11px] text-slate-400">
                {"输出样式只改变字体、配色与页边距，不会套用数据模板的封面和版式；PDF 仅应用水印。"}
              </div>
            </div>

            {hasPdf ? (
              <div className="mb-4">
                <span className="label">PDF 密码（可选）</span>
                <div className="relative">
                  <LockKeyhole size={15} className="absolute left-3 top-3 text-slate-400" />
                  <input
                    className="input pl-9"
                    type="password"
                    placeholder="加密 PDF 才需要填写"
                    value={pdfPassword}
                    onChange={(event) => setPdfPassword(event.target.value)}
                  />
                </div>
              </div>
            ) : null}
            {hasCsv && target === 'xlsx' ? (
              <label className="mb-4 flex cursor-pointer items-start gap-2.5 rounded-xl border border-slate-100 bg-slate-50/60 px-3 py-2.5">
                <input
                  type="checkbox"
                  className="mt-0.5 accent-brand-600"
                  checked={csvAllText}
                  onChange={(event) => setCsvAllText(event.target.checked)}
                />
                <span>
                  <span className="block text-xs font-medium text-slate-700">全部按文本导入</span>
                  <span className="mt-0.5 block text-[11px] leading-4 text-slate-400">避免编号前导零被 Excel 当作数字处理</span>
                </span>
              </label>
            ) : null}

            {hasXlsx && ['csv', 'docx', 'md', 'pdf'].includes(target) ? (
              <div className="mb-4">
                <span className="label">指定工作表（可选）</span>
                <input
                  className="input"
                  placeholder="留空使用第一个工作表"
                  value={sheetName}
                  onChange={(event) => setSheetName(event.target.value)}
                />
              </div>
            ) : null}

            <div className="mb-5 rounded-xl border border-slate-100 bg-slate-50/70 px-3.5 py-3">
              <div className="flex items-center justify-between text-xs text-slate-500">
                <span>待转换</span>
                <span className="font-medium text-slate-700">{files.length} 个文件</span>
              </div>
              <div className="mt-1.5 flex items-center gap-2 text-xs text-slate-500">
                <span>
                  {activeSources.length
                    ? [...new Set(activeSources.map((item) => formatLabels[item]))].join(' / ')
                    : '—'}
                </span>
                <ArrowRight size={12} />
                <span className="font-medium text-brand-700">{formatLabels[target]}</span>
              </div>
            </div>

            {error ? (
              <Notice tone="error" className="mb-4">
                {error}
              </Notice>
            ) : null}

            <button
              className="btn-primary w-full !py-3"
              disabled={mutation.isPending || files.length === 0 || !targetCapability}
              onClick={start}
            >
              {mutation.isPending ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />}
              {mutation.isPending ? '正在创建任务…' : '开始转换'}
            </button>
            <div className="mt-2 text-center text-[11px] text-slate-400">文件只在本机处理，不会上传到云端</div>
          </Card>

          <div className="rounded-2xl border border-brand-100 bg-gradient-to-br from-brand-50 to-white p-5">
            <div className="text-sm font-semibold text-brand-900">小提示</div>
            <ul className="mt-2 space-y-1.5 text-xs leading-5 text-brand-800/70">
              <li>· PDF 扫描件 v1 不做 OCR，可编辑模式会提示，高保真模式保留图片版。</li>
              <li>· WPS 转换前请关闭已打开的 WPS 文档，避免抢占用户会话。</li>
              <li>· 批量结果会在完成后归档到设置中的输出目录。</li>
            </ul>
          </div>
        </div>
      </div>

      {activeTask ? <ActiveTask initial={activeTask} onReset={() => setActiveTask(null)} /> : null}
    </div>
  )
}
