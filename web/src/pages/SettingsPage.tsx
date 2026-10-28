import { useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Download, FolderOpen, HardDrive, RefreshCw, Server, Trash2 } from 'lucide-react'
import { apiClient, ApiError, download } from '../api'
import { Card, Notice, PageHeader, Spinner } from '../components/UI'
import { cn } from '../utils'

export default function SettingsPage() {
  const queryClient = useQueryClient()
  const { data: settings, isLoading } = useQuery({ queryKey: ['settings'], queryFn: apiClient.settings })
  const { data: engines, isLoading: loadingEngines, refetch: refetchEngines } = useQuery({
    queryKey: ['engines'],
    queryFn: apiClient.engines,
  })
  const [outputDir, setOutputDir] = useState('')
  const [retention, setRetention] = useState(7)
  const [csvAllText, setCsvAllText] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    if (!settings) return
    setOutputDir(settings.output_dir)
    setRetention(settings.retention_days)
    setCsvAllText(settings.csv_all_text)
  }, [settings])

  const save = useMutation({
    mutationFn: () =>
      apiClient.updateSettings({ output_dir: outputDir, retention_days: retention, csv_all_text: csvAllText }),
    onSuccess: () => {
      setMessage('设置已保存')
      setError('')
      queryClient.invalidateQueries({ queryKey: ['settings'] })
    },
    onError: (err) => {
      setMessage('')
      setError(err instanceof ApiError ? err.message : String(err))
    },
  })

  const probe = useMutation({
    mutationFn: apiClient.probeEngines,
    onSuccess: (data) => {
      queryClient.setQueryData(['engines'], data)
      queryClient.invalidateQueries({ queryKey: ['capabilities'] })
      setMessage('引擎状态已重新探测')
    },
  })

  const cleanup = useMutation({
    mutationFn: apiClient.cleanup,
    onSuccess: (result) => {
      setMessage(`清理完成，移除 ${result.task_dirs_removed} 个过期任务目录`)
      setError('')
    },
    onError: (err) => setError(err instanceof ApiError ? err.message : String(err)),
  })

  return (
    <div>
      <PageHeader title="设置" description="输出位置、历史文件保留策略与本地转换引擎状态。" />

      {isLoading ? (
        <Card>
          <Spinner label="加载设置…" />
        </Card>
      ) : (
        <div className="grid grid-cols-12 gap-5">
          <div className="col-span-7 space-y-5">
            <Card>
              <div className="section-title mb-4">基础设置</div>
              <div className="space-y-4">
                <div>
                  <span className="label">默认输出目录</span>
                  <div className="relative">
                    <FolderOpen size={15} className="absolute left-3 top-3 text-slate-400" />
                    <input className="input pl-9" value={outputDir} onChange={(event) => setOutputDir(event.target.value)} />
                  </div>
                  <div className="mt-1.5 text-[11px] leading-5 text-slate-400">
                    每个任务会在此目录下创建带时间戳的子目录；转换完成时自动归档。
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <span className="label">任务文件保留期</span>
                    <select className="input" value={retention} onChange={(event) => setRetention(Number(event.target.value))}>
                      <option value={1}>1 天</option>
                      <option value={7}>7 天</option>
                      <option value={30}>30 天</option>
                      <option value={90}>90 天</option>
                    </select>
                  </div>
                  <div className="flex items-end pb-2">
                    <label className="flex items-center gap-2 text-xs text-slate-600">
                      <input
                        type="checkbox"
                        className="accent-brand-600"
                        checked={csvAllText}
                        onChange={(event) => setCsvAllText(event.target.checked)}
                      />
                      CSV 默认全部按文本处理
                    </label>
                  </div>
                </div>

                {message ? <Notice tone="success">{message}</Notice> : null}
                {error ? <Notice tone="error">{error}</Notice> : null}

                <button className="btn-primary" disabled={save.isPending} onClick={() => save.mutate()}>
                  保存设置
                </button>
              </div>
            </Card>

            <Card>
              <div className="flex items-center justify-between">
                <div>
                  <div className="section-title">本地数据</div>
                  <div className="mt-0.5 text-xs text-slate-400">清理过期任务文件，不删除已归档到输出目录的副本。</div>
                </div>
                <button className="btn-danger" disabled={cleanup.isPending} onClick={() => cleanup.mutate()}>
                  <Trash2 size={15} /> 立即清理
                </button>
              </div>
              <div className="mt-4 flex items-center justify-between rounded-xl border border-slate-100 bg-slate-50/60 px-4 py-3">
                <div className="flex items-center gap-3">
                  <HardDrive size={18} className="text-slate-400" />
                  <div>
                    <div className="text-sm text-slate-700">导出诊断包</div>
                    <div className="text-[11px] text-slate-400">包含日志、引擎版本与设置，不含文档内容。</div>
                  </div>
                </div>
                <button className="btn-secondary" onClick={() => void download('/api/v1/system/diagnostics', 'diagnostics.zip')}>
                  <Download size={15} /> 下载
                </button>
              </div>
            </Card>
          </div>

          <div className="col-span-5 space-y-5">
            <Card>
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="section-title">转换引擎</div>
                  <div className="mt-0.5 text-xs text-slate-400">WPS 用于 .wps 与高保真导出，LibreOffice 作为兜底。</div>
                </div>
                <button className="btn-ghost !p-2" disabled={probe.isPending} onClick={() => probe.mutate()} title="重新探测">
                  <RefreshCw size={16} className={cn(probe.isPending && 'animate-spin')} />
                </button>
              </div>

              {loadingEngines ? (
                <Spinner label="探测引擎…" />
              ) : (
                <div className="space-y-3">
                  {Object.values(engines || {}).map((engine) => (
                    <div key={engine.id} className="rounded-xl border border-slate-100 bg-slate-50/50 p-3.5">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2.5">
                          <Server size={16} className="text-slate-400" />
                          <span className="text-sm font-medium text-slate-800">{engine.label}</span>
                        </div>
                        <span
                          className={cn(
                            'badge',
                            engine.available && !engine.busy
                              ? 'bg-emerald-50 text-emerald-700'
                              : engine.busy
                                ? 'bg-amber-50 text-amber-700'
                                : 'bg-slate-100 text-slate-500',
                          )}
                        >
                          {engine.available && !engine.busy ? '可用' : engine.busy ? '忙碌' : '未安装'}
                        </span>
                      </div>
                      <div className="mt-2 space-y-1 text-[11px] leading-5 text-slate-400">
                        {engine.version ? <div>版本：{engine.version}</div> : null}
                        {engine.path ? <div className="break-all">路径：{engine.path}</div> : null}
                        {engine.error ? <div className="text-rose-500">{engine.error}</div> : null}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Card>

            <Card>
              <div className="section-title mb-3">关于</div>
              <div className="space-y-1.5 text-xs leading-6 text-slate-500">
                <div>格式工坊 · Format Studio 1.0.0</div>
                <div>本地单用户桌面工具，不联网、不上传文件。</div>
                <div>v1 不含 OCR；扫描件会明确提示，可用高保真模式保留图片版。</div>
                <div>闭源商用构建，依赖均为宽松许可；LibreOffice 以 MPL-2.0 随包分发。</div>
              </div>
            </Card>
          </div>
        </div>
      )}
    </div>
  )
}
