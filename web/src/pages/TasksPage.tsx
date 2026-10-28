import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Download, FolderOpen, RefreshCw, RotateCcw, XCircle } from 'lucide-react'
import { useNavigate, useParams } from 'react-router-dom'
import { apiClient, download } from '../api'
import { Card, EmptyState, Modal, Notice, PageHeader, ProgressBar, Spinner, StatusBadge } from '../components/UI'
import type { Task } from '../types'
import { formatLabels, formatTime, metaSummary } from '../utils'

function TaskDetail({ taskId, onClose }: { taskId: string; onClose: () => void }) {
  const queryClient = useQueryClient()
  const { data: task, isLoading } = useQuery({
    queryKey: ['task', taskId],
    queryFn: () => apiClient.task(taskId),
    refetchInterval: (query) => {
      const status = query.state.data?.status
      return status && ['queued', 'running'].includes(status) ? 1200 : false
    },
  })

  const cancel = useMutation({
    mutationFn: () => apiClient.cancelTask(taskId),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ['task', taskId] }),
  })
  const retry = useMutation({
    mutationFn: () => apiClient.retryTask(taskId),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ['task', taskId] }),
  })

  return (
    <Modal open title="任务详情" onClose={onClose} width="max-w-3xl">
      {isLoading || !task ? (
        <Spinner label="加载任务详情…" />
      ) : (
        <div className="space-y-5">
          <div className="flex items-start justify-between gap-4">
            <div>
              <div className="flex items-center gap-2.5">
                <StatusBadge status={task.status} />
                <span className="text-xs text-slate-400">{task.id}</span>
              </div>
              <div className="mt-2 text-xs text-slate-400">创建于 {formatTime(task.created_at)}</div>
            </div>
            <div className="flex flex-wrap items-center justify-end gap-2">
              {['queued', 'running'].includes(task.status) ? (
                <button className="btn-danger" disabled={cancel.isPending} onClick={() => cancel.mutate()}>
                  <XCircle size={15} /> 取消
                </button>
              ) : null}
              {['failed', 'partial', 'canceled', 'interrupted'].includes(task.status) ? (
                <button className="btn-secondary" disabled={retry.isPending} onClick={() => retry.mutate()}>
                  <RotateCcw size={15} /> 重试
                </button>
              ) : null}
              {['succeeded', 'partial'].includes(task.status) ? (
                <>
                  <button className="btn-secondary" onClick={() => apiClient.openFolder(task.id)}>
                    <FolderOpen size={15} /> 打开目录
                  </button>
                  <button className="btn-primary" onClick={() => void download(`/api/v1/tasks/${task.id}/download`, `\u4efb\u52a1_${task.id.slice(0, 6)}_\u8f6c\u6362\u7ed3\u679c.zip`)}>
                    <Download size={15} /> 全部下载
                  </button>
                </>
              ) : null}
            </div>
          </div>

          <div>
            <div className="mb-2 flex justify-between text-xs text-slate-500">
              <span>{task.status === 'queued' ? `排队中 · 第 ${task.queue_position} 位` : '总体进度'}</span>
              <span>{task.progress}%</span>
            </div>
            <ProgressBar value={task.progress} />
          </div>

          <div className="space-y-3">
            {task.items.map((item) => (
              <div key={item.id} className="rounded-xl border border-slate-100 bg-slate-50/40 p-4">
                <div className="flex items-start justify-between gap-4">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="truncate text-sm font-medium text-slate-800">{item.source_name}</span>
                      <StatusBadge status={item.status} />
                    </div>
                    <div className="mt-1 text-xs text-slate-400">
                      {formatLabels[item.source_format]} → {formatLabels[item.target_format]}
                      {item.engine ? ` · 引擎 ${item.engine}` : ''}
                      {item.progress && item.status === 'running' ? ` · ${item.progress}%` : ''}
                    </div>
                    {metaSummary(item.meta).length ? (
                      <div className="mt-1 text-[11px] text-slate-400">{metaSummary(item.meta).join(' · ')}</div>
                    ) : null}
                  </div>
                  {item.status === 'succeeded' ? (
                    <button className="btn-secondary !px-3 !py-1.5 text-xs" onClick={() => void download(item.download_url, item.download_name)}>
                      <Download size={13} /> 下载
                    </button>
                  ) : null}
                </div>
                {item.error_code ? (
                  <Notice tone="error" className="mt-3">
                    <span className="font-medium">{item.error_code}</span>
                    <span className="ml-2">{item.error_message}</span>
                  </Notice>
                ) : null}
                {item.warning ? (
                  <Notice tone="warning" className="mt-3">
                    {item.warning}
                  </Notice>
                ) : null}
              </div>
            ))}
          </div>

          {cancel.isError || retry.isError ? (
            <Notice tone="error">{String((cancel.error || retry.error) as Error)}</Notice>
          ) : null}
        </div>
      )}
    </Modal>
  )
}

export default function TasksPage() {
  const navigate = useNavigate()
  const { id } = useParams()
  const { data: tasks = [], isLoading } = useQuery({
    queryKey: ['tasks'],
    queryFn: apiClient.tasks,
    refetchInterval: 2500,
  })

  return (
    <div>
      <PageHeader
        title="任务中心"
        description="查看队列与历史任务。转换会串行执行，单个任务最多 50 个文件。"
        actions={
          <button className="btn-secondary" onClick={() => navigate('/')}>
            新建转换
          </button>
        }
      />

      {isLoading ? (
        <Card>
          <Spinner label="加载任务…" />
        </Card>
      ) : tasks.length === 0 ? (
        <EmptyState
          title="还没有转换任务"
          description="回到工作台选择文件，开始第一次转换。"
          action={
            <button className="btn-primary" onClick={() => navigate('/')}>
              去工作台
            </button>
          }
        />
      ) : (
        <Card className="!p-0">
          <div className="grid grid-cols-[1.6fr_1fr_0.9fr_1.2fr_0.8fr] gap-4 border-b border-slate-100 px-6 py-3 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
            <span>文件 / 方向</span>
            <span>状态</span>
            <span>进度</span>
            <span>时间</span>
            <span className="text-right">操作</span>
          </div>
          <div className="divide-y divide-slate-50">
            {tasks.map((task) => {
              const first = task.items[0]
              return (
                <div
                  key={task.id}
                  className="grid cursor-pointer grid-cols-[1.6fr_1fr_0.9fr_1.2fr_0.8fr] items-center gap-4 px-6 py-3.5 transition hover:bg-slate-50/60"
                  onClick={() => navigate(`/tasks/${task.id}`)}
                >
                  <div className="min-w-0">
                    <div className="truncate text-sm font-medium text-slate-800">
                      {first ? first.source_name : '任务'}
                      {task.items.length > 1 ? <span className="ml-2 text-xs text-slate-400">等 {task.items.length} 个</span> : null}
                    </div>
                    <div className="mt-0.5 text-[11px] text-slate-400">
                      {first ? `${formatLabels[first.source_format]} → ${formatLabels[first.target_format]}` : ''}
                      {first?.engine ? ` · ${first.engine}` : ''}
                    </div>
                  </div>
                  <div>
                    <StatusBadge status={task.status} />
                  </div>
                  <div>
                    <div className="mb-1 text-[11px] text-slate-400">{task.progress}%</div>
                    <ProgressBar value={task.progress} />
                  </div>
                  <div className="text-xs text-slate-400">{formatTime(task.created_at)}</div>
                  <div className="flex items-center justify-end gap-1">
                    {['succeeded', 'partial'].includes(task.status) ? (
                      <button
                        className="rounded-lg p-2 text-slate-400 hover:bg-brand-50 hover:text-brand-600"
                        title="下载全部"
                        onClick={(event) => {
                          event.stopPropagation()
                          void download(`/api/v1/tasks/${task.id}/download`, `\u4efb\u52a1_${task.id.slice(0, 6)}_\u8f6c\u6362\u7ed3\u679c.zip`)
                        }}
                      >
                        <Download size={15} />
                      </button>
                    ) : null}
                    <button
                      className="rounded-lg p-2 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                      title="刷新"
                      onClick={(event) => {
                        event.stopPropagation()
                        navigate(`/tasks/${task.id}`)
                      }}
                    >
                      <RefreshCw size={15} />
                    </button>
                  </div>
                </div>
              )
            })}
          </div>
        </Card>
      )}

      {id ? <TaskDetail taskId={id} onClose={() => navigate('/tasks')} /> : null}
    </div>
  )
}