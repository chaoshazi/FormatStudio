import { useQuery } from '@tanstack/react-query'
import { ArrowLeftRight, Layers, Settings, Sparkles } from 'lucide-react'
import { NavLink, Outlet } from 'react-router-dom'
import { apiClient } from '../api'
import { cn } from '../utils'

const navItems = [
  { to: '/', label: '转换工作台', icon: ArrowLeftRight, end: true },
  { to: '/tasks', label: '任务中心', icon: Layers },
  { to: '/templates', label: '模板中心', icon: Sparkles },
  { to: '/settings', label: '设置', icon: Settings },
]

export default function Layout() {
  const { data } = useQuery({ queryKey: ['capabilities'], queryFn: apiClient.capabilities, staleTime: 30_000 })
  const engines = data?.engines || {}
  const engineList = Object.values(engines)

  return (
    <div className="flex h-full min-h-screen bg-slate-50">
      <aside className="fixed inset-y-0 left-0 flex w-60 flex-col border-r border-slate-200/80 bg-white">
        <div className="flex items-center gap-3 px-6 py-6">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-brand-600 text-lg font-bold text-white shadow-lg shadow-brand-600/20">
            格
          </div>
          <div>
            <div className="text-[15px] font-semibold tracking-tight text-slate-900">格式工坊</div>
            <div className="text-[11px] text-slate-400">Format Studio</div>
          </div>
        </div>
        <nav className="flex-1 space-y-1 px-3">
          {navItems.map((item) => {
            const Icon = item.icon
            return (
              <NavLink
                key={item.to}
                to={item.to}
                end={item.end}
                className={({ isActive }) =>
                  cn(
                    'flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition',
                    isActive
                      ? 'bg-brand-50 text-brand-700'
                      : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900',
                  )
                }
              >
                <Icon size={18} strokeWidth={1.9} />
                {item.label}
              </NavLink>
            )
          })}
        </nav>
        <div className="mx-3 mb-4 rounded-xl border border-slate-100 bg-slate-50/70 p-3">
          <div className="mb-2 text-[10px] font-semibold uppercase tracking-widest text-slate-400">转换引擎</div>
          {engineList.length === 0 ? (
            <div className="text-xs text-slate-400">正在探测…</div>
          ) : (
            engineList.map((engine) => (
              <div key={engine.id} className="flex items-center justify-between py-1 text-xs">
                <span className="text-slate-600">{engine.label}</span>
                <span
                  className={cn(
                    'flex items-center gap-1.5',
                    engine.available && !engine.busy ? 'text-emerald-600' : engine.busy ? 'text-amber-600' : 'text-slate-400',
                  )}
                  title={engine.error || engine.version || ''}
                >
                  <span
                    className={cn(
                      'h-1.5 w-1.5 rounded-full',
                      engine.available && !engine.busy ? 'bg-emerald-500' : engine.busy ? 'bg-amber-500' : 'bg-slate-300',
                    )}
                  />
                  {engine.busy ? '忙碌' : engine.available ? '可用' : '未安装'}
                </span>
              </div>
            ))
          )}
        </div>
        <div className="border-t border-slate-100 px-6 py-4 text-[11px] text-slate-400">本地处理 · 文件不出机器</div>
      </aside>
      <main className="ml-60 min-h-screen flex-1">
        <div className="mx-auto max-w-[1240px] px-8 py-8">
          <Outlet />
        </div>
      </main>
    </div>
  )
}