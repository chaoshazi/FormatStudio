# AGENTS.md

本仓库是 Windows 上运行的本地多格式转换桌面应用「格式工坊 / Format Studio」。后端 FastAPI，前端 React + TypeScript + Vite + Tailwind，桌面壳 pywebview。

## 项目结构

- `app/`：FastAPI 应用、数据库、能力矩阵、任务队列。
- `app/services/`：转换引擎、PDF/Office/Markdown/表格转换、模板与样式服务。
- `web/`：React 前端；`web/dist` 为构建产物，由后端托管。
- `tests/`：pytest 测试，样本在测试运行时动态生成。
- `scripts/`：开发、打包、UI 冒烟脚本。
- `installer/`：Inno Setup 安装脚本；`vendor/` 放置 LibreOffice Portable。

## 常用命令

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe run_server.py
cd web; npm run dev
cd web; npm run build
.\.venv\Scripts\python.exe desktop.py
```

开发后端默认端口 8765；可用 `FORMAT_STUDIO_PORT` 覆盖。桌面模式自动生成 API Token，开发可设置 `FORMAT_STUDIO_DISABLE_TOKEN=1`。
## 工程约定

- 转换能力以 `app/capabilities.py` 为唯一事实来源；新增方向必须同步前端能力矩阵与测试。
- 转换逻辑放 `app/services/`，API 层只做校验、持久化与调度。
- 任务串行执行（`ThreadPoolExecutor(max_workers=1)`）；WPS COM 与 LibreOffice 都通过子进程/超时隔离，不要在请求线程内直接长跑。
- 输出文件一律写入任务目录，成功后再复制到用户输出目录；下载接口只接受 task/item ID，禁止原始路径参数。
- 文件名、模板字段、ZIP 路径都要做清洗与路径穿越防护。
- 不提交 `*.docx/xlsx/pptx/pdf` 等真实业务样本；测试样本运行时生成。
- 前端使用中文界面与浅色专业风；新增页面复用 `components/UI.tsx` 的 Card / Notice / Modal / StatusBadge。

## 引擎与许可

- 原生库负责结构转换；Office → PDF 使用 WPS COM 优先、LibreOffice 兜底，xlsx/csv 优先 LibreOffice。
- `.wps` 仅 WPS COM；检测到用户可见的 WPS 窗口时标记忙碌，不得抢占用户会话。
- 闭源商用构建禁止引入 AGPL/强著佐权依赖（pdf2docx、PyMuPDF、docxtpl 等）。
- 新增依赖前确认许可证与 Windows 兼容性；LibreOffice 以 MPL-2.0 随包分发。

## 验证

- 提交前运行 `pytest` 与 `cd web; npm run build`。
- PDF 高保真模式至少验证页数与文本位置；模板功能至少验证跨 run 占位符、缺失字段与批量 ZIP。
- UI 改动建议运行 `scripts/ui_smoke.py`（需要本机 Edge）。