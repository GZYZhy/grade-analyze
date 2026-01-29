# grade-analyze

一款成绩分析工具（Web 方式，跨平台运行）。

## 运行环境

- Python 3.10+（建议使用 3.11）
- Windows / Linux / macOS 均可

## 安装

1. 克隆仓库并进入项目目录：

```powershell
git clone <repo> && cd grade-analyze
```

1. 创建虚拟环境并安装依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1   # Windows PowerShell

## 启动
本地快速启动：

```powershell
streamlit run app.py
```

## 功能概览

- 导入两种格式的成绩单：输入成绩单（考试明细）和通常成绩单（按学生行）
- 支持 .xlsx 与 .xls 文件（注意：不接受含宏的 .xlsm/.xlsb）

- 成绩趋势 / 学生对比 / 统计分析（支持百分位、阈值筛选、两场考试对比）
- 表格导出（Excel）、图表导出（PNG）与按学生批量报告导出

- 数据管理：考试顺序配置、修正成绩、删除记录、导出数据库
- 支持样例下载（samples/目录下的 Excel 文件）

- 简单的应用层密码保护（可由环境变量提供密码哈希以便部署时配置）

## 快速使用指南

- 样例文件：`samples/输入成绩单.xlsx` 和 `samples/通常成绩单.xlsx`，在“导入成绩”页面可以点击下载样例。
- 导入时可选择工作表并映射列名，应用会对常见列做猜测并允许手动调整映射。

- 在“数据管理”中可以配置考试显示顺序（上/下移动），删除考试与学生，以及打开单个学生的“修正成绩”编辑器进行手工修改。

## 配置与环境变量

建议把敏感配置放到环境变量或 secret 管理器，不要把密码写入仓库。

- GRADE_ANALYZE_PASSWORD_HASH: （可选）bcrypt 哈希字符串，若设置会替代数据库中保存的密码哈希，用于部署时统一管理。
- MAX_UPLOAD_MB: 上传文件最大体积（MB），默认 10

示例（Linux / systemd 环境）:

```bash
export GRADE_ANALYZE_PASSWORD_HASH='$2b$12$...'
export MAX_UPLOAD_MB=20
```

## 数据文件与目录

- 数据库：`data/grade_analyze.db`（SQLite）
- 导出目录：`exports/`（报告与导出文件）

请确保这些目录对运行用户具有适当的读写权限，并且数据库文件不被 Web 可访问（非 wwwroot）。

## 安全建议（简要）

- 在反向代理层终端 TLS 并注入安全头（HSTS、X-Frame-Options、X-Content-Type-Options、CSP 等）。
- 限制上传大小并在代理与应用层双重限制（已在应用支持 MAX_UPLOAD_MB）。

- 运行服务时使用最小权限的专用用户，并设置合适的文件权限。
- 定期备份 `data/grade_analyze.db` 与 `exports/`。

- 对外网部署时考虑使用 WAF / fail2ban / 登录限流，或通过 OAuth2/OpenID Connect 集成更完善的认证系统。

## 导出与兼容性

- 导出的 Excel/图片功能保持与数据导出一致；统计页面展示会将列名翻译为中文以便查看（导出文件仍使用内部字段列名以便兼容老的分析脚本）。

如果你希望导出文件也使用中文列名与序号，请在 Issues 中告知，我可以调整导出逻辑以与界面完全一致。

## 故障排查

- 无法读取 .xls 文件：请确保安装了支持旧格式的依赖（如 `xlrd`）；推荐使用 `.xlsx`。
- 上传被拒绝：检查 `MAX_UPLOAD_MB` 环境变量和代理的 `client_max_body_size` 设置。

- 权限/数据库问题：检查运行用户对 `data/` 目录的读写权限。

## 开发与贡献

- 代码基于 Python + Streamlit；欢迎提交 issue / PR。请在提交前运行基本测试并保持依赖清单更新。

## 许可证

请查看仓库根目录下的 `LICENSE` 文件。

# grade-analyze

一款成绩分析工具（Web 方式，跨平台运行）。

## 运行环境

- Python 3.9
- macOS / Windows

## 安装依赖

在项目根目录执行：

```
pip install -r requirements.txt
```

## 启动应用

```
streamlit run app.py
```

## 功能概览

- 导入输入成绩单 / 通常成绩单
- 成绩趋势展示（PPT 风格循环）
- 导出分析报告（单文件或按学生拆分）
- 学生对比
- 统计分析（阈值、进退步、百分位线、导出 Excel/图片）
- 数据管理（修正、删除、导出全量数据）
- 多班级切换
- Highlight 科目
- 软件打开密码保护
