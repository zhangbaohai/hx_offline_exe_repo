# 零配置自动产出 EXE（v2）

这个仓库已内置 GitHub Actions 工作流。你只要把它推到 GitHub：

1. 新建仓库（main 或 master 均可），把本目录所有文件上传。
2. 打开仓库 **Actions**，如提示先点 **Enable**。
3. 每次 push 自动构建；也可在 Actions 里点 **Run workflow** 手动触发。
4. 构建完成后到该 workflow 运行页 **Artifacts** 下载：
   - `hx-offline-editor-dist-v2`（目录版，更稳）
   - `hx-offline-editor-exe-v2`（单文件版，如生成）

已包含：
- `app_exact.py`（v2：支持 **TXT 导入** & **新增/编辑对话框**）
- `db_helper.py`、`requirements.txt`、`README.txt`
- `icon.ico`（可替换为你的企业图标）
- `app_exact.spec`（PyInstaller 打包规约，收集 pandas/openpyxl 隐式依赖）
- `.github/workflows/build.yml`（Windows 打包流水线）

> 如需预置行号库，把你的 IBPS/CNAPS 源文件放到仓库里，并在 `app_exact.py` 中首次运行时加载写入本地 `codebook.db`（我可应你需求补上该逻辑）。