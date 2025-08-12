
# 零操作自动出 EXE（GitHub Actions）

1. 把整个文件夹推到一个新的 GitHub 仓库（默认分支 main/master 均可）。
2. 打开仓库的 **Actions** → 找到 "Build Windows EXE" → 首次点击 **Enable**（如提示）。
3. 之后每次 push 会自动触发构建；或在 Actions 页面点 **Run workflow** 立即构建。
4. 构建完成后，在 **Artifacts** 区域下载：
   - `hx-offline-editor-dist`（目录版，更稳）
   - `hx-offline-editor-exe`（单文件版，如有生成）

> 如需 Release 自动发布，我也可以把工作流改成 push tag 自动创建 Release 并附带 exe。
