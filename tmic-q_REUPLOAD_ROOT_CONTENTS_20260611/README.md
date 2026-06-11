# TMIC-Q 插件市场

这个仓库包含一个 Codex 本地插件市场，用于分发和安装 `tmic-q` 插件。

## 目录结构

- `.agents/plugins/marketplace.json`：插件市场配置
- `plugins/tmic-q`：插件源码、技能说明、脚本、参考文档和 TMIC 问卷模板资产

## 安装方式

同事拉取或下载这个仓库后，在仓库根目录运行：

```bash
codex plugin marketplace add .
codex plugin add tmic-q@tmic-team
```

安装或更新插件后，请新开一个 Codex 对话，让 Codex 读取最新插件能力。

## 使用方式

在 Codex 中输入或提及 `@tmic-q`，然后提供概念卡图片、TMIC 模板相关文件，或修改后的完整问卷工作簿。

插件会生成完整 TMIC 问卷工作簿和上传专用单表。默认流程只到表格文件交付为止，不会接管 Chrome，不会进入 TMIC 后台设置网页逻辑、保存或发布问卷。
