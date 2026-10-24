# Codex PPTX

自用 PowerPoint 工具包：创建、读取、编辑、合并、渲染和检查 PPTX。

## 内容

- `skills/pptx/SKILL.md`：入口与工作约定。
- `skills/pptx/references/`：生成、视觉设计、技术对接和质量检查指南。
- `skills/pptx/scripts/`：PptxGenJS 启动、LibreOffice 渲染、缩略图、Office XML 解包、打包和校验工具。
- `plugin.json`：插件清单。
- `.agents/plugins/marketplace.json`：仓库插件目录。

## 使用

在有权限访问此私有仓库的机器上，先配置 GitHub 身份，再添加插件来源：

```bash
codex plugin marketplace add yss-off/codex-pptx --ref main
```

随后在支持本地插件的桌面客户端插件目录中选择 Personal PPTX 并安装。
这里仅提供打包和分发文件，尚未在另一台机器上验证安装。

也可将 `skills/pptx` 复制到工具实际使用的技能目录；已有同名技能时先比较并备份。

## 依赖

Python、Node.js/npm、LibreOffice、Poppler，以及：

```bash
python -m pip install 'markitdown[pptx]' Pillow lxml defusedxml
npm install -g pptxgenjs
```

可选图标依赖：`react-icons`、`react`、`react-dom`、`sharp`。
中文演示文稿需要运行机器上可用的中文字体。

## 来源和许可

本仓库保存本机现用 PPTX 技能的快照，排除了 Python 缓存；原始技能和脚本未作修改。
原 `SKILL.md` 标有 `Proprietary. LICENSE.txt has complete terms`，但本机源目录未提供 `LICENSE.txt`。
本仓库保留这一声明，不为原有内容重新授予开源许可；当前作为私有自用备份保存。
第三方 Office XML schema 保留其原有文件内容与声明。

插件结构依据 [OpenAI 官方打包文档](https://developers.openai.com/plugins/build/plugins)。
