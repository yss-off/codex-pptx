# 0.1.1 验证记录

验证日期：2026-10-24。环境：Linux、Codex CLI 0.163.0、Python 3.13.12、
Node.js 24.16.0、PptxGenJS 4.0.1，以及本机 LibreOffice、Poppler 和中文字体。

## 回归检查

```bash
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

16 个测试通过。修复前的测试先复现了原有故障；测试还覆盖正常文件打包、
实际删除页面，以及图表复制后清理、校验和重新打包的完整链路。

另已实际编译 socket 兼容库，确认构建目录权限为 `0700`，并通过该库完成
LibreOffice PDF 转换。带隐藏页面的三页演示文稿也完成了缩略图渲染。

## Codex 安装与调用

在隔离的 `CODEX_HOME` 中，通过 CLI 添加本地仓库来源并安装
`codex-pptx@personal-pptx`。安装列表确认版本 `0.1.1`、已安装且已启用。
安装缓存中的 65 个技能文件与待提交源码逐字节一致。

在隔离工作目录开启新的 `codex exec` 会话，调用 `codex-pptx:pptx`，
实际读取安装缓存中的技能指南和脚本，而非原独立技能。会话完成：

- 创建两页中文 PPTX。
- 生成原生可编辑条形图，类别 A/B、数值 10/20，包内包含嵌入工作簿。
- 使用已安装的生成和渲染脚本，产出 PDF 和两张页面图片。
- 通过 MarkItDown 核对正文、图表数值和顺序，并目视检查两页渲染。
- 输出验证报告，新会话正常结束。

## 验证边界

本次使用现有机器的系统依赖，隔离了 Codex 配置、插件缓存和工作目录。
未验证另一台机器的依赖安装、Windows/macOS 或桌面图形界面的安装操作。
普通 LibreOffice 渲染有不阻塞输出的 `javaldx` 警告。
GitHub Actions 配置在 Python 3.10 和 3.13 上执行同一组回归测试。
