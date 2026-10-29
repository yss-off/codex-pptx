# 贡献说明

欢迎改进模板编辑、生成工作流和跨环境兼容性。

## 提交问题

请说明 Codex、Python、Node.js 和操作系统版本，提供执行命令、错误输出和
可公开的最小复现文件。不要上传真实凭据或未获授权公开的演示文稿。
排版问题请同时提供渲染截图；文本提取结果不能证明排版正确。

## 修改与验证

```bash
python -m pip install -r requirements-dev.txt
npm install -g pptxgenjs@4.0.1
python -m unittest discover -s tests -v
git diff --check
```

修复行为问题时，增加能捕获该故障的回归测试；脚本涉及内容修改时，
同时验证正常文件、错误输入和资源引用。生成或渲染链路的修改还需要实际
打开或渲染示例，记录无法验证的环境。

请保持改动聚焦，保留第三方 schema 的原有声明。文档或展示素材修改通常
无需增加行为测试；检查链接、命令和图片是否正确即可。
