<p align="center">
  <img src="docs/assets/banner.svg" alt="Codex PPTX — Create. Edit. Render. Verify." width="100%">
</p>

<p align="center">
  <a href="https://github.com/yss-off/codex-pptx/actions/workflows/tests.yml"><img src="https://github.com/yss-off/codex-pptx/actions/workflows/tests.yml/badge.svg?branch=main" alt="Regression tests"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-167D8D" alt="MIT License"></a>
  <a href="plugin.json"><img src="https://img.shields.io/badge/version-0.1.1-16324F" alt="Version 0.1.1"></a>
  <a href="requirements-dev.txt"><img src="https://img.shields.io/badge/Python-3.10%2B-16324F" alt="Python 3.10 or later"></a>
</p>

<p align="center"><strong>Turn content into editable, inspectable PowerPoint presentations with Codex.</strong><br>A packaged skill and local tools for creating, editing, rendering, and checking slides.</p>

<p align="center"><a href="README.md">中文</a> · English · <a href="#quick-start">Quick start</a> · <a href="docs/validation.md">Validation record</a></p>

## Workflows

| Workflow | What it provides |
|---|---|
| **Create** | Editable text, shapes, tables, and native charts with PptxGenJS |
| **Edit a template** | Unpack, duplicate, reorder, and clean slides while retaining the template structure |
| **Inspect** | Extract text, render PDF and slide images, and validate Office XML and resource references |

No additional MCP server is required. The plugin supplies the workflow; content review and visual inspection remain part of producing a good presentation.

## Actual output

These slides were created by the installed plugin in a fresh Codex session and rendered with LibreOffice.
The chart contains explicitly labeled example data, not benchmark results.

<table>
  <tr>
    <td width="50%"><img src="docs/assets/preview-cover.jpg" alt="Chinese presentation cover created by the installed plugin"></td>
    <td width="50%"><img src="docs/assets/preview-chart.jpg" alt="Editable native chart with example data"></td>
  </tr>
</table>

[Download the editable example PPTX](examples/codex-pptx-example.pptx)

## Quick start

### 1. Dependencies

Install Codex, Python 3.10+, Node.js/npm, LibreOffice, and Poppler.

```bash
python -m pip install 'markitdown[pptx]' Pillow lxml defusedxml
npm install -g pptxgenjs
```

On Ubuntu / Debian, system dependencies can be installed with:

```bash
sudo apt install libreoffice poppler-utils fontconfig fonts-noto-cjk
```

Use an installed CJK font for Chinese slides. Optional icon dependencies are `react-icons`, `react`, `react-dom`, and `sharp`.
The LibreOffice socket compatibility branch on restricted Linux hosts also needs `gcc`.

### 2. Install from GitHub

```bash
codex plugin marketplace add yss-off/codex-pptx --ref main
codex plugin add codex-pptx@personal-pptx
codex plugin list --json --marketplace personal-pptx
```

The desktop plugin directory can also expose the **Personal PPTX** source. CLI installation has been tested; the desktop installation UI has not been separately verified.

If you already have a standalone `pptx` skill, back it up outside the skill discovery directories and keep one active entry point. Do not keep the backup under `skills/`.

### 3. Use a fresh chat

```text
Use codex-pptx:pptx to turn the following content into a presentation.
Keep charts natively editable, verify the content, render every slide,
and deliver the final PPTX.
```

You can also ask Codex to edit an existing template or review a deck. Existing chats may retain the previous skill context; verify the loaded source in a fresh chat after switching.

## Validation behavior

`pack.py` validates by default. `--original` compares XSD errors with the source file; it does not enable validation.
Only an explicit `--validate false` skips validation. Duplicated slides receive independent chart parts and editable dependencies; speaker notes are not inherited automatically.

## Development

```bash
python -m pip install -r requirements-dev.txt
npm install -g pptxgenjs@4.0.1
python -m unittest discover -s tests -v
```

Regression tests run in temporary directories without a model call or GPU.
Generation integration tests require Node.js and PptxGenJS; missing dependencies
produce explicit skips, not full validation. CI installs the pinned generator.
GitHub Actions runs the suite on Python 3.10 and 3.13. Linux CLI installation and invocation have been verified; Windows/macOS have not.

See the [validation record](docs/validation.md), [contribution guide](CONTRIBUTING.md), and [changelog](CHANGELOG.md).
Issues and focused pull requests are welcome.

## License

Original project content is available under the [MIT License](LICENSE).
Third-party Office XML schemas retain their original notices; this project's license does not replace their terms.

Plugin packaging follows the [official OpenAI documentation](https://developers.openai.com/plugins/build/plugins).
