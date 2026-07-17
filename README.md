# PPT Bulk Background Replacer / PPT批量背景替换工具

> Replace background images on every slide of PowerPoint (.pptx) files while preserving all content.
> 批量替换PPT文件每一页的背景图，保留所有原有内容。

[![Python](https://img.shields.io/badge/python-3.7+-blue)](https://python.org)

---

## Features / 功能特性

- **Replace every slide background** — Inserts new full-slide background on every page
- **Preserves all content** — Text, images, charts, animations, media all remain intact
- **Batch directory processing** — Process entire directory trees, preserving folder structure
- **Safe INSERT-only approach** — Does NOT corrupt content images (fixed: no more SWAP with extension mismatch)
- **Unicode/Chinese path support** — Reliable handling of international filenames

- **替换每页PPT背景** — 每张幻灯片插入新的全幅背景图
- **保留所有内容** — 文字、图片、图表、动画、媒体全部保留
- **批量处理目录** — 递归处理整个目录，保持文件夹结构
- **安全的INSERT方式** — 不会破坏内容图片（已修复：不再使用SWAP方式导致扩展名不匹配）
- **支持中文路径** — 可靠处理中文字符文件名和路径

## Requirements / 环境要求

- **Python 3.7+**
- **lxml** (pip install lxml)

## Usage / 使用方法

### Single file / 单文件

`ash
python scripts/replace_ppt_bg.py input.pptx background.jpg
# → input_bg_replaced.pptx

python scripts/replace_ppt_bg.py input.pptx background.jpg output.pptx
`

### Batch directory / 批量目录

`ash
# Auto output dir: input_dir_bg_replaced/
python scripts/replace_ppt_bg.py ./ppt_dir/ background.jpg

# Custom output dir
python scripts/replace_ppt_bg.py ./ppt_dir/ background.jpg ./output_dir/
`

## How It Works / 工作原理

For every slide:
1. Insert a new full-slide picture at bottom z-order
2. Add <p:bg> with explicit background image reference
3. Add <a:noFill/> to existing full-slide shapes to hide them
4. Remove fills from group shapes
5. All original content preserved

每个幻灯片：
1. 在底层插入新的全幅图片
2. 添加 <p:bg> 显式背景引用
3. 对原有全幅形状添加 <a:noFill/> 隐藏填充
4. 移除组形状的填充
5. 所有原有内容完好保留

## Project Structure / 项目结构

`
ppt-bulk-bg-replacer/
├── SKILL.md                    # Codex skill definition
├── README.md                   # This file
├── agents/
│   └── openai.yaml             # Agent configuration
└── scripts/
    └── replace_ppt_bg.py       # Main replacement script
`

## License

[MIT License](LICENSE)
