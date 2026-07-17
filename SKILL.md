---
name: ppt-bg-replacer
description: "批量替换PPT文件每一页的背景图。Use when: (1) 用户需要将PPT所有页面的背景统一替换为一张图片, (2) 需要批量处理多个PPT文件的背景, (3) 需要在保留所有文字、图表、形状等原有内容的前提下更换幻灯片背景。Triggers: '替换PPT背景', 'PPT换背景', 'change PPT background', 'replace slide background', '批量替换PPT'"
---

# PPT Background Replacer (PPT背景替换工具)

## Overview

Replace the background image on every slide of a PowerPoint presentation while preserving all existing content. Directly manipulates the PPTX ZIP structure for maximum control and compatibility.

**Key improvements over v1:**
- **INSERT-only approach** — No more corrupted content images from binary SWAP
- **Correct file extensions** — No more .png containing JPEG data
- **Batch directory processing** — Process entire directory trees preserving folder structure
- **Reliable Unicode support** — Handles Chinese/Unicode file paths correctly

## Quick Start

`ash
pip install lxml
`

### Single file
`ash
python scripts/replace_ppt_bg.py input.pptx background.jpg
# creates input_bg_replaced.pptx
`

### Custom output path
`ash
python scripts/replace_ppt_bg.py input.pptx background.jpg output.pptx
`

### Batch process entire directory
`ash
python scripts/replace_ppt_bg.py ./ppt_dir/ background.jpg
# creates ppt_dir_bg_replaced/ with same subfolder structure

python scripts/replace_ppt_bg.py ./ppt_dir/ background.jpg ./output_dir/
`

## How It Works

For every slide, the script:
1. **Inserts** a new full-slide picture at the bottom z-order
2. **Adds** <p:bg> with explicit background reference
3. **Adds** <a:noFill/> to existing full-slide elements to hide their fills
4. **Removes** fills from group shapes
5. **Preserves** all content images, text, charts, animations, and media

## Requirements

- Python 3.7+
- lxml (pip install lxml)

## Parameters

| Argument | Required | Description |
|---|---|---|
| input.pptx / input_dir | Yes | Source PowerPoint file or directory |
| background_image | Yes | Background image (jpg, png, etc.) |
| output.pptx / output_dir | No | Output path (defaults: input_bg_replaced.pptx / input_dir_bg_replaced/) |
