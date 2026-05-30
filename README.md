# 🎬 视频 → 分镜头脚本 PPT 生成器

**Video-to-Storyboard PPT Generator**

输入一个视频链接（B站/YouTube/抖音/小红书等），自动下载视频、场景检测、提取关键帧，生成导演级 **分镜头脚本 PPT**。

## ✨ 效果预览

| 格式 | 说明 |
|------|------|
| 🖼️ **尺寸** | 16:9 宽屏 |
| ⚫ **配色** | 黑白灰调，干净专业 |
| 📊 **表格** | 8列 — 镜号 \| 画面截图 \| 景别 \| 时长 \| 内容(三行) \| 文案 \| 运镜 \| 道具 |
| 🔠 **文字** | 居中放大，清晰易读 |
| 🖼️ **截图** | 保持原始宽高比，不变形 |

## 🚀 快速开始

### 安装依赖

```bash
# macOS (Homebrew)
brew install yt-dlp ffmpeg

# Python 包
pip install python-pptx Pillow
```

### 使用

```bash
python3 storyboard_generator.py <视频链接>
```

**示例：**

```bash
python3 storyboard_generator.py https://www.bilibili.com/video/BV1xx411c7mD
python3 storyboard_generator.py https://www.xiaohongshu.com/explore/xxx
python3 storyboard_generator.py https://youtu.be/xxxxxx
```

PPT 会自动生成到桌面：`~/Desktop/<视频标题>-分镜头脚本.pptx`

### 补充字段

生成后打开 PPT，需要手动补充以下字段（AI 自动识别有局限）：

| 字段 | 说明 |
|------|------|
| **景别** | 远景/全景/中景/近景/特写 |
| **运镜** | 固定/推/拉/摇/移/跟/环绕 |
| **场景/人物/动作** | 画面描述三行 |
| **道具** | 画面中的关键道具 |

## 📁 输出示例

```
~/Desktop/
├── AUM所向自然-分镜头脚本.pptx
├── 城市漫步-分镜头脚本.pptx
└── ...
```

## 🛠️ 技术栈

| 工具 | 用途 |
|------|------|
| [yt-dlp](https://github.com/yt-dlp/yt-dlp) | 视频下载 + 元数据提取 |
| [ffmpeg](https://ffmpeg.org/) | 场景检测 + 关键帧提取 |
| [python-pptx](https://python-pptx.readthedocs.io/) | PPT 生成 |
| [Pillow](https://python-pillow.org/) | 图片尺寸读取（保比例缩放） |

## 📄 许可证

MIT
