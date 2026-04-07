# WOS Spider - Web of Science 论文信息爬虫

从 Web of Science (WOS) 网站抓取论文详细信息的 Python 爬虫工具，配备 Web 可视化界面。

## 项目简介

本项目提供完整的 WOS 论文爬取、数据合并与可视化分析工作流：

- **爬虫引擎**: 基于 Playwright 的自动化爬虫，支持暂停/继续/停止和断点续传
- **Web UI**: Flask 驱动的浏览器界面，实时查看爬取日志和进度
- **数据管理**: CSV 文件预览、多文件合并、按国家/年份分组统计
- **可视化**: ECharts 交互式图表 + matplotlib 学术图表导出

## 快速开始

### 环境要求

- Python 3.8+
- Chrome / Chromium 浏览器（Playwright 自动管理）

### 安装

```bash
pip install -r requirements.txt
playwright install chromium
```

### 启动

```bash
python app.py
```

浏览器打开 http://127.0.0.1:5000

### 使用流程

1. 在 **爬虫控制** 页面填入 WOS 搜索结果 URL，点击「启动浏览器」
2. 在弹出的 Chromium 窗口中手动登录 WOS 并导航到目标页面
3. 回到 Web UI 点击「开始爬取」，爬虫自动逐篇抓取
4. 在 **数据管理** 页面预览、合并 CSV 数据
5. 在 **可视化** 页面生成统计图表

## 项目结构

```
WOS_Spider/
├── app.py                    # Flask 应用入口
├── config.yaml               # 全局配置
├── requirements.txt          # Python 依赖
├── wos_spider/               # 核心模块
│   ├── parser.py             # HTML 解析器（多策略提取）
│   ├── spider.py             # Playwright 爬虫引擎
│   ├── merger.py             # CSV/XLSX 数据合并
│   └── visualizer.py         # matplotlib + ECharts 图表
├── web/                      # Web 层
│   ├── api.py                # REST API + SSE 实时日志
│   ├── templates/index.html  # 单页面应用
│   └── static/               # CSS + JS
├── output/                   # 输出目录
│   ├── csv/                  # 爬取结果
│   └── charts/               # 生成的图表
└── tests/                    # 测试
    ├── test_parser.py
    ├── test_merger.py
    └── test_visualizer.py
```

## 功能特性

- 多策略 HTML 解析，WOS 页面改版时只需调整选择器
- SSE 实时日志推送，浏览器中即时查看爬取状态
- 断点续传，读取已有 CSV 自动从上次中断处继续
- 自动重试（可配置次数），失败条目标记跳过
- 跨平台，不依赖 Windows 专属模块
- YAML 配置文件，无需改代码即可调整参数

## 配置

编辑 `config.yaml`：

```yaml
spider:
  max_papers: 10000
  wait_timeout: 5
  pause_interval: 0.1
  max_retries: 3
  headless: false

output:
  dir: output/csv
  charts_dir: output/charts
  default_filename: result.csv

server:
  host: 127.0.0.1
  port: 5000
```

## 运行测试

```bash
python -m pytest tests/ -v
```

## 作者

- **Chream** (yaom7917@gmail.com)

## 声明

此工具仅用于教育目的，使用时应遵守 Web of Science 的服务条款和使用政策。
