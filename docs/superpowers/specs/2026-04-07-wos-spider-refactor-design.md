# WOS Spider 重构设计文档

**日期**: 2026-04-07
**作者**: Chream
**状态**: 已确认

## 概述

将 WOS Spider 从单脚本爬虫重构为标准化的 Web 应用工具，包含 Flask 后端、浏览器可视化 UI、Playwright 爬虫引擎、ECharts + matplotlib 双图表系统。

## 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| 后端框架 | Flask | 轻量、Python 生态成熟 |
| 爬虫引擎 | Playwright | 比 Selenium 更快、API 更简洁、自动管理浏览器 |
| 前端 | 原生 HTML/CSS/JS | 不引入重框架，保持轻量 |
| 实时通信 | SSE (Server-Sent Events) | 单向推送日志/进度，比 WebSocket 简单 |
| 交互图表 | ECharts | 丰富的图表类型、交互性强 |
| 学术图表 | matplotlib | 学术出版标准、支持导出高清图片 |
| 数据存储 | CSV | 与原项目一致、轻量 |
| 配置 | YAML | 人类可读、支持嵌套结构 |

## 项目结构

```
WOS_Spider/
├── app.py                    # Flask 应用入口（启动服务器）
├── requirements.txt          # Python 依赖
├── config.yaml               # 全局配置（URL、超时、重试等）
├── wos_spider/
│   ├── __init__.py
│   ├── spider.py             # Playwright 爬虫核心引擎
│   ├── parser.py             # HTML 解析器（提取论文字段）
│   ├── merger.py             # CSV/Excel 数据合并
│   └── visualizer.py         # matplotlib 学术图表生成
├── web/
│   ├── __init__.py
│   ├── api.py                # Flask REST API + SSE 路由
│   ├── static/
│   │   ├── css/style.css     # 简约白主题样式
│   │   └── js/
│   │       ├── app.js        # 主逻辑、Tab 切换、SSE 连接
│   │       └── charts.js     # ECharts 图表封装
│   └── templates/
│       └── index.html        # 单页面应用
├── output/                   # 默认输出目录
│   ├── csv/                  # 爬取结果 CSV
│   └── charts/               # 生成的图表
└── tests/
    ├── test_parser.py
    ├── test_merger.py
    └── test_visualizer.py
```

## API 路由

### 配置
- `GET /api/config` — 获取当前配置
- `POST /api/config` — 更新配置

### 爬虫控制
- `POST /api/spider/start` — 启动爬虫（返回任务 ID）
- `POST /api/spider/pause` — 暂停爬取
- `POST /api/spider/resume` — 继续爬取
- `POST /api/spider/stop` — 停止爬取
- `GET /api/spider/status` — 获取当前状态（进度、速度、成功/失败数）
- `GET /api/spider/logs` — SSE 实时日志流

### 数据管理
- `GET /api/data/files` — 列出所有 CSV 文件
- `GET /api/data/preview?file=<name>` — 预览某文件前 100 行
- `POST /api/data/merge` — 合并多个 CSV（按国家/年份分组）
- `GET /api/data/download?file=<name>` — 下载文件

### 图表
- `POST /api/charts/echarts` — 生成 ECharts 配置 JSON
- `POST /api/charts/matplotlib` — 生成 matplotlib 图表并返回图片

## UI 设计

### 视觉风格：简约白

明亮浅色背景，干净的留白，蓝色强调色。参考现代 SaaS 工具的简洁风格。

### 页面布局

单页面应用，顶部导航 Tab 切换四个功能区：

**Tab 1: 爬虫控制**
- 配置表单：WOS URL、最大篇数、等待超时、输出文件名
- 控制按钮：启动 / 暂停 / 停止
- 进度条：已完成数 / 总数 / 百分比 / 速度
- 实时日志面板：彩色分级（info/success/warning/error）

**Tab 2: 数据管理**
- CSV 文件列表
- 数据预览表格
- 合并功能（按国家/年份分组求和）
- 导出为 Excel

**Tab 3: 可视化**
- ECharts 交互式图表（按国家分布、年份趋势、期刊排名等）
- matplotlib 学术图表生成（可选图表类型、配色方案）
- 图表导出（PNG/SVG）

**Tab 4: 设置**
- 浏览器配置（有头/无头模式、代理）
- 爬取策略（超时时间、重试次数、暂停间隔）
- 输出路径配置

## 爬虫核心流程

```
1. 用户在 Web UI 填写配置并点击"启动"
2. Flask 后台启动 Playwright 浏览器线程
3. 打开 Chromium，用户手动登录 WOS 并导航到搜索结果页
4. 用户在 Web UI 点击"开始爬取"
5. 爬虫逐篇点击论文详情页
6. parser.py 解析 HTML，提取字段
7. 每篇完成后: SSE 推送日志 + 更新进度 + 追加写入 CSV
8. 错误处理: 自动重试（最多 3 次），失败则标记跳过
9. 支持: 暂停/继续/停止/断点续传（读取已有 CSV 继续）
```

### 提取字段

| 字段 | 提取策略 |
|------|----------|
| 标题 | CSS 选择器 → XPath 兜底 |
| 引用次数 | CSS class 匹配 |
| 国家/地区 | 地址字段末尾提取 |
| 期刊名 | CSS class 匹配 |
| 发表年份 | XPath 定位 |
| 当前页码 | aria-label 解析 |

所有字段都有多套选择器策略和容错处理。

## 解析器设计

`parser.py` 与爬虫引擎解耦。解析器接收 HTML 字符串，返回结构化字典。每类字段定义多套提取规则（CSS/XPath），按优先级尝试，全部失败则返回默认值。WOS 改版时只需修改解析规则，不影响爬虫逻辑。

## 数据合并

`merger.py` 修复原项目 bug（只合并了最后一个文件），正确实现：
1. 读取目录下所有 CSV/XLSX 文件
2. 按指定列（国家/年份）分组求和
3. 降序排列
4. 输出合并结果

## 配置管理

`config.yaml` 管理所有可配置项：

```yaml
spider:
  max_papers: 10000
  wait_timeout: 5
  pause_interval: 0.1
  max_retries: 3
  headless: false

output:
  dir: output/csv
  default_filename: result.csv

server:
  host: 127.0.0.1
  port: 5000
```

## 跨平台

- 不使用 `winsound`、`keyboard` 等 Windows 专属模块
- 音频提示改为 Web UI 内通知
- 键盘控制改为 Web UI 按钮
- Playwright 自动管理浏览器驱动，无需手动下载 WebDriver

## 依赖

```
flask
playwright
beautifulsoup4
pandas
openpyxl
pyyaml
matplotlib
lxml
```
