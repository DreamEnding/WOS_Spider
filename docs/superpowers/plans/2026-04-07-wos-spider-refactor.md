# WOS Spider 重构实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 WOS Spider 从单脚本重构为 Flask + Playwright + ECharts 的标准化 Web 应用。

**Architecture:** Flask 后端提供 REST API + SSE 实时推送，Playwright 异步引擎处理爬虫，前端为原生 HTML/CSS/JS 单页面应用。模块分为 parser（解析）、spider（爬虫）、merger（合并）、visualizer（图表），通过 API 层与前端交互。

**Tech Stack:** Python 3.8+, Flask, Playwright, BeautifulSoup4, pandas, pyyaml, matplotlib, ECharts (CDN)

---

## File Structure

| File | Responsibility |
|------|---------------|
| `app.py` | Flask 应用入口，启动服务器 |
| `config.yaml` | 全局配置 |
| `requirements.txt` | Python 依赖 |
| `wos_spider/__init__.py` | 包初始化 |
| `wos_spider/parser.py` | HTML 解析，提取论文字段 |
| `wos_spider/spider.py` | Playwright 爬虫引擎 |
| `wos_spider/merger.py` | CSV/XLSX 数据合并 |
| `wos_spider/visualizer.py` | matplotlib 学术图表 |
| `web/__init__.py` | 包初始化 |
| `web/api.py` | Flask 路由（REST + SSE） |
| `web/static/css/style.css` | 简约白主题样式 |
| `web/static/js/app.js` | 前端主逻辑 |
| `web/static/js/charts.js` | ECharts 封装 |
| `web/templates/index.html` | 单页面 HTML |
| `output/csv/.gitkeep` | 输出目录占位 |
| `output/charts/.gitkeep` | 图表目录占位 |
| `tests/test_parser.py` | parser 测试 |
| `tests/test_merger.py` | merger 测试 |
| `tests/test_visualizer.py` | visualizer 测试 |

---

### Task 1: 项目脚手架

**Files:**
- Create: `requirements.txt`
- Create: `config.yaml`
- Create: `wos_spider/__init__.py`
- Create: `web/__init__.py`
- Create: `output/csv/.gitkeep`
- Create: `output/charts/.gitkeep`

- [ ] **Step 1: 创建目录结构**

```bash
mkdir -p wos_spider web/static/css web/static/js web/templates output/csv output/charts tests
```

- [ ] **Step 2: 创建 requirements.txt**

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

- [ ] **Step 3: 创建 config.yaml**

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

- [ ] **Step 4: 创建包初始化文件和 .gitkeep**

`wos_spider/__init__.py`:
```python
from .parser import parse_paper_html
from .spider import SpiderEngine
from .merger import merge_files
from .visualizer import generate_matplotlib_chart
```

`web/__init__.py`:
```python
```

`output/csv/.gitkeep` and `output/charts/.gitkeep`: 空文件

- [ ] **Step 5: 安装依赖**

```bash
pip install -r requirements.txt
playwright install chromium
```

- [ ] **Step 6: Commit**

```bash
git add requirements.txt config.yaml wos_spider/__init__.py web/__init__.py output/ tests/
git commit -m "feat: scaffold project structure"
```

---

### Task 2: Parser 模块

**Files:**
- Create: `wos_spider/parser.py`
- Create: `tests/test_parser.py`

- [ ] **Step 1: 写 parser 的失败测试**

`tests/test_parser.py`:
```python
import pytest
from wos_spider.parser import parse_paper_html


def sample_html(title="Test Paper Title", citation="42",
                address="Tsinghua Univ, Beijing, Peoples R China",
                journal="Nature", year="2024", page_label="Page 5 of 100"):
    """构建一个模拟 WOS 详情页的 HTML 片段"""
    return f"""
    <html><body>
    <div class="cdx-two-column-grid-container">
        <div class="cdx-grid-label">By</div>
        <div class="cdx-grid-data">Author One</div>
    </div>
    <h2 class="title text--large cdx-title">{title}</h2>
    <a class="mat-tooltip-trigger medium-link-number link ng-star-inserted">{citation}</a>
    <span class="ng-star-inserted" id="FRAOrgTa-RepAddressFull-0">{address}</span>
    <a class="mat-focus-indicator mat-tooltip-trigger font-size-14 summary-source-title-link remove-space no-left-padding mat-button mat-button-base mat-primary font-size-16 ng-star-inserted">{journal}</a>
    <span class="frozen-page-links-label ng-star-inserted">{year}</span>
    <input class="wos-input-underline page-box" aria-label="{page_label}">
    </body></html>
    """


class TestParsePaperHtml:
    def test_extracts_title(self):
        html = sample_html(title="Deep Learning for Control")
        index, data = parse_paper_html(html)
        assert data["title"] == "Deep Learning for Control"

    def test_extracts_citation(self):
        html = sample_html(citation="99")
        index, data = parse_paper_html(html)
        assert data["citation"] == "99"

    def test_citation_default_zero(self):
        html = sample_html()
        # 移除引用元素
        html = html.replace('mat-tooltip-trigger medium-link-number link ng-star-inserted', 'removed')
        index, data = parse_paper_html(html)
        assert data["citation"] == "0"

    def test_extracts_country_from_address(self):
        html = sample_html(address="MIT, Cambridge, USA")
        index, data = parse_paper_html(html)
        assert data["country"] == "USA"

    def test_extracts_journal(self):
        html = sample_html(journal="Science")
        index, data = parse_paper_html(html)
        assert data["journal"] == "Science"

    def test_extracts_page_index(self):
        html = sample_html(page_label="Page 42 of 500")
        index, data = parse_paper_html(html)
        assert index == 42

    def test_returns_dict_on_partial_html(self):
        """即使部分字段缺失也不应抛异常"""
        html = "<html><body><h2 class='title text--large cdx-title'>Only Title</h2></body></html>"
        index, data = parse_paper_html(html)
        assert data["title"] == "Only Title"
        assert "citation" not in data or data.get("citation") == "0"

    def test_empty_html_returns_defaults(self):
        html = "<html><body></body></html>"
        index, data = parse_paper_html(html)
        assert isinstance(data, dict)
```

- [ ] **Step 2: 运行测试确认失败**

```bash
python -m pytest tests/test_parser.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'wos_spider'` 或 `ImportError`

- [ ] **Step 3: 实现 parser.py**

`wos_spider/parser.py`:
```python
"""WOS 论文详情页 HTML 解析器。

接收 HTML 字符串，返回 (页码索引, 字段字典)。
所有字段使用多套提取规则，按优先级尝试，失败返回默认值。
"""
from bs4 import BeautifulSoup


def _try_extract(soup, strategies, default=""):
    """按优先级尝试多套提取策略，返回第一个成功的结果。"""
    for strategy in strategies:
        try:
            result = strategy(soup)
            if result:
                return result.strip()
        except Exception:
            continue
    return default


def parse_paper_html(html):
    """解析 WOS 论文详情页 HTML。

    Args:
        html: 页面 HTML 字符串

    Returns:
        tuple: (页码索引 int, 字段字典 dict)
    """
    soup = BeautifulSoup(html, "lxml")
    data = {}

    # --- 通用容器字段 ---
    _extract_container_fields(soup, data)

    # --- 标题 ---
    data["title"] = _try_extract(soup, [
        lambda s: s.find(class_="title text--large cdx-title").text,
    ], default="未获取到标题")

    # --- 引用次数 ---
    data["citation"] = _try_extract(soup, [
        lambda s: s.find(class_="mat-tooltip-trigger medium-link-number link ng-star-inserted").text,
    ], default="0")

    # --- 国家/地区 ---
    data["country"] = _try_extract(soup, [
        lambda s: s.find("span", class_="ng-star-inserted",
                          id="FRAOrgTa-RepAddressFull-0").text.split(",")[-1],
        lambda s: s.find("span", class_="value padding-right-5--reversible").text.split(",")[-1],
    ], default="未获取到国家")

    # --- 期刊 ---
    data["journal"] = _try_extract(soup, [
        lambda s: s.find(class_="mat-focus-indicator mat-tooltip-trigger font-size-14 summary-source-title-link remove-space no-left-padding mat-button mat-button-base mat-primary font-size-16 ng-star-inserted").text,
    ], default="未获取到期刊")

    # --- 页码索引 ---
    index = _extract_page_index(soup)

    return index, data


def _extract_container_fields(soup, data):
    """提取 cdx-two-column-grid-container 中的通用字段。"""
    try:
        containers = soup.find_all("div", class_="cdx-two-column-grid-container")
        for container in containers:
            labels = container.find_all(class_="cdx-grid-label")
            datas = container.find_all(class_="cdx-grid-data")
            if labels and datas:
                label = labels[0].text.strip()
                text = "\n".join(d.text.strip() for d in datas)
                data[label] = text
    except Exception:
        pass


def _extract_page_index(soup):
    """从 aria-label 提取当前页码。"""
    try:
        input_box = soup.find(class_="wos-input-underline page-box")
        if input_box and input_box.get("aria-label"):
            return int(input_box["aria-label"].split()[-1].replace(",", ""))
    except Exception:
        pass
    return 0
```

- [ ] **Step 4: 运行测试确认通过**

```bash
python -m pytest tests/test_parser.py -v
```
Expected: ALL PASS

- [ ] **Step 5: Commit**

```bash
git add wos_spider/parser.py tests/test_parser.py
git commit -m "feat: add HTML parser with multi-strategy extraction"
```

---

### Task 3: Merger 模块

**Files:**
- Create: `wos_spider/merger.py`
- Create: `tests/test_merger.py`

- [ ] **Step 1: 写 merger 的失败测试**

`tests/test_merger.py`:
```python
import os
import pytest
import pandas as pd
from wos_spider.merger import merge_files


@pytest.fixture
def csv_dir(tmp_path):
    """创建包含多个 CSV 的临时目录。"""
    df1 = pd.DataFrame({"country": ["China", "USA", "Japan"], "count": [100, 80, 60]})
    df2 = pd.DataFrame({"country": ["China", "UK", "USA"], "count": [50, 40, 30]})
    df1.to_csv(tmp_path / "batch1.csv", index=False)
    df2.to_csv(tmp_path / "batch2.csv", index=False)
    return tmp_path


class TestMergeFiles:
    def test_merges_multiple_csv(self, csv_dir):
        result = merge_files(str(csv_dir), group_col="country", value_col="count")
        assert isinstance(result, pd.DataFrame)
        assert len(result) == 4  # China, USA, Japan, UK

    def test_sums_values(self, csv_dir):
        result = merge_files(str(csv_dir), group_col="country", value_col="count")
        row = result[result["country"] == "China"].iloc[0]
        assert row["count"] == 150  # 100 + 50

    def test_sorted_descending(self, csv_dir):
        result = merge_files(str(csv_dir), group_col="country", value_col="count")
        counts = result["count"].tolist()
        assert counts == sorted(counts, reverse=True)

    def test_output_to_file(self, csv_dir, tmp_path):
        out = str(tmp_path / "merged.csv")
        merge_files(str(csv_dir), group_col="country", value_col="count", output_path=out)
        assert os.path.exists(out)
        df = pd.read_csv(out)
        assert len(df) == 4

    def test_empty_directory(self, tmp_path):
        result = merge_files(str(tmp_path), group_col="country", value_col="count")
        assert result.empty

    def test_single_file(self, tmp_path):
        df = pd.DataFrame({"country": ["China"], "count": [10]})
        df.to_csv(tmp_path / "only.csv", index=False)
        result = merge_files(str(tmp_path), group_col="country", value_col="count")
        assert len(result) == 1
        assert result.iloc[0]["count"] == 10
```

- [ ] **Step 2: 运行测试确认失败**

```bash
python -m pytest tests/test_merger.py -v
```
Expected: FAIL — `ImportError`

- [ ] **Step 3: 实现 merger.py**

`wos_spider/merger.py`:
```python
"""CSV/XLSX 数据合并工具。

读取目录下所有数据文件，按指定列分组求和并排序。
修复原项目 bug：原 merge_xlsx.py 只合并了最后一个文件。
"""
import os
import glob
import pandas as pd


def merge_files(directory, group_col, value_col, output_path=None):
    """合并目录下所有 CSV/XLSX 文件。

    Args:
        directory: 包含数据文件的目录路径
        group_col: 分组列名（如 "country"）
        value_col: 求和列名（如 "count"）
        output_path: 可选，输出文件路径。不传则只返回 DataFrame

    Returns:
        pd.DataFrame: 合并后的数据，按 value_col 降序排列
    """
    csv_files = glob.glob(os.path.join(directory, "*.csv"))
    xlsx_files = glob.glob(os.path.join(directory, "*.xlsx"))
    all_files = csv_files + xlsx_files

    if not all_files:
        return pd.DataFrame(columns=[group_col, value_col])

    frames = []
    for f in all_files:
        try:
            if f.endswith(".csv"):
                df = pd.read_csv(f)
            else:
                df = pd.read_excel(f)
            frames.append(df)
        except Exception as e:
            print(f"Warning: 跳过文件 {f}: {e}")
            continue

    if not frames:
        return pd.DataFrame(columns=[group_col, value_col])

    merged = pd.concat(frames, ignore_index=True)

    # 统一列名：取第一列和第二列
    if group_col not in merged.columns or value_col not in merged.columns:
        cols = merged.columns.tolist()
        if len(cols) >= 2:
            merged = merged.rename(columns={cols[0]: group_col, cols[1]: value_col})

    result = merged.groupby(group_col)[value_col].sum().reset_index()
    result = result.sort_values(by=value_col, ascending=False).reset_index(drop=True)

    if output_path:
        if output_path.endswith(".xlsx"):
            result.to_excel(output_path, index=False)
        else:
            result.to_csv(output_path, index=False)

    return result
```

- [ ] **Step 4: 运行测试确认通过**

```bash
python -m pytest tests/test_merger.py -v
```
Expected: ALL PASS

- [ ] **Step 5: Commit**

```bash
git add wos_spider/merger.py tests/test_merger.py
git commit -m "feat: add data merger with multi-file support"
```

---

### Task 4: Visualizer 模块

**Files:**
- Create: `wos_spider/visualizer.py`
- Create: `tests/test_visualizer.py`

- [ ] **Step 1: 写 visualizer 的失败测试**

`tests/test_visualizer.py`:
```python
import os
import pytest
import pandas as pd
import matplotlib
matplotlib.use("Agg")
from wos_spider.visualizer import generate_matplotlib_chart, generate_echarts_config


@pytest.fixture
def sample_df():
    return pd.DataFrame({
        "country": ["China", "USA", "UK", "Japan", "Germany"],
        "count": [150, 120, 80, 60, 40],
    })


class TestMatplotlibChart:
    def test_generates_bar_chart(self, sample_df, tmp_path):
        out = str(tmp_path / "bar.png")
        generate_matplotlib_chart(
            sample_df, chart_type="bar",
            x_col="country", y_col="count",
            title="Papers by Country",
            output_path=out,
        )
        assert os.path.exists(out)
        assert os.path.getsize(out) > 0

    def test_generates_pie_chart(self, sample_df, tmp_path):
        out = str(tmp_path / "pie.png")
        generate_matplotlib_chart(
            sample_df, chart_type="pie",
            labels_col="country", values_col="count",
            title="Distribution",
            output_path=out,
        )
        assert os.path.exists(out)

    def test_invalid_chart_type_raises(self, sample_df, tmp_path):
        with pytest.raises(ValueError):
            generate_matplotlib_chart(
                sample_df, chart_type="invalid",
                x_col="country", y_col="count",
                output_path=str(tmp_path / "bad.png"),
            )


class TestEChartsConfig:
    def test_generates_bar_config(self, sample_df):
        config = generate_echarts_config(
            sample_df, chart_type="bar",
            x_col="country", y_col="count",
            title="Papers by Country",
        )
        assert config["title"]["text"] == "Papers by Country"
        assert config["series"][0]["type"] == "bar"
        assert len(config["xAxis"]["data"]) == 5

    def test_generates_pie_config(self, sample_df):
        config = generate_echarts_config(
            sample_df, chart_type="pie",
            names_col="country", values_col="count",
            title="Distribution",
        )
        assert config["series"][0]["type"] == "pie"
        assert len(config["series"][0]["data"]) == 5

    def test_generates_line_config(self, sample_df):
        config = generate_echarts_config(
            sample_df, chart_type="line",
            x_col="country", y_col="count",
            title="Trend",
        )
        assert config["series"][0]["type"] == "line"
```

- [ ] **Step 2: 运行测试确认失败**

```bash
python -m pytest tests/test_visualizer.py -v
```
Expected: FAIL — `ImportError`

- [ ] **Step 3: 实现 visualizer.py**

`wos_spider/visualizer.py`:
```python
"""图表生成模块。

提供 matplotlib 学术图表和 ECharts 配置生成。
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd


ACADEMIC_STYLE = {
    "font.size": 11,
    "axes.titlesize": 13,
    "axes.labelsize": 11,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "figure.dpi": 150,
    "savefig.bbox": "tight",
}


def generate_matplotlib_chart(df, chart_type, output_path, title="",
                               x_col=None, y_col=None,
                               labels_col=None, values_col=None,
                               style=None):
    """生成 matplotlib 图表并保存为图片。

    Args:
        df: 数据 DataFrame
        chart_type: "bar", "pie", "line"
        output_path: 输出图片路径
        title: 图表标题
        x_col/labels_col: X 轴/标签列名
        y_col/values_col: Y 轴/数值列名
        style: 可选 matplotlib rcParams 覆盖
    """
    plt.rcParams.update(ACADEMIC_STYLE)
    if style:
        plt.rcParams.update(style)

    fig, ax = plt.subplots()

    if chart_type == "bar":
        x = x_col or labels_col
        y = y_col or values_col
        ax.bar(df[x], df[y], color="#3b82f6", edgecolor="white")
        ax.set_xlabel(x.capitalize())
        ax.set_ylabel(y.capitalize())
        ax.set_title(title)
        plt.xticks(rotation=45, ha="right")

    elif chart_type == "pie":
        labels_col = labels_col or x_col
        values_col = values_col or y_col
        ax.pie(df[values_col], labels=df[labels_col], autopct="%1.1f%%",
               colors=plt.cm.Set3.colors[:len(df)])
        ax.set_title(title)

    elif chart_type == "line":
        x = x_col or labels_col
        y = y_col or values_col
        ax.plot(df[x], df[y], marker="o", color="#3b82f6", linewidth=2)
        ax.set_xlabel(x.capitalize())
        ax.set_ylabel(y.capitalize())
        ax.set_title(title)
        plt.xticks(rotation=45, ha="right")

    else:
        raise ValueError(f"Unsupported chart type: {chart_type}")

    fig.savefig(output_path)
    plt.close(fig)
    return output_path


def generate_echarts_config(df, chart_type, title="", **kwargs):
    """生成 ECharts 配置字典。

    Args:
        df: 数据 DataFrame
        chart_type: "bar", "pie", "line"
        title: 图表标题

    Returns:
        dict: ECharts option 配置
    """
    if chart_type == "pie":
        names_col = kwargs.get("names_col") or kwargs.get("x_col")
        values_col = kwargs.get("values_col") or kwargs.get("y_col")
        data = [{"name": row[names_col], "value": int(row[values_col])}
                for _, row in df.iterrows()]
        return {
            "title": {"text": title, "left": "center"},
            "tooltip": {"trigger": "item"},
            "legend": {"orient": "vertical", "left": "left"},
            "series": [{
                "type": "pie",
                "radius": "60%",
                "data": data,
                "emphasis": {
                    "itemStyle": {"shadowBlur": 10, "shadowOffsetX": 0, "shadowColor": "rgba(0,0,0,0.5)"}
                },
            }],
        }

    # bar / line
    x_col = kwargs.get("x_col")
    y_col = kwargs.get("y_col")
    return {
        "title": {"text": title, "left": "center"},
        "tooltip": {"trigger": "axis"},
        "xAxis": {"type": "category", "data": df[x_col].tolist(), "axisLabel": {"rotate": 30}},
        "yAxis": {"type": "value"},
        "series": [{
            "type": chart_type,
            "data": df[y_col].tolist(),
            "itemStyle": {"color": "#3b82f6"},
        }],
    }
```

- [ ] **Step 4: 运行测试确认通过**

```bash
python -m pytest tests/test_visualizer.py -v
```
Expected: ALL PASS

- [ ] **Step 5: Commit**

```bash
git add wos_spider/visualizer.py tests/test_visualizer.py
git commit -m "feat: add chart generation (matplotlib + ECharts config)"
```

---

### Task 5: Spider Engine (Playwright)

**Files:**
- Create: `wos_spider/spider.py`

- [ ] **Step 1: 实现 spider.py**

`wos_spider/spider.py`:
```python
"""Playwright 爬虫引擎。

管理浏览器生命周期、爬取循环、暂停/继续/停止、断点续传。
"""
import os
import time
import logging
import pandas as pd
from threading import Event, Thread
from playwright.sync_api import sync_playwright
from .parser import parse_paper_html

logger = logging.getLogger(__name__)


class SpiderEngine:
    """WOS 爬虫引擎，支持暂停/继续/停止和断点续传。"""

    def __init__(self, config):
        self.config = config
        self._stop_event = Event()
        self._pause_event = Event()
        self._pause_event.set()  # 初始不暂停
        self._status = "idle"
        self._progress = {"done": 0, "failed": 0, "total": 0, "current_index": 0}
        self._log_queue = []  # 日志缓冲
        self._playwright = None
        self._browser = None
        self._page = None
        self._thread = None

    @property
    def status(self):
        return {
            "status": self._status,
            "progress": self._progress,
        }

    def get_logs(self, since=0):
        """获取指定索引之后的所有日志。"""
        return self._log_queue[since:]

    def start_browser(self):
        """启动 Playwright 浏览器，供用户手动登录。"""
        self._playwright = sync_playwright().start()
        launch_args = ["--disable-images", "--disable-plugins"]
        self._browser = self._playwright.chromium.launch(
            headless=self.config.get("headless", False),
            args=launch_args,
        )
        self._context = self._browser.new_context()
        self._page = self._context.new_page()
        self._status = "browser_ready"
        self._log("info", "浏览器已启动，请手动登录 WOS 并导航到搜索结果页")

    def navigate(self, url):
        """导航到指定 URL。"""
        if self._page:
            self._page.goto(url)
            self._log("info", f"已导航到: {url}")

    def start_crawl(self, url_root, max_papers, file_path):
        """在后台线程中启动爬取任务。"""
        self._stop_event.clear()
        self._pause_event.set()
        self._status = "running"
        self._progress = {"done": 0, "failed": 0, "total": max_papers, "current_index": 0}

        # 断点续传
        start_index = 0
        if os.path.exists(file_path):
            try:
                df = pd.read_csv(file_path, index_col=0)
                start_index = int(df.index[-1].split("_")[-1])
                self._progress["done"] = start_index
                self._log("info", f"断点续传：从第 {start_index + 1} 篇继续")
            except Exception:
                pass

        self._thread = Thread(
            target=self._crawl_loop,
            args=(url_root, max_papers, file_path, start_index),
            daemon=True,
        )
        self._thread.start()

    def pause(self):
        self._pause_event.clear()
        self._status = "paused"
        self._log("info", "爬虫已暂停")

    def resume(self):
        self._pause_event.set()
        self._status = "running"
        self._log("info", "爬虫已继续")

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()  # 解除暂停以便线程退出
        self._status = "stopped"
        self._log("info", "爬虫已停止")

    def close(self):
        """关闭浏览器和 Playwright。"""
        self.stop()
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()
        self._status = "idle"

    def _log(self, level, message):
        self._log_queue.append({
            "level": level,
            "message": message,
            "time": time.strftime("%H:%M:%S"),
        })

    def _crawl_loop(self, url_root, max_papers, file_path, start_index):
        """爬取主循环，在后台线程中运行。"""
        try:
            # 如果浏览器未打开，先打开
            if not self._page:
                self.start_browser()
                self.navigate(url_root)

            wait_timeout = self.config.get("wait_timeout", 5) * 1000
            pause_interval = self.config.get("pause_interval", 0.1)
            max_retries = self.config.get("max_retries", 3)

            index = start_index
            df = pd.DataFrame()
            if os.path.exists(file_path):
                try:
                    df = pd.read_csv(file_path, index_col=0)
                except Exception:
                    df = pd.DataFrame()

            while index < max_papers and not self._stop_event.is_set():
                self._pause_event.wait()

                if self._stop_event.is_set():
                    break

                self._log("info", f"正在处理第 {index + 1} 篇论文")
                self._progress["current_index"] = index

                retries = 0
                success = False
                while retries < max_retries and not success:
                    try:
                        self._page.wait_for_selector(
                            "app-full-record", timeout=wait_timeout
                        )
                        html = self._page.content()

                        # 点击下一篇
                        try:
                            next_btn = self._page.locator(
                                'button[aria-label="Next Page"]'
                            )
                            if next_btn.is_visible():
                                next_btn.click()
                        except Exception:
                            pass

                        page_index, data = parse_paper_html(html)
                        row_index = f"Row_{page_index}"

                        if row_index in df.index:
                            df.loc[row_index] = pd.Series(data)
                        else:
                            new_row = pd.DataFrame([data], index=[row_index])
                            df = pd.concat([df, new_row])

                        df.to_csv(file_path, index=True)
                        self._progress["done"] = index + 1
                        self._log("success", f"第 {index + 1} 篇完成: {data.get('title', '')[:50]}")
                        success = True
                        index += 1

                    except Exception as e:
                        retries += 1
                        self._log("warning", f"第 {index + 1} 篇出错 (重试 {retries}/{max_retries}): {e}")
                        if retries < max_retries:
                            time.sleep(pause_interval)
                            try:
                                self._page.go_back()
                                time.sleep(pause_interval)
                            except Exception:
                                pass

                if not success:
                    self._progress["failed"] += 1
                    self._log("error", f"第 {index + 1} 篇失败，跳过")
                    index += 1

            if not self._stop_event.is_set():
                self._status = "completed"
                self._log("success", f"爬取完成！共处理 {self._progress['done']} 篇")

        except Exception as e:
            self._status = "error"
            self._log("error", f"爬虫异常终止: {e}")

    # --- XPath 选择器（兜底方案） ---
    XPATH_TITLE = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[1]/div/div/div/h2'
    XPATH_YEAR = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[7]/div[2]/span'
```

- [ ] **Step 2: Commit**

```bash
git add wos_spider/spider.py
git commit -m "feat: add Playwright spider engine with pause/resume/stop"
```

---

### Task 6: Flask API 路由

**Files:**
- Create: `web/api.py`

- [ ] **Step 1: 实现 api.py**

`web/api.py`:
```python
"""Flask REST API + SSE 路由。"""
import os
import json
import time
import yaml
import pandas as pd
from flask import Blueprint, request, jsonify, Response, send_file, render_template
from wos_spider.spider import SpiderEngine
from wos_spider.merger import merge_files
from wos_spider.visualizer import generate_matplotlib_chart, generate_echarts_config

api_bp = Blueprint("api", __name__)

# 全局状态
_spider = None
_config = {}
_config_path = ""


def init_api(config_path="config.yaml"):
    """初始化 API 模块，加载配置。"""
    global _config, _config_path, _spider
    _config_path = config_path
    _load_config()
    spider_cfg = _config.get("spider", {})
    _spider = SpiderEngine(spider_cfg)


def _load_config():
    global _config
    try:
        with open(_config_path, "r", encoding="utf-8") as f:
            _config = yaml.safe_load(f) or {}
    except FileNotFoundError:
        _config = {}


def _save_config():
    with open(_config_path, "w", encoding="utf-8") as f:
        yaml.dump(_config, f, allow_unicode=True, default_flow_style=False)


# --- Pages ---

@api_bp.route("/")
def index():
    return render_template("index.html")


# --- Config ---

@api_bp.route("/api/config", methods=["GET"])
def get_config():
    safe_config = {k: v for k, v in _config.items() if k != "secrets"}
    return jsonify(safe_config)


@api_bp.route("/api/config", methods=["POST"])
def update_config():
    global _config
    updates = request.json
    _config.update(updates)
    _save_config()
    return jsonify({"ok": True})


# --- Spider Control ---

@api_bp.route("/api/spider/start", methods=["POST"])
def spider_start():
    params = request.json or {}
    url_root = params.get("url", "")
    max_papers = params.get("max_papers", _config.get("spider", {}).get("max_papers", 10000))
    file_path = params.get("file_path",
        os.path.join(_config.get("output", {}).get("dir", "output/csv"),
                     _config.get("output", {}).get("default_filename", "result.csv")))

    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    _spider.start_browser()
    if url_root:
        _spider.navigate(url_root)

    return jsonify({"ok": True, "status": _spider.status})


@api_bp.route("/api/spider/crawl", methods=["POST"])
def spider_crawl():
    params = request.json or {}
    url_root = params.get("url", "")
    max_papers = params.get("max_papers", _config.get("spider", {}).get("max_papers", 10000))
    file_path = params.get("file_path",
        os.path.join(_config.get("output", {}).get("dir", "output/csv"),
                     _config.get("output", {}).get("default_filename", "result.csv")))

    _spider.start_crawl(url_root, max_papers, file_path)
    return jsonify({"ok": True})


@api_bp.route("/api/spider/pause", methods=["POST"])
def spider_pause():
    _spider.pause()
    return jsonify({"ok": True, "status": _spider.status})


@api_bp.route("/api/spider/resume", methods=["POST"])
def spider_resume():
    _spider.resume()
    return jsonify({"ok": True, "status": _spider.status})


@api_bp.route("/api/spider/stop", methods=["POST"])
def spider_stop():
    _spider.stop()
    return jsonify({"ok": True, "status": _spider.status})


@api_bp.route("/api/spider/status", methods=["GET"])
def spider_status():
    return jsonify(_spider.status)


@api_bp.route("/api/spider/logs")
def spider_logs():
    """SSE 实时日志流。"""
    since = int(request.args.get("since", 0))

    def generate():
        last_idx = since
        while True:
            logs = _spider.get_logs(since=last_idx)
            if logs:
                for log in logs:
                    yield f"data: {json.dumps(log, ensure_ascii=False)}\n\n"
                last_idx += len(logs)
            else:
                # 无新日志时发送心跳
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"
            time.sleep(0.5)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# --- Data Management ---

@api_bp.route("/api/data/files", methods=["GET"])
def data_files():
    output_dir = _config.get("output", {}).get("dir", "output/csv")
    os.makedirs(output_dir, exist_ok=True)
    files = []
    for f in sorted(os.listdir(output_dir)):
        if f.endswith((".csv", ".xlsx")):
            path = os.path.join(output_dir, f)
            files.append({
                "name": f,
                "size": os.path.getsize(path),
                "modified": time.strftime("%Y-%m-%d %H:%M:%S",
                                          time.localtime(os.path.getmtime(path))),
            })
    return jsonify(files)


@api_bp.route("/api/data/preview", methods=["GET"])
def data_preview():
    filename = request.args.get("file", "")
    if not filename:
        return jsonify({"error": "file parameter required"}), 400
    output_dir = _config.get("output", {}).get("dir", "output/csv")
    path = os.path.join(output_dir, filename)
    if not os.path.exists(path):
        return jsonify({"error": "file not found"}), 404
    try:
        df = pd.read_csv(path, index_col=0, nrows=100)
        return jsonify({
            "columns": df.columns.tolist(),
            "rows": df.reset_index().values.tolist()[:100],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@api_bp.route("/api/data/merge", methods=["POST"])
def data_merge():
    params = request.json or {}
    group_col = params.get("group_col", "country")
    value_col = params.get("value_col", "count")
    output_filename = params.get("output_filename", "merged.csv")

    output_dir = _config.get("output", {}).get("dir", "output/csv")
    output_path = os.path.join(output_dir, output_filename)

    result = merge_files(output_dir, group_col, value_col, output_path)
    return jsonify({
        "ok": True,
        "rows": len(result),
        "data": result.to_dict(orient="records"),
    })


@api_bp.route("/api/data/download", methods=["GET"])
def data_download():
    filename = request.args.get("file", "")
    if not filename:
        return jsonify({"error": "file parameter required"}), 400
    output_dir = _config.get("output", {}).get("dir", "output/csv")
    path = os.path.join(output_dir, filename)
    if not os.path.exists(path):
        return jsonify({"error": "file not found"}), 404
    return send_file(path, as_attachment=True)


# --- Charts ---

@api_bp.route("/api/charts/echarts", methods=["POST"])
def charts_echarts():
    params = request.json or {}
    chart_type = params.get("chart_type", "bar")
    file = params.get("file", "")
    group_col = params.get("group_col", "country")
    value_col = params.get("value_col", "count")
    title = params.get("title", "")

    df = _load_data_for_chart(file, group_col, value_col)
    if df is None:
        return jsonify({"error": "no data"}), 400

    config = generate_echarts_config(df, chart_type, title=title,
                                      x_col=group_col, y_col=value_col,
                                      names_col=group_col, values_col=value_col)
    return jsonify(config)


@api_bp.route("/api/charts/matplotlib", methods=["POST"])
def charts_matplotlib():
    params = request.json or {}
    chart_type = params.get("chart_type", "bar")
    file = params.get("file", "")
    group_col = params.get("group_col", "country")
    value_col = params.get("value_col", "count")
    title = params.get("title", "")

    charts_dir = _config.get("output", {}).get("charts_dir", "output/charts")
    os.makedirs(charts_dir, exist_ok=True)
    output_path = os.path.join(charts_dir, f"{chart_type}_{int(time.time())}.png")

    df = _load_data_for_chart(file, group_col, value_col)
    if df is None:
        return jsonify({"error": "no data"}), 400

    generate_matplotlib_chart(df, chart_type, output_path, title=title,
                               x_col=group_col, y_col=value_col)
    return send_file(output_path, as_attachment=True)


def _load_data_for_chart(file, group_col, value_col):
    """加载 CSV 数据并按指定列分组统计。"""
    output_dir = _config.get("output", {}).get("dir", "output/csv")
    if file:
        path = os.path.join(output_dir, file)
        if not os.path.exists(path):
            return None
        df = pd.read_csv(path, index_col=0)
    else:
        # 合并所有文件
        df = merge_files(output_dir, group_col, value_col)

    if group_col in df.columns and value_col in df.columns:
        # 已经是聚合数据
        if file:
            grouped = df.groupby(group_col).size().reset_index(name=value_col)
            return grouped.sort_values(value_col, ascending=False)
        return df

    # 尝试从原始数据聚合
    if "country" in df.columns:
        grouped = df.groupby("country").size().reset_index(name="count")
        return grouped.sort_values("count", ascending=False)

    return None
```

- [ ] **Step 2: Commit**

```bash
git add web/api.py
git commit -m "feat: add Flask REST API and SSE routes"
```

---

### Task 7: Flask 应用入口

**Files:**
- Create: `app.py`

- [ ] **Step 1: 实现 app.py**

`app.py`:
```python
"""WOS Spider Web 应用入口。"""
import os
from flask import Flask
from web.api import api_bp, init_api


def create_app(config_path="config.yaml"):
    """创建 Flask 应用。"""
    app = Flask(
        __name__,
        template_folder=os.path.join(os.path.dirname(__file__), "web", "templates"),
        static_folder=os.path.join(os.path.dirname(__file__), "web", "static"),
    )
    app.register_blueprint(api_bp)
    init_api(config_path)
    return app


if __name__ == "__main__":
    import yaml
    config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    server_cfg = cfg.get("server", {"host": "127.0.0.1", "port": 5000})

    app = create_app(config_path)
    print(f"WOS Spider 启动: http://{server_cfg['host']}:{server_cfg['port']}")
    app.run(host=server_cfg.get("host", "127.0.0.1"),
            port=server_cfg.get("port", 5000),
            debug=False, threaded=True)
```

- [ ] **Step 2: 验证启动**

```bash
python app.py
```
Expected: 输出 `WOS Spider 启动: http://127.0.0.1:5000`，浏览器访问可打开页面（此时页面空白）

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add Flask app entry point"
```

---

### Task 8: 前端 HTML + CSS（简约白主题）

**Files:**
- Create: `web/templates/index.html`
- Create: `web/static/css/style.css`

- [ ] **Step 1: 创建 style.css**

`web/static/css/style.css`:
```css
/* WOS Spider — 简约白主题 */
:root {
    --bg: #f8fafc;
    --surface: #ffffff;
    --border: #e2e8f0;
    --text: #1e293b;
    --text-secondary: #64748b;
    --primary: #3b82f6;
    --primary-hover: #2563eb;
    --success: #22c55e;
    --success-hover: #16a34a;
    --danger: #ef4444;
    --danger-hover: #dc2626;
    --warning: #f59e0b;
    --info: #3b82f6;
    --radius: 8px;
    --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
}

* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
}

/* Header */
.header {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 24px;
    display: flex;
    align-items: center;
    height: 56px;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: var(--shadow);
}

.header .logo {
    font-weight: 700;
    font-size: 18px;
    color: var(--primary);
    margin-right: 32px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.header .logo span {
    width: 8px; height: 8px;
    background: var(--success);
    border-radius: 50%;
}

/* Tabs */
.tabs {
    display: flex;
    gap: 4px;
}

.tab-btn {
    padding: 8px 16px;
    border: none;
    background: none;
    color: var(--text-secondary);
    font-size: 14px;
    cursor: pointer;
    border-radius: 6px;
    transition: all 0.15s;
}

.tab-btn:hover { background: var(--bg); }
.tab-btn.active {
    background: #eff6ff;
    color: var(--primary);
    font-weight: 500;
}

.status-indicator {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 13px;
    color: var(--text-secondary);
}

/* Main Content */
.main { padding: 24px; max-width: 1200px; margin: 0 auto; }

.tab-panel { display: none; }
.tab-panel.active { display: block; }

/* Cards */
.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px;
    margin-bottom: 16px;
    box-shadow: var(--shadow);
}

.card-title {
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--text);
}

/* Form */
.form-group { margin-bottom: 12px; }

.form-label {
    display: block;
    font-size: 12px;
    color: var(--text-secondary);
    margin-bottom: 4px;
}

.form-input {
    width: 100%;
    padding: 8px 12px;
    border: 1px solid var(--border);
    border-radius: 6px;
    font-size: 13px;
    background: var(--bg);
    color: var(--text);
    outline: none;
    transition: border-color 0.15s;
}

.form-input:focus { border-color: var(--primary); }

.form-row { display: flex; gap: 12px; }
.form-row .form-group { flex: 1; }

/* Buttons */
.btn-group { display: flex; gap: 8px; margin-top: 16px; }

.btn {
    flex: 1;
    padding: 10px 16px;
    border: none;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s;
}

.btn-primary { background: var(--primary); color: white; }
.btn-primary:hover { background: var(--primary-hover); }
.btn-success { background: var(--success); color: white; }
.btn-success:hover { background: var(--success-hover); }
.btn-danger { background: var(--danger); color: white; }
.btn-danger:hover { background: var(--danger-hover); }
.btn-secondary { background: var(--border); color: var(--text); }
.btn-secondary:hover { background: #cbd5e1; }
.btn:disabled { opacity: 0.5; cursor: not-allowed; }

/* Progress */
.progress-bar {
    background: var(--border);
    border-radius: 4px;
    height: 8px;
    overflow: hidden;
    margin: 8px 0;
}

.progress-fill {
    background: linear-gradient(90deg, var(--primary), var(--success));
    height: 100%;
    border-radius: 4px;
    transition: width 0.3s;
}

.progress-info {
    display: flex;
    justify-content: space-between;
    font-size: 12px;
    color: var(--text-secondary);
}

/* Log Panel */
.log-panel {
    background: #1e293b;
    border-radius: var(--radius);
    padding: 12px;
    font-family: "SF Mono", "Fira Code", monospace;
    font-size: 12px;
    line-height: 1.7;
    height: 380px;
    overflow-y: auto;
    color: #94a3b8;
}

.log-panel::-webkit-scrollbar { width: 6px; }
.log-panel::-webkit-scrollbar-thumb { background: #475569; border-radius: 3px; }

.log-info { color: #94a3b8; }
.log-success { color: #4ade80; }
.log-warning { color: #fbbf24; }
.log-error { color: #f87171; }

/* Two Column Layout */
.two-col { display: flex; gap: 16px; }
.two-col > * { flex: 1; }

/* Table */
.data-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

.data-table th {
    background: var(--bg);
    padding: 8px 12px;
    text-align: left;
    font-weight: 600;
    color: var(--text-secondary);
    border-bottom: 1px solid var(--border);
    position: sticky;
    top: 0;
}

.data-table td {
    padding: 8px 12px;
    border-bottom: 1px solid var(--border);
    color: var(--text);
}

.data-table tr:hover td { background: var(--bg); }

.table-wrap {
    max-height: 400px;
    overflow-y: auto;
    border: 1px solid var(--border);
    border-radius: var(--radius);
}

/* File List */
.file-list { list-style: none; }

.file-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 12px;
    border-bottom: 1px solid var(--border);
    cursor: pointer;
    transition: background 0.15s;
}

.file-item:hover { background: var(--bg); }
.file-item:last-child { border-bottom: none; }

.file-name { font-weight: 500; color: var(--primary); }
.file-meta { font-size: 12px; color: var(--text-secondary); }

/* Chart Container */
.chart-container {
    width: 100%;
    height: 400px;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--surface);
}

/* Settings */
.setting-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px 0;
    border-bottom: 1px solid var(--border);
}

.setting-row:last-child { border-bottom: none; }
.setting-label { font-size: 14px; }
.setting-desc { font-size: 12px; color: var(--text-secondary); }

/* Notification */
.notification {
    position: fixed;
    top: 70px;
    right: 24px;
    padding: 12px 20px;
    border-radius: var(--radius);
    color: white;
    font-size: 13px;
    font-weight: 500;
    z-index: 200;
    animation: slideIn 0.3s ease;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

.notification.success { background: var(--success); }
.notification.error { background: var(--danger); }
.notification.warning { background: var(--warning); color: var(--text); }

@keyframes slideIn {
    from { transform: translateX(100px); opacity: 0; }
    to { transform: translateX(0); opacity: 1; }
}

/* Responsive */
@media (max-width: 768px) {
    .two-col { flex-direction: column; }
    .main { padding: 16px; }
    .form-row { flex-direction: column; gap: 0; }
}
```

- [ ] **Step 2: 创建 index.html**

`web/templates/index.html`:
```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WOS Spider</title>
    <link rel="stylesheet" href="/static/css/style.css">
    <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
</head>
<body>
    <!-- Header -->
    <header class="header">
        <div class="logo">
            <span></span> WOS Spider
        </div>
        <nav class="tabs">
            <button class="tab-btn active" data-tab="control">爬虫控制</button>
            <button class="tab-btn" data-tab="data">数据管理</button>
            <button class="tab-btn" data-tab="charts">可视化</button>
            <button class="tab-btn" data-tab="settings">设置</button>
        </nav>
        <div class="status-indicator">
            <span id="statusText">就绪</span>
        </div>
    </header>

    <main class="main">
        <!-- Tab 1: 爬虫控制 -->
        <div id="tab-control" class="tab-panel active">
            <div class="two-col">
                <div>
                    <div class="card">
                        <div class="card-title">爬取配置</div>
                        <div class="form-group">
                            <label class="form-label">WOS 搜索结果 URL</label>
                            <input id="cfgUrl" class="form-input" placeholder="https://www.webofscience.com/...">
                        </div>
                        <div class="form-row">
                            <div class="form-group">
                                <label class="form-label">最大篇数</label>
                                <input id="cfgMax" class="form-input" type="number" value="10000">
                            </div>
                            <div class="form-group">
                                <label class="form-label">等待超时(秒)</label>
                                <input id="cfgTimeout" class="form-input" type="number" value="5">
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="form-label">输出文件名</label>
                            <input id="cfgFile" class="form-input" value="result.csv">
                        </div>
                    </div>
                    <div class="btn-group">
                        <button id="btnStart" class="btn btn-success" onclick="startSpider()">启动浏览器</button>
                        <button id="btnCrawl" class="btn btn-primary" onclick="startCrawl()" disabled>开始爬取</button>
                        <button id="btnPause" class="btn btn-secondary" onclick="pauseSpider()" disabled>暂停</button>
                        <button id="btnStop" class="btn btn-danger" onclick="stopSpider()" disabled>停止</button>
                    </div>
                    <div class="card" style="margin-top:16px">
                        <div class="card-title">进度</div>
                        <div class="progress-bar"><div id="progressFill" class="progress-fill" style="width:0%"></div></div>
                        <div class="progress-info">
                            <span id="progressText">0 / 0</span>
                            <span id="progressStats">已完成: 0 | 失败: 0</span>
                        </div>
                    </div>
                </div>
                <div>
                    <div class="card" style="height:100%">
                        <div class="card-title">实时日志</div>
                        <div id="logPanel" class="log-panel"></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab 2: 数据管理 -->
        <div id="tab-data" class="tab-panel">
            <div class="two-col">
                <div>
                    <div class="card">
                        <div class="card-title">数据文件</div>
                        <ul id="fileList" class="file-list">
                            <li class="file-item"><span class="file-meta">加载中...</span></li>
                        </ul>
                    </div>
                    <div class="card">
                        <div class="card-title">数据合并</div>
                        <div class="form-row">
                            <div class="form-group">
                                <label class="form-label">分组列</label>
                                <select id="mergeGroup" class="form-input">
                                    <option value="country">国家</option>
                                    <option value="publication_year">年份</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label class="form-label">输出文件名</label>
                                <input id="mergeOutput" class="form-input" value="merged.csv">
                            </div>
                        </div>
                        <button class="btn btn-primary" style="margin-top:12px" onclick="mergeData()">合并数据</button>
                    </div>
                </div>
                <div>
                    <div class="card">
                        <div class="card-title">数据预览 <span id="previewFile" style="color:var(--text-secondary);font-weight:400"></span></div>
                        <div class="table-wrap">
                            <table id="previewTable" class="data-table">
                                <thead><tr></tr></thead>
                                <tbody></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab 3: 可视化 -->
        <div id="tab-charts" class="tab-panel">
            <div class="card">
                <div class="card-title">图表配置</div>
                <div class="form-row">
                    <div class="form-group">
                        <label class="form-label">数据文件</label>
                        <select id="chartFile" class="form-input"><option value="">全部数据</option></select>
                    </div>
                    <div class="form-group">
                        <label class="form-label">图表类型</label>
                        <select id="chartType" class="form-input">
                            <option value="bar">柱状图</option>
                            <option value="pie">饼图</option>
                            <option value="line">折线图</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label class="form-label">分组列</label>
                        <select id="chartGroup" class="form-input">
                            <option value="country">国家</option>
                            <option value="publication_year">年份</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label class="form-label">图表标题</label>
                        <input id="chartTitle" class="form-input" value="论文数量统计">
                    </div>
                </div>
                <div class="btn-group" style="max-width:400px">
                    <button class="btn btn-primary" onclick="loadEChart()">生成交互图表</button>
                    <button class="btn btn-secondary" onclick="loadMatplotlib()">导出学术图表</button>
                </div>
            </div>
            <div class="card">
                <div id="echartContainer" class="chart-container"></div>
            </div>
        </div>

        <!-- Tab 4: 设置 -->
        <div id="tab-settings" class="tab-panel">
            <div class="card">
                <div class="card-title">浏览器配置</div>
                <div class="setting-row">
                    <div>
                        <div class="setting-label">无头模式</div>
                        <div class="setting-desc">不显示浏览器窗口（需要已登录 Cookie）</div>
                    </div>
                    <input id="setHeadless" type="checkbox">
                </div>
            </div>
            <div class="card">
                <div class="card-title">爬取策略</div>
                <div class="setting-row">
                    <div>
                        <div class="setting-label">最大重试次数</div>
                        <div class="setting-desc">单篇论文失败后最大重试次数</div>
                    </div>
                    <input id="setRetries" class="form-input" type="number" value="3" style="width:80px">
                </div>
                <div class="setting-row">
                    <div>
                        <div class="setting-label">暂停间隔(秒)</div>
                        <div class="setting-desc">每篇论文之间的等待时间</div>
                    </div>
                    <input id="setInterval" class="form-input" type="number" value="0.1" step="0.1" style="width:80px">
                </div>
            </div>
            <div class="card">
                <div class="card-title">服务器配置</div>
                <div class="setting-row">
                    <div>
                        <div class="setting-label">监听端口</div>
                    </div>
                    <input id="setPort" class="form-input" type="number" value="5000" style="width:100px">
                </div>
            </div>
            <button class="btn btn-primary" onclick="saveSettings()">保存设置</button>
        </div>
    </main>

    <script src="/static/js/app.js"></script>
    <script src="/static/js/charts.js"></script>
</body>
</html>
```

- [ ] **Step 3: Commit**

```bash
git add web/templates/index.html web/static/css/style.css
git commit -m "feat: add frontend HTML and CSS (light theme)"
```

---

### Task 9: 前端主逻辑 (app.js)

**Files:**
- Create: `web/static/js/app.js`

- [ ] **Step 1: 实现 app.js**

`web/static/js/app.js`:
```javascript
/* WOS Spider 前端主逻辑 */
let sseSource = null;
let currentLogIndex = 0;

// --- Tab 切换 ---
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('tab-' + btn.dataset.tab).classList.add('active');

        if (btn.dataset.tab === 'data') loadFileList();
        if (btn.dataset.tab === 'charts') loadChartFileOptions();
    });
});

// --- 通知 ---
function notify(message, type = 'success') {
    const el = document.createElement('div');
    el.className = 'notification ' + type;
    el.textContent = message;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 3000);
}

// --- 爬虫控制 ---
async function api(method, path, body) {
    const opts = { method, headers: { 'Content-Type': 'application/json' } };
    if (body) opts.body = JSON.stringify(body);
    const res = await fetch(path, opts);
    return res.json();
}

async function startSpider() {
    const url = document.getElementById('cfgUrl').value;
    if (!url) { notify('请输入 WOS URL', 'warning'); return; }
    const max = parseInt(document.getElementById('cfgMax').value) || 10000;
    const file = document.getElementById('cfgFile').value || 'result.csv';

    const res = await api('POST', '/api/spider/start', { url, max_papers: max, file_path: file });
    if (res.ok) {
        notify('浏览器已启动，请手动登录');
        document.getElementById('btnStart').disabled = true;
        document.getElementById('btnCrawl').disabled = false;
        document.getElementById('btnStop').disabled = false;
        updateStatus('browser_ready', '浏览器已启动');
        startLogStream();
    }
}

async function startCrawl() {
    const url = document.getElementById('cfgUrl').value;
    const max = parseInt(document.getElementById('cfgMax').value) || 10000;
    const file = document.getElementById('cfgFile').value || 'result.csv';

    const res = await api('POST', '/api/spider/crawl', { url, max_papers: max, file_path: file });
    if (res.ok) {
        notify('爬取已开始');
        document.getElementById('btnCrawl').disabled = true;
        document.getElementById('btnPause').disabled = false;
        updateStatus('running', '正在爬取');
    }
}

async function pauseSpider() {
    const res = await api('POST', '/api/spider/pause');
    if (res.ok) {
        notify('已暂停');
        document.getElementById('btnPause').disabled = true;
        document.getElementById('btnCrawl').disabled = false;
        document.getElementById('btnCrawl').textContent = '继续';
        document.getElementById('btnCrawl').onclick = resumeSpider;
        updateStatus('paused', '已暂停');
    }
}

async function resumeSpider() {
    const res = await api('POST', '/api/spider/resume');
    if (res.ok) {
        notify('已继续');
        document.getElementById('btnPause').disabled = false;
        document.getElementById('btnCrawl').disabled = true;
        updateStatus('running', '正在爬取');
    }
}

async function stopSpider() {
    const res = await api('POST', '/api/spider/stop');
    if (res.ok) {
        notify('已停止');
        document.getElementById('btnStart').disabled = false;
        document.getElementById('btnCrawl').disabled = true;
        document.getElementById('btnPause').disabled = true;
        document.getElementById('btnStop').disabled = true;
        updateStatus('stopped', '已停止');
    }
}

function updateStatus(status, text) {
    document.getElementById('statusText').textContent = text;
}

// --- SSE 日志流 ---
function startLogStream() {
    if (sseSource) sseSource.close();
    sseSource = new EventSource('/api/spider/logs?since=' + currentLogIndex);
    sseSource.onmessage = function(event) {
        const data = JSON.parse(event.data);
        if (data.type === 'heartbeat') return;
        currentLogIndex++;
        appendLog(data);
        updateProgress();
    };
    sseSource.onerror = function() {
        sseSource.close();
        setTimeout(startLogStream, 5000);
    };
}

function appendLog(entry) {
    const panel = document.getElementById('logPanel');
    const line = document.createElement('div');
    line.className = 'log-' + (entry.level || 'info');
    line.textContent = '[' + entry.time + '] ' + entry.message;
    panel.appendChild(line);
    panel.scrollTop = panel.scrollHeight;
}

async function updateProgress() {
    const res = await api('GET', '/api/spider/status');
    const p = res.progress || {};
    const total = p.total || 0;
    const done = p.done || 0;
    const pct = total > 0 ? (done / total * 100).toFixed(1) : 0;

    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressText').textContent = done + ' / ' + total;
    document.getElementById('progressStats').textContent =
        '已完成: ' + done + ' | 失败: ' + (p.failed || 0);

    if (res.status === 'completed') {
        updateStatus('completed', '爬取完成');
        notify('爬取完成！');
        document.getElementById('btnStart').disabled = false;
        document.getElementById('btnCrawl').disabled = true;
        document.getElementById('btnPause').disabled = true;
        document.getElementById('btnStop').disabled = true;
    }
}

// --- 数据管理 ---
async function loadFileList() {
    const res = await api('GET', '/api/data/files');
    const list = document.getElementById('fileList');
    list.innerHTML = '';
    if (!res.length) {
        const li = document.createElement('li');
        li.className = 'file-item';
        const span = document.createElement('span');
        span.className = 'file-meta';
        span.textContent = '暂无数据文件';
        li.appendChild(span);
        list.appendChild(li);
        return;
    }
    res.forEach(f => {
        const li = document.createElement('li');
        li.className = 'file-item';
        const nameDiv = document.createElement('div');
        const nameSpan = document.createElement('span');
        nameSpan.className = 'file-name';
        nameSpan.textContent = f.name;
        nameDiv.appendChild(nameSpan);
        const metaDiv = document.createElement('div');
        metaDiv.className = 'file-meta';
        metaDiv.textContent = formatSize(f.size) + ' | ' + f.modified;
        li.appendChild(nameDiv);
        li.appendChild(metaDiv);
        li.onclick = () => loadPreview(f.name);
        list.appendChild(li);
    });
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

async function loadPreview(filename) {
    document.getElementById('previewFile').textContent = filename;
    const res = await api('GET', '/api/data/preview?file=' + encodeURIComponent(filename));
    const table = document.getElementById('previewTable');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = '';
    const thIdx = document.createElement('th');
    thIdx.textContent = '#';
    thead.appendChild(thIdx);
    (res.columns || []).forEach(c => {
        const th = document.createElement('th');
        th.textContent = c;
        thead.appendChild(th);
    });
    tbody.innerHTML = '';
    (res.rows || []).forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

async function mergeData() {
    const groupCol = document.getElementById('mergeGroup').value;
    const output = document.getElementById('mergeOutput').value;
    const res = await api('POST', '/api/data/merge', {
        group_col: groupCol, value_col: 'count', output_filename: output
    });
    if (res.ok) {
        notify('合并完成，共 ' + res.rows + ' 条记录');
        loadFileList();
    }
}

// --- 设置 ---
async function saveSettings() {
    const config = {
        spider: {
            max_retries: parseInt(document.getElementById('setRetries').value) || 3,
            pause_interval: parseFloat(document.getElementById('setInterval').value) || 0.1,
            headless: document.getElementById('setHeadless').checked,
        },
        server: {
            port: parseInt(document.getElementById('setPort').value) || 5000,
        },
    };
    const res = await api('POST', '/api/config', config);
    if (res.ok) notify('设置已保存');
}

// --- 初始化 ---
(async function init() {
    const res = await api('GET', '/api/config');
    if (res.spider) {
        document.getElementById('setRetries').value = res.spider.max_retries || 3;
        document.getElementById('setInterval').value = res.spider.pause_interval || 0.1;
        document.getElementById('setHeadless').checked = res.spider.headless || false;
    }
    if (res.server) {
        document.getElementById('setPort').value = res.server.port || 5000;
    }
})();
```

- [ ] **Step 2: Commit**

```bash
git add web/static/js/app.js
git commit -m "feat: add frontend JS logic with SSE and tab switching"
```

---

### Task 10: ECharts 前端封装 (charts.js)

**Files:**
- Create: `web/static/js/charts.js`

- [ ] **Step 1: 实现 charts.js**

`web/static/js/charts.js`:
```javascript
/* ECharts 图表封装 */
let echartInstance = null;

async function loadChartFileOptions() {
    const res = await (await fetch('/api/data/files')).json();
    const sel = document.getElementById('chartFile');
    // 保留第一个 "全部数据" 选项
    while (sel.options.length > 1) sel.remove(1);
    res.forEach(f => {
        const opt = document.createElement('option');
        opt.value = f.name;
        opt.textContent = f.name;
        sel.appendChild(opt);
    });
}

async function loadEChart() {
    const container = document.getElementById('echartContainer');
    if (!echartInstance) {
        echartInstance = echarts.init(container);
    }

    const params = {
        chart_type: document.getElementById('chartType').value,
        file: document.getElementById('chartFile').value,
        group_col: document.getElementById('chartGroup').value,
        value_col: 'count',
        title: document.getElementById('chartTitle').value,
    };

    try {
        const res = await fetch('/api/charts/echarts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(params),
        });
        const option = await res.json();
        if (option.error) {
            notify(option.error, 'error');
            return;
        }
        echartInstance.setOption(option, true);
    } catch (e) {
        notify('图表加载失败: ' + e.message, 'error');
    }
}

async function loadMatplotlib() {
    const params = {
        chart_type: document.getElementById('chartType').value,
        file: document.getElementById('chartFile').value,
        group_col: document.getElementById('chartGroup').value,
        value_col: 'count',
        title: document.getElementById('chartTitle').value,
    };

    try {
        const res = await fetch('/api/charts/matplotlib', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(params),
        });
        if (!res.ok) throw new Error('请求失败');
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'chart.png';
        a.click();
        URL.revokeObjectURL(url);
        notify('图表已导出');
    } catch (e) {
        notify('导出失败: ' + e.message, 'error');
    }
}

window.addEventListener('resize', () => {
    if (echartInstance) echartInstance.resize();
});
```

- [ ] **Step 2: Commit**

```bash
git add web/static/js/charts.js
git commit -m "feat: add ECharts frontend wrapper"
```

---

### Task 11: .gitignore + 最终集成

**Files:**
- Create: `.gitignore`

- [ ] **Step 1: 创建 .gitignore**

```
__pycache__/
*.pyc
.env
output/csv/*.csv
output/csv/*.xlsx
output/charts/*.png
.superpowers/
node_modules/
```

- [ ] **Step 2: 运行全部测试**

```bash
python -m pytest tests/ -v
```
Expected: ALL PASS (test_parser, test_merger, test_visualizer)

- [ ] **Step 3: 启动应用验证**

```bash
python app.py
```
Expected: 浏览器打开 http://127.0.0.1:5000 可看到完整 UI（四个 Tab 均可切换）

- [ ] **Step 4: Commit**

```bash
git add .gitignore
git commit -m "feat: add .gitignore and finalize project structure"
```
