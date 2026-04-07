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

    _extract_container_fields(soup, data)

    data["title"] = _try_extract(soup, [
        lambda s: s.find(class_="title text--large cdx-title").text,
    ], default="未获取到标题")

    data["citation"] = _try_extract(soup, [
        lambda s: s.find(class_="mat-tooltip-trigger medium-link-number link ng-star-inserted").text,
    ], default="0")

    data["country"] = _try_extract(soup, [
        lambda s: s.find("span", class_="ng-star-inserted",
                          id="FRAOrgTa-RepAddressFull-0").text.split(",")[-1],
        lambda s: s.find("span", class_="value padding-right-5--reversible").text.split(",")[-1],
    ], default="未获取到国家")

    data["journal"] = _try_extract(soup, [
        lambda s: s.find(class_="mat-focus-indicator mat-tooltip-trigger font-size-14 summary-source-title-link remove-space no-left-padding mat-button mat-button-base mat-primary font-size-16 ng-star-inserted").text,
    ], default="未获取到期刊")

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
            return int(input_box["aria-label"].split()[1].replace(",", ""))
    except Exception:
        pass
    return 0
