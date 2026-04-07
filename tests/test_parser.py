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
        html = "<html><body><h2 class='title text--large cdx-title'>Only Title</h2></body></html>"
        index, data = parse_paper_html(html)
        assert data["title"] == "Only Title"
        assert "citation" not in data or data.get("citation") == "0"

    def test_empty_html_returns_defaults(self):
        html = "<html><body></body></html>"
        index, data = parse_paper_html(html)
        assert isinstance(data, dict)
