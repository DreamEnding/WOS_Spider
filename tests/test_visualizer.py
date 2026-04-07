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
