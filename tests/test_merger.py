import os
import pytest
import pandas as pd
from wos_spider.merger import merge_files


@pytest.fixture
def csv_dir(tmp_path):
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
