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
    """生成 matplotlib 图表并保存为图片。"""
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
    """生成 ECharts 配置字典。"""
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
