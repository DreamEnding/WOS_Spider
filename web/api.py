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


@api_bp.route("/")
def index():
    return render_template("index.html")


@api_bp.route("/api/config", methods=["GET"])
def get_config():
    safe_config = {k: v for k, v in _config.items()}
    return jsonify(safe_config)


@api_bp.route("/api/config", methods=["POST"])
def update_config():
    global _config
    updates = request.json
    _config.update(updates)
    _save_config()
    return jsonify({"ok": True})


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
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"
            time.sleep(0.5)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


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
        df = merge_files(output_dir, group_col, value_col)

    if group_col in df.columns and value_col in df.columns:
        if file:
            grouped = df.groupby(group_col).size().reset_index(name=value_col)
            return grouped.sort_values(value_col, ascending=False)
        return df

    if "country" in df.columns:
        grouped = df.groupby("country").size().reset_index(name="count")
        return grouped.sort_values("count", ascending=False)

    return None
