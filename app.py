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
