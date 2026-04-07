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
        self._log_queue = []
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
        self._pause_event.set()
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

    XPATH_TITLE = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[1]/div/div/div/h2'
    XPATH_YEAR = '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[7]/div[2]/span'
