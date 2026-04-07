/* WOS Spider 前端主逻辑 */
var sseSource = null;
var currentLogIndex = 0;

// --- Tab 切换 ---
document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.addEventListener('click', function() {
        document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
        document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('active'); });
        btn.classList.add('active');
        document.getElementById('tab-' + btn.dataset.tab).classList.add('active');

        if (btn.dataset.tab === 'data') loadFileList();
        if (btn.dataset.tab === 'charts') loadChartFileOptions();
    });
});

// --- 通知 ---
function notify(message, type) {
    type = type || 'success';
    var el = document.createElement('div');
    el.className = 'notification ' + type;
    el.textContent = message;
    document.body.appendChild(el);
    setTimeout(function() { el.remove(); }, 3000);
}

// --- API helper ---
function api(method, path, body) {
    var opts = { method: method, headers: { 'Content-Type': 'application/json' } };
    if (body) opts.body = JSON.stringify(body);
    return fetch(path, opts).then(function(res) { return res.json(); });
}

// --- 爬虫控制 ---
function startSpider() {
    var url = document.getElementById('cfgUrl').value;
    if (!url) { notify('请输入 WOS URL', 'warning'); return; }
    var max = parseInt(document.getElementById('cfgMax').value) || 10000;
    var file = document.getElementById('cfgFile').value || 'result.csv';

    api('POST', '/api/spider/start', { url: url, max_papers: max, file_path: file })
    .then(function(res) {
        if (res.ok) {
            notify('浏览器已启动，请手动登录');
            document.getElementById('btnStart').disabled = true;
            document.getElementById('btnCrawl').disabled = false;
            document.getElementById('btnStop').disabled = false;
            updateStatus('browser_ready', '浏览器已启动');
            startLogStream();
        }
    });
}

function startCrawl() {
    var url = document.getElementById('cfgUrl').value;
    var max = parseInt(document.getElementById('cfgMax').value) || 10000;
    var file = document.getElementById('cfgFile').value || 'result.csv';

    api('POST', '/api/spider/crawl', { url: url, max_papers: max, file_path: file })
    .then(function(res) {
        if (res.ok) {
            notify('爬取已开始');
            document.getElementById('btnCrawl').disabled = true;
            document.getElementById('btnPause').disabled = false;
            updateStatus('running', '正在爬取');
        }
    });
}

function pauseSpider() {
    api('POST', '/api/spider/pause').then(function(res) {
        if (res.ok) {
            notify('已暂停');
            document.getElementById('btnPause').disabled = true;
            document.getElementById('btnCrawl').disabled = false;
            document.getElementById('btnCrawl').textContent = '继续';
            document.getElementById('btnCrawl').onclick = resumeSpider;
            updateStatus('paused', '已暂停');
        }
    });
}

function resumeSpider() {
    api('POST', '/api/spider/resume').then(function(res) {
        if (res.ok) {
            notify('已继续');
            document.getElementById('btnPause').disabled = false;
            document.getElementById('btnCrawl').disabled = true;
            updateStatus('running', '正在爬取');
        }
    });
}

function stopSpider() {
    api('POST', '/api/spider/stop').then(function(res) {
        if (res.ok) {
            notify('已停止');
            document.getElementById('btnStart').disabled = false;
            document.getElementById('btnCrawl').disabled = true;
            document.getElementById('btnPause').disabled = true;
            document.getElementById('btnStop').disabled = true;
            updateStatus('stopped', '已停止');
        }
    });
}

function updateStatus(status, text) {
    document.getElementById('statusText').textContent = text;
}

// --- SSE 日志流 ---
function startLogStream() {
    if (sseSource) sseSource.close();
    sseSource = new EventSource('/api/spider/logs?since=' + currentLogIndex);
    sseSource.onmessage = function(event) {
        var data = JSON.parse(event.data);
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
    var panel = document.getElementById('logPanel');
    var line = document.createElement('div');
    line.className = 'log-' + (entry.level || 'info');
    line.textContent = '[' + entry.time + '] ' + entry.message;
    panel.appendChild(line);
    panel.scrollTop = panel.scrollHeight;
}

function updateProgress() {
    api('GET', '/api/spider/status').then(function(res) {
        var p = res.progress || {};
        var total = p.total || 0;
        var done = p.done || 0;
        var pct = total > 0 ? (done / total * 100).toFixed(1) : 0;

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
    });
}

// --- 数据管理 ---
function clearElement(el) {
    while (el.firstChild) el.removeChild(el.firstChild);
}

function loadFileList() {
    api('GET', '/api/data/files').then(function(res) {
        var list = document.getElementById('fileList');
        clearElement(list);
        if (!res.length) {
            var li = document.createElement('li');
            li.className = 'file-item';
            var span = document.createElement('span');
            span.className = 'file-meta';
            span.textContent = '暂无数据文件';
            li.appendChild(span);
            list.appendChild(li);
            return;
        }
        res.forEach(function(f) {
            var li = document.createElement('li');
            li.className = 'file-item';
            var nameDiv = document.createElement('div');
            var nameSpan = document.createElement('span');
            nameSpan.className = 'file-name';
            nameSpan.textContent = f.name;
            nameDiv.appendChild(nameSpan);
            var metaDiv = document.createElement('div');
            metaDiv.className = 'file-meta';
            metaDiv.textContent = formatSize(f.size) + ' | ' + f.modified;
            li.appendChild(nameDiv);
            li.appendChild(metaDiv);
            li.onclick = function() { loadPreview(f.name); };
            list.appendChild(li);
        });
    });
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

function loadPreview(filename) {
    document.getElementById('previewFile').textContent = filename;
    api('GET', '/api/data/preview?file=' + encodeURIComponent(filename)).then(function(res) {
        var table = document.getElementById('previewTable');
        var thead = table.querySelector('thead tr');
        var tbody = table.querySelector('tbody');
        clearElement(thead);
        var thIdx = document.createElement('th');
        thIdx.textContent = '#';
        thead.appendChild(thIdx);
        (res.columns || []).forEach(function(c) {
            var th = document.createElement('th');
            th.textContent = c;
            thead.appendChild(th);
        });
        clearElement(tbody);
        (res.rows || []).forEach(function(row) {
            var tr = document.createElement('tr');
            row.forEach(function(cell) {
                var td = document.createElement('td');
                td.textContent = cell || '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    });
}

function mergeData() {
    var groupCol = document.getElementById('mergeGroup').value;
    var output = document.getElementById('mergeOutput').value;
    api('POST', '/api/data/merge', {
        group_col: groupCol, value_col: 'count', output_filename: output
    }).then(function(res) {
        if (res.ok) {
            notify('合并完成，共 ' + res.rows + ' 条记录');
            loadFileList();
        }
    });
}

// --- 设置 ---
function saveSettings() {
    var config = {
        spider: {
            max_retries: parseInt(document.getElementById('setRetries').value) || 3,
            pause_interval: parseFloat(document.getElementById('setInterval').value) || 0.1,
            headless: document.getElementById('setHeadless').checked,
        },
        server: {
            port: parseInt(document.getElementById('setPort').value) || 5000,
        },
    };
    api('POST', '/api/config', config).then(function(res) {
        if (res.ok) notify('设置已保存');
    });
}

// --- 初始化 ---
(function init() {
    api('GET', '/api/config').then(function(res) {
        if (res.spider) {
            document.getElementById('setRetries').value = res.spider.max_retries || 3;
            document.getElementById('setInterval').value = res.spider.pause_interval || 0.1;
            document.getElementById('setHeadless').checked = res.spider.headless || false;
        }
        if (res.server) {
            document.getElementById('setPort').value = res.server.port || 5000;
        }
    });
})();
