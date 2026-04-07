/* ECharts 图表封装 */
var echartInstance = null;

function loadChartFileOptions() {
    fetch('/api/data/files').then(function(res) { return res.json(); }).then(function(res) {
        var sel = document.getElementById('chartFile');
        // 保留第一个 "全部数据" 选项
        while (sel.options.length > 1) sel.remove(1);
        res.forEach(function(f) {
            var opt = document.createElement('option');
            opt.value = f.name;
            opt.textContent = f.name;
            sel.appendChild(opt);
        });
    });
}

function loadEChart() {
    var container = document.getElementById('echartContainer');
    if (!echartInstance) {
        echartInstance = echarts.init(container);
    }

    var params = {
        chart_type: document.getElementById('chartType').value,
        file: document.getElementById('chartFile').value,
        group_col: document.getElementById('chartGroup').value,
        value_col: 'count',
        title: document.getElementById('chartTitle').value,
    };

    fetch('/api/charts/echarts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
    })
    .then(function(res) { return res.json(); })
    .then(function(option) {
        if (option.error) {
            notify(option.error, 'error');
            return;
        }
        echartInstance.setOption(option, true);
    })
    .catch(function(e) {
        notify('图表加载失败: ' + e.message, 'error');
    });
}

function loadMatplotlib() {
    var params = {
        chart_type: document.getElementById('chartType').value,
        file: document.getElementById('chartFile').value,
        group_col: document.getElementById('chartGroup').value,
        value_col: 'count',
        title: document.getElementById('chartTitle').value,
    };

    fetch('/api/charts/matplotlib', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
    })
    .then(function(res) {
        if (!res.ok) throw new Error('请求失败');
        return res.blob();
    })
    .then(function(blob) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'chart.png';
        a.click();
        URL.revokeObjectURL(url);
        notify('图表已导出');
    })
    .catch(function(e) {
        notify('导出失败: ' + e.message, 'error');
    });
}

window.addEventListener('resize', function() {
    if (echartInstance) echartInstance.resize();
});
