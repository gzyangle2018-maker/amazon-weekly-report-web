/**
 * 主应用逻辑
 */
class App {
    constructor() {
        this.parser = new ReportParser();
        this.processor = new DataProcessor();
        this.exporter = new ExcelExporter();
        this.files = {
            sales: null,
            profit: null,
            lastSales: null,
            lastProfit: null,
        };
        this.resultData = null;
        this.init();
    }

    init() {
        this.setupNavigation();
        this.setupDropZones();
        this.setupPeriodType();
        this.setupButtons();
        this.setDefaultDateRange();
    }

    setupNavigation() {
        const navItems = document.querySelectorAll('.nav-item');
        navItems.forEach(item => {
            item.addEventListener('click', (e) => {
                e.preventDefault();
                const page = item.dataset.page;
                this.switchPage(page);
                navItems.forEach(n => n.classList.remove('active'));
                item.classList.add('active');
            });
        });
    }

    switchPage(page) {
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        const target = document.querySelector(`.page-${page}`);
        if (target) target.classList.add('active');
    }

    setupDropZones() {
        const zones = [
            { dropId: 'salesDrop', inputId: 'salesFile', statusId: 'salesStatus', key: 'sales' },
            { dropId: 'profitDrop', inputId: 'profitFile', statusId: 'profitStatus', key: 'profit' },
            { dropId: 'lastSalesDrop', inputId: 'lastSalesFile', statusId: 'lastSalesStatus', key: 'lastSales' },
            { dropId: 'lastProfitDrop', inputId: 'lastProfitFile', statusId: 'lastProfitStatus', key: 'lastProfit' },
        ];

        for (const z of zones) {
            const drop = document.getElementById(z.dropId);
            const input = document.getElementById(z.inputId);
            const status = document.getElementById(z.statusId);

            if (!drop || !input) continue;

            drop.addEventListener('dragover', (e) => {
                e.preventDefault();
                drop.classList.add('drag-over');
            });
            drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
            drop.addEventListener('drop', (e) => {
                e.preventDefault();
                drop.classList.remove('drag-over');
                const files = e.dataTransfer.files;
                if (files.length) {
                    input.files = files;
                    this.handleFile(files[0], z.key, drop, status);
                }
            });
            input.addEventListener('change', () => {
                if (input.files.length) {
                    this.handleFile(input.files[0], z.key, drop, status);
                }
            });
        }
    }

    handleFile(file, key, dropEl, statusEl) {
        this.files[key] = file;
        dropEl.classList.add('has-file');
        const title = dropEl.querySelector('.drop-zone-title');
        if (title) title.textContent = file.name;
        const subtitle = dropEl.querySelector('.drop-zone-subtitle');
        if (subtitle) subtitle.textContent = '点击或拖拽可替换文件';
        statusEl.textContent = file.name;
        statusEl.classList.add('has-file');
    }

    setupPeriodType() {
        const select = document.getElementById('periodType');
        const input = document.getElementById('dateRange');
        const tip = document.getElementById('dateTip');

        select.addEventListener('change', () => {
            const type = select.value;
            if (type === 'weekly') {
                input.placeholder = '如: 4.27-5.3';
                tip.textContent = '格式: M.D-M.D';
            } else if (type === 'monthly') {
                input.placeholder = '如: 2026-04';
                tip.textContent = '格式: YYYY-MM';
            } else {
                input.placeholder = '如: 4.1-5.3';
                tip.textContent = '格式: M.D-M.D';
            }
            input.value = this.processor.getDefaultDateRange(type);
        });
    }

    setDefaultDateRange() {
        const input = document.getElementById('dateRange');
        input.value = this.processor.getDefaultDateRange('weekly');
    }

    setupButtons() {
        document.getElementById('btnProcess').addEventListener('click', () => this.process());
        document.getElementById('btnClear').addEventListener('click', () => this.clearAll());
        document.getElementById('btnExport').addEventListener('click', () => this.exportExcel());
    }

    showProgress(percent) {
        const bar = document.getElementById('progressBar');
        const fill = document.getElementById('progressFill');
        bar.style.display = 'block';
        fill.style.width = percent + '%';
    }

    hideProgress() {
        document.getElementById('progressBar').style.display = 'none';
    }

    async process() {
        if (!this.files.sales) {
            alert('请先选择销量报表');
            return;
        }

        const dateRange = document.getElementById('dateRange').value.trim();
        if (!dateRange) {
            alert('请输入报表周期');
            return;
        }

        const btn = document.getElementById('btnProcess');
        btn.disabled = true;
        btn.textContent = '处理中...';
        this.showProgress(30);

        try {
            this.resultData = await this.processor.process(
                this.files.sales,
                this.files.profit,
                this.files.lastSales,
                this.files.lastProfit,
                dateRange
            );

            this.showProgress(100);
            setTimeout(() => this.hideProgress(), 500);

            this.renderReport(dateRange);
            this.switchPage('report');
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            document.querySelector('[data-page="report"]').classList.add('active');
        } catch (err) {
            alert('处理失败: ' + err.message);
            console.error(err);
        } finally {
            btn.disabled = false;
            btn.textContent = '开始处理';
        }
    }

    renderReport(dateRange) {
        document.getElementById('reportTitle').textContent = `周报预览 - ${dateRange}`;

        const stats = this.processor.getSummaryStats();
        document.getElementById('statAsin').textContent = stats['总ASIN数'] || 0;
        document.getElementById('statRevenue').textContent = '$' + (stats['总7天销售额'] || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        document.getElementById('statUnits').textContent = (stats['总7天销量'] || 0).toLocaleString();
        document.getElementById('statAdSpend').textContent = '$' + (stats['总广告费用'] || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        document.getElementById('statTacos').textContent = (stats['平均Tacos'] || 0).toFixed(2) + '%';

        const tbody = document.getElementById('tableBody');
        if (!this.resultData || !this.resultData.length) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="11">暂无数据</td></tr>';
            return;
        }

        const displayCols = [
            { key: 'ASIN', fmt: v => v },
            { key: '店铺', fmt: v => v },
            { key: '7天销售额', fmt: v => '$' + (v || 0).toLocaleString('en-US', { minimumFractionDigits: 2 }) },
            { key: '7天销量', fmt: v => (v || 0).toLocaleString() },
            { key: '30天销量', fmt: v => (v || 0).toLocaleString() },
            { key: '均价', fmt: v => '$' + (v || 0).toFixed(2) },
            { key: '广告费用', fmt: v => '$' + (v || 0).toLocaleString('en-US', { minimumFractionDigits: 2 }) },
            { key: 'Tacos', fmt: v => (v || 0).toFixed(2) + '%' },
            { key: '广告转化率', fmt: v => (v || 0).toFixed(2) + '%' },
            { key: '自然转化率', fmt: v => (v || 0).toFixed(2) + '%' },
            { key: 'CPA', fmt: v => '$' + (v || 0).toFixed(2) },
        ];

        let html = '';
        const wowCols = ['7天销售额_环比', '7天销量_环比', '30天销量_环比', '广告费用_环比', 'Tacos_环比'];

        for (const row of this.resultData.slice(0, 500)) {
            html += '<tr>';
            for (const col of displayCols) {
                let val = col.fmt(row[col.key]);
                let cls = '';
                // 环比颜色
                const wowKey = col.key + '_环比';
                if (row[wowKey] !== undefined) {
                    const wow = parseFloat(row[wowKey]);
                    if (wow > 0) cls = 'wow-up';
                    else if (wow < 0) cls = 'wow-down';
                }
                html += `<td class="${cls}">${val}</td>`;
            }
            html += '</tr>';
        }
        tbody.innerHTML = html;
    }

    exportExcel() {
        if (!this.resultData || !this.resultData.length) {
            alert('没有可导出的数据');
            return;
        }
        const dateRange = document.getElementById('dateRange').value.trim();
        this.exporter.generate(this.resultData, dateRange);
    }

    clearAll() {
        this.files = { sales: null, profit: null, lastSales: null, lastProfit: null };
        this.resultData = null;

        const clears = [
            { drop: 'salesDrop', status: 'salesStatus', title: '拖拽销量报表到此处', sub: '支持 优卖云 / 赛狐ERP / 通用格式' },
            { drop: 'profitDrop', status: 'profitStatus', title: '拖拽利润报表到此处', sub: '支持 优卖云 / 赛狐ERP / 通用格式' },
            { drop: 'lastSalesDrop', status: 'lastSalesStatus', title: '拖拽上周销量报表', sub: '可选' },
            { drop: 'lastProfitDrop', status: 'lastProfitStatus', title: '拖拽上周利润报表', sub: '可选' },
        ];

        for (const c of clears) {
            const drop = document.getElementById(c.drop);
            const status = document.getElementById(c.status);
            if (drop) {
                drop.classList.remove('has-file');
                const title = drop.querySelector('.drop-zone-title');
                const sub = drop.querySelector('.drop-zone-subtitle');
                if (title) title.textContent = c.title;
                if (sub) sub.textContent = c.sub;
            }
            if (status) {
                status.textContent = '未选择文件';
                status.classList.remove('has-file');
            }
        }

        document.getElementById('salesFile').value = '';
        document.getElementById('profitFile').value = '';
        document.getElementById('lastSalesFile').value = '';
        document.getElementById('lastProfitFile').value = '';

        this.setDefaultDateRange();
        this.hideProgress();
    }
}

// 启动应用
document.addEventListener('DOMContentLoaded', () => {
    window.app = new App();
});
