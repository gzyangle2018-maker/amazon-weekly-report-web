/**
 * Excel 导出器 - 生成格式化的周报 Excel
 */
class ExcelExporter {
    constructor() {
        this.COLORS = {
            headerBg: 'E8F4FD',
            headerFont: '1C1C1E',
            asinBgOdd: 'F0F9FF',
            asinBgEven: 'FFFFFF',
            borderColor: 'D1D5DB',
            accentBlue: '007AFF',
            accentGreen: '34C759',
            accentRed: 'FF3B30',
            textPrimary: '1C1C1E',
            textSecondary: '8E8E93',
            liquidGlass: 'EBEBF0',
        };
        this.METRICS = [
            { name: '7天销售额', unit: 'usd', currency: true },
            { name: '7天销量', unit: '', currency: false },
            { name: '30天销量', unit: '', currency: false },
            { name: '均价', unit: 'ERP', currency: true },
            { name: '广告费用', unit: '后台sp+sb+sd', currency: true },
            { name: 'Tacos', unit: '%', percent: true },
            { name: '30天退款', unit: '', currency: false },
            { name: '90天退款率', unit: '%', percent: true },
            { name: '广告转化率', unit: '%', percent: true },
            { name: '自然转化率', unit: '%', percent: true },
            { name: 'CPA', unit: 'usd', currency: true },
        ];
    }

    generate(data, dateRange, includeWow = true) {
        const wb = XLSX.utils.book_new();
        const wsData = [];

        // 标题行
        wsData.push([`亚马逊周报销量汇总 - ${dateRange}${includeWow ? ' (含环比)' : ''}`]);

        // 表头
        const headers = ['ASIN', '店铺', '类型', dateRange];
        const wowMetrics = ['7天销售额', '7天销量', '30天销量', '广告费用', 'Tacos', '广告转化率', '自然转化率', 'CPA'];
        if (includeWow) {
            for (const m of wowMetrics) headers.push(`${m}环比`);
        }
        wsData.push(headers);

        // 数据行
        for (const row of data) {
            const asin = row.ASIN || '';
            const store = row.店铺 || '';
            for (const metric of this.METRICS) {
                const rowData = [asin, store, metric.name];
                let val = row[metric.name] || 0;
                if (metric.percent) val = val / 100;
                rowData.push(val);

                if (includeWow) {
                    for (const wm of wowMetrics) {
                        const wowVal = row[`${wm}_环比`] || 0;
                        rowData.push(wowVal / 100);
                    }
                }
                wsData.push(rowData);
            }
        }

        // 汇总
        const stats = this._calcStats(data);
        wsData.push([]);
        wsData.push(['汇总统计', '']);
        for (const [k, v] of Object.entries(stats)) {
            wsData.push([k, '', '', v]);
        }

        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // 设置列宽
        ws['!cols'] = [
            { wch: 18 }, { wch: 14 }, { wch: 22 }, { wch: 16 },
            ...wowMetrics.map(() => ({ wch: 14 })),
        ];

        // 合并标题单元格
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
        ];

        XLSX.utils.book_append_sheet(wb, ws, `周报 ${dateRange}`);
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `亚马逊周报_${dateRange}.xlsx`.replace(/\./g, '_');
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    _calcStats(data) {
        const sum = key => data.reduce((a, b) => a + (b[key] || 0), 0);
        const mean = key => data.length ? sum(key) / data.length : 0;
        return {
            '总ASIN数': data.length,
            '7天总销售额': sum('7天销售额'),
            '7天总销量': sum('7天销量'),
            '30天总销量': sum('30天销量'),
            '总广告费用': sum('广告费用'),
            '平均Tacos': mean('Tacos'),
            '平均广告转化率': mean('广告转化率'),
            '平均自然转化率': mean('自然转化率'),
            '平均CPA': mean('CPA'),
        };
    }
}
