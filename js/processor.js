/**
 * 数据处理核心模块
 * 计算周报所需所有指标、环比、广告转化率、自然转化率、CPA
 */
class DataProcessor {
    constructor() {
        this.resultData = null;
        this.dateRange = '';
    }

    calculateMetrics(data) {
        return data.map(row => {
            const r = { ...row };

            // TACOS = 广告费用 / 7天销售额
            if (!r.Tacos || r.Tacos === 0) {
                r.Tacos = r['7天销售额'] > 0 ? (r['广告费用'] / r['7天销售额']) * 100 : 0;
            }

            // ACOS = 广告费用 / 广告销售额
            if (!r.ACOS || r.ACOS === 0) {
                r.ACOS = r['广告销售额'] > 0 ? (r['广告费用'] / r['广告销售额']) * 100 : 0;
            }

            // 广告转化率 = 广告订单 / 广告点击
            r['广告转化率'] = r['广告点击'] > 0 ? (r['广告订单'] / r['广告点击']) * 100 : 0;

            // 自然转化率 = 自然订单 / 自然流量
            r['自然转化率'] = r['自然流量'] > 0 ? (r['自然订单'] / r['自然流量']) * 100 : 0;
            if (r['自然转化率'] === 0 && r['总流量'] > 0) {
                r['自然转化率'] = (r['自然订单'] / r['总流量']) * 100;
            }

            // CPA = 广告费用 / 广告订单
            r.CPA = r['广告订单'] > 0 ? r['广告费用'] / r['广告订单'] : 0;

            // 30天退款率
            if (!r['30天退款率'] || r['30天退款率'] === 0) {
                r['30天退款率'] = r['30天销量'] > 0 ? (r['退款量'] / r['30天销量']) * 100 : 0;
            }

            // 利润额
            if ((!r['利润'] || r['利润'] === 0) && r['利润率']) {
                r['利润'] = r['7天销售额'] * (r['利润率'] / 100);
            }

            // 毛利率估算
            if (!r['毛利率']) {
                r['毛利率'] = r['7天销售额'] > 0 ? ((r['7天销售额'] - r['广告费用']) / r['7天销售额']) * 100 : 0;
            }

            return r;
        });
    }

    calculateWeekOverWeek(currentData, lastData) {
        if (!lastData || !lastData.length) {
            return currentData.map(r => ({
                ...r,
                '7天销售额_环比': 0,
                '7天销量_环比': 0,
                '30天销量_环比': 0,
                '广告费用_环比': 0,
                'Tacos_环比': 0,
                '广告转化率_环比': 0,
                '自然转化率_环比': 0,
                'CPA_环比': 0,
            }));
        }

        const lastByAsin = {};
        for (const row of lastData) {
            lastByAsin[row.ASIN] = row;
        }

        return currentData.map(r => {
            const last = lastByAsin[r.ASIN];
            const calc = (cur, prev) => {
                if (prev > 0) return ((cur - prev) / prev) * 100;
                return cur > 0 ? 100 : 0;
            };

            return {
                ...r,
                '7天销售额_环比': last ? calc(r['7天销售额'], last['7天销售额']) : 0,
                '7天销量_环比': last ? calc(r['7天销量'], last['7天销量']) : 0,
                '30天销量_环比': last ? calc(r['30天销量'], last['30天销量']) : 0,
                '广告费用_环比': last ? calc(r['广告费用'], last['广告费用']) : 0,
                'Tacos_环比': last ? calc(r.Tacos, last.Tacos) : 0,
                '广告转化率_环比': last ? calc(r['广告转化率'], last['广告转化率']) : 0,
                '自然转化率_环比': last ? calc(r['自然转化率'], last['自然转化率']) : 0,
                'CPA_环比': last ? calc(r.CPA, last.CPA) : 0,
            };
        });
    }

    process(salesFile, profitFile, lastSalesFile, lastProfitFile, dateRange) {
        const parser = new ReportParser();
        this.dateRange = dateRange || '';

        return Promise.all([
            parser.parseSalesReport(salesFile),
            profitFile ? parser.parseProfitReport(profitFile) : Promise.resolve(null),
            lastSalesFile ? parser.parseSalesReport(lastSalesFile) : Promise.resolve(null),
            lastProfitFile ? parser.parseProfitReport(lastProfitFile) : Promise.resolve(null),
        ]).then(([sales, profit, lastSales, lastProfit]) => {
            const current = parser.mergeReports(sales, profit);
            let currentCalc = this.calculateMetrics(current);

            let lastCalc = null;
            if (lastSales) {
                const lastMerged = parser.mergeReports(lastSales, lastProfit);
                lastCalc = this.calculateMetrics(lastMerged);
            }

            this.resultData = this.calculateWeekOverWeek(currentCalc, lastCalc);
            return this.resultData;
        });
    }

    getSummaryStats() {
        if (!this.resultData || !this.resultData.length) return {};
        const df = this.resultData;
        const sum = key => df.reduce((a, b) => a + (b[key] || 0), 0);
        const mean = key => df.reduce((a, b) => a + (b[key] || 0), 0) / df.length;

        return {
            '总ASIN数': df.length,
            '总7天销售额': sum('7天销售额'),
            '总7天销量': sum('7天销量'),
            '总30天销量': sum('30天销量'),
            '总广告费用': sum('广告费用'),
            '平均Tacos': mean('Tacos'),
            '总广告订单': sum('广告订单'),
            '总自然订单': sum('自然订单'),
            '平均广告转化率': mean('广告转化率'),
            '平均自然转化率': mean('自然转化率'),
            '平均CPA': mean('CPA'),
        };
    }

    getDefaultDateRange(type) {
        const today = new Date();
        if (type === 'monthly') {
            const firstDay = new Date(today.getFullYear(), today.getMonth() - 1, 1);
            return firstDay.toISOString().slice(0, 7);
        }
        const day = today.getDay();
        const lastMonday = new Date(today);
        lastMonday.setDate(today.getDate() - day - 7);
        const lastSunday = new Date(lastMonday);
        lastSunday.setDate(lastMonday.getDate() + 6);
        return `${lastMonday.getMonth() + 1}.${lastMonday.getDate()}-${lastSunday.getMonth() + 1}.${lastSunday.getDate()}`;
    }
}
