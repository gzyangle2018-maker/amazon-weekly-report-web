/**
 * 报表解析器 - 智能解析ERP销量表和利润报表
 * 支持优卖云、赛狐ERP及通用Excel格式
 */
class ReportParser {
    constructor() {
        this.COLUMN_MAPPINGS = {
            asin: ['ASIN', 'asin', 'Asin', '产品ASIN', '商品ASIN', '子ASIN', 'Child ASIN'],
            store: ['店铺', 'Store', 'store', '店铺名', '店铺名称', 'Seller', '卖家', '店铺站点'],
            sales_7d: ['7天销量', '近7天销量', '7天订单', '近7天订单', 'Last 7 Days Sales', '7 Days Sales', '销量'],
            sales_30d: ['30天销量', '近30天销量', '30天订单', '近30天订单', 'Last 30 Days Sales', '30 Days Sales'],
            units_7d: ['7天销量', '近7天销量', '7天Units', 'Units 7d'],
            units_30d: ['30天销量', '近30天销量', '30天Units', 'Units 30d'],
            revenue_7d: ['7天销售额', '近7天销售额', '7天收入', 'Last 7 Days Revenue', '7 Days Revenue', '销售额7天'],
            revenue_30d: ['30天销售额', '近30天销售额', '30天收入', 'Last 30 Days Revenue', '30 Days Revenue'],
            avg_price: ['均价', '平均单价', '平均价格', 'Avg Price', 'Average Price', 'ERP均价', '均价（ERP）'],
            ad_spend: ['广告费用', '广告花费', '广告支出', 'Ad Spend', 'Spend', 'spend', '广告费', '广告花费'],
            ad_spend_sp: ['SP广告费', 'SP花费', 'SP Spend', 'Sponsored Products Spend'],
            ad_spend_sb: ['SB广告费', 'SB花费', 'SB Spend', 'Sponsored Brands Spend'],
            ad_spend_sd: ['SD广告费', 'SD花费', 'SD Spend', 'Sponsored Display Spend'],
            ad_sales: ['广告销售额', 'Ad Sales', '广告收入', '广告单销售额'],
            ad_orders: ['广告订单', 'Ad Orders', '广告单量', '广告订单数'],
            ad_clicks: ['广告点击', 'Ad Clicks', 'Clicks', '点击量', 'clicks'],
            ad_impressions: ['广告曝光', 'Ad Impressions', 'Impressions', '曝光量', 'impressions'],
            acos: ['ACOS', 'acos', 'Acos'],
            tacos: ['TACOS', 'tacos', 'Tacos', 'tacos(%)', 'TACOS(%)'],
            refund_30d: ['30天退款', '近30天退款', '30天退货', 'Refunds 30d', '退货数', '退款量'],
            refund_rate_30d: ['30天退款率', '退款率30天', '退货率', 'Refund Rate', 'refund_rate'],
            refund_rate_90d: ['90天退款率', '退款率90天', '90天退货率', 'Refund Rate 90d', '90天退款率'],
            profit: ['利润', '净利润', 'Profit', 'Net Profit', '毛利'],
            profit_margin: ['利润率', '利润比例', 'Profit Margin'],
            roi: ['ROI', 'roi'],
            organic_sales: ['自然销售额', 'Organic Sales', '自然单销售额'],
            organic_orders: ['自然订单', 'Organic Orders', '自然单量', '自然订单数'],
            organic_sessions: ['自然流量', '自然会话', 'Organic Sessions', 'Sessions'],
            total_sessions: ['总流量', '总会话', 'Total Sessions', 'Sessions'],
            total_orders: ['总订单', '订单总数', 'Total Orders', '订单量'],
            total_sales: ['总销售额', 'Total Sales', '总收入', '销售额'],
            store_site: ['店铺站点', '站点', '店铺-站点', 'Store-Site'],
            fba_fee: ['FBA费用', 'FBA Fee', 'FBA运费'],
            commission: ['亚马逊佣金', 'Commission', '佣金', '平台佣金'],
            storage_fee: ['仓储费', 'Storage Fee', '仓库费用'],
            procurement_cost: ['设置商品采购成本', '采购成本', '商品成本'],
            logistics_cost: ['设置商品物流成本', '物流成本', '头程运费'],
            refund_amount: ['退款金额', 'Refund Amount', '退货金额'],
            amazon_compensation: ['亚马逊赔偿', 'Compensation', '赔偿'],
        };
    }

    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    // 选择数据最多的sheet
                    let bestSheet = null;
                    let maxRows = 0;
                    for (const sheetName of workbook.SheetNames) {
                        const sheet = workbook.Sheets[sheetName];
                        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                        if (json.length > maxRows) {
                            maxRows = json.length;
                            bestSheet = json;
                        }
                    }
                    if (!bestSheet || bestSheet.length < 2) {
                        reject(new Error('文件中没有有效数据'));
                        return;
                    }
                    // 第一行是header
                    const headers = bestSheet[0].map(h => String(h || '').trim());
                    const rows = bestSheet.slice(1).map(row => {
                        const obj = {};
                        headers.forEach((h, i) => {
                            obj[h] = row[i];
                        });
                        return obj;
                    });
                    resolve({ headers, rows, raw: bestSheet });
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    findColumn(headers, field) {
        const candidates = this.COLUMN_MAPPINGS[field] || [field];
        for (const col of headers) {
            const colStr = String(col).trim();
            if (candidates.includes(colStr)) return colStr;
            for (const c of candidates) {
                if (c.toLowerCase() === colStr.toLowerCase()) return colStr;
                if (colStr.toLowerCase().includes(c.toLowerCase())) return colStr;
                if (c.toLowerCase().includes(colStr.toLowerCase())) return colStr;
            }
        }
        return null;
    }

    cleanNumeric(val) {
        if (val === null || val === undefined || val === '') return 0;
        if (typeof val === 'number') return val;
        const str = String(val)
            .replace(/[$￥,£€%\s]/g, '')
            .replace(/US\$/, '')
            .replace(/nan|NaN|None|null|-/gi, '');
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    extractColumn(rows, headers, field, defaultVal = 0) {
        const col = this.findColumn(headers, field);
        if (!col) return rows.map(() => defaultVal);
        return rows.map(r => this.cleanNumeric(r[col]) || defaultVal);
    }

    extractStringColumn(rows, headers, field, defaultVal = '') {
        const col = this.findColumn(headers, field);
        if (!col) return rows.map(() => defaultVal);
        return rows.map(r => {
            const v = r[col];
            return v !== undefined && v !== null ? String(v).trim() : defaultVal;
        });
    }

    parseSalesReport(file) {
        return this.readFile(file).then(({ headers, rows }) => {
            const result = [];

            const asinCol = this.findColumn(headers, 'asin');
            if (!asinCol) throw new Error('无法识别ASIN列，请检查表格格式');

            const asins = this.extractStringColumn(rows, headers, 'asin');
            const stores = this.extractStringColumn(rows, headers, 'store', '默认店铺');
            const sales7d = this.extractColumn(rows, headers, 'sales_7d', 0);
            const sales30d = this.extractColumn(rows, headers, 'sales_30d', 0);
            const units7d = this.extractColumn(rows, headers, 'units_7d', 0);
            const units30d = this.extractColumn(rows, headers, 'units_30d', 0);
            const revenue7d = this.extractColumn(rows, headers, 'revenue_7d', 0);
            const avgPrice = this.extractColumn(rows, headers, 'avg_price', 0);
            const totalOrders = this.extractColumn(rows, headers, 'total_orders', 0);
            const adOrders = this.extractColumn(rows, headers, 'ad_orders', 0);
            const adClicks = this.extractColumn(rows, headers, 'ad_clicks', 0);
            const organicSessions = this.extractColumn(rows, headers, 'organic_sessions', 0);
            const totalSessions = this.extractColumn(rows, headers, 'total_sessions', 0);
            const adImpressions = this.extractColumn(rows, headers, 'ad_impressions', 0);

            for (let i = 0; i < rows.length; i++) {
                const qty7d = sales7d[i] || units7d[i];
                const qty30d = sales30d[i] || units30d[i];
                const orders = totalOrders[i] || qty7d;
                const adOrd = adOrders[i];
                const organicOrd = Math.max(0, orders - adOrd);
                const organicSess = organicSessions[i] || Math.max(0, (totalSessions[i] || 0) - adClicks[i]);

                let price = avgPrice[i];
                if (!price && qty7d > 0) price = revenue7d[i] / qty7d;

                result.push({
                    ASIN: asins[i].toUpperCase(),
                    店铺: stores[i],
                    '7天销量': qty7d,
                    '30天销量': qty30d,
                    '7天销售额': revenue7d[i],
                    '均价': price,
                    '总订单': orders,
                    '广告订单': adOrd,
                    '自然订单': organicOrd,
                    '广告点击': adClicks[i],
                    '自然流量': organicSess,
                    '总流量': totalSessions[i] || adClicks[i] + organicSess,
                    '广告曝光': adImpressions[i],
                });
            }
            return result;
        });
    }

    parseProfitReport(file) {
        return this.readFile(file).then(({ headers, rows }) => {
            const result = [];

            const asinCol = this.findColumn(headers, 'asin');
            const storeSiteCol = this.findColumn(headers, 'store_site');
            const storeCol = this.findColumn(headers, 'store');

            let keyType = 'ASIN';
            let keys = [];
            if (asinCol) {
                keys = this.extractStringColumn(rows, headers, 'asin');
            } else if (storeSiteCol || storeCol) {
                keyType = '店铺';
                keys = this.extractStringColumn(rows, headers, storeSiteCol ? 'store_site' : 'store');
                keys = keys.map(k => k.split('-')[0].trim());
            } else {
                throw new Error('无法识别利润报表格式：既无ASIN列也无店铺站点列');
            }

            const adSpend = this.extractColumn(rows, headers, 'ad_spend', 0);
            const adSpendSp = this.extractColumn(rows, headers, 'ad_spend_sp', 0);
            const adSpendSb = this.extractColumn(rows, headers, 'ad_spend_sb', 0);
            const adSpendSd = this.extractColumn(rows, headers, 'ad_spend_sd', 0);
            const adSales = this.extractColumn(rows, headers, 'ad_sales', 0);
            const profit = this.extractColumn(rows, headers, 'profit', 0);
            const profitMargin = this.extractColumn(rows, headers, 'profit_margin', 0);
            const roi = this.extractColumn(rows, headers, 'roi', 0);
            const refund30d = this.extractColumn(rows, headers, 'refund_30d', 0);
            const refundRate30d = this.extractColumn(rows, headers, 'refund_rate_30d', 0);
            const refundRate90d = this.extractColumn(rows, headers, 'refund_rate_90d', 0);
            const refundAmount = this.extractColumn(rows, headers, 'refund_amount', 0);
            const tacos = this.extractColumn(rows, headers, 'tacos', 0);
            const acos = this.extractColumn(rows, headers, 'acos', 0);
            const fbaFee = this.extractColumn(rows, headers, 'fba_fee', 0);
            const commission = this.extractColumn(rows, headers, 'commission', 0);
            const storageFee = this.extractColumn(rows, headers, 'storage_fee', 0);
            const procurementCost = this.extractColumn(rows, headers, 'procurement_cost', 0);
            const logisticsCost = this.extractColumn(rows, headers, 'logistics_cost', 0);
            const amazonCompensation = this.extractColumn(rows, headers, 'amazon_compensation', 0);

            for (let i = 0; i < rows.length; i++) {
                let spend = adSpend[i];
                if (!spend) spend = (adSpendSp[i] || 0) + (adSpendSb[i] || 0) + (adSpendSd[i] || 0);

                result.push({
                    [keyType]: keys[i].toUpperCase(),
                    '广告费用': spend,
                    '广告销售额': adSales[i],
                    '利润': profit[i],
                    '利润率': profitMargin[i],
                    'ROI': roi[i],
                    '退款量': refund30d[i],
                    '退货率': refundRate30d[i],
                    '90天退款率': refundRate90d[i],
                    '退款金额': refundAmount[i],
                    'Tacos': tacos[i],
                    'ACOS': acos[i],
                    'FBA费用': fbaFee[i],
                    '亚马逊佣金': commission[i],
                    '仓储费': storageFee[i],
                    '采购成本': procurementCost[i],
                    '物流成本': logisticsCost[i],
                    '亚马逊赔偿': amazonCompensation[i],
                });
            }
            return result;
        });
    }

    mergeReports(salesData, profitData) {
        if (!salesData || !salesData.length) throw new Error('销量报表为空');
        const merged = salesData.map(s => ({ ...s }));
        if (!profitData || !profitData.length) return merged;

        const profitByKey = {};
        const hasAsin = profitData[0].hasOwnProperty('ASIN');
        for (const p of profitData) {
            const key = hasAsin ? p.ASIN : p.店铺;
            profitByKey[key] = p;
        }

        for (const m of merged) {
            const key = hasAsin ? m.ASIN : m.店铺;
            const p = profitByKey[key];
            if (p) {
                Object.keys(p).forEach(k => {
                    if (k !== 'ASIN' && k !== '店铺') {
                        m[k] = p[k];
                    }
                });
            }
        }

        const numericCols = ['广告费用','退款量','Tacos','ACOS','广告销售额','利润',
            '退款金额','FBA费用','亚马逊佣金','仓储费','采购成本','物流成本','亚马逊赔偿'];
        for (const row of merged) {
            for (const col of numericCols) {
                if (row[col] === undefined || row[col] === null || row[col] === '') {
                    row[col] = 0;
                }
            }
        }
        return merged;
    }
}
