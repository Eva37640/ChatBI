"""
ChatBI 智能商业分析系统 - 终极版
支持28种查询意图，包括跨表JOIN、聚合运算、时间分析等
"""
from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
from difflib import get_close_matches

app = Flask(__name__)

# ========== 全局数据 ==========
orders_df = None
details_df = None
logistics_df = None
all_customers = []
all_products = []
all_operators = []
all_drivers = []
all_warehouses = []

def load_data():
    """加载所有数据"""
    global orders_df, details_df, logistics_df
    global all_customers, all_products, all_operators, all_drivers, all_warehouses
    
    print("=" * 70)
    print("ChatBI 终极版 - 正在加载数据...")
    print("=" * 70)
    
    # 使用相对路径
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    data_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')] if os.path.exists(data_dir) else []
    
    if not data_files:
        print(f"⚠️ 数据目录 {data_dir} 中没有找到Excel文件")
        return
    
    orders_list, details_list, logistics_list = [], [], []
    
    for fname in sorted(data_files):
        file_path = os.path.join(data_dir, fname)
        print(f"\n📁 加载: {fname}")
        
        try:
            df_o = pd.read_excel(file_path, sheet_name='订单表')
            df_o = df_o.rename(columns={
                '订单单号': 'order_id', '订单类型': 'order_type',
                '货主编码': 'customer_code', '货主': 'customer_name',
                '仓库': 'warehouse', '收货门店': 'store', '省市区': 'region',
                '求和项:预计发货数量EA': 'quantity_ea', '预计总箱数': 'box_count',
                '创建人': 'creator', '创建时间': 'create_time'
            }, errors='ignore')
            df_o['create_time'] = pd.to_datetime(df_o['create_time'], errors='coerce')
            df_o['month'] = df_o['create_time'].dt.month
            df_o['year'] = df_o['create_time'].dt.year
            df_o['date'] = df_o['create_time'].dt.date
            df_o['week'] = df_o['create_time'].dt.isocalendar().week.astype(int)
            df_o['weekday'] = df_o['create_time'].dt.weekday
            df_o['hour'] = df_o['create_time'].dt.hour
            df_o['province'] = df_o['region'].apply(lambda x: str(x).split('-')[0] if pd.notna(x) else '未知')
            orders_list.append(df_o)
            print(f"   ✓ 订单: {len(df_o)} 条")
            
            df_d = pd.read_excel(file_path, sheet_name='订单明细')
            col_map = {}
            for col in df_d.columns:
                if '预计发货数量' in col and 'EA' in col.upper():
                    col_map[col] = 'quantity_ea'
                elif '预计发货数量' in col:
                    col_map[col] = 'quantity'
            col_map.update({
                '订单单号': 'order_id', '商品编码': 'product_code',
                '商品名称': 'product_name', '温区': 'temp_zone', '单位': 'unit'
            })
            df_d = df_d.rename(columns=col_map, errors='ignore')
            df_d['category'] = df_d['product_name'].apply(extract_category)
            df_d['is_frozen'] = df_d['temp_zone'].apply(lambda x: '冷冻' in str(x) if pd.notna(x) else False)
            df_d['is_cold'] = df_d['temp_zone'].apply(lambda x: '冷藏' in str(x) if pd.notna(x) else False)
            df_d['is_normal'] = df_d['temp_zone'].apply(lambda x: '常温' in str(x) if pd.notna(x) else False)
            details_list.append(df_d)
            print(f"   ✓ 明细: {len(df_d)} 条")
            
            df_l = pd.read_excel(file_path, sheet_name='物流信息')
            df_l = df_l.rename(columns={
                '订单号': 'order_id', '操作时间': 'operation_time',
                '操作记录': 'operation_record', '操作人': 'operator',
                '司机': 'driver', '车牌': 'vehicle', '状态': 'status'
            }, errors='ignore')
            df_l['operation_time'] = pd.to_datetime(df_l['operation_time'], errors='coerce')
            df_l['op_type'] = df_l['operation_record'].apply(extract_op_type)
            df_l['op_date'] = df_l['operation_time'].dt.date
            df_l['op_month'] = df_l['operation_time'].dt.month
            logistics_list.append(df_l)
            print(f"   ✓ 物流: {len(df_l)} 条")
            
        except Exception as e:
            print(f"   ✗ 错误: {e}")
    
    if orders_list:
        orders_df = pd.concat(orders_list, ignore_index=True)
        all_customers = orders_df['customer_name'].dropna().unique().tolist()
        all_warehouses = orders_df['warehouse'].dropna().unique().tolist()
    
    if details_list:
        details_df = pd.concat(details_list, ignore_index=True)
        all_products = details_df['product_name'].dropna().unique().tolist()
    
    if logistics_list:
        logistics_df = pd.concat(logistics_list, ignore_index=True)
        all_operators = logistics_df['operator'].dropna().unique().tolist()
        if 'driver' in logistics_df.columns and logistics_df['driver'].notna().any():
            all_drivers = logistics_df['driver'].dropna().unique().tolist()
        else:
            driver_records = logistics_df[logistics_df['operation_record'].str.contains('司机', na=False)]
            if len(driver_records) > 0:
                driver_names = driver_records['operation_record'].str.extract(r'司机[:：\s]*([^\s,，。]+)')[0].dropna().unique().tolist()
                all_drivers = [n for n in driver_names if len(n) >= 2]
            else:
                all_drivers = []
            logistics_df['driver'] = np.nan
    
    print("\n" + "=" * 70)
    print(f"✅ 数据加载完成!")
    print(f"   订单: {len(orders_df) if orders_df is not None else 0:,}")
    print(f"   明细: {len(details_df) if details_df is not None else 0:,}")
    print(f"   物流: {len(logistics_df) if logistics_df is not None else 0:,}")
    print(f"   客户: {len(all_customers)} | 商品: {len(all_products)} | 司机: {len(all_drivers)}")
    print("=" * 70)

def extract_category(name):
    if pd.isna(name): return '其他'
    name = str(name)
    cats = {'鸡腿': '禽肉', '鸡翅': '禽肉', '鸡肉': '禽肉', '牛肉': '牛羊肉', '羊肉': '牛羊肉',
            '猪肉': '猪肉', '底料': '调料', '酱料': '调料', '青豆': '蔬菜', '玉米': '蔬菜',
            '碗': '餐具', '杯': '餐具', '鱼': '水产', '虾': '水产', 'iPhone': '电子产品'}
    for k, v in cats.items():
        if k in name: return v
    return '其他'

def extract_op_type(record):
    if pd.isna(record): return '其他'
    r = str(record)
    if '创建' in r: return '创建订单'
    elif '审核' in r: return '审核'
    elif '拣货' in r: return '拣货'
    elif '出库' in r or '发货' in r: return '出库发货'
    elif '运输' in r or '配送' in r: return '运输中'
    elif '签收' in r or '送达' in r: return '签收'
    elif '司机' in r: return '司机相关'
    else: return '其他'

# ========== 智能查询引擎 ==========
class ChatBIEngine:
    def __init__(self, query):
        self.query = query.strip()
        self.thinking = []
    
    def execute(self):
        self.thinking.append(f"📝 查询: {self.query}")
        intent = self._identify_intent()
        self.thinking.append(f"🔍 意图: {intent['desc']}")
        handler = intent['handler']
        return handler()
    
    def _identify_intent(self):
        """识别查询意图"""
        q = self.query
        ql = q.lower()
        
        # 1. 司机相关
        if '司机' in q:
            return {'desc': '司机相关查询', 'handler': self._handle_driver}
        
        # 2. 订单详情
        order_match = re.search(r'([A-Z]{2,}\d+)', q)
        if order_match and ('订单' in q or '操作' in q or '进度' in q):
            return {'desc': '订单详情查询', 'handler': self._handle_order_detail}
        
        # 3. 跨表对比
        if any(w in q for w in ['对比', '比较', 'vs']) and any(w in q for w in ['订单', '操作', '发货']):
            return {'desc': '跨表对比分析', 'handler': self._handle_cross_table}
        
        # 4. 温区销量变化
        if any(w in q for w in ['温区', '常温', '冷藏', '冷冻']) and any(w in q for w in ['变化', '对比', '环比', '增长']):
            return {'desc': '温区销量变化分析', 'handler': self._handle_temp_zone_change}
        
        # 5. 商品只在某温区
        if '只' in q and '温区' in q:
            return {'desc': '温区唯一性分析', 'handler': self._handle_unique_temp_zone}
        
        # 6. 冷冻商品
        if '冷冻' in q and not any(w in q for w in ['变化', '对比', '环比']):
            return {'desc': '冷冻商品分析', 'handler': self._handle_frozen}
        
        # 7. 条件筛选聚合
        if any(w in q for w in ['超过', '大于', '小于', '多于', '少于']) and any(w in q for w in ['条', '个', 'EA', '件']):
            return {'desc': '条件聚合查询', 'handler': self._handle_filter_agg}
        
        # 8. 客户排名
        if '客户' in q and any(w in q for w in ['排名', '前', 'top', '最多', '最高']):
            return {'desc': '客户排名统计', 'handler': self._handle_customer_rank}
        
        # 9. 订单处理时长
        if any(w in q for w in ['处理时长', '处理时间', '耗时最长', '时长最长', '创建到签收']):
            return {'desc': '订单处理时长分析', 'handler': self._handle_process_time}
        
        # 10. 商品销量排名
        if ('商品' in q or '产品' in q) and any(w in q for w in ['最高', '前', 'top', '排名', '最多']):
            return {'desc': '商品销量排名', 'handler': self._handle_product_rank}
        
        # 11. 环比增长率
        if any(w in q for w in ['环比', '增长率', '增长趋势']):
            return {'desc': '环比增长率计算', 'handler': self._handle_mom_growth}
        
        # 12. 订单处理效率
        if '效率' in q:
            return {'desc': '订单处理效率分析', 'handler': self._handle_efficiency}
        
        # 13. 上周订单
        if '上周' in q:
            return {'desc': '上周订单统计', 'handler': self._handle_last_week}
        
        # 14. 签收率
        if '签收率' in q:
            return {'desc': '签收率分析', 'handler': self._handle_sign_rate}
        
        # 15. 操作人跨仓库
        if '操作人' in q and any(w in q for w in ['多个仓库', '跨仓库', '不同仓库']):
            return {'desc': '操作人跨仓库分析', 'handler': self._handle_operator_cross_warehouse}
        
        # 16. 无操作订单
        if any(w in q for w in ['没有操作', '无操作', '未操作', '一直没有操作']):
            return {'desc': '无操作订单查询', 'handler': self._handle_no_op_orders}
        
        # 17. 操作人处理多客户
        if '操作人' in q and any(w in q for w in ['超过', '多于', '100', '不同客户']):
            return {'desc': '操作人客户数分析', 'handler': self._handle_operator_customers}
        
        # 18. 冷冻+数量筛选
        if '冷冻' in q and any(w in q for w in ['超过', '大于', '500']):
            return {'desc': '冷冻商品条件筛选', 'handler': self._handle_frozen_filter}
        
        # 19. 月份对比
        if any(w in q for w in ['2月', '3月', '1月']) and any(w in q for w in ['表现', '对比', '变化']):
            return {'desc': '月份对比分析', 'handler': self._handle_month_compare}
        
        # 20. 明细行数筛选
        if '明细' in q and any(w in q for w in ['超过', '大于', '多于']):
            return {'desc': '明细行数筛选', 'handler': self._handle_detail_count_filter}
        
        # 21. 客户平均操作数
        if '客户' in q and any(w in q for w in ['平均操作', '操作最多', '操作最高']):
            return {'desc': '客户平均操作数排名', 'handler': self._handle_customer_avg_ops}
        
        # 22. 品类月度增长
        if '品类' in q and any(w in q for w in ['增长最快', '增长', '变化']):
            return {'desc': '品类月度增长分析', 'handler': self._handle_category_monthly_growth}
        
        # 23. 每周哪天订单最多
        if any(w in q for w in ['每周', '星期', '周几', '哪天']) and any(w in q for w in ['最多', '订单量']):
            return {'desc': '每周订单分析', 'handler': self._handle_weekday_analysis}
        
        # 24. 时段分析
        if any(w in q for w in ['时段', '上午', '下午', '晚上', '几点']) and any(w in q for w in ['最多', '订单']):
            return {'desc': '时段订单分析', 'handler': self._handle_hour_analysis}
        
        # 25. 趋势分析
        if '趋势' in q and any(w in q for w in ['下单', '订单', '变化']):
            return {'desc': '趋势分析', 'handler': self._handle_trend}
        
        # 26. 温区查询
        if '温区' in q:
            return {'desc': '温区查询', 'handler': self._handle_temp_zone}
        
        # 27. 商品查询
        if any(w in ql for w in ['商品', '产品', 'iphone']):
            return {'desc': '商品查询', 'handler': self._handle_product}
        
        return {'desc': '未知查询', 'handler': self._handle_unknown}
    
    def _extract_driver_name(self, query):
        for name in all_drivers:
            if name and len(name) >= 2 and name in query:
                return name
        for name in all_operators:
            if name and len(name) >= 2 and name in query:
                return name
        match = re.search(r'司机[的负责]*([^\s，。？?!的负责]{2,4})', query)
        if match:
            search_name = match.group(1)
            for w in ['负责', '订单', '进度', '怎么', '样了']:
                search_name = search_name.replace(w, '')
            if len(search_name) >= 2:
                all_names = all_drivers + all_operators
                close = get_close_matches(search_name, all_names, n=3, cutoff=0.4)
                if close:
                    return close[0]
                return search_name
        return None
    
    # ========== 处理函数 ==========
    
    def _handle_driver(self):
        html = '<h3>🚚 司机订单查询</h3>'
        driver_name = self._extract_driver_name(self.query)
        
        if not driver_name:
            html += '<p>请指定司机姓名。</p>'
            if all_drivers:
                html += '<h4>可用司机:</h4><p>' + ', '.join(all_drivers[:20]) + '</p>'
            html += '<h4>操作人列表（部分）:</h4><p>' + ', '.join(all_operators[:20]) + '</p>'
            return {'success': True, 'html': html, 'thinking': self.thinking}
        
        self.thinking.append(f"🚚 查询司机: {driver_name}")
        
        logs = logistics_df[
            (logistics_df['driver'].fillna('').str.contains(driver_name, na=False)) |
            (logistics_df['operator'].fillna('').str.contains(driver_name, na=False))
        ]
        
        if len(logs) == 0:
            all_names = list(set(all_drivers + all_operators))
            close = get_close_matches(driver_name, all_names, n=5, cutoff=0.4)
            if close:
                html += f'<p>未找到"{driver_name}"，您是否想找：</p><ul>'
                for c in close:
                    html += f'<li>{c}</li>'
                html += '</ul>'
            else:
                html += f'<p>未找到"{driver_name}"的相关记录</p>'
            return {'success': True, 'html': html, 'thinking': self.thinking}
        
        order_ids = logs['order_id'].unique()
        html += f'<p>司机/操作人 <b>{driver_name}</b> 涉及 <b>{len(order_ids)}</b> 个订单，<b>{len(logs)}</b> 条操作记录</p>'
        
        op_types = logs['op_type'].value_counts()
        html += '<h4>📊 操作类型分布</h4>'
        html += self._df_to_html(op_types.reset_index().rename(columns={'index': '操作类型', 'op_type': '次数'}))
        
        html += '<h4>📋 订单列表及进度（前20个）</h4>'
        progress = []
        for oid in order_ids[:20]:
            oid_logs = logs[logs['order_id'] == oid].sort_values('operation_time')
            latest_op = oid_logs.iloc[-1]['op_type'] if len(oid_logs) > 0 else '未知'
            order_info = orders_df[orders_df['order_id'] == oid]
            customer = order_info.iloc[0]['customer_name'] if len(order_info) > 0 else '未知'
            qty = order_info.iloc[0]['quantity_ea'] if len(order_info) > 0 else 0
            progress.append({
                '订单编号': oid, '客户': customer, '预计数量(EA)': qty,
                '最新状态': latest_op, '操作次数': len(oid_logs)
            })
        html += self._df_to_html(pd.DataFrame(progress))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_cross_table(self):
        html = '<h3>📊 订单预计发货数量 vs 操作记录数量对比</h3>'
        self.thinking.append("📥 提取订单数据 + 物流统计...")
        
        orders = orders_df[['order_id', 'customer_name', 'quantity_ea', 'warehouse']].copy()
        op_stats = logistics_df.groupby('order_id').agg(
            op_count=('operation_time', 'count'),
            operators=('operator', lambda x: ', '.join(x.dropna().unique()))
        ).reset_index()
        
        merged = orders.merge(op_stats, on='order_id', how='left')
        merged['op_count'] = merged['op_count'].fillna(0).astype(int)
        merged['qty_per_op'] = np.where(merged['op_count'] > 0, (merged['quantity_ea'] / merged['op_count']).round(2), 0)
        
        html += self._stats_cards([
            ("总订单数", f"{len(merged):,}"),
            ("有操作订单", f"{(merged['op_count'] > 0).sum():,}"),
            ("平均操作数", f"{merged['op_count'].mean():.1f}"),
            ("最大操作数", f"{merged['op_count'].max()}")
        ])
        
        display = merged[['order_id', 'customer_name', 'warehouse', 'quantity_ea', 'op_count', 'qty_per_op']].head(50)
        display.columns = ['订单编号', '客户', '仓库', '预计发货数量(EA)', '操作记录数', '每操作处理量']
        html += '<h4>📋 对比详情（前50条）</h4>'
        html += self._df_to_html(display)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_filter_agg(self):
        html = '<h3>📋 条件聚合查询</h3>'
        
        match = re.search(r'(超过|大于|>|多于|少于|小于|<)(\d+)\s*(条|个|EA|件)?', self.query)
        if not match:
            return {'success': True, 'html': '<p>请指定筛选条件</p>', 'thinking': self.thinking}
        
        op_word = match.group(1)
        threshold = int(match.group(2))
        unit = match.group(3) or '条'
        
        is_greater = op_word in ['超过', '大于', '>', '多于']
        op_desc = '超过' if is_greater else '少于'
        
        self.thinking.append(f"📋 条件: 操作记录数 {'>' if is_greater else '<'} {threshold}")
        
        if unit in ['EA', '件', '个'] and '操作' not in self.query:
            filtered = orders_df[orders_df['quantity_ea'] > threshold] if is_greater else orders_df[orders_df['quantity_ea'] < threshold]
            html += f'<p>预计发货数量{op_desc}{threshold}{unit}的订单: <b>{len(filtered)}</b> 个</p>'
            display = filtered[['order_id', 'customer_name', 'quantity_ea', 'warehouse']].head(50)
            display.columns = ['订单编号', '客户', '预计发货数量(EA)', '仓库']
        else:
            op_counts = logistics_df.groupby('order_id').size().reset_index(name='op_count')
            filtered = op_counts[op_counts['op_count'] > threshold] if is_greater else op_counts[op_counts['op_count'] < threshold]
            result = filtered.merge(orders_df[['order_id', 'customer_name', 'quantity_ea', 'warehouse']], on='order_id', how='left')
            html += f'<p>操作记录{op_desc}{threshold}条的订单: <b>{len(result)}</b> 个</p>'
            display = result[['order_id', 'customer_name', 'warehouse', 'quantity_ea', 'op_count']].head(50)
            display.columns = ['订单编号', '客户', '仓库', '预计发货数量(EA)', '操作记录数']
        
        html += self._df_to_html(display)
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_temp_zone_change(self):
        html = '<h3>🌡️ 温区销量变化分析</h3>'
        self.thinking.append("📊 关联订单+明细，按温区和月份聚合...")
        
        merged = orders_df[['order_id', 'month']].merge(
            details_df[['order_id', 'temp_zone', 'quantity_ea']], on='order_id', how='inner')
        result = merged.groupby(['month', 'temp_zone'])['quantity_ea'].sum().reset_index()
        pivot = result.pivot(index='temp_zone', columns='month', values='quantity_ea').fillna(0)
        
        html += '<h4>温区-月份销量透视表</h4>'
        html += self._df_to_html(pivot.round(0))
        
        months = sorted(pivot.columns)
        if len(months) >= 2:
            html += '<h4>📈 相邻月份变化</h4>'
            html += '<table style="width:100%;border-collapse:collapse;">'
            html += '<tr style="background:#667eea;color:white;"><th>温区</th>'
            for i in range(len(months) - 1):
                html += f'<th>{months[i]}月→{months[i+1]}月 变化</th><th>变化率</th>'
            html += '</tr>'
            for tz in pivot.index:
                html += f'<tr><td>{tz}</td>'
                for i in range(len(months) - 1):
                    v1, v2 = pivot.loc[tz, months[i]], pivot.loc[tz, months[i+1]]
                    change = v2 - v1
                    pct = (change / v1 * 100) if v1 > 0 else 0
                    html += f'<td>{change:+,.0f}</td><td>{pct:+.1f}%</td>'
                html += '</tr>'
            html += '</table>'
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_frozen(self):
        html = '<h3>❄️ 冷冻商品销量分析</h3>'
        self.thinking.append("❄️ 筛选冷冻温区商品...")
        
        frozen = details_df[details_df['is_frozen'] == True]
        total_qty = frozen['quantity_ea'].sum()
        total_all = details_df['quantity_ea'].sum()
        pct = (total_qty / total_all * 100) if total_all > 0 else 0
        
        if '多不多' in self.query or '多少' in self.query:
            if pct > 30:
                judgment = f'冷冻商品销量占比 <b>{pct:.1f}%</b>，销量<b>较多</b>。'
            elif pct > 15:
                judgment = f'冷冻商品销量占比 <b>{pct:.1f}%</b>，销量<b>中等</b>。'
            else:
                judgment = f'冷冻商品销量占比 <b>{pct:.1f}%</b>，销量<b>较少</b>。'
            html += f'<p>{judgment}</p>'
        
        html += self._stats_cards([
            ("冷冻商品总量(EA)", f"{total_qty:,.0f}"),
            ("总占比", f"{pct:.1f}%"),
            ("涉及订单数", f"{frozen['order_id'].nunique()}"),
            ("商品种类", f"{frozen['product_name'].nunique()}")
        ])
        
        top = frozen.groupby('product_name')['quantity_ea'].sum().nlargest(10).reset_index()
        html += '<h4>🏆 TOP 10 冷冻商品</h4>'
        html += self._df_to_html(top)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_customer_rank(self):
        html = '<h3>🏆 客户订单排名</h3>'
        self.thinking.append("📊 按客户聚合...")
        
        stats = orders_df.groupby('customer_name').agg(
            订单数=('order_id', 'count'), 总数量=('quantity_ea', 'sum')
        ).reset_index().sort_values('总数量', ascending=False)
        stats['占比%'] = (stats['总数量'] / stats['总数量'].sum() * 100).round(2)
        
        n = 5
        match = re.search(r'前(\d+)', self.query)
        if match:
            n = int(match.group(1))
        
        html += f'<h4>TOP {n} 客户</h4>'
        html += self._df_to_html(stats.head(n))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_process_time(self):
        html = '<h3>⏱️ 订单处理时长分析</h3>'
        self.thinking.append("🔗 关联订单创建时间和物流操作时间...")
        
        order_create = orders_df[['order_id', 'create_time', 'customer_name']].copy()
        
        if '签收' in self.query:
            sign_times = logistics_df[logistics_df['op_type'] == '签收'].groupby('order_id')['operation_time'].min().reset_index()
            sign_times.columns = ['order_id', 'end_time']
            merged = order_create.merge(sign_times, on='order_id', how='inner')
            label = '签收'
        else:
            first_op = logistics_df.groupby('order_id')['operation_time'].min().reset_index()
            first_op.columns = ['order_id', 'end_time']
            merged = order_create.merge(first_op, on='order_id', how='inner')
            label = '第一次操作'
        
        merged['process_hours'] = (merged['end_time'] - merged['create_time']).dt.total_seconds() / 3600
        merged = merged[merged['process_hours'] >= 0]
        
        n = 5
        match = re.search(r'前(\d+)|(\d+)个', self.query)
        if match:
            n = int(match.group(1) or match.group(2))
        
        top = merged.nlargest(n, 'process_hours')
        display = top[['order_id', 'customer_name', 'create_time', 'end_time', 'process_hours']].copy()
        display['create_time'] = display['create_time'].dt.strftime('%Y-%m-%d %H:%M')
        display['end_time'] = display['end_time'].dt.strftime('%Y-%m-%d %H:%M')
        display['process_hours'] = display['process_hours'].round(2)
        display.columns = ['订单编号', '客户', '创建时间', f'{label}时间', '处理时长(小时)']
        
        html += f'<h4>处理时长最长的{n}个订单</h4>'
        html += self._df_to_html(display)
        html += f'<p>平均: {merged["process_hours"].mean():.1f}h | 中位数: {merged["process_hours"].median():.1f}h</p>'
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_product_rank(self):
        html = '<h3>📦 商品销量排名</h3>'
        
        month = None
        for m in range(1, 13):
            if f'{m}月' in self.query:
                month = m
                break
        
        self.thinking.append(f"📊 按商品聚合{'（筛选' + str(month) + '月）' if month else ''}...")
        
        if month:
            merged = orders_df[orders_df['month'] == month][['order_id']].merge(
                details_df[['order_id', 'product_name', 'quantity_ea']], on='order_id', how='inner')
        else:
            merged = details_df[['order_id', 'product_name', 'quantity_ea']]
        
        stats = merged.groupby('product_name')['quantity_ea'].sum().nlargest(10).reset_index()
        stats.columns = ['商品名称', '总数量(EA)']
        
        n = 3
        match = re.search(r'前(\d+)', self.query)
        if match:
            n = int(match.group(1))
        
        html += f'<h4>🏆 TOP {n} 商品</h4>'
        html += self._df_to_html(stats.head(n))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_mom_growth(self):
        html = '<h3>📈 环比增长率</h3>'
        self.thinking.append("📊 按月统计并计算环比...")
        
        monthly = orders_df.groupby('month').agg(
            订单数=('order_id', 'count'), 总数量=('quantity_ea', 'sum')
        ).reset_index().sort_values('month')
        monthly['订单环比%'] = monthly['订单数'].pct_change().mul(100).round(2).fillna(0)
        monthly['数量环比%'] = monthly['总数量'].pct_change().mul(100).round(2).fillna(0)
        
        html += self._df_to_html(monthly)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_order_detail(self):
        match = re.search(r'([A-Z]{2,}\d+)', self.query)
        if not match:
            return {'success': True, 'html': '<p>请提供订单编号</p>', 'thinking': self.thinking}
        
        order_id = match.group(1)
        html = f'<h3>📋 订单 {order_id} 详情</h3>'
        self.thinking.append(f"📋 查询订单: {order_id}")
        
        order = orders_df[orders_df['order_id'] == order_id]
        if len(order) == 0:
            return {'success': True, 'html': f'<p>未找到订单 {order_id}</p>', 'thinking': self.thinking}
        
        o = order.iloc[0]
        html += f'<div style="background:#f8f9fa;padding:15px;border-radius:8px;margin:15px 0;">'
        html += f'<p><b>客户:</b> {o.get("customer_name", "N/A")}</p>'
        html += f'<p><b>仓库:</b> {o.get("warehouse", "N/A")}</p>'
        html += f'<p><b>预计发货数量:</b> {o.get("quantity_ea", 0):,.0f} EA</p>'
        html += f'<p><b>创建时间:</b> {o.get("create_time", "N/A")}</p></div>'
        
        logs = logistics_df[logistics_df['order_id'] == order_id].sort_values('operation_time')
        html += f'<h4>📦 物流操作记录（共{len(logs)}条）</h4>'
        if len(logs) > 0:
            display = logs[['operation_time', 'operation_record', 'operator']].copy()
            display['operation_time'] = display['operation_time'].dt.strftime('%Y-%m-%d %H:%M:%S')
            display.columns = ['操作时间', '操作记录', '操作人']
            html += self._df_to_html(display)
        else:
            html += '<p>暂无物流操作记录</p>'
        
        details = details_df[details_df['order_id'] == order_id]
        if len(details) > 0:
            html += f'<h4>📦 订单明细（共{len(details)}条）</h4>'
            html += self._df_to_html(details[['product_name', 'temp_zone', 'quantity_ea']].rename(
                columns={'product_name': '商品名称', 'temp_zone': '温区', 'quantity_ea': '数量(EA)'}))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_efficiency(self):
        html = '<h3>⚡ 订单处理效率分析</h3>'
        self.thinking.append("⏱️ 计算处理时长并分级...")
        
        order_create = orders_df[['order_id', 'create_time']].copy()
        first_op = logistics_df.groupby('order_id')['operation_time'].min().reset_index()
        first_op.columns = ['order_id', 'first_op_time']
        
        merged = order_create.merge(first_op, on='order_id', how='inner')
        merged['hours'] = (merged['first_op_time'] - merged['create_time']).dt.total_seconds() / 3600
        merged = merged[merged['hours'] >= 0]
        
        merged['等级'] = pd.cut(merged['hours'], bins=[0, 1, 4, 12, 24, float('inf')],
            labels=['极快(<1h)', '快(1-4h)', '正常(4-12h)', '慢(12-24h)', '极慢(>24h)'])
        
        html += self._stats_cards([
            ("平均时长", f"{merged['hours'].mean():.1f}h"),
            ("中位数", f"{merged['hours'].median():.1f}h"),
            ("最快", f"{merged['hours'].min():.1f}h"),
            ("最慢", f"{merged['hours'].max():.1f}h")
        ])
        
        eff = merged['等级'].value_counts().reset_index()
        eff.columns = ['效率等级', '订单数']
        html += '<h4>📊 效率分布</h4>'
        html += self._df_to_html(eff)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_last_week(self):
        html = '<h3>📅 上周订单统计</h3>'
        latest = orders_df['create_time'].max()
        start = latest - timedelta(days=7)
        
        wk = orders_df[(orders_df['create_time'] >= start) & (orders_df['create_time'] <= latest)]
        html += f'<p>周期: {start.strftime("%Y-%m-%d")} 至 {latest.strftime("%Y-%m-%d")}</p>'
        html += self._stats_cards([
            ("订单数", f"{len(wk):,}"),
            ("总数量(EA)", f"{wk['quantity_ea'].sum():,.0f}"),
            ("日均订单", f"{len(wk)/7:.0f}")
        ])
        
        daily = wk.groupby(wk['create_time'].dt.date).agg(订单数=('order_id', 'count'), 总数量=('quantity_ea', 'sum')).reset_index()
        html += '<h4>每日统计</h4>'
        html += self._df_to_html(daily)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_customer_avg_ops(self):
        html = '<h3>📊 客户平均操作记录数排名</h3>'
        self.thinking.append("🔗 客户→订单→物流 三表关联...")
        
        customer_orders = orders_df[['order_id', 'customer_name']]
        op_counts = logistics_df.groupby('order_id').size().reset_index(name='op_count')
        merged = customer_orders.merge(op_counts, on='order_id', how='inner')
        stats = merged.groupby('customer_name')['op_count'].mean().reset_index()
        stats.columns = ['客户', '平均操作数']
        stats['平均操作数'] = stats['平均操作数'].round(2)
        stats = stats.sort_values('平均操作数', ascending=False)
        
        n = 3
        match = re.search(r'前(\d+)', self.query)
        if match: n = int(match.group(1))
        
        html += f'<h4>TOP {n} 客户</h4>'
        html += self._df_to_html(stats.head(n))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_sign_rate(self):
        html = '<h3>📦 仓库订单签收率</h3>'
        self.thinking.append("📊 统计每个仓库的签收情况...")
        
        signed = logistics_df[logistics_df['op_type'] == '签收']['order_id'].unique()
        
        stats = orders_df.groupby('warehouse').agg(
            总订单=('order_id', 'count'),
            已签收=('order_id', lambda x: x.isin(signed).sum())
        ).reset_index()
        stats['签收率%'] = (stats['已签收'] / stats['总订单'] * 100).round(2)
        stats = stats.sort_values('签收率%', ascending=False)
        
        html += self._df_to_html(stats)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_operator_cross_warehouse(self):
        html = '<h3>🔄 跨仓库操作人分析</h3>'
        self.thinking.append("🔗 操作人→订单→仓库 三表关联...")
        
        op_orders = logistics_df[['order_id', 'operator']].drop_duplicates('order_id')
        merged = op_orders.merge(orders_df[['order_id', 'warehouse']], on='order_id', how='inner')
        
        stats = merged.groupby('operator')['warehouse'].nunique().reset_index()
        stats.columns = ['操作人', '涉及仓库数']
        cross = stats[stats['涉及仓库数'] > 1].sort_values('涉及仓库数', ascending=False)
        
        html += f'<p>共有 <b>{len(cross)}</b> 个操作人涉及多个仓库</p>'
        html += self._df_to_html(cross.head(20))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_no_op_orders(self):
        html = '<h3>⚠️ 无操作订单</h3>'
        self.thinking.append("🔍 LEFT JOIN检测无物流记录的订单...")
        
        op_order_ids = logistics_df['order_id'].unique()
        no_op = orders_df[~orders_df['order_id'].isin(op_order_ids)]
        
        html += f'<p>共 <b>{len(no_op)}</b> 个订单创建了但一直没有操作记录</p>'
        if len(no_op) > 0:
            display = no_op[['order_id', 'customer_name', 'quantity_ea', 'create_time', 'warehouse']].head(30)
            display.columns = ['订单编号', '客户', '预计数量(EA)', '创建时间', '仓库']
            html += self._df_to_html(display)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_unique_temp_zone(self):
        html = '<h3>🌡️ 温区唯一性分析</h3>'
        self.thinking.append("🔍 分析商品出现的温区...")
        
        product_temp = details_df.groupby('product_name')['temp_zone'].apply(lambda x: set(x.dropna().unique())).reset_index()
        product_temp['温区数'] = product_temp['temp_zone'].apply(len)
        
        single = product_temp[product_temp['温区数'] == 1].copy()
        single['温区'] = single['temp_zone'].apply(lambda x: list(x)[0] if x else '未知')
        
        for tz in single['温区'].unique():
            items = single[single['温区'] == tz]['product_name'].tolist()
            html += f'<h4>只在【{tz}】出现的商品（{len(items)}种）</h4>'
            html += '<p>' + ', '.join(items[:20]) + '</p>'
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_operator_customers(self):
        html = '<h3>👤 操作人客户数分析</h3>'
        self.thinking.append("🔗 操作人→订单→客户 关联...")
        
        op_orders = logistics_df[['order_id', 'operator']].drop_duplicates('order_id')
        merged = op_orders.merge(orders_df[['order_id', 'customer_name']], on='order_id', how='inner')
        
        stats = merged.groupby('operator')['customer_name'].nunique().reset_index()
        stats.columns = ['操作人', '涉及客户数']
        stats = stats.sort_values('涉及客户数', ascending=False)
        
        threshold = 100
        match = re.search(r'超过(\d+)', self.query)
        if match: threshold = int(match.group(1))
        
        qualified = stats[stats['涉及客户数'] > threshold]
        html += f'<p>涉及超过{threshold}个不同客户的操作人: <b>{len(qualified)}</b> 个</p>'
        html += self._df_to_html(qualified.head(20))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_frozen_filter(self):
        html = '<h3>❄️ 冷冻商品条件筛选</h3>'
        
        match = re.search(r'超过|大于|>(\d+)', self.query)
        threshold = int(match.group(1)) if match else 500
        
        frozen = details_df[details_df['is_frozen'] == True]
        product_stats = frozen.groupby('product_name')['quantity_ea'].sum().reset_index()
        filtered = product_stats[product_stats['quantity_ea'] > threshold].sort_values('quantity_ea', ascending=False)
        
        html += f'<p>冷冻温区且总数量超过{threshold}EA的商品: <b>{len(filtered)}</b> 种</p>'
        html += self._df_to_html(filtered)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_month_compare(self):
        html = '<h3>📅 月份对比分析</h3>'
        self.thinking.append("📊 提取月份并对比...")
        
        months = []
        for m in range(1, 13):
            if f'{m}月' in self.query:
                months.append(m)
        
        if len(months) < 2:
            months = sorted(orders_df['month'].unique())
        
        for i in range(len(months) - 1):
            m1, m2 = months[i], months[i + 1]
            self.thinking.append(f"📊 对比 {m1}月 vs {m2}月...")
            
            s1 = orders_df[orders_df['month'] == m1].groupby('customer_name')['quantity_ea'].sum()
            s2 = orders_df[orders_df['month'] == m2].groupby('customer_name')['quantity_ea'].sum()
            
            compare = pd.DataFrame({f'{m1}月': s1, f'{m2}月': s2}).fillna(0)
            compare['变化'] = compare[f'{m2}月'] - compare[f'{m1}月']
            compare['变化率%'] = np.where(compare[f'{m1}月'] > 0, (compare['变化'] / compare[f'{m1}月'] * 100).round(2), 0)
            compare = compare.sort_values(f'{m1}月', ascending=False)
            
            html += f'<h4>{m1}月 TOP客户 在{m2}月的表现</h4>'
            html += self._df_to_html(compare.head(10))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_detail_count_filter(self):
        html = '<h3>📋 明细行数筛选</h3>'
        
        match = re.search(r'超过|大于|>(\d+)', self.query)
        threshold = int(match.group(1)) if match else 10
        
        detail_counts = details_df.groupby('order_id').size().reset_index(name='detail_count')
        filtered = detail_counts[detail_counts['detail_count'] > threshold]
        result = filtered.merge(orders_df[['order_id', 'customer_name', 'quantity_ea']], on='order_id', how='left')
        
        html += f'<p>明细行数超过{threshold}条的订单: <b>{len(result)}</b> 个</p>'
        display = result[['order_id', 'customer_name', 'quantity_ea', 'detail_count']].head(30)
        display.columns = ['订单编号', '客户', '预计数量(EA)', '明细行数']
        html += self._df_to_html(display)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_category_monthly_growth(self):
        html = '<h3>📈 品类月度销量增长分析</h3>'
        self.thinking.append("📊 关联订单+明细，按品类和月份聚合...")
        
        qty_col = 'quantity_ea' if 'quantity_ea' in details_df.columns else None
        detail_cols = ['order_id', 'category']
        if qty_col:
            detail_cols.append(qty_col)
        
        merged = orders_df[['order_id', 'month']].merge(details_df[detail_cols], on='order_id', how='inner')
        
        if qty_col:
            result = merged.groupby(['month', 'category'])[qty_col].sum().reset_index(name='qty')
        else:
            result = merged.groupby(['month', 'category']).size().reset_index(name='qty')
        
        months = sorted(result['month'].unique())
        if len(months) < 2:
            return {'success': True, 'html': '<p>数据不足</p>', 'thinking': self.thinking}
        
        m1, m2 = months[-2], months[-1]
        pivot = result.pivot(index='category', columns='month', values='qty').fillna(0)
        
        if m1 in pivot.columns and m2 in pivot.columns:
            pivot['增长量'] = pivot[m2] - pivot[m1]
            pivot['增长率%'] = np.where(pivot[m1] > 0, (pivot['增长量'] / pivot[m1] * 100).round(2), 0)
            pivot = pivot.sort_values('增长率%', ascending=False)
            
            html += f'<h4>{m1}月→{m2}月 品类增长排名</h4>'
            html += self._df_to_html(pivot[[m1, m2, '增长量', '增长率%']].head(15))
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_weekday_analysis(self):
        html = '<h3>📅 每周订单分析</h3>'
        self.thinking.append("📊 按星期几聚合...")
        
        weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
        stats = orders_df.groupby('weekday').agg(
            订单数=('order_id', 'count'), 总数量=('quantity_ea', 'sum')
        ).reset_index()
        stats['星期'] = stats['weekday'].apply(lambda x: weekday_names[x] if x < 7 else '未知')
        stats = stats.sort_values('订单数', ascending=False)
        
        html += self._df_to_html(stats[['星期', '订单数', '总数量']])
        
        max_day = stats.iloc[0]
        html += f'<p>📊 订单量最多的是 <b>{max_day["星期"]}</b>，共 <b>{int(max_day["订单数"]):,}</b> 单</p>'
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_hour_analysis(self):
        html = '<h3>🕐 时段订单分析</h3>'
        self.thinking.append("📊 按小时聚合...")
        
        def hour_period(h):
            if 6 <= h < 12: return '上午(6-12点)'
            elif 12 <= h < 18: return '下午(12-18点)'
            elif 18 <= h < 24: return '晚上(18-24点)'
            else: return '凌晨(0-6点)'
        
        stats = orders_df.groupby('hour').agg(订单数=('order_id', 'count')).reset_index()
        stats['时段'] = stats['hour'].apply(hour_period)
        
        period_stats = stats.groupby('时段')['订单数'].sum().reset_index().sort_values('订单数', ascending=False)
        html += '<h4>各时段订单量</h4>'
        html += self._df_to_html(period_stats)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_trend(self):
        html = '<h3>📈 下单量变化趋势</h3>'
        self.thinking.append("📊 按月统计趋势...")
        
        monthly = orders_df.groupby('month').agg(
            订单数=('order_id', 'count'), 总数量=('quantity_ea', 'sum'),
            客户数=('customer_name', 'nunique')
        ).reset_index().sort_values('month')
        monthly['订单环比%'] = monthly['订单数'].pct_change().mul(100).round(2)
        
        html += self._df_to_html(monthly)
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_product(self):
        html = '<h3>📦 商品查询</h3>'
        product_name = None
        if 'iphone' in self.query.lower():
            product_name = 'iPhone'
        else:
            match = re.search(r'商品\s*([\u4e00-\u9fa5a-zA-Z0-9]+)', self.query)
            if match: product_name = match.group(1)
        
        if not product_name:
            return {'success': True, 'html': html + '<p>请指定商品名称</p>', 'thinking': self.thinking}
        
        self.thinking.append(f"📦 查询商品: {product_name}")
        matched = details_df[details_df['product_name'].str.contains(product_name, na=False, case=False)]
        
        if len(matched) == 0:
            close = get_close_matches(product_name, all_products, n=5, cutoff=0.3)
            html += f'<p>未找到商品"{product_name}"</p>'
            if close: html += '<p>您是否想找:</p><ul>' + ''.join([f'<li>{c}</li>' for c in close]) + '</ul>'
            return {'success': True, 'html': html, 'thinking': self.thinking}
        
        html += self._stats_cards([
            ("预计发货总量(EA)", f"{matched['quantity_ea'].sum():,.0f}"),
            ("涉及订单数", f"{matched['order_id'].nunique()}")
        ])
        
        names = matched['product_name'].unique()
        if len(names) > 1:
            html += '<h4>匹配到的商品</h4><p>' + ', '.join(names[:10]) + '</p>'
        
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_temp_zone(self):
        html = '<h3>🌡️ 温区销量统计</h3>'
        merged = orders_df[['order_id']].merge(details_df[['order_id', 'temp_zone', 'quantity_ea']], on='order_id', how='inner')
        stats = merged.groupby('temp_zone').agg(总数量=('quantity_ea', 'sum'), 订单数=('order_id', 'nunique')).reset_index().sort_values('总数量', ascending=False)
        html += self._df_to_html(stats)
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    def _handle_unknown(self):
        html = '<h3>💡 欢迎使用ChatBI</h3><p>您可以尝试以下查询：</p><ul>'
        suggestions = [
            "司机王快乐负责的订单进度怎么样了？",
            "上周的订单总量是多少？",
            "冷冻商品卖得多不多？",
            "订单处理效率怎么样？",
            "订单编号CO039276924的每一次操作时间分别是什么？",
            "哪些订单的操作记录条数超过5条？",
            "计算每个温区在3月和4月的销量变化",
            "统计每个客户的订单总金额排名（前5名）",
            "找出订单处理时长最长的5个订单",
            "哪个商品在4月的预计发货数量最高？",
            "计算每个月订单数量的环比增长率",
            "哪些客户的订单平均操作记录数最高？",
            "每个仓库的订单签收率是多少？",
            "有没有订单创建了但一直没有操作的？",
            "哪些商品只在冷冻温区出现？",
            "每周哪一天的订单量最多？",
            "哪个时段创建的订单最多？",
        ]
        for s in suggestions:
            html += f'<li>{s}</li>'
        html += '</ul>'
        return {'success': True, 'html': html, 'thinking': self.thinking}
    
    # ========== 工具函数 ==========
    
    def _stats_cards(self, stats):
        html = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:15px;margin:20px 0;">'
        for label, value in stats:
            html += f'<div style="background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:15px;border-radius:10px;text-align:center;"><div style="font-size:24px;font-weight:bold;">{value}</div><div style="font-size:12px;opacity:0.9;">{label}</div></div>'
        html += '</div>'
        return html
    
    def _df_to_html(self, df):
        if len(df) == 0: return '<p>无数据</p>'
        html = '<table style="width:100%;border-collapse:collapse;margin:15px 0;font-size:13px;">'
        html += '<thead><tr style="background:linear-gradient(135deg,#667eea,#764ba2);color:white;">'
        for col in df.columns:
            html += f'<th style="padding:10px;text-align:left;border:1px solid #ddd;">{col}</th>'
        html += '</tr></thead><tbody>'
        for idx, row in df.iterrows():
            bg = '#f9f9f9' if idx % 2 == 0 else 'white'
            html += f'<tr style="background:{bg}">'
            for val in row:
                if isinstance(val, float): val = f'{val:,.2f}' if val != int(val) else f'{int(val):,}'
                elif isinstance(val, (int, np.integer)): val = f'{int(val):,}'
                elif isinstance(val, (np.floating,)): val = f'{float(val):,.2f}'
                html += f'<td style="padding:8px;border:1px solid #ddd;">{val}</td>'
            html += '</tr>'
        html += '</tbody></table>'
        return html


# ========== Flask路由 ==========

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/chat', methods=['POST'])
def chat():
    try:
        data = request.json
        query = data.get('message', '').strip()
        if not query:
            return jsonify({'error': '请输入查询内容'}), 400
        
        print(f"\n📝 查询: {query}")
        engine = ChatBIEngine(query)
        result = engine.execute()
        
        thinking_html = '<div style="background:#f0f7ff;border-left:4px solid #4a90e2;padding:15px;margin:15px 0;"><h4 style="margin-top:0;color:#4a90e2;">🧠 推理过程</h4><ol>'
        for step in result['thinking']:
            thinking_html += f'<li>{step}</li>'
        thinking_html += '</ol></div>'
        
        final_html = f'<div style="background:#f5f5f5;padding:10px 15px;border-radius:4px;margin-bottom:15px;font-weight:bold;">💬 {query}</div>'
        final_html += thinking_html + result['html']
        
        return jsonify({'success': True, 'html': final_html})
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/suggestions', methods=['GET'])
def suggestions():
    return jsonify({'suggestions': [
        "司机王快乐负责的订单进度怎么样了？",
        "上周的订单总量是多少？",
        "冷冻商品卖得多不多？",
        "订单处理效率怎么样？",
        "订单编号CO039276924的每一次操作时间分别是什么？",
        "哪些订单的操作记录条数超过5条？",
        "计算每个温区在3月和4月的销量变化",
        "统计每个客户的订单总金额排名（前5名）",
        "找出订单处理时长最长的5个订单",
        "哪个商品在4月的预计发货数量最高？",
        "计算每个月订单数量的环比增长率",
        "哪些客户的订单平均操作记录数最高？",
        "每个仓库的订单签收率是多少？",
        "有没有订单创建了但一直没有操作的？",
        "哪些商品只在冷冻温区出现？",
        "每周哪一天的订单量最多？",
        "哪个时段创建的订单最多？",
    ]})

@app.route('/api/stats', methods=['GET'])
def stats():
    return jsonify({
        'orders': len(orders_df) if orders_df is not None else 0,
        'details': len(details_df) if details_df is not None else 0,
        'logistics': len(logistics_df) if logistics_df is not None else 0,
        'customers': len(all_customers), 'products': len(all_products), 'drivers': len(all_drivers)
    })

if __name__ == '__main__':
    load_data()
    print("\n" + "=" * 70)
    print("🚀 ChatBI 终极版已启动！")
    print("=" * 70)
    print("🌐 访问: http://localhost:5000")
    print("=" * 70)
    app.run(debug=False, host='0.0.0.0', port=5000, use_reloader=False)
