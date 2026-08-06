import openpyxl
import json
from collections import defaultdict, OrderedDict
from datetime import datetime

# ========== 1. 解析用户提供的102家店白名单 ==========
whitelist_raw = """东莞零跑中心塘厦车城店	东莞市	方任昊
东莞体验中心东城店	东莞市	方任昊
东莞体验中心松山湖店	东莞市	方任昊
东莞体验中心石排大道店	东莞市	方任昊
东莞零跑中心常朗路店	东莞市	方任昊
东莞体验中心高埗车城店	东莞市	方任昊
东莞零跑中心虎门车城店	东莞市	方任昊
东莞零跑中心寮步车城店	东莞市	方任昊
东莞零跑中心莞太路店	东莞市	方任昊
东莞体验中心厚街店	东莞市	方任昊
东莞零跑中心长安振安东路店	东莞市	方任昊
东莞体验中心大岭山镇店	东莞市	方任昊
广州零跑中心金马汽车城店	广州市	何佳欢
广州零跑中心亚运大道店	广州市	何佳欢
广州零跑中心美轮汽车园店	广州市	何佳欢
广州零跑中心沥滘海心沙店	广州市	何佳欢
广州体验中心天河汇彩路店	广州市	何佳欢
广州体验中心天河广汕路店	广州市	何佳欢
广州体验中心番禺大道店	广州市	何佳欢
惠州零跑中心金山汽车城店	惠州市	胡浩
惠州体验中心惠阳大道店	惠州市	胡浩
惠州零跑中心惠南汽车城店	惠州市	胡浩
惠州零跑中心仲恺汽车城店	惠州市	胡浩
惠州体验中心河南岸车城店	惠州市	胡浩
汕尾体验中心汕尾大道店	汕尾市	胡浩
韶关体验中心浈江大道店	韶关市	胡浩
韶关零跑中心沐溪大道店	韶关市	胡浩
广州体验中心为正车城店	广州市	黄伟峰
广州零跑中心花都建设路店	广州市	黄伟峰
广州零跑中心白云大道店	广州市	黄伟峰
广州体验中心广花路店	广州市	黄伟峰
广州体验中心花都北站店	广州市	黄伟峰
广州体验中心白云机场路店	广州市	黄伟峰
广州零跑中心新塘汽车城店	广州市	黄伟峰
广州零跑中心科学城店	广州市	黄伟峰
清远零跑中心港鸿汽车城店	清远市	黄伟峰
清远体验中心奇晟汽车城店	清远市	黄伟峰
深圳体验中心龙岗信义店	深圳市	贾迪赫
深圳体验中心锦龙大道店	深圳市	贾迪赫
深圳零跑中心百世汽车城店	深圳市	贾迪赫
深圳零跑中心南山嘉进隆店	深圳市	贾迪赫
深圳体验中心福田瑞鹏达店	深圳市	贾迪赫
深圳零跑中心芙蓉路店	深圳市	贾迪赫
三亚体验中心迎宾路汽车城店	三亚市	李博恩
三亚零跑中心榆亚路店	三亚市	李博恩
海口零跑中心琼山大道店	海口市	李博恩
海口体验中心江东店	海口市	李博恩
海口体验中心海口东站店	海口市	李博恩
海口零跑中心海盛路店	海口市	李博恩
海口零跑中心南海大道店（销售）	海口市	李博恩
海口零跑中心南海大道东店	海口市	李博恩
海口体验中心海甸城店	海口市	李博恩
海口体验中心上邦百汇城店	海口市	李博恩
海口零跑中心美安科技城店	海口市	李博恩
中山零跑中心小榄菊城大道店	中山市	林伟龙
中山零跑中心金宁汽车城店	中山市	林伟龙
中山体验中心港口店	中山市	林伟龙
中山零跑中心中山六路店	中山市	林伟龙
中山零跑中心彩虹大道店	中山市	林伟龙
珠海体验中心美满车城店	珠海市	林伟龙
珠海零跑中心南屏珠海大道店	珠海市	林伟龙
珠海体验中心西部车城店	珠海市	林伟龙
珠海零跑中心上冲车城店	珠海市	林伟龙
深圳体验中心杰鹏车城店	深圳市	罗捷
深圳体验中心龙华中心店	深圳市	罗捷
深圳零跑中心光明车城店	深圳市	罗捷
深圳零跑中心万国城旗舰店	深圳市	罗捷
深圳零跑中心远望车城店	深圳市	罗捷
深圳零跑中心宝运达汽车城店	深圳市	罗捷
佛山体验中心海八西路店	佛山市	沈祖福
佛山零跑中心南庄汽车城店	佛山市	沈祖福
佛山零跑中心禅城国际汽车城店	佛山市	沈祖福
佛山体验中心南海平洲店	佛山市	沈祖福
佛山体验中心广佛汽车城店	佛山市	沈祖福
佛山零跑中心海八路汽车城店	佛山市	沈祖福
揭阳零跑中心荣通汽车城店	揭阳市	熊俊晖
揭阳体验中心普宁万泰新天地商场店	揭阳市	熊俊晖
梅州零跑中心剑英大道店	梅州市	熊俊晖
汕头体验中心广汕路店	汕头市	熊俊晖
汕头零跑中心金凤路店	汕头市	熊俊晖
汕头零跑中心汕汾路店	汕头市	熊俊晖
河源零跑中心河源大道店	河源市	熊俊晖
河源体验中心坚基购物中心店	河源市	熊俊晖
潮州零跑中心潮汕路店	潮州市	熊俊晖
佛山体验中心容桂车城店	佛山市	余子恩
佛山零跑中心顺德大良车城店	佛山市	余子恩
佛山体验中心北滘车城店	佛山市	余子恩
佛山体验中心高明荷香路店	佛山市	余子恩
佛山零跑中心三水友好车城店	佛山市	余子恩
佛山体验中心狮山汽车城店	佛山市	余子恩
江门体验中心开平车城店	江门市	余子恩
江门零跑中心冈州大道店	江门市	余子恩
江门体验中心蓬江华鸿店	江门市	余子恩
江门零跑中心建设路店	江门市	余子恩
云浮体验中心环市中路店	云浮市	庄文迪
湛江零跑中心海田车城店	湛江市	庄文迪
湛江体验中心粤西车城店	湛江市	庄文迪
湛江体验中心瀚龙车城店	湛江市	庄文迪
肇庆零跑中心肇庆大道店	肇庆市	庄文迪
茂名零跑中心茂名大道店	茂名市	庄文迪
茂名体验中心茂水路店	茂名市	庄文迪
阳江零跑中心溢信汽车城店	阳江市	庄文迪"""

# 解析白名单：门店名称 -> 经理映射
store_to_manager = {}
manager_stores = defaultdict(list)
for line in whitelist_raw.strip().split('\n'):
    parts = line.split('\t')
    if len(parts) >= 3:
        store_name = parts[0].strip()
        manager = parts[2].strip()
        store_to_manager[store_name] = manager
        manager_stores[manager].append(store_name)

whitelist = set(store_to_manager.keys())
print(f"白名单门店数: {len(whitelist)}")
print(f"经理数: {len(manager_stores)}")
for m in manager_stores:
    print(f"  {m}: {len(manager_stores[m])}家")

# ========== 2. 读取Excel审核数据 ==========
excel_path = "/Users/mei/Downloads/晨夕会审核列表2026-06-24.xlsx"
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# 统计周期：昨天往前7天 (6/17 - 6/23)
start_date = datetime(2026, 6, 17).date()
end_date = datetime(2026, 6, 23).date()

# 只保留白名单中的门店记录
store_meetings = {}  # store -> {'morning': [bool*7], 'evening': [bool*7]}

for store in whitelist:
    store_meetings[store] = {
        'morning': [False] * 7,
        'evening': [False] * 7
    }

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[3]:
        continue
    
    store_name = str(row[3]).strip()
    
    # 严格过滤：不在白名单中的不计数
    if store_name not in whitelist:
        continue
    
    meeting_type = str(row[1]).strip() if row[1] else ""
    meeting_date = row[6]
    
    if meeting_date is None:
        continue
        
    if isinstance(meeting_date, str):
        try:
            dt = datetime.strptime(meeting_date[:10], '%Y-%m-%d').date()
        except:
            continue
    elif isinstance(meeting_date, datetime):
        dt = meeting_date.date()
    else:
        continue
    
    if start_date <= dt <= end_date:
        idx = (dt - start_date).days
        if '晨会' in meeting_type:
            store_meetings[store_name]['morning'][idx] = True
        elif '夕会' in meeting_type:
            store_meetings[store_name]['evening'][idx] = True

print(f"\n已读取Excel数据，匹配到 {sum(1 for s in whitelist if any(store_meetings[s]['morning']) or any(store_meetings[s]['evening']))} 家有数据")

# ========== 3. 按经理统计 ==========
date_labels = ['6/17', '6/18', '6/19', '6/20', '6/21', '6/22', '6/23']

manager_stats = []
for manager in sorted(manager_stores.keys()):
    stores = manager_stores[manager]
    n_stores = len(stores)
    
    # 每天开展数
    morning_daily = [0] * 7
    evening_daily = [0] * 7
    
    # 门店明细
    store_details = []
    
    for store in stores:
        m_data = store_meetings[store]['morning']
        e_data = store_meetings[store]['evening']
        
        for i in range(7):
            if m_data[i]:
                morning_daily[i] += 1
            if e_data[i]:
                evening_daily[i] += 1
        
        store_details.append({
            'name': store,
            'morning': m_data,
            'evening': e_data
        })
    
    # 计算昨日开展率（第7天，index=6）
    m_rate = round(morning_daily[6] / n_stores * 100) if n_stores > 0 else 0
    e_rate = round(evening_daily[6] / n_stores * 100) if n_stores > 0 else 0
    
    # 每天开展率百分比
    morning_rates = [round(m / n_stores * 100) if n_stores > 0 else 0 for m in morning_daily]
    evening_rates = [round(e / n_stores * 100) if n_stores > 0 else 0 for e in evening_daily]
    
    manager_stats.append({
        'name': manager,
        'total': n_stores,
        'morning': morning_daily,
        'evening': evening_daily,
        'morning_rate': f"{m_rate}%",
        'evening_rate': f"{e_rate}%",
        'store_details': store_details
    })

# 总体统计
total_morning = [0] * 7
total_evening = [0] * 7
for ms in manager_stats:
    for i in range(7):
        total_morning[i] += ms['morning'][i]
        total_evening[i] += ms['evening'][i]

total_n = len(whitelist)
overall_m_rate = round(total_morning[6] / total_n * 100) if total_n > 0 else 0
overall_e_rate = round(total_evening[6] / total_n * 100) if total_n > 0 else 0

print(f"\n总体统计:")
print(f"  白名单总门店数: {total_n}")
print(f"  昨日(6/23)晨会: {total_morning[6]}/{total_n} ({overall_m_rate}%)")
print(f"  昨日(6/23)夕会: {total_evening[6]}/{total_n} ({overall_e_rate}%)")
print(f"  7天晨会数据: {total_morning}")
print(f"  7天夕会数据: {total_evening}")

# ========== 4. 生成完整HTML海报 ==========
html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1400">
<title>华南大区晨夕会开展情况</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
  background: #f0f2f5;
  display: flex;
  justify-content: center;
  padding: 40px 20px;
}
.poster {
  width: 1340px;
  background: #fff;
  border-radius: 12px;
  overflow: hidden;
  box-shadow: 0 4px 24px rgba(0,0,0,.08);
}
.header {
  background: linear-gradient(135deg, #1a3c6e 0%, #2d5fa0 100%);
  padding: 28px 40px 22px;
  display: flex;
  align-items: center;
  gap: 20px;
}
.header-left { display: flex; align-items: center; gap: 14px; }
.header-dot {
  width: 8px; height: 8px;
  background: #f9a826;
  border-radius: 50%;
  box-shadow: 0 0 8px #f9a82688;
}
.header h1 { color: #fff; font-size: 26px; font-weight: 700; letter-spacing: 2px; }
.header-date {
  color: #bfd4f0; font-size: 14px; margin-left: 10px;
  padding: 4px 14px; background: rgba(255,255,255,.1); border-radius: 20px;
}
.header-stats { margin-left: auto; display: flex; gap: 24px; }
.stat-item { text-align: center; }
.stat-num { color: #f9a826; font-size: 28px; font-weight: 700; }
.stat-label { color: #bfd4f0; font-size: 12px; margin-top: 2px; }
.body { padding: 20px 40px 30px; }

/* 经理统计表格 */
.section-label {
  display: flex; align-items: center; gap: 10px;
  margin: 18px 0 10px;
}
.section-label .bar { width: 4px; height: 18px; border-radius: 2px; }
.bar.morning { background: #f9a826; }
.bar.evening { background: #5b8def; }
.section-label span { font-size: 16px; font-weight: 600; color: #333; }

.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th {
  background: #f5f7fa; padding: 10px 8px; text-align: center;
  font-weight: 600; color: #666; white-space: nowrap;
  border-bottom: 2px solid #e8ecf0;
}
td {
  padding: 8px; text-align: center; border-bottom: 1px solid #f0f2f5;
  white-space: nowrap;
}
.label-col { text-align: left !important; font-weight: 500; color: #333; min-width: 80px; }
.gap-col { width: 20px; background: transparent !important; border: none !important; }
.row-total { background: #fafbfc; font-weight: 600; }
.rate-tag {
  display: inline-block; padding: 2px 10px; border-radius: 10px;
  font-size: 12px; font-weight: 600;
}
.rate-high { background: #e8f5e9; color: #2e7d32; }
.rate-mid { background: #fff3e0; color: #ef6c00; }
.rate-low { background: #fce4ec; color: #c62828; }

/* 门店明细 */
.store-section { margin-top: 30px; }
.store-table { font-size: 11px; }
.store-table th { font-size: 11px; padding: 8px 6px; }
.store-table td { padding: 6px; }
.store-col { text-align: left !important; min-width: 220px; color: #444; }
.manager-col { min-width: 60px; color: #888; font-size: 11px; }
.check-col { padding: 4px 2px !important; }
.check-yes { color: #2e7d32; }
.check-no { color: #bdbdbd; }
</style>
</head>
<body>
<div class="poster">
<div class="header">
  <div class="header-left"><div class="header-dot"></div><h1>华南大区晨夕会开展情况</h1></div>
  <span class="header-date">''' + date_labels[0] + '-' + date_labels[-1] + '''</span>
  <div class="header-stats">
    <div class="stat-item"><div class="stat-num">''' + str(overall_m_rate) + '''%</div><div class="stat-label">昨日晨会率</div></div>
    <div class="stat-item"><div class="stat-num">''' + str(overall_e_rate) + '''%</div><div class="stat-label">昨日夕会率</div></div>
  </div>
</div>

<div class="body">

<!-- 经理统计表 -->
<div class="section-label"><div class="bar morning"></div><span>各经理晨夕会开展情况</span></div>
<table id="managerTable"></table>

<!-- 门店明细 -->
<div class="section-label store-section"><div class="bar morning"></div><span>门店明细（共''' + str(total_n) + '''家）</span></div>
<div class="table-wrap"><table id="storeTable" class="store-table"></table></div>

</div>
</div>

<script>
// 数据定义
const DATE_LABELS = ''' + json.dumps(date_labels, ensure_ascii=False) + ''';
const TOTAL_STORES = ''' + str(total_n) + ''';

const MANAGERS_DATA = ''' + json.dumps(manager_stats, ensure_ascii=False) + ''';

function rateClass(rateStr) {
  const v = parseInt(rateStr);
  if (v >= 90) return 'rate-high';
  if (v >= 70) return 'rate-mid';
  return 'rate-low';
}

// 渲染经理统计表
function renderManagerTable() {
  const table = document.getElementById('managerTable');
  let html = '<thead><tr>';
  html += '<th>零售高级经理</th><th>门店数</th>';
  for (let d of DATE_LABELS) html += '<th>' + d + '</th>';
  html += '<th>晨会率</th><th class="gap-col"></th>';
  for (let d of DATE_LABELS) html += '<th>' + d + '</th>';
  html += '<th>夕会率</th>';
  html += '</tr></thead><tbody>';

  // 总计行
  let tm = ''' + json.dumps(total_morning) + ''';
  let te = ''' + json.dumps(total_evening) + ''';
  html += '<tr class="row-total">';
  html += '<td class="label-col">华南总计</td><td>' + TOTAL_STORES + '</td>';
  for (let v of tm) html += '<td>' + v + '</td>';
  html += '<td><span class="rate-tag ' + rateClass("''' + str(overall_m_rate) + '''%") + '">''' + str(overall_m_rate) + '''%</span></td>';
  html += '<td class="gap-col"></td>';
  for (let v of te) html += '<td>' + v + '</td>';
  html += '<td><span class="rate-tag ' + rateClass("''' + str(overall_e_rate) + '''%") + '">''' + str(overall_e_rate) + '''%</span></td>';
  html += '</tr>';

  // 各经理行
  for (let m of MANAGERS_DATA) {
    html += '<tr>';
    html += '<td class="label-col">' + m.name + '</td><td>' + m.total + '</td>';
    for (let v of m.morning) html += '<td>' + v + '</td>';
    html += '<td><span class="rate-tag ' + rateClass(m.morning_rate) + '">' + m.morning_rate + '</span></td>';
    html += '<td class="gap-col"></td>';
    for (let v of m.evening) html += '<td>' + v + '</td>';
    html += '<td><span class="rate-tag ' + rateClass(m.evening_rate) + '">' + m.evening_rate + '</span></td>';
    html += '</tr>';
  }
  html += '</tbody>';
  table.innerHTML = html;
}

// 渲染门店明细表
function renderStoreTable() {
  const table = document.getElementById('storeTable');
  let html = '<thead><tr>';
  html += '<th class="store-col">门店名称</th><th class="manager-col">零售高级经理</th>';
  for (let d of DATE_LABELS) html += '<th>晨' + d.slice(-2) + '</th>';
  html += '<th class="gap-col"></th>';
  for (let d of DATE_LABELS) html += '<th>夕' + d.slice(-2) + '</th>';
  html += '</tr></thead><tbody>';

  for (let m of MANAGERS_DATA) {
    for (let s of m.store_details) {
      html += '<tr>';
      html += '<td class="store-col">' + s.name + '</td><td class="manager-col">' + m.name + '</td>';
      for (let v of s.morning) html += '<td class="check-col"><span class="' + (v ? 'check-yes' : 'check-no') + '">' + (v ? '✅' : '❌') + '</span></td>';
      html += '<td class="gap-col"></td>';
      for (let v of s.evening) html += '<td class="check-col"><span class="' + (v ? 'check-yes' : 'check-no') + '">' + (v ? '✅' : '❌') + '</span></td>';
      html += '</tr>';
    }
  }
  html += '</tbody>';
  table.innerHTML = html;
}

renderManagerTable();
renderStoreTable();
</script>
</body>
</html>'''

with open('晨夕会海报.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("\n✅ 海报已生成: 晨夕会海报.html")
