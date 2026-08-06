import openpyxl
import json
from collections import defaultdict
from datetime import datetime, timedelta

# 读取Excel文件
excel_path = "/Users/mei/Downloads/晨夕会审核列表2026-06-24.xlsx"
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# 统计周期
yesterday = datetime(2026, 6, 23).date()
start_date = datetime(2026, 6, 17).date()

# 读取并过滤记录
store_meetings = defaultdict(lambda: {'morning': [], 'evening': []})
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[3]:
        store_name = str(row[3]).strip()
        
        # 过滤掉"深圳零跑中心万国城店"
        if store_name == '深圳零跑中心万国城店':
            continue
        
        meeting_type = str(row[1]).strip() if row[1] else ""
        meeting_date = row[6]
        
        if meeting_date:
            if isinstance(meeting_date, str):
                try:
                    meeting_date = datetime.strptime(meeting_date[:10], '%Y-%m-%d')
                except:
                    continue
            
            if isinstance(meeting_date, datetime):
                meeting_date = meeting_date.date()
                if start_date <= meeting_date <= yesterday:
                    date_idx = (meeting_date - start_date).days
                    if '晨会' in meeting_type:
                        if date_idx not in store_meetings[store_name]['morning']:
                            store_meetings[store_name]['morning'].append(date_idx)
                    elif '夕会' in meeting_type:
                        if date_idx not in store_meetings[store_name]['evening']:
                            store_meetings[store_name]['evening'].append(date_idx)

print(f"统计门店数: {len(store_meetings)}")

# 门店-经理映射
manager_map = {
    "方任昊": ["惠州零跑中心惠博大道店", "惠州体验中心华贸店", "深圳零跑中心综上汽城店", "深圳零跑中心万国城旗舰店", "深圳零跑中心宝安大道店", "深圳零跑中心龙岗大道店", "深圳零跑中心泰然店", "深圳零跑中心光明店", "深圳零跑中心坪山店", "东莞零跑中心寮步店", "东莞零跑中心厚街店"],
    "何佳欢": ["东莞零跑中心虎门店", "东莞体验中心东城店", "东莞体验中心民盈国贸店", "东莞体验中心松山湖店", "广州零跑中心天河店", "广州零跑中心海珠店", "广州零跑中心番禺店"],
    "胡浩": ["广州零跑中心花都店", "广州零跑中心南沙店", "广州零跑中心增城店", "广州零跑中心从化店", "广州体验中心美林天地店", "佛山零跑中心顺德店", "佛山零跑中心南海店", "佛山零跑中心禅城店"],
    "黄伟峰": ["佛山体验中心岭南天地店", "佛山体验中心桂城店", "湛江零跑中心麻章店", "湛江体验中心人民大道店", "茂名零跑中心茂名大道店", "肇庆零跑中心端州店", "清远零跑中心清城店", "韶关零跑中心沐溪大道店"],
    "贾迪赫": ["河源零跑中心建设大道店", "梅州零跑中心梅县店", "梅州体验中心万象店", "汕头零跑中心泰山路店", "汕头体验中心万象城店", "潮州零跑中心潮汕公路店"],
    "李博恩": ["揭阳零跑中心榕城店", "阳江零跑中心东风三路店", "阳江体验中心百利广场店", "云浮零跑中心浩林东路店", "桂林零跑中心八里街店", "桂林体验中心临桂店", "福州零跑中心自贸区店", "福州零跑中心闽侯店"],
    "林伟龙": ["福州体验中心东二环泰禾店", "厦门零跑中心集美大道店", "厦门零跑中心海沧店", "厦门体验中心SM城市广场店", "泉州零跑中心晋江店", "泉州零跑中心鲤城店", "泉州体验中心中骏世界城店", "漳州零跑中心龙文店"],
    "罗捷": ["漳州体验中心碧湖万达店", "莆田零跑中心荔城店", "三明零跑中心三元店", "南平零跑中心建阳店", "龙岩零跑中心新罗店", "宁德零跑中心东侨店"],
    "沈祖福": ["贵阳零跑中心花溪店", "贵阳体验中心中大国际店", "贵阳体验中心万象汇店", "遵义零跑中心汇川店", "遵义体验中心吾悦广场店", "昆明零跑中心经开区店"],
    "熊俊晖": ["昆明零跑中心西山店", "昆明体验中心同德广场店", "昆明体验中心七彩云南第壹城店", "南宁零跑中心荔滨大道店", "南宁体验中心五象航洋城店", "海口零跑中心南海大道东店"],
    "余子恩": ["海口体验中心明珠广场店", "三亚零跑中心榆亚路店", "柳州零跑中心广汽路店", "柳州体验中心万象城店", "南宁零跑中心武鸣店", "柳州零跑中心柳江店"],
    "庄文迪": ["玉林零跑中心玉州店", "北海零跑中心海城店", "桂林零跑中心灵川店", "梧州零跑中心长洲店", "钦州零跑中心钦南店", "贵港零跑中心港北店"]
}

# 生成STORE_DATA
store_data = []
for manager, stores in manager_map.items():
    manager_data = {'manager': manager, 'stores': []}
    for store in stores:
        if store in store_meetings:
            morning_list = [i in store_meetings[store]['morning'] for i in range(7)]
            evening_list = [i in store_meetings[store]['evening'] for i in range(7)]
            manager_data['stores'].append({
                'name': store,
                'morning': morning_list,
                'evening': evening_list
            })
        else:
            # 如果门店没有数据，全部设为false
            manager_data['stores'].append({
                'name': store,
                'morning': [False] * 7,
                'evening': [False] * 7
            })
    store_data.append(manager_data)

# 计算总体统计
total_stores = len(store_meetings)
morning_count = 0
evening_count = 0
morning_rates = []
evening_rates = []

for i in range(7):
    m_count = sum(1 for s in store_meetings if i in store_meetings[s]['morning'])
    e_count = sum(1 for s in store_meetings if i in store_meetings[s]['evening'])
    morning_rates.append(round(m_count / total_stores * 100) if total_stores > 0 else 0)
    evening_rates.append(round(e_count / total_stores * 100) if total_stores > 0 else 0)

# 昨日开展率
morning_yesterday = sum(1 for s in store_meetings if 6 in store_meetings[s]['morning'])
evening_yesterday = sum(1 for s in store_meetings if 6 in store_meetings[s]['evening'])
morning_rate = round(morning_yesterday / total_stores * 100) if total_stores > 0 else 0
evening_rate = round(evening_yesterday / total_stores * 100) if total_stores > 0 else 0

# 生成DATA变量
data_js = f"""const DATA = {{
  morning: {morning_rates},
  evening: {evening_rates},
  morning_rate: "{morning_rate}%",
  evening_rate: "{evening_rate}%",
  total: {total_stores}
}};"""

# 生成STORE_DATA变量
store_data_js = f"const STORE_DATA = {json.dumps(store_data, ensure_ascii=False)};"

# 读取原HTML文件
with open('晨夕会海报.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

# 替换DATA变量
import re
html_content = re.sub(r'const DATA = \{.*?\};', data_js, html_content, flags=re.DOTALL)

# 替换STORE_DATA变量（找到STORE_DATA = [...]的部分）
# 由于STORE_DATA可能很长，使用更精确的正则
pattern = r'const STORE_DATA = \[.*?\];'
html_content = re.sub(pattern, store_data_js, html_content, flags=re.DOTALL)

# 保存更新后的HTML
with open('晨夕会海报.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\nHTML海报已更新！")
print(f"  - 总门店数: {total_stores}")
print(f"  - 昨日晨会开展率: {morning_rate}%")
print(f"  - 昨日夕会开展率: {evening_rate}%")
print(f"  - '深圳零跑中心万国城店'已过滤")

