import openpyxl
import json
from collections import defaultdict
from datetime import datetime, timedelta

# 完整的白名单（103家店）
whitelist = [
    "惠州零跑中心惠博大道店", "惠州体验中心华贸店", "深圳零跑中心综上汽城店",
    "深圳零跑中心万国城旗舰店", "深圳零跑中心宝安大道店", "深圳零跑中心龙岗大道店",
    "深圳零跑中心泰然店", "深圳零跑中心光明店", "深圳零跑中心坪山店",
    "东莞零跑中心寮步店", "东莞零跑中心厚街店", "东莞零跑中心虎门店",
    "东莞体验中心东城店", "东莞体验中心民盈国贸店", "东莞体验中心松山湖店",
    "广州零跑中心天河店", "广州零跑中心海珠店", "广州零跑中心番禺店",
    "广州零跑中心花都店", "广州零跑中心南沙店", "广州零跑中心增城店",
    "广州零跑中心从化店", "广州体验中心美林天地店", "佛山零跑中心顺德店",
    "佛山零跑中心南海店", "佛山零跑中心禅城店", "佛山体验中心岭南天地店",
    "佛山体验中心桂城店", "湛江零跑中心麻章店", "湛江体验中心人民大道店",
    "茂名零跑中心茂名大道店", "肇庆零跑中心端州店", "清远零跑中心清城店",
    "韶关零跑中心沐溪大道店", "河源零跑中心建设大道店", "梅州零跑中心梅县店",
    "梅州体验中心万象店", "汕头零跑中心泰山路店", "汕头体验中心万象城店",
    "潮州零跑中心潮汕公路店", "揭阳零跑中心榕城店", "阳江零跑中心东风三路店",
    "阳江体验中心百利广场店", "云浮零跑中心浩林东路店", "桂林零跑中心八里街店",
    "桂林体验中心临桂店", "柳州零跑中心广汽路店", "柳州体验中心万象城店",
    "南宁零跑中心荔滨大道店", "南宁体验中心五象航洋城店", "海口零跑中心南海大道东店",
    "海口体验中心明珠广场店", "三亚零跑中心榆亚路店", "福州零跑中心自贸区店",
    "福州零跑中心闽侯店", "福州体验中心东二环泰禾店", "厦门零跑中心集美大道店",
    "厦门零跑中心海沧店", "厦门体验中心SM城市广场店", "泉州零跑中心晋江店",
    "泉州零跑中心鲤城店", "泉州体验中心中骏世界城店", "漳州零跑中心龙文店",
    "漳州体验中心碧湖万达店", "莆田零跑中心荔城店", "三明零跑中心三元店",
    "南平零跑中心建阳店", "龙岩零跑中心新罗店", "宁德零跑中心东侨店",
    "贵阳零跑中心花溪店", "贵阳体验中心中大国际店", "贵阳体验中心万象汇店",
    "遵义零跑中心汇川店", "遵义体验中心吾悦广场店", "昆明零跑中心经开区店",
    "昆明零跑中心西山店", "昆明体验中心同德广场店", "昆明体验中心七彩云南第壹城店",
    "南宁零跑中心武鸣店", "柳州零跑中心柳江店", "玉林零跑中心玉州店",
    "北海零跑中心海城店", "桂林零跑中心灵川店", "梧州零跑中心长洲店",
    "钦州零跑中心钦南店", "贵港零跑中心港北店", "百色零跑中心右江店",
    "崇左零跑中心江州店", "来宾零跑中心兴宾店", "贺州零跑中心八步店",
    "河池零跑中心金城江店", "防城港零跑中心港口店", "贺州体验中心灵峰广场店",
    "玉林体验中心玉州店", "北海体验中心海城店", "海口零跑中心秀英店",
    "三亚体验中心天涯店", "儋州零跑中心那大店", "昆明零跑中心盘龙店",
    "昆明零跑中心官渡店", "昆明体验中心五华店", "贵阳零跑中心观山湖店",
    "贵阳零跑中心白云店", "遵义零跑中心播州店", "黔东南零跑中心凯里店",
    "黔南零跑中心都匀店", "黔西南零跑中心兴义店", "六盘水零跑中心钟山店",
    "安顺零跑中心西秀店", "毕节零跑中心七星关店", "铜仁零跑中心碧江店"
]

print(f"白名单门店数: {len(whitelist)}")

# 门店-经理映射（根据白名单）
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

# 读取Excel文件
excel_path = "/Users/mei/Downloads/晨夕会审核列表2026-06-24.xlsx"
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# 统计周期：昨天往前7天
yesterday = datetime.now() - timedelta(days=1)
seven_days_ago = yesterday - timedelta(days=6)
start_date = seven_days_ago.date()
end_date = yesterday.date()

print(f"\n统计周期: {start_date} 到 {end_date}")

# 读取并过滤记录
filtered_records = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[3]:
        store_name = str(row[3]).strip()
        
        # 只保留白名单中的门店
        if store_name not in whitelist:
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
                if start_date <= meeting_date <= end_date:
                    filtered_records.append({
                        'store': store_name,
                        'date': meeting_date,
                        'type': meeting_type
                    })

print(f"过滤后记录数: {len(filtered_records)}")

# 统计每个门店的晨夕会情况
store_meetings = defaultdict(lambda: {'morning': set(), 'evening': set()})
for r in filtered_records:
    store = r['store']
    date = r['date']
    if '晨会' in r['type']:
        store_meetings[store]['morning'].add(date)
    elif '夕会' in r['type']:
        store_meetings[store]['evening'].add(date)

# 按经理统计
manager_stats = {}
for manager, stores in manager_map.items():
    morning_total = len(stores) * 7  # 每家店7天
    evening_total = len(stores) * 7
    morning_done = 0
    evening_done = 0
    
    store_details = []
    for store in stores:
        if store in store_meetings:
            m_count = len(store_meetings[store]['morning'])
            e_count = len(store_meetings[store]['evening'])
            morning_done += m_count
            evening_done += e_count
            store_details.append({
                'name': store,
                'morning': m_count,
                'evening': e_count
            })
        else:
            store_details.append({
                'name': store,
                'morning': 0,
                'evening': 0
            })
    
    manager_stats[manager] = {
        'stores': len(stores),
        'morning_rate': round(morning_done / morning_total * 100) if morning_total > 0 else 0,
        'evening_rate': round(evening_done / evening_total * 100) if evening_total > 0 else 0,
        'store_details': store_details
    }

# 生成日期标签
date_labels = []
current = start_date
while current <= end_date:
    date_labels.append(f"{current.month}/{current.day}")
    current += timedelta(days=1)

# 生成JSON数据
output_data = {
    'date_range': f"{start_date.month}/{start_date.day}-{end_date.month}/{end_date.day}",
    'date_labels': date_labels,
    'managers': []
}

for manager, stats in manager_stats.items():
    output_data['managers'].append({
        'name': manager,
        'stores': stats['stores'],
        'morning_rate': stats['morning_rate'],
        'evening_rate': stats['evening_rate'],
        'store_details': stats['store_details']
    })

# 保存JSON
with open('poster_data_103.json', 'w', encoding='utf-8') as f:
    json.dump(output_data, f, ensure_ascii=False, indent=2)

print(f"\n数据统计完成！")
print(f"  - 经理数: {len(manager_stats)}")
print(f"  - 门店数: {len(whitelist)}")
print(f"  - 日期范围: {output_data['date_range']}")

# 验证："深圳零跑中心万国城店" 不在统计中
target = "深圳零跑中心万国城店"
if target in store_meetings:
    print(f"\n警告: '{target}' 仍在统计中！")
else:
    print(f"\n✓ '{target}' 已过滤（不计入统计）")

