# 晨夕会海报项目记忆

## 核心规则（必须遵守）
1. **严格按白名单统计**：脚本内 WHITELIST_MAP 是唯一数据源（2026-10-01 版：108家门店）
2. **Excel中但不在白名单的门店不计数**（如"深圳零跑中心万国城店"）
3. **白名单中的正确名称**："深圳零跑中心万国城旗舰店"（罗捷）
4. **统计周期**：昨天往前7天（动态计算，脚本自动过滤日期范围外的数据）
5. **海报结构**：上方=11位经理汇总，下方=108家门店明细

## 脚本修复记录
- **2026-07-05**：修复日期逻辑bug
  - 问题：脚本之前统计Excel中所有日期（可能8天），导致数据错误
  - 修复：在`process_excel()`中添加日期范围过滤，只统计"昨天往前7天"的数据
  - 影响：昨日夕会率从错误显示1%修正为正确的81%
- **2026-08-18**：门店名零宽字符修复（normalize_name 去除 \u200b\u200c\u200d\ufeff\u00a0）
- **2026-10-01**：白名单大改版（103→108家，12→11经理）

## 白名单概要（2026-10-01 版）
- 108家门店，11位零售高级经理
- 经理分布: 黄伟峰(12) 熊俊晖(12) 罗捷(7) 贾迪赫(6) 余子恩(19) 方任昊(12) 秦启超(9) 胡浩(5) 陈恒惕(6) 李传华(9) 李博恩(11)
- 经理显示顺序（MANAGER_ORDER）: 黄伟峰, 熊俊晖, 罗捷, 贾迪赫, 余子恩, 方任昊, 秦启超, 胡浩, 陈恒惕, 李传华, 李博恩
- 关键变化: 庄文迪→李传华；熊俊晖(代)→陈恒惕；汕尾→陈恒惕；韶关2家→余子恩；新增门店: 广州南站、观澜湖新城、陈宜禧路、梅州环市西路、阳江新达城

## Web 自动下载通道（2026-10-11 建成，端到端验证通过）
- **一键脚本**: `web_update_poster.sh`（workspace 根目录）——登录→拉数据→生成xlsx→跑海报→push，全自动
- **系统**: 零跑partner「零小跑」 https://partner.leapmotor.com/partner/morningEveningManagement/auditList
  - 登录页 /login，纯账号密码**无验证码**，可自动化；凭证在 `.leap_cred.json`（600权限，已gitignore，绝不进仓库）
  - auditList 主区是 qiankun 微前端，首次渲染慢需 sleep≈14s 或重开一次
- **数据接口**: POST `https://psg.leapmotor.com/brand/meetingList/pcPageList`
  - body: `{page,size:300,meetingTypeList:[],meetingStatusList:[],regionCodeList:[],districtCodeList:[]}`
  - 必带头: `Authorization: Bearer <JWT>` + `X-CROSS-NODE-ID: <同JWT>` + `Accepte: application/vnd.leapx.v2+json`
  - JWT 有效期 7 天（iss=leapmotor_portal），token 从页面 XHR hook（setRequestHeader）捕获，不落盘
  - 记录字段映射: meetingTypeLabel→B列, storeName→D列, uploadTime→G列(会议日期)；约150条/天
- **配套**: `web_fetch_to_xlsx.py` 合并 /tmp/leap_p*.raw → 晨夕会审核列表<date>_web.xlsx（gitignore）
- **验证**: 2026-10-11 web 通道与手工 Excel 结果完全一致（晨会84%/夕会79%）
- **已知风险**: git push 凭据走 macOS 钥匙串，沙箱安全中心可能拦截（Authentication failed/failed to erase）；失败时需内嵌 token 的 remote URL 或用户授权
- **运行环境**: 脚本须用 venv python `/Users/mei/.workbuddy/binaries/python/envs/default/bin/python`（openpyxl 所在）

## 部署信息
- GitHub仓库: Dorameimeimon/leapmotor-poster
- GitHub Pages链接: https://dorameimeimon.github.io/leapmotor-poster/
- Git推送格式: https://dorameimeimon:{token}@github.com/Dorameimeimon/leapmotor-poster.git main
- 数据处理脚本: auto_generate_poster.py（自动识别日期、过滤白名单、计算完成率）

## Excel数据结构
- 文件位置: /Users/mei/Downloads/晨夕会审核列表2026-06-24.xlsx
- 关键列: B=晨夕会类型, D=门店, G=会议日期
- 日期格式: datetime对象或字符串

## 输出偏好（2026-08-07 用户明确要求）
- 给出在线网址（GitHub Pages 链接）时，**只发链接本身，后面不要跟任何说明文字**（如"如未刷新，按…强刷"等）
- 其他数据说明（周期、完成率、门店数等）照常放在链接之前或单独表述，但链接行保持纯净