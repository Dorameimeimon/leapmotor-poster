#!/bin/zsh
# ============================================================
# 每日晨夕会海报自动更新（web 通道，全自动）
# 流程: 登录零跑partner系统 -> 带日期过滤拉取晨夕会审核数据
#       -> 生成xlsx -> auto_generate_poster.py -> git push GitHub Pages
# 口径: 会议日期 = 昨天往前7天 (如今天10/11 -> 取10/04~10/10)
#       服务端参数: uploadTimeStart="<起> 00:00:00", uploadTimeEnd="<止> 23:59:59"
# 依赖: .leap_cred.json (凭证, 已gitignore) / agent-browser / openpyxl
# ============================================================
set -e
setopt nonomatch 2>/dev/null || true   # zsh: 空glob不报错
WS="/Users/mei/WorkBuddy/2026-06-23-10-23-34"
export PATH="/Users/mei/.workbuddy/binaries/node/versions/22.22.2-6/bin:$PATH"
AG="/Users/mei/.workbuddy/binaries/node/versions/22.22.2-6/bin/agent-browser"
PY="/Users/mei/.workbuddy/binaries/python/envs/default/bin/python"
cd "$WS"
TODAY=$(date +%F)
DEND=$(date -v-1d +%Y-%m-%d)      # 昨天
DSTA=$(date -v-7d +%Y-%m-%d)      # 昨天往前第7天
AUDIT_URL="https://partner.leapmotor.com/partner/morningEveningManagement/auditList"
echo ">>> 统计窗口: ${DSTA} ~ ${DEND}"

# 注入请求钩子（捕获 Authorization / X-CROSS-NODE-ID）
HOOK_JS="window.__hdrs=[];var osr=XMLHttpRequest.prototype.setRequestHeader;XMLHttpRequest.prototype.setRequestHeader=function(k,v){window.__hdrs.push(k+': '+v);return osr.apply(this,arguments)};'hooked'"
# 点击查询（优先 button 文本，回退到微应用内元素）
CLICK_JS="(function(){var b=[].slice.call(document.querySelectorAll('button')).find(function(x){return x.innerText.trim()==='查询'});if(b){b.click();return 'clicked'}var mc=document.getElementById('microAppContainer');if(mc){var el=[].slice.call(mc.querySelectorAll('*')).filter(function(e){return (e.textContent||'').trim()==='查询'&&e.children.length<=2}).pop();if(el){el.click();return 'clicked-fallback'}}return 'NOBTN'})()"

# --- 1. 打开系统，检测登录态 ---
OUT=$($AG open "$AUDIT_URL" 2>&1 | tail -2) || true
echo "$OUT" | tail -2
if echo "$OUT" | grep -q "/login"; then
  echo ">>> 未登录，自动登录..."
  $AG wait --load load > /dev/null 2>&1 || true
  CRED=$($PY -c "import json;d=json.load(open('$WS/.leap_cred.json'));print(d['user']);print(d['pass'])")
  LUSER=$(echo "$CRED" | sed -n 1p)
  LPASS=$(echo "$CRED" | sed -n 2p)
  REFS=$($AG snapshot -i 2>/dev/null)
  ACC_REF=$(echo "$REFS" | grep -oE 'textbox "请输入账号" \[ref=(e[0-9]+)\]' | grep -oE 'e[0-9]+' | head -1)
  PWD_REF=$(echo "$REFS" | grep -oE 'textbox "请输入密码" \[ref=(e[0-9]+)\]' | grep -oE 'e[0-9]+' | head -1)
  BTN_REF=$(echo "$REFS" | grep -oE 'button "登录" \[ref=(e[0-9]+)\]' | grep -oE 'e[0-9]+' | head -1)
  $AG type "$ACC_REF" "$LUSER" > /dev/null 2>&1
  $AG type "$PWD_REF" "$LPASS" > /dev/null 2>&1
  $AG click "$BTN_REF" > /dev/null 2>&1
  sleep 8
  # 登录后重新打开目标页（此时应已带登录态）
  $AG open "$AUDIT_URL" > /dev/null 2>&1 || true
fi

# --- 2. 注入钩子 + 轮询渲染 + 捕获鉴权头（最多5轮，每轮重载） ---
HASAUTH="NO"
for ATT in 1 2 3 4 5; do
  echo ">>> [轮 $ATT] 注入钩子并等待微应用渲染..."
  $AG eval "$HOOK_JS" > /dev/null 2>&1 || true
  sleep 16
  RENDER=$($AG eval "var mc=document.getElementById('microAppContainer');var t=(mc&&mc.innerText)||'';(t.indexOf('审核ID')>=0||t.indexOf('查询')>=0)?'OK':'BLANK'" 2>/dev/null | tr -d '"')
  if [ "$RENDER" != "OK" ]; then
    echo ">>> [轮 $ATT] 微应用未渲染, 重载..."
    $AG open "$AUDIT_URL" > /dev/null 2>&1 || true
    $AG wait --load load > /dev/null 2>&1 || true
    sleep 4
    continue
  fi
  # 微应用已渲染：触发一次查询，确保请求发出（初始自动加载可能已发，这里双保险）
  $AG eval "$CLICK_JS" > /dev/null 2>&1 || true
  sleep 5
  HASAUTH=$($AG eval "window.__hdrs.some(function(h){return h.indexOf('Authorization:')===0})?'YES':'NO'" 2>/dev/null | tr -d '"')
  if [ "$HASAUTH" = "YES" ]; then
    echo ">>> [轮 $ATT] 鉴权头已捕获"
    break
  fi
  echo ">>> [轮 $ATT] 鉴权头未捕获, 重载..."
  $AG open "$AUDIT_URL" > /dev/null 2>&1 || true
  $AG wait --load load > /dev/null 2>&1 || true
  sleep 4
done
[ "$HASAUTH" = "YES" ] || { echo "ERROR: 未能捕获鉴权头"; $AG close > /dev/null 2>&1 || true; exit 1; }

# --- 3. 分页拉取（服务端按会议日期过滤, 空页提前终止, 最多6页） ---
rm -f /tmp/leap_p*.raw
for P in 1 2 3 4 5 6; do
  $AG eval "(function(){var ah=window.__hdrs.find(function(h){return h.indexOf('Authorization:')===0});var xh=window.__hdrs.find(function(h){return h.indexOf('X-CROSS-NODE-ID:')===0});var a=ah.split(':').slice(1).join(':').trim();var x=xh.split(':').slice(1).join(':').trim();fetch('https://psg.leapmotor.com/brand/meetingList/pcPageList',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json','Authorization':a,'X-CROSS-NODE-ID':x,'Accepte':'application/vnd.leapx.v2+json'},body:JSON.stringify({page:$P,size:300,meetingTypeList:[],meetingStatusList:[],regionCodeList:[],districtCodeList:[],uploadTimeStart:'${DSTA} 00:00:00',uploadTimeEnd:'${DEND} 23:59:59'})}).then(function(r){return r.text()}).then(function(t){window.__pg=t}).catch(function(e){window.__pg='ERR:'+e});return 'go'})()" > /dev/null 2>&1
  sleep 4
  $AG eval "window.__pg || 'EMPTY'" > /tmp/leap_p$P.raw 2>&1
  SZ=$(wc -c < /tmp/leap_p$P.raw | tr -d ' ')
  echo ">>> page $P: $SZ bytes"
  # 空页/异常页提前终止
  if [ "$SZ" -lt 200 ] || grep -q '"records":\[\]' /tmp/leap_p$P.raw; then
    echo ">>> 第 $P 页为空, 停止翻页"
    break
  fi
done

# --- 4. 合并生成 xlsx ---
$PY "$WS/web_fetch_to_xlsx.py" "$TODAY"

# --- 5. 跑海报脚本 ---
$PY "$WS/auto_generate_poster.py" "$WS/晨夕会审核列表${TODAY}_web.xlsx" 2>&1 | tail -8

# --- 6. 提交推送 ---
cp "$WS/晨夕会海报.html" "$WS/index.html"
cp "$WS/晨夕会海报.html" "$WS/deploy/index.html"
git add -A
git commit -m "自动更新晨夕会海报 ${TODAY} (周期 ${DSTA}-${DEND}, web通道)" > /dev/null 2>&1 || echo ">>> 无变更"
if git -c http.proxy= -c https.proxy= -c http.version=HTTP/1.1 push origin main > /tmp/push_log.txt 2>&1; then
  echo ">>> PUSH_OK"
else
  echo ">>> PUSH_FAILED:"; tail -3 /tmp/push_log.txt
fi

# --- 7. 关浏览器释放资源 ---
$AG close > /dev/null 2>&1 || true
echo "DONE"
