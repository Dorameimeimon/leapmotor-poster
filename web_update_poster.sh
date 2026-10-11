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
  $AG open "$AUDIT_URL" > /dev/null 2>&1 || true
fi

# --- 2. 等微应用渲染（空页面自动重载一次） ---
sleep 14
RENDER=$($AG eval "var mc=document.getElementById('microAppContainer');var t=(mc&&mc.innerText)||'';(t.indexOf('审核ID')>=0||t.indexOf('查询')>=0)?'OK':'BLANK'" 2>/dev/null | tr -d '"')
if [ "$RENDER" != "OK" ]; then
  echo ">>> 微应用未渲染, 重载重试..."
  $AG open "$AUDIT_URL" > /dev/null 2>&1 || true
  $AG wait --load load > /dev/null 2>&1 || true
  sleep 15
fi

# --- 3. 注入请求钩子 -> 点查询 -> 捕获鉴权头 ---
$AG eval "window.__hdrs=[];var osr=XMLHttpRequest.prototype.setRequestHeader;XMLHttpRequest.prototype.setRequestHeader=function(k,v){window.__hdrs.push(k+': '+v);return osr.apply(this,arguments)};'hooked'" > /dev/null 2>&1
CLICK_JS="(function(){var b=[].slice.call(document.querySelectorAll('button')).find(function(x){return x.innerText.trim()==='查询'});if(b){b.click();return 'clicked'}var mc=document.getElementById('microAppContainer');if(mc){var el=[].slice.call(mc.querySelectorAll('*')).filter(function(e){return (e.textContent||'').trim()==='查询'&&e.children.length<=2}).pop();if(el){el.click();return 'clicked-fallback'}}return 'NOBTN'})()"
$AG eval "$CLICK_JS" > /dev/null 2>&1
sleep 5
HASAUTH=$($AG eval "window.__hdrs.some(function(h){return h.indexOf('Authorization:')===0})?'YES':'NO'" 2>/dev/null | tr -d '"')
if [ "$HASAUTH" != "YES" ]; then
  echo ">>> 鉴权头未捕获, 重试查询..."
  $AG eval "$CLICK_JS" > /dev/null 2>&1
  sleep 6
  HASAUTH=$($AG eval "window.__hdrs.some(function(h){return h.indexOf('Authorization:')===0})?'YES':'NO'" 2>/dev/null | tr -d '"')
fi
[ "$HASAUTH" = "YES" ] || { echo "ERROR: 未能捕获鉴权头"; $AG close > /dev/null 2>&1 || true; exit 1; }
echo ">>> 鉴权头已捕获"

# --- 4. 分页拉取（服务端按会议日期过滤, 空页提前终止, 最多6页） ---
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

# --- 5. 合并生成 xlsx ---
$PY "$WS/web_fetch_to_xlsx.py" "$TODAY"

# --- 6. 跑海报脚本 ---
$PY "$WS/auto_generate_poster.py" "$WS/晨夕会审核列表${TODAY}_web.xlsx" 2>&1 | tail -8

# --- 7. 提交推送 ---
cp "$WS/晨夕会海报.html" "$WS/index.html"
cp "$WS/晨夕会海报.html" "$WS/deploy/index.html"
git add -A
git commit -m "自动更新晨夕会海报 ${TODAY} (周期 ${DSTA}-${DEND}, web通道)" > /dev/null 2>&1 || echo ">>> 无变更"
git -c http.proxy= -c https.proxy= -c http.version=HTTP/1.1 push origin main 2>&1 | tail -2 || echo ">>> PUSH_FAILED(凭据)"

# --- 8. 关浏览器释放资源 ---
$AG close > /dev/null 2>&1 || true
echo "DONE"
