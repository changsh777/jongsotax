"""신고내역조회 팝업 빠른 DOM 검사"""
import sys, io, json, time, requests

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

CDP_PORT = 9222

# 1. 열린 탭 목록
tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
print(f"탭 {len(tabs)}개:")
for t in tabs:
    print(f"  [{t.get('type')}] {t.get('url','')[:80]}")
    print(f"     ws: {t.get('webSocketDebuggerUrl','')[:60]}")

# 홈택스 탭 찾기
ht_tab = None
for t in tabs:
    if "hometax.go.kr" in t.get("url",""):
        ht_tab = t
        break

if not ht_tab:
    print("홈택스 탭 없음!")
    sys.exit(1)

print(f"\n홈택스 탭: {ht_tab['url'][:80]}")
ws_url = ht_tab["webSocketDebuggerUrl"]

# 2. WebSocket으로 DOM 검사
import websocket
import threading

results = {}
done = threading.Event()

def on_message(ws_conn, message):
    msg = json.loads(message)
    if msg.get("id") == 1:
        results["dom"] = msg.get("result", {})
        done.set()

def on_error(ws_conn, error):
    print(f"WS 오류: {error}")
    done.set()

ws_conn = websocket.WebSocketApp(ws_url,
    on_message=on_message,
    on_error=on_error,
    header={"Origin": "http://localhost:9222"})

t = threading.Thread(target=ws_conn.run_forever, daemon=True)
t.start()
time.sleep(1)

# JS 실행: input[type=button] 목록
ws_conn.send(json.dumps({
    "id": 1,
    "method": "Runtime.evaluate",
    "params": {
        "expression": """(function() {
    var btns = Array.from(document.querySelectorAll("input[type=button], button"));
    return btns.map(function(el) {
        var bg = window.getComputedStyle(el).backgroundColor;
        return {
            tag: el.tagName,
            id: el.id || '',
            value: (el.value || el.innerText || '').trim().slice(0,50),
            cls: (el.className || '').slice(0,50),
            bg: bg,
            vis: el.offsetWidth > 0
        };
    });
})()""",
        "returnByValue": True
    }
}))

if not done.wait(timeout=15):
    print("타임아웃!")
    sys.exit(1)

ws_conn.close()

dom_result = results.get("dom", {})
if "result" in dom_result and "value" in dom_result["result"]:
    buttons = dom_result["result"]["value"]
    print(f"\n버튼 {len(buttons)}개 발견:")
    for b in buttons:
        print(f"  {b['tag']} id={b['id'][:50]:50s} val={b['value'][:30]:30s} bg={b['bg']}")
else:
    print("결과:", json.dumps(dom_result, ensure_ascii=False)[:500])
