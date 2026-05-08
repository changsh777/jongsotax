"""신고내역 팝업 첫 행 col[9~14] 링크 구조 검사 — 팝업이 이미 열려 있을 때 실행"""
import asyncio, json, websockets, requests, sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

SELECT_ROWNUM = "mf_txppWframe_UTERNAAZ0Z31_wframe_edtGrdRowNum"

async def main():
    tabs = requests.get('http://localhost:9222/json').json()
    ht = next((t for t in tabs if 'websquare.html' in t.get('url','') and 'sesw.' not in t.get('url','')), None)
    print('홈택스 탭:', ht['url'][:80] if ht else None)

    async with websockets.connect(ht['webSocketDebuggerUrl']) as ws:
        async def ev(code, cid=1):
            await ws.send(json.dumps({'id':cid,'method':'Runtime.evaluate',
                'params':{'expression':code,'returnByValue':True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=30))
                if r.get('id')==cid:
                    return r.get('result',{}).get('result',{}).get('value')

        # 팝업 열려있는지 확인
        sel_check = await ev(f'document.getElementById("{SELECT_ROWNUM}") ? "open" : "closed"', 1)
        print('팝업 상태:', sel_check)

        if sel_check != 'open':
            print('팝업이 닫혀있어요. 홈택스 신고내역조회 팝업을 먼저 열어주세요.')
            return

        # 데이터 행 확인
        row_count = await ev(r"""(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    return container.querySelectorAll('table tbody tr').length;
})()""", 2)
        print('행 수:', row_count)

        # 첫 번째 데이터 행 col[9~14] 링크 구조
        result = await ev(r"""(function(){
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(r){ return r.querySelectorAll('td').length >= 38; });
    if (!rows.length) return 'no_rows_with_38cols, total=' + container.querySelectorAll('table tbody tr').length;
    var row = rows[0];
    var tds = row.querySelectorAll('td');
    var info = [];
    var cols = [9,10,11,12,13,14,36,37,38,39];
    for (var i=0; i<cols.length; i++) {
        var ci = cols[i];
        var td = tds[ci];
        if (!td) { info.push('col['+ci+']: MISSING'); continue; }
        var txt = td.innerText.trim().slice(0,30);
        var els = Array.from(td.querySelectorAll('a, input, button'));
        var elInfo = els.map(function(el){
            return el.tagName
                + '|id:' + (el.id||'').slice(0,50)
                + '|val:' + (el.value||el.innerText||el.textContent||'').trim().slice(0,20)
                + '|onclick:' + (el.getAttribute('onclick')||'none').slice(0,60)
                + '|href:' + (el.getAttribute('href')||'none').slice(0,30);
        });
        info.push('col['+ci+']: ['+txt+'] '+( elInfo.join(' | ') || '(no elements)' ));
    }
    return info.join('\n');
})()""", 3)
        print(result)

asyncio.run(main())
