// Read the BridgeTest page state over CDP (port 9333) using Node's built-in WebSocket.
const http = require('http');

function getTargets() {
    return new Promise((resolve, reject) => {
        http.get('http://localhost:9333/json', res => {
            let d = '';
            res.on('data', c => d += c);
            res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { reject(e); } });
        }).on('error', reject);
    });
}

function evaluate(wsUrl, expr) {
    return new Promise((resolve, reject) => {
        const ws = new WebSocket(wsUrl);
        let idc = 0;
        const pending = {};
        ws.onmessage = (ev) => {
            const m = JSON.parse(ev.data);
            if (m.id && pending[m.id]) { pending[m.id](m.result || {}); delete pending[m.id]; }
        };
        ws.onerror = e => reject(new Error('ws error'));
        ws.onopen = () => {
            idc++;
            const myId = idc;
            pending[myId] = (r) => { resolve(r.result ? r.result.value : null); try{ws.close();}catch(_){}};
            ws.send(JSON.stringify({ id: myId, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true } }));
        };
        setTimeout(() => reject(new Error('eval timeout')), 8000);
    });
}

(async () => {
    const targets = await getTargets();
    const page = targets.find(t => t.type === 'page' && /test\.html/.test(t.url)) || targets.find(t => t.type === 'page');
    if (!page) { console.log('NO PAGE TARGET. Targets:', JSON.stringify(targets.map(t=>({type:t.type,url:t.url})))); process.exit(2); }
    const expr = `(() => {
        const rows = [...document.querySelectorAll('#results tr')].slice(1).map(tr => [...tr.children].map(td=>td.textContent.trim()).join(' | '));
        const stats = ['recv='+stRecv.textContent,'matched='+stMatched.textContent,'pushes='+stPushes.textContent,'listener='+stListener.textContent];
        return JSON.stringify({ url: location.href, title: document.title, rows, stats });
    })()`;
    try {
        console.log(await evaluate(page.webSocketDebuggerUrl, expr));
    } catch (e) { console.error('EVAL FAILED:', e.message); process.exit(3); }
})();
