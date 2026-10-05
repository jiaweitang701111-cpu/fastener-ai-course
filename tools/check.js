// 版面自動檢查：逐頁找文字重疊、壓到分隔線、超出色塊或版面邊界。
// 用法（先在專案資料夾啟動 python -m http.server 8322）：
//   chrome --headless --disable-gpu --virtual-time-budget=30000 --dump-dom "http://localhost:8322/index.html?check=1"
// 結果寫在頁面最後的 <pre id="check-out">。
const pre = document.body.appendChild(Object.assign(document.createElement('pre'), { id: 'check-out', textContent: 'RUNNING' }));
try { run(); } catch (e) { pre.textContent = 'ERROR ' + e.message + ' / ' + e.stack; }
function run() {
  const D = window.__deck, out = [];
  const tcount = {}, sizes = [];
  for (let i = 0; i < D.S.length; i++) {
    const s = D.S[i];
    if (s.classList.contains('ix')) continue;
    D.go(i, true);
    if (s.dataset.tmin) tcount[s.dataset.tmin] = (tcount[s.dataset.tmin] || 0) + 1;
    // 內文（原字級小於 21px 的文字）排版後實際的最小、最大字級
    const px = [...s.querySelectorAll('[data-c] span')].filter(sp => sp.textContent.trim().length > 3 && parseFloat(getComputedStyle(sp).getPropertyValue('--z')) < 21).map(sp => parseFloat(getComputedStyle(sp).fontSize));
    if (px.length) sizes.push(s.dataset.src + ':' + Math.round(Math.min(...px)) + '-' + Math.round(Math.max(...px)));
    const sr = s.getBoundingClientRect();
    const P = q => ({ l: q.left - sr.left, t: q.top - sr.top, r: q.right - sr.left, b: q.bottom - sr.top });
    const tbs = [...s.querySelectorAll('.tb')].map(tb => {
      let l = 1e9, t = 1e9, r = -1e9, b = -1e9;
      tb.querySelectorAll('span').forEach(sp => { for (const q of sp.getClientRects()) { l = Math.min(l, q.left); t = Math.min(t, q.top); r = Math.max(r, q.right); b = Math.max(b, q.bottom); } });
      return { tb, box: P(tb.getBoundingClientRect()), c: P({ left: l, top: t, right: r, bottom: b }), txt: tb.textContent.slice(0, 12) };
    });
    const rules = [...s.querySelectorAll('i.e')].map(el => P(el.getBoundingClientRect())).filter(r => r.b - r.t < 7 || r.r - r.l < 7);
    const issues = [];
    tbs.forEach((a, x) => {
      if (a.c.b > 712) issues.push('超出下緣:' + a.txt);
      if (a.c.r > 1262) issues.push('超出右緣:' + a.txt);
      if (a.tb.style.background && (a.c.b > a.box.b + 1 || a.c.t < a.box.t - 1)) issues.push('超出色塊:' + a.txt);
      rules.forEach(r => { if (Math.min(a.c.r, r.r) - Math.max(a.c.l, r.l) > 3 && Math.min(a.c.b - 5, r.b) - Math.max(a.c.t + 5, r.t) > 0) issues.push('壓到分隔線:' + a.txt); });
      tbs.forEach((b, y) => {
        if (y <= x) return;
        const ww = Math.min(a.c.r, b.c.r) - Math.max(a.c.l, b.c.l), hh = Math.min(a.c.b, b.c.b) - Math.max(a.c.t, b.c.t);
        if (ww > 3 && hh > 4) issues.push('文字重疊:' + a.txt + '|' + b.txt);
      });
    });
    // 內容不能壓到頁尾來源列；卡片裡的字不能超出卡片下緣
    const src = tbs.filter(a => /^(資料|案例)?來源|^資料分級|^依 OWASP|^Skill /.test(a.tb.textContent) && a.c.t > 560)[0];
    const cards = [...s.querySelectorAll('i.e[data-g]')].map(el => P(el.getBoundingClientRect())).filter(r => r.b - r.t > 40 && r.r - r.l > 60);
    [...s.querySelectorAll('[data-g]')].forEach(el => { const r = P(el.getBoundingClientRect()); if (src && r.b > src.c.t + 1 && r.t < src.c.t) issues.push('壓到來源列:' + (el.textContent || el.tagName).slice(0, 10)); });
    tbs.forEach(a => { if (!a.tb.dataset.g) return;
      const card = cards.filter(r => a.c.l >= r.l - 2 && a.c.r <= r.r + 2 && a.c.t >= r.t - 2 && a.c.t < r.b).sort((x, y) => (x.r - x.l) * (x.b - x.t) - (y.r - y.l) * (y.b - y.t))[0];
      if (card && a.c.b > card.b + 1) issues.push('超出卡片:' + a.txt); });
    s.querySelectorAll('.tbl').forEach(t => {
      const tr = P(t.getBoundingClientRect());
      if (tr.b > 690) issues.push('表格超出下緣');
      tbs.forEach(a => { if (a.c.t < tr.b - 3 && a.c.b > tr.t + 3 && a.c.l < tr.r && a.c.r > tr.l) issues.push('文字壓到表格:' + a.txt); });
    });
    if (issues.length) out.push('p.' + (i + 1) + '（原 p.' + s.dataset.src + '，t=' + (s.dataset.tmin || '-') + '）' + [...new Set(issues)].slice(0, 6).join('；'));
  }
  pre.textContent = 'DONE 頁數 ' + D.S.length + '；各頁最小放大程度 t 的分布 ' + JSON.stringify(tcount) + '；有問題的頁 ' + out.length + '\n' + out.join('\n') + '\n原頁碼:內文字級px（最小-最大） ' + sizes.join(' ');
}
