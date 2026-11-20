/* Extra tools: Data check, Excel export, Backup/Restore, carry-forward, Install app (PWA) */
(function () {
  const $id = id => document.getElementById(id);
  const peso = n => (typeof money === 'function' ? money(n) : '₱' + Number(n).toFixed(2));
  const live = () => S.transactions.filter(t => t.status !== 'void' && t.date);
  const ymOf = d => d.slice(0, 7);
  const monthName = ym => new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  const dow = d => new Date(d + 'T00:00:00').getDay();
  const nextYm = ym => { const d = new Date(+ym.slice(0, 4), +ym.slice(5, 7), 1); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); };
  const prevYm = ym => { const d = new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 2, 1); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); };

  function monthStats() {
    const m = {};
    live().forEach(t => {
      const k = ymOf(t.date);
      const s = m[k] ||= { begin: 0, hasBegin: false, offer: 0, genExp: 0, alloc: 0, allowExp: 0 };
      if (t.type === 'income') {
        if (isAllowanceInc(t)) s.alloc += t.amount;
        else if (t.category === 'Beginning Balance') { s.begin += t.amount; s.hasBegin = true; }
        else s.offer += t.amount;
      } else if (t.status === 'approved' || !t.status) {
        if (isAllowanceExp(t)) s.allowExp += t.amount; else s.genExp += t.amount;
      }
    });
    Object.values(m).forEach(s => { s.net = s.begin + s.offer - s.genExp; s.allowLeft = s.alloc - s.allowExp; });
    return m;
  }

  /* ---------- Data check ---------- */
  function findIssues() {
    const issues = [], stats = monthStats(), tx = live();
    tx.forEach(t => {
      if (t.type === 'income' && t.category === 'Youth Night Offering' && dow(t.date) !== 5)
        issues.push({ lvl: 'warn', msg: `Youth Night Offering ${peso(t.amount)} is dated <b>${t.date}</b> which is not a Friday.` });
      if (t.type === 'income' && t.category === 'Joint Fellowship Offering' && dow(t.date) !== 0)
        issues.push({ lvl: 'info', msg: `Joint Fellowship ${peso(t.amount)} is dated <b>${t.date}</b> which is not a Sunday.` });
    });
    const seen = {};
    tx.forEach(t => { const k = [t.type, t.date, t.amount, t.category, t.fund || ''].join('|'); (seen[k] ||= []).push(t); });
    Object.values(seen).filter(a => a.length > 1).forEach(a =>
      issues.push({ lvl: 'warn', msg: `Possible duplicate ×${a.length}: ${a[0].type} ${peso(a[0].amount)} · ${a[0].category} · ${a[0].date}` }));
    Object.keys(stats).sort().forEach(k => {
      const s = stats[k], nk = nextYm(k), n = stats[nk];
      if (s.allowLeft < 0) issues.push({ lvl: 'bad', msg: `<b>${monthName(k)}</b>: allowance overspent by ${peso(-s.allowLeft)}.` });
      if (n && n.hasBegin && Math.abs(n.begin - s.net) > 0.009)
        issues.push({ lvl: 'warn', msg: `<b>${monthName(nk)}</b> starts with ${peso(n.begin)} but ${monthName(k)} ended with ${peso(s.net)} (difference ${peso(n.begin - s.net)}).` });
      if (n && !n.hasBegin)
        issues.push({ lvl: 'info', msg: `<b>${monthName(nk)}</b> has no beginning balance. Carry forward ${peso(s.net)} from ${monthName(k)}?`, carry: { ym: nk, amount: s.net } });
    });
    return issues;
  }
  window.__applyCarry = async (ym, amount) => {
    await window.saveBatchOfferings([{ type: 'income', fund: 'general', date: ym + '-01', amount, category: 'Beginning Balance', notes: 'Carried forward from ' + monthName(prevYm(ym)) }]);
    setTimeout(window.dataCheck, 400);
  };
  window.dataCheck = function () {
    const issues = findIssues(), icon = { bad: '🔴', warn: '🟠', info: '🔵' };
    openModal(`<h2>🩺 Data Check</h2>
      <p class="mute" style="margin-bottom:12px">${issues.length ? issues.length + ' item(s) to review.' : 'Everything looks consistent ✔'}</p>
      <div style="display:grid;gap:8px;max-height:55vh;overflow:auto">${issues.map(i => `
        <div style="background:rgba(255,255,255,.05);padding:10px 12px;border-radius:10px;font-size:13px;display:flex;gap:10px;align-items:center;justify-content:space-between">
          <span>${icon[i.lvl]} ${i.msg}</span>
          ${i.carry ? `<button class="btn sm" onclick="__applyCarry('${i.carry.ym}',${i.carry.amount})">Apply</button>` : ''}
        </div>`).join('')}</div>
      <div class="row" style="margin-top:16px"><button class="btn ghost" onclick="closeModal()">Close</button></div>`);
  };

  /* ---------- Excel export (opens in Excel) ---------- */
  window.exportExcel = function () {
    const stats = monthStats(), tx = live().sort((a, b) => a.date.localeCompare(b.date));
    const keys = Object.keys(stats).sort();
    if (!keys.length) return alert('No transactions to export yet.');
    const e = s => String(s ?? '').replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
    let h = `<html><head><meta charset="utf-8"></head><body><h2>Tagbilaran Bible Training Ministries, Inc.</h2><h3>COYAP Cash Accountability Report</h3>`;
    h += `<table border="1"><tr><th>Month</th><th>Beginning</th><th>Offerings</th><th>Expenses</th><th>Net Cash on Hand</th><th>Allowance Allocated</th><th>Allowance Spent</th><th>Allowance Left</th></tr>` +
      keys.map(k => { const s = stats[k]; return `<tr><td>${monthName(k)}</td><td>${s.begin}</td><td>${s.offer}</td><td>${s.genExp}</td><td>${s.net}</td><td>${s.alloc}</td><td>${s.allowExp}</td><td>${s.allowLeft}</td></tr>`; }).join('') + '</table><br>';
    keys.forEach(k => {
      h += `<h3>${monthName(k)}</h3><table border="1"><tr><th>Date</th><th>Type</th><th>Fund</th><th>Category</th><th>Notes</th><th>Amount</th></tr>` +
        tx.filter(t => ymOf(t.date) === k).map(t => `<tr><td>${t.date}</td><td>${t.type}</td><td>${(isAllowanceInc(t) || isAllowanceExp(t)) ? 'Allowance' : 'General'}</td><td>${e(t.category)}</td><td>${e(cleanNotes(t.notes))}</td><td>${t.type === 'expense' ? '-' : ''}${t.amount}</td></tr>`).join('') + '</table><br>';
    });
    download(`BTC-Youth-Report-${new Date().toISOString().slice(0, 10)}.xls`, h + '</body></html>', 'application/vnd.ms-excel');
  };

  /* ---------- Backup / Restore ---------- */
  function download(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }
  window.backupData = function () {
    const skippedSlots = typeof getSkipReasons === 'function' ? getSkipReasons() : {};
    const data = { app: 'btc-youth', version: 1, exported: new Date().toISOString(), currency: S.currency, transactions: S.transactions, events: S.events, members: S.members, incomeCats: S.incomeCats, expenseCats: S.expenseCats, skippedSlots };
    download(`BTC-Youth-backup-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(data, null, 2), 'application/json');
  };
  window.restoreData = function () {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.json,application/json';
    inp.onchange = async () => {
      try {
        const file = inp.files[0];
        if (!file) return;
        const d = JSON.parse(await file.text());
        if (d.app !== 'btc-youth' || !Array.isArray(d.transactions)) throw new Error('Not a valid BTC Youth JSON backup file.');

        const txs = d.transactions || [];
        const evs = d.events || [];
        const mems = d.members || [];

        if (typeof sb === 'undefined') {
          if (!confirm(`Restore "${file.name}"?\n\nThis will replace current local data with ${txs.length} transactions, ${evs.length} events, and ${mems.length} members.`)) return;
          Object.assign(S, { currency: d.currency || S.currency, transactions: txs, events: evs, members: mems, incomeCats: d.incomeCats || S.incomeCats, expenseCats: d.expenseCats || S.expenseCats });
          if (d.skippedSlots && typeof d.skippedSlots === 'object') {
            localStorage.setItem('btc-skipped-slots', JSON.stringify(d.skippedSlots));
          }
          save();
          if (typeof showToast === 'function') showToast(`✔ Restored ${txs.length} transactions`);
          return;
        }

        // Cloud (Supabase) import modal: Merge or Replace
        openModal(`<h2>♻️ Import Backup (JSON)</h2>
          <p class="mute" style="margin-bottom:10px">Loaded <b>${esc(file.name)}</b>:</p>
          <div style="background:rgba(255,255,255,.05);padding:12px 14px;border-radius:12px;font-size:13px;margin-bottom:14px">
            • <b>${txs.length}</b> transactions<br>
            • <b>${evs.length}</b> events<br>
            • <b>${mems.length}</b> members
          </div>
          <p class="mute" id="json-imp-status" style="margin-bottom:12px">Choose how to import into your Supabase database:</p>
          <div style="display:grid;gap:10px">
            <button type="button" class="btn" id="json-imp-merge" style="background:linear-gradient(135deg,#06b6d4,#3b82f6)">➕ Merge (Add new entries, skip duplicates)</button>
            <button type="button" class="btn alt" id="json-imp-replace">🔄 Replace All Cloud Data with Backup</button>
            <button type="button" class="btn ghost" onclick="closeModal()">Cancel</button>
          </div>`);

        const runCloudImport = async (replaceAll) => {
          const st = document.getElementById('json-imp-status');
          const b1 = document.getElementById('json-imp-merge');
          const b2 = document.getElementById('json-imp-replace');
          if (b1) b1.disabled = true;
          if (b2) b2.disabled = true;
          try {
            if (replaceAll) {
              if (!confirm('Replace ALL current Supabase transactions, events, and members with this backup?')) {
                if (b1) b1.disabled = false;
                if (b2) b2.disabled = false;
                return;
              }
              if (st) st.textContent = 'Clearing existing cloud data…';
              await sb.from('audit_log').delete().not('id', 'is', null);
              await sb.from('transactions').delete().not('id', 'is', null);
              await sb.from('events').delete().not('id', 'is', null);
              await sb.from('members').delete().not('id', 'is', null);
              const { data: leftover } = await sb.from('transactions').select('id').neq('status', 'void');
              if (leftover && leftover.length) {
                await sb.from('transactions').update({ status: 'void' }).in('id', leftover.map(r => r.id));
              }
            }

            if (st) st.textContent = 'Importing categories, events, and members…';
            // Categories
            for (const [type, mine, theirs] of [['income', S.incomeCats, d.incomeCats || []], ['expense', S.expenseCats, d.expenseCats || []]]) {
              const add = theirs.filter(c => c && !mine.includes(c));
              if (add.length) await sb.from('categories').insert(add.map(name => ({ type, name })));
            }

            // Events
            const evMap = {};
            const existingEvByName = Object.fromEntries((replaceAll ? [] : S.events).map(e => [e.name.toLowerCase(), e.id]));
            for (const e of evs) {
              const hit = existingEvByName[String(e.name || '').toLowerCase()];
              if (hit) { evMap[e.id] = hit; continue; }
              const res = await sb.from('events').insert({ name: e.name, budget: Number(e.budget) || 0 }).select('id').single();
              if (res.error) throw res.error;
              evMap[e.id] = res.data.id;
            }

            // Members
            const memMap = {};
            const existingMemByName = Object.fromEntries((replaceAll ? [] : S.members).map(m => [m.name.toLowerCase(), m.id]));
            for (const m of mems) {
              const hit = existingMemByName[String(m.name || '').toLowerCase()];
              if (hit) { memMap[m.id] = hit; continue; }
              const res = await sb.from('members').insert({ name: m.name, contact: m.contact || '', monthly: Number(m.monthly) || 0, due: Number(m.due) || 0 }).select('id').single();
              if (res.error) throw res.error;
              memMap[m.id] = res.data.id;
            }

            // Transactions
            const dupKey = t => [t.type, t.date, Number(t.amount).toFixed(2), t.category || ''].join('|');
            const have = replaceAll ? new Set() : new Set(S.transactions.filter(t => t.status !== 'void').map(dupKey));
            const rows = [];
            for (const t of txs) {
              if (!t.date || !(Number(t.amount) > 0) || t.status === 'void') continue;
              if (!replaceAll && have.has(dupKey(t))) continue;
              let notes = t.notes || '';
              if (t.fund === 'allowance' && !notes.includes('[Allowance]')) {
                notes = notes ? `[Allowance] ${notes}` : '[Allowance]';
              }
              rows.push({
                type: t.type === 'expense' ? 'expense' : 'income',
                date: t.date,
                amount: Number(t.amount),
                category: t.category || (t.type === 'expense' ? 'Other' : 'Offering'),
                event_id: evMap[t.eventId] || null,
                member_id: memMap[t.memberId] || null,
                notes,
                status: 'approved'
              });
            }

            if (st) st.textContent = `Uploading ${rows.length} transactions…`;
            const uploadedIds = [];
            for (let i = 0; i < rows.length; i += 100) {
              const chunk = rows.slice(i, i + 100);
              const res = await sb.from('transactions').insert(chunk).select('id');
              if (res.error) throw res.error;
              if (res.data) uploadedIds.push(...res.data.map(r => r.id));
            }
            if (uploadedIds.length && typeof window.recordUploadBatch === 'function') {
              window.recordUploadBatch(uploadedIds, `JSON Backup (${uploadedIds.length})`);
            }

            if (d.skippedSlots && typeof d.skippedSlots === 'object') {
              localStorage.setItem('btc-skipped-slots', JSON.stringify(d.skippedSlots));
              if (typeof window.onPrefsChanged === 'function') window.onPrefsChanged();
            }
            if (d.currency) {
              S.currency = d.currency;
              localStorage.setItem('btc-cloud-currency', d.currency);
            }
            closeModal();
            if (typeof refresh === 'function') await refresh();
            if (typeof showToast === 'function') showToast(`✔ Imported ${rows.length} transactions from backup`);
          } catch (err) {
            if (st) { st.style.color = 'var(--bad)'; st.textContent = 'Import failed: ' + (err.message || err); }
            if (b1) b1.disabled = false;
            if (b2) b2.disabled = false;
          }
        };

        document.getElementById('json-imp-merge').onclick = () => runCloudImport(false);
        document.getElementById('json-imp-replace').onclick = () => runCloudImport(true);
      } catch (err) { alert('Could not restore: ' + err.message); }
    };
    inp.click();
  };

  /* ---------- Install app (PWA) ---------- */
  let deferred = null;
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferred = e; });
  window.installApp = async function () {
    if (deferred) { deferred.prompt(); await deferred.userChoice; deferred = null; return; }
    if (window.matchMedia('(display-mode: standalone)').matches) return alert('The app is already installed.');
    openModal(`<h2>📲 Install on your phone</h2>
      <p class="mute" style="margin:8px 0">Open this site in your phone browser, then:</p>
      <p><b>Android (Chrome):</b> tap ⋮ menu → <b>Install app</b> / <b>Add to Home screen</b>.</p>
      <p style="margin-top:8px"><b>iPhone (Safari):</b> tap Share ⬆ → <b>Add to Home Screen</b>.</p>
      <p class="mute" style="margin-top:10px;font-size:12px">Installing needs the site to be opened over https (or localhost).</p>
      <div class="row" style="margin-top:16px"><button class="btn ghost" onclick="closeModal()">Close</button></div>`);
  };
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol))
    navigator.serviceWorker.register('sw.js').catch(() => { });

  /* Tools are now presented inside the clean "Tools & Data" modal in app-core.js */
})();
