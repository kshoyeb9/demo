"""
make_cash_page.py — Builds cash.html, the DAILY cash dashboard shell.

The stylesheet and the chart library are lifted verbatim from index.html so the
two pages stay visually identical siblings. Re-run this after changing the
monthly page's design to carry the change across.

    python make_cash_page.py            # writes cash.html with a DATA placeholder
    python etl_cash.py --file ...       # fills that placeholder in

Usage:
    python make_cash_page.py [--from index.html] [--out cash.html]
"""

import argparse, re, sys
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--from", dest="src", default="index.html")
ap.add_argument("--out", default="cash.html")
a = ap.parse_args()

src = Path(a.src).read_text(encoding="utf-8")

# ── lift the shared pieces ────────────────────────────────────────────────────
# The published file begins with the platform's own reset <style>; take the block
# that actually carries our design tokens, not the first one in the file.
blocks = re.findall(r"<style>.*?</style>", src, re.DOTALL)
STYLE = next((b for b in blocks if "--dark-blue" in b), None)
if STYLE is None:
    sys.exit(f"No design-token <style> block found in {a.src} "
             f"({len(blocks)} style block(s) present)")

i = src.index("/* ───────── helpers ───────── */")
j = src.index("/* ═════════ period model")
HELPERS = src[i:j].rstrip()

fonts = re.search(r'(<link rel="preconnect".*?&display=swap">)', src, re.DOTALL)
FONTS = fonts.group(1) if fonts else ""

PAGE = f"""<title>Noon Cash Position</title>
{FONTS}
{STYLE}
<style>
  /* Daily-only additions */
  .mast h1 .kicker{{display:block;font-size:11px;font-weight:700;letter-spacing:.12em;
    text-transform:uppercase;color:var(--accent);margin-bottom:2px}}
  .asat{{font-size:12.5px;color:var(--band-muted)}}
  .asat b{{color:var(--band-ink);font-weight:600}}
  .fc-table td:first-child{{white-space:nowrap}}
  tr.trough td{{background:var(--bad-bg)!important;font-weight:600}}
  tr.trough td:first-child::after{{content:"LOW";margin-left:8px;font-size:9.5px;font-weight:700;
    color:var(--bad-ink);border:1px solid var(--orange);padding:0 4px;vertical-align:1px}}
</style>

<header class="mast">
  <div class="mast-in">
    <div>
      <div class="eyebrow">Noon Academy · Treasury</div>
      <h1>Cash Position</h1>
      <div class="sub">Bank balance, daily movement and the thirteen-week outlook. Updated each
        business day, independently of the monthly management pack.</div>
      <div class="mast-right" style="margin-top:8px">All figures in
        <b id="m-cur">USD millions</b> · refreshed <b id="m-gen"></b></div>
      <div class="cur-wrap">
        <label for="cur-usd">Currency</label>
        <span class="quick" id="cur-toggle" role="group" aria-label="Display currency">
          <button type="button" id="cur-usd" data-cur="USD" aria-pressed="true">USD</button>
          <button type="button" data-cur="SAR" aria-pressed="false">SAR</button>
        </span>
        <span class="cur-rate" id="cur-rate"></span>
      </div>
    </div>
    <div class="period">
      <div class="asat">Position as at <b id="m-asat"></b></div>
      <div class="asat" id="m-cover"></div>
    </div>
  </div>
</header>

<div class="frame">
  <nav class="rail" aria-label="Sections">
    <a href="#s0"><span class="n">—</span>Position today</a>
    <a href="#s1"><span class="n">1</span>Daily movement</a>
    <a href="#s2"><span class="n">2</span>13-week outlook</a>
    <a href="#s3"><span class="n">3</span>Receivables</a>
    <a href="#s4"><span class="n">4</span>Notes</a>
    <div class="rail-foot">Actuals run to the as-at date. Everything from the next Monday on is
      forecast and only as good as its last revision.</div>
  </nav>

  <main>
    <section class="sec" id="s0">
      <div class="sec-head"><span class="n">—</span><h2>Position today</h2>
        <span class="scope" id="es-scope"></span></div>
      <div class="panels">
        <div class="panel wide">
          <div class="panel-h"><span class="eyebrow">Cash &amp; liquidity</span>
            <span class="when" id="p-when"></span></div>
          <div class="tiles six" id="tiles"></div>
        </div>
      </div>
    </section>

    <section class="sec" id="s1">
      <div class="sec-head"><span class="n">1</span><h2>Daily movement</h2>
        <span class="scope" id="led-scope"></span></div>
      <div class="card">
        <div class="card-h"><h3>Closing bank balance</h3><span class="sub" id="bal-scope"></span></div>
        <div class="legend"><span><i class="line" style="background:var(--s1)"></i>Closing balance</span>
          <span><i class="dash" style="color:var(--orange)"></i>Operating floor</span></div>
        <div class="chart" id="ch-bal"></div>
      </div>
      <div class="grid g2" style="margin-top:16px">
        <div class="card">
          <div class="card-h"><h3>Receipts vs payments</h3><span class="sub" id="rp-scope"></span></div>
          <div class="legend"><span><i style="background:var(--s1)"></i>Receipts</span>
            <span><i style="background:var(--s3)"></i>Payments</span></div>
          <div class="chart" id="ch-rp"></div>
        </div>
        <div class="card">
          <div class="card-h"><h3>Net movement by week</h3><span class="sub" id="nw-scope"></span></div>
          <div class="legend"><span><i style="background:var(--s2)"></i>Net movement</span></div>
          <div class="chart" id="ch-net"></div>
        </div>
      </div>
    </section>

    <section class="sec" id="s2">
      <div class="sec-head"><span class="n">2</span><h2>13-week outlook</h2>
        <span class="scope" id="fc-scope"></span></div>
      <div class="card">
        <div class="card-h"><h3>Projected closing balance</h3><span class="sub" id="fcb-scope"></span></div>
        <div class="legend"><span><i class="line" style="background:var(--s1)"></i>Projected balance</span>
          <span><i class="dash" style="color:var(--orange)"></i>Operating floor</span></div>
        <div class="chart" id="ch-fc"></div>
        <div class="note" id="fc-note"></div>
      </div>
      <div class="card" style="margin-top:16px">
        <h3 class="tbl-title">Week by week</h3>
        <div class="tbl-wrap"><table id="t-fc" class="fc-table"></table></div>
        <div class="note">Receipts and payments are expected amounts. The lowest projected balance
          is marked. Roll the workbook forward each Monday.</div>
      </div>
    </section>

    <section class="sec" id="s3">
      <div class="sec-head"><span class="n">3</span><h2>Receivables</h2>
        <span class="scope" id="ar-scope"></span></div>
      <div class="grid g-aging">
        <div class="card">
          <div class="card-h"><h3>Aging</h3><span class="sub" id="ag-scope"></span></div>
          <div class="legend"><span><i style="background:var(--s1)"></i>Current</span>
            <span><i style="background:var(--s2)"></i>31–60 days</span>
            <span><i style="background:var(--s3)"></i>60+ days</span></div>
          <div class="chart" id="ch-ag"></div>
        </div>
        <div class="card">
          <div class="card-h"><h3>By contract</h3><span class="sub" id="ct-scope"></span></div>
          <div class="legend"><span><i style="background:var(--s1)"></i>Outstanding balance</span></div>
          <div class="chart" id="ch-ct"></div>
        </div>
      </div>
    </section>

    <section class="sec" id="s4">
      <div class="sec-head"><span class="n">4</span><h2>Notes</h2>
        <span class="scope">Written alongside the daily update</span></div>
      <div class="updates" id="updates"></div>
    </section>

    <footer>
      <span>Source: <span id="f-src"></span> · refreshed <span id="f-gen"></span></span>
      <span>Noon Academy · internal management information</span>
    </footer>
  </main>
</div>
<div id="tip" role="status" aria-live="polite"></div>

<script>
const DATA = {{}};

{HELPERS}

/* ═════════ currency ═════════
   Kept identical to the monthly page: the workbook is USD, and the dataset is
   converted once per toggle so chart scales land on round numbers either way.
   Ratios and day counts are not currency and are never scaled. */
const FX={{USD:{{rate:1,sym:'$',unit:'USD'}}, SAR:{{rate:3.75,sym:'SAR ',unit:'SAR'}}}};
let CUR='USD';
const SY=()=>FX[CUR].sym, UNIT=()=>FX[CUR].unit+' M';
const D0=DATA, M=D0.meta;
let D=D0;

function convertFX(src,k,sym){{
  if(k===1) return src;
  const d=structuredClone(src);
  const arr=a=>Array.isArray(a)?a.map(v=>typeof v==='number'?v*k:v):a;
  const mul=(o,ks)=>{{ if(o) ks.forEach(f=>{{ if(typeof o[f]==='number') o[f]*=k; }}); }};
  const prose=s=>String(s).replace(/\\$\\s?(\\d+(?:\\.\\d+)?)\\s?M/g,(_,n)=>sym+(parseFloat(n)*k).toFixed(2)+'M');
  d.ledger.receipts=arr(d.ledger.receipts); d.ledger.payments=arr(d.ledger.payments);
  d.ledger.net=arr(d.ledger.net); d.ledger.balance=arr(d.ledger.balance);
  d.weekly.net=arr(d.weekly.net);
  d.forecast.forEach(w=>{{ mul(w,['collections','other_receipts','payroll','suppliers','tax',
                                 'debt','other_payments','net','closing']); }});
  d.ar_aging.forEach(b=>mul(b,['amount'])); d.ar_aging_total*=k;
  d.ar_by_contract.forEach(c=>mul(c,['amount']));
  mul(d.kpi,['balance','burn_day','net_wtd','collections_mtd','ar_total','floor','trough']);
  d.notes.forEach(n=>n.text=prose(n.text));
  return d;
}}

(function buildCurrency(){{
  const box=$('#cur-toggle');
  box.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{{
    CUR=b.dataset.cur;
    D=convertFX(D0,FX[CUR].rate,FX[CUR].sym);
    box.querySelectorAll('button').forEach(x=>x.setAttribute('aria-pressed',x.dataset.cur===CUR));
    render();
  }}));
}})();

/* ═════════ render ═════════ */
function render(){{
  const K=D.kpi, L=D.ledger, W=D.weekly, F=D.forecast;
  $('#m-gen').textContent=M.generated; $('#f-gen').textContent=M.generated;
  $('#f-src').textContent=M.source;    $('#m-asat').textContent=M.asat;
  $('#m-cur').textContent=`${{FX[CUR].unit}} millions`;
  $('#cur-rate').textContent = CUR==='USD' ? '' : `converted at 1 USD = ${{FX.SAR.rate}} SAR`;
  $('#m-cover').textContent=`Ledger from ${{M.ledger_from}} · ${{L.labels.length}} business days`;
  $('#es-scope').textContent=`as at ${{M.asat}}`;
  $('#p-when').textContent=M.asat;

  const runwayTxt = K.runway_months==null ? '—' : K.runway_months.toFixed(1);
  const burnPos = K.burn_day>=0;
  $('#tiles').innerHTML =
    tile('Cash balance', fM(K.balance,3), `as at ${{M.asat}}`,
         chip(K.net_wtd>=0?'good':'bad', `${{K.net_wtd>=0?'▲':'▼'}} ${{fMs(K.net_wtd,3)}} this week`), K.net_wtd<0) +
    tile('Runway', `${{runwayTxt}}<small>mo</small>`, `at ${{fM(Math.abs(K.burn_day),3)}} a day`,
         chip(K.runway_months!=null&&K.runway_months<12?'bad':'good',
              K.runway_months==null?'● cash is building':`${{K.runway_months<12?'▼ under':'▲ over'}} 12 months`),
         K.runway_months!=null&&K.runway_months<12) +
    tile('Daily burn', fMs(K.burn_day,3), `20-day average`,
         chip(burnPos?'good':'bad', burnPos?'▲ net inflow':'▼ net outflow'), !burnPos) +
    tile('Collections', fM(K.collections_mtd,3), `month to date`,
         chip('flat', `${{L.labels.length}} days in ledger`)) +
    tile('AR outstanding', fM(K.ar_total,3), `${{fP(K.ar_over60_share,0)}} over 60 days`,
         chip(K.ar_over60_share>0.25?'bad':'flat',
              K.ar_over60_share>0.25?'▼ concentration in 60+':'● within tolerance'),
         K.ar_over60_share>0.25) +
    tile('Forecast low', fM(K.trough,3), K.trough_week||'—',
         chip(K.trough<=K.floor?'bad':'good',
              K.trough<=K.floor?'▼ breaches floor':'▲ clears floor'), K.trough<=K.floor);

  /* 1 · daily movement */
  $('#led-scope').textContent=`${{UNIT()}} · ${{L.labels.length}} business days`;
  $('#bal-scope').textContent=`${{UNIT()}} · closing · floor ${{fM(K.floor,2)}}`;
  lineChart('#ch-bal', L.sparse, [{{name:'Closing balance', values:L.balance, color:'--s1', area:true}}],
            {{yFmt:v=>SY()+f2(v)+'M', minZero:false, frame:'full', floor:K.floor}});
  $('#rp-scope').textContent=`${{UNIT()}} · last ${{W.recent.length}} weeks`;
  columnChart('#ch-rp', W.recent, [
    {{name:'Receipts', values:W.receipts, color:'--s1'}},
    {{name:'Payments', values:W.payments, color:'--s3'}}], {{fmt:v=>fM(v,2)}});
  $('#nw-scope').textContent=`${{UNIT()}} · weekly net`;
  columnChart('#ch-net', W.recent, [{{name:'Net movement', values:W.net,
              color:'--s2'}}], {{fmt:v=>fM(v,2)}});

  /* 2 · forecast */
  $('#fc-scope').textContent=`${{F.length}} weeks from ${{M.forecast_from}}`;
  $('#fcb-scope').textContent=`${{UNIT()}} · projected closing`;
  lineChart('#ch-fc', F.map(w=>w.week), [{{name:'Projected balance',
            values:F.map(w=>w.closing), color:'--s1', area:true}}],
            {{yFmt:v=>SY()+f2(v)+'M', minZero:false, frame:'full', floor:K.floor}});
  $('#fc-note').textContent = K.trough<=K.floor
    ? `The projection dips to ${{fM(K.trough,3)}} in ${{K.trough_week}}, below the ${{fM(K.floor,2)}} operating floor. Collections timing is the main lever.`
    : `The projection holds above the ${{fM(K.floor,2)}} operating floor throughout, with its low point of ${{fM(K.trough,3)}} in ${{K.trough_week}}.`;

  simpleTable('#t-fc', ['Week','Collections','Other','Payroll','Suppliers','Tax','Debt','Other pmts','Net','Closing'],
    F.map(w=>({{cls: w.closing===K.trough?'trough':'',
      cells:[esc(w.week), f2(w.collections), f2(w.other_receipts), f2(w.payroll),
             f2(w.suppliers), f2(w.tax), f2(w.debt), f2(w.other_payments),
             f2s(w.net), f2(w.closing)],
      cellCls:['','','','','','','','', w.net<0?'neg':'pos','']}})));

  /* 3 · receivables */
  $('#ar-scope').textContent=`as at ${{M.asat}} · total ${{fM(D.ar_aging_total,2)}}`;
  $('#ag-scope').textContent=`total ${{fM(D.ar_aging_total,2)}}`;
  hbarChart('#ch-ag', D.ar_aging.map(b=>({{label:b.bucket.replace(' — ',' · '),
            amount:b.amount, share:b.share}})),
            {{frame:'aging', padL:200, colors:['--s1','--s2','--s3'], total:D.ar_aging_total}});
  const ctTot=sum(D.ar_by_contract.map(c=>c.amount));
  $('#ct-scope').textContent=`total ${{fM(ctTot,2)}}`;
  hbarChart('#ch-ct', D.ar_by_contract.map(c=>({{label:c.contract, amount:c.amount,
            share:c.amount/ctTot}})), {{frame:'aging', padL:200, total:ctTot}});

  /* 4 · notes */
  $('#updates').innerHTML=D.notes.map(n=>
    `<div class="upd"><div class="eyebrow">${{esc(n.topic)}}</div><p>${{esc(n.text)}}</p></div>`).join('');
}}
render();

(function(){{ const links=[...document.querySelectorAll('nav.rail a')],
    secs=[...document.querySelectorAll('section.sec')];
  const io=new IntersectionObserver(es=>es.forEach(e=>{{ if(e.isIntersecting)
    links.forEach(l=>l.classList.toggle('on', l.getAttribute('href')==='#'+e.target.id)); }}),
    {{rootMargin:'-20% 0px -70% 0px'}});
  secs.forEach(s=>io.observe(s)); }})();
</script>
"""

Path(a.out).write_text(PAGE, encoding="utf-8")
print(f"✓ Created {a.out} (style + chart library lifted from {a.src})")
print(f"  Now run:  python etl_cash.py --file noon_cash_daily.xlsx")
