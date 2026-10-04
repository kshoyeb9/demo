#!/usr/bin/env python3
"""Add a fiscal-year switch, and the FY2025/26 year, to the published P&L page.

The live dashboard carries one year as `const DATA = {...}` and derives a
handful of module-level constants from it. This rewrites the page so it holds
both years and can rebind those constants when the year changes, leaving every
chart, table and the period selector exactly as they were.

It also makes the comparison series self-describing. FY2026/27 is held against
budget; FY2025/26 is a closed year with no budget, so it is held against
FY2024/25 actual, and the page now takes that label from the data rather than
saying 'Budget' everywhere.

    python build_two_year_page.py <live.html> <fy2526_page.json> <out.html>
"""
import json, re, sys


def main(live, fy_json, out, aging_json=None):
    html = open(live, encoding="utf-8").read()
    fy2526 = json.load(open(fy_json))
    aging = json.load(open(aging_json)) if aging_json else None

    # ── 1 · two datasets in place of one ───────────────────────────────────
    m = re.search(r"const DATA = (\{.*?\});\n", html, re.DOTALL)
    if not m:
        sys.exit("could not find 'const DATA = {...};' in the live page.")
    live_data = json.loads(m.group(1))
    live_fy = live_data["meta"]["fiscal_year"]
    # The live year is held against budget; say so explicitly now that the
    # label is data rather than hard-coded.
    live_data["meta"].setdefault("comparator", "Budget")
    live_data["meta"].setdefault("closed", False)

    # Ageing and counterparty detail are reported per month, so they are keyed
    # by month rather than held as one list at a fixed date. The flat keys stay
    # empty: a month with no schedule should read as not supplied rather than
    # inherit another month's balances.
    if aging:
        live_data["aging_by_month"] = aging["by_month"]
        for k in ("ar_aging", "ar_by_contract", "ap_aging", "ap_by_vendor"):
            live_data[k] = []
        live_data["ar_aging_total"] = live_data["ap_aging_total"] = None

    datasets = {fy2526["meta"]["fiscal_year"]: fy2526, live_fy: live_data}
    order = sorted(datasets)                       # oldest year first
    block = (
        "const DATASETS = " + json.dumps(datasets, ensure_ascii=False, separators=(",", ":")) + ";\n"
        "const FY_ORDER = " + json.dumps(order) + ";\n"
        "/* The most recent year is what the page opens on. */\n"
        "let FY = " + json.dumps(order[-1]) + ";\n"
        "let DATA = DATASETS[FY];\n"
    )
    html = html[:m.start()] + block + html[m.end():]

    # ── 2 · rebindable year constants ──────────────────────────────────────
    old = ("const D0=DATA, M=D0.meta, MONTHS=M.months, LAST=M.period_index, FYS=M.fy_start_month||7;\n"
           "const LBL = M.month_labels || MONTHS.map(m=>`${m} ${M.year}`);\n")
    assert html.count(old) == 1, "year constants not found"
    new = ("""/* These follow the selected year, so they are rebound rather than fixed.
   CMP is what the comparison series holds: budget for a year that has one,
   prior-year actual for a closed year that does not. */
let D0, M, MONTHS, LAST, FYS, LBL, CMP, CMPSHORT;
function bindYear(){
  DATA = DATASETS[FY]; D0 = DATA; M = D0.meta;
  MONTHS = M.months; LAST = M.period_index; FYS = M.fy_start_month || 7;
  LBL = M.month_labels || MONTHS.map(m=>`${m} ${M.year}`);
  CMP = M.comparator || 'Budget';
  CMPSHORT = CMP === 'Budget' ? 'Bdgt' : CMP;
}
bindYear();
""")
    html = html.replace(old, new)

    # ── 3 · the period selector is rebuilt when the year changes ───────────
    old = "/* period selector */\n(function buildSelector(){"
    assert html.count(old) == 1
    html = html.replace(old, "/* period selector */\nfunction buildSelector(){")
    old = """  function sync(){ from.value=R.i0; to.value=R.i1; quick.querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed', +b.dataset.a===R.i0&&+b.dataset.b===R.i1)); }
  sync();
})();"""
    assert html.count(old) == 1
    html = html.replace(old, """  function sync(){ from.value=R.i0; to.value=R.i1; quick.querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed', +b.dataset.a===R.i0&&+b.dataset.b===R.i1)); }
  sync();
}
buildSelector();""")
    # the month options and quick-picks are per-year, so clear them first
    old = "  const from=$('#p-from'), to=$('#p-to'), quick=$('#quick');\n  MONTHS.forEach("
    assert html.count(old) == 1
    html = html.replace(old, "  const from=$('#p-from'), to=$('#p-to'), quick=$('#quick');\n"
                             "  from.replaceChildren(); to.replaceChildren(); quick.replaceChildren();\n"
                             "  MONTHS.forEach(")

    # ── 4 · the year toggle itself ─────────────────────────────────────────
    old = "let R={i0:LAST, i1:LAST};"
    assert html.count(old) == 1
    html = html.replace(old, """let R={i0:LAST, i1:LAST};

/* Fiscal-year switch. Changing the year rebinds everything derived from it,
   reselects the full year, and rebuilds the period selector, because the month
   labels and the closed-month count both change with it. */
(function buildYearToggle(){
  const box=$('#fy-toggle');
  FY_ORDER.forEach(y=>{
    const b=document.createElement('button'); b.type='button'; b.dataset.fy=y;
    b.textContent=y.replace('FY','FY ');
    const d=DATASETS[y];
    b.title = d.meta.closed ? y+' — closed year, all twelve months actual'
                            : y+' — '+d.meta.ytd_label+' actual to date';
    b.setAttribute('aria-pressed', String(y===FY));
    b.addEventListener('click',()=>{
      if(y===FY) return;
      FY=y; bindYear();
      D = CUR==='USD' ? D0 : convertFX(D0, FX[CUR].rate, FX[CUR].sym);
      R={i0:0, i1:LAST};               /* open a switched-to year on its full span */
      box.querySelectorAll('button').forEach(x=>x.setAttribute('aria-pressed', String(x.dataset.fy===FY)));
      buildSelector(); render();
    });
    box.appendChild(b);
  });
})();""")

    # the toggle's markup, beside the currency control
    old = """        <span class="cur-rate" id="cur-rate"></span>
      </div>"""
    assert html.count(old) == 1
    html = html.replace(old, """        <span class="cur-rate" id="cur-rate"></span>
        <label for="fy-toggle" style="margin-left:14px">Fiscal year</label>
        <span class="quick" id="fy-toggle" role="group" aria-label="Fiscal year"></span>
      </div>""")

    # ── 5 · the comparison series names itself ─────────────────────────────
    subs = [
        # masthead and card headings
        ('· actual vs budget · refreshed',
         '· actual vs <span class="cmp-label">budget</span> · refreshed'),
        ('<h3>P&amp;L summary — actual vs budget</h3>',
         '<h3>P&amp;L summary — actual vs <span class="cmp-label">budget</span></h3>'),
        ('variance = actual − budget · margins in % of revenue',
         'variance = actual − <span class="cmp-label">budget</span> · margins in % of revenue'),
        ('<i class="mark" style="background:var(--s2)"></i>Budget</span></div>',
         '<i class="mark" style="background:var(--s2)"></i><span class="cmp-label">Budget</span></span></div>'),
        # charts
        ("""svg.appendChild(txt(VB,padT-9,'BUDGET',{'text-anchor':'end',class:'axis',style:'letter-spacing:.08em'}));""",
         """svg.appendChild(txt(VB,padT-9,CMP.toUpperCase(),{'text-anchor':'end',class:'axis',style:'letter-spacing:.08em'}));"""),
        ("""{label:'Budget',value:fM(b),color:markCol},""",
         """{label:CMP,value:fM(b),color:markCol},"""),
        ("""fP(v.pct)+' of budget')}]);""", """fP(v.pct)+' of '+CMP)}]);"""),
        # tables
        ("""<tr><th>Actual</th><th>Budget</th><th>Var</th><th>% Bdgt</th><th>Actual</th><th>Budget</th><th>Var</th><th>% Bdgt</th></tr></thead>`;""",
         """<tr><th>Actual</th><th>${esc(CMP)}</th><th>Var</th><th>% ${esc(CMPSHORT)}</th>"""
         """<th>Actual</th><th>${esc(CMP)}</th><th>Var</th><th>% ${esc(CMPSHORT)}</th></tr></thead>`;"""),
        # chips and tiles
        ("""(pct==null?`${fMs(v)} vs budget`:`${fP(pct)} of budget`)""",
         """(pct==null?`${fMs(v)} vs ${CMP}`:`${fP(pct)} of ${CMP}`)"""),
        ("""`${d>0?'▲':d<0?'▼':'●'} ${fPp(d)} vs budget`""",
         """`${d>0?'▲':d<0?'▼':'●'} ${fPp(d)} vs ${CMP}`"""),
        ("""`Budget ${fM(rv.budget)}`""", """`${CMP} ${fM(rv.budget)}`"""),
        ("""`Budget ${fM(ct.budget)}`""", """`${CMP} ${fM(ct.budget)}`"""),
        ("""`${fP(gm.actual)} margin · budget ${fP(gm.budget)}`""",
         """`${fP(gm.actual)} margin · ${CMP} ${fP(gm.budget)}`"""),
        ("""`${fP(em.actual)} margin · budget ${fP(em.budget)}`""",
         """`${fP(em.actual)} margin · ${CMP} ${fP(em.budget)}`"""),
        # With no burn there is no runway to state, so the tile reports the cash
        # generated rather than a dash against a 'burn' that is really an inflow.
        ("""    tile('Runway', `${Rw.months==null?'—':Rw.months.toFixed(1)}<small>mo</small>`,
         `at ${fM(Math.abs(Rw.burn_3m))} monthly burn`,""",
         """    tile('Runway', Rw.months==null?'n/a':`${Rw.months.toFixed(1)}<small>mo</small>`,
         Rw.months==null
           ? `cash generated at ${fM(Math.abs(Rw.burn_3m))} a month`
           : `at ${fM(Math.abs(Rw.burn_3m))} monthly burn`,"""),
        ("""+ `<span><i class="mark" style="background:var(--s2)"></i>Budget</span>`;""",
         """+ `<span><i class="mark" style="background:var(--s2)"></i>${esc(CMP)}</span>`;"""),
    ] + [
        ("  d.ap_by_vendor.forEach(r=>mul(r,['amount']));\n",
         "  d.ap_by_vendor.forEach(r=>mul(r,['amount']));\n  Object.values(d.aging_by_month||{}).forEach(g=>{\n    ['ar_aging','ar_by_contract','ap_aging','ap_by_vendor'].forEach(key=>(g[key]||[]).forEach(r=>mul(r,['amount'])));\n    mul(g,['ar_aging_total','ar_by_contract_total','ap_aging_total','ap_unreconciled','ap_balance_sheet']);\n  });\n"),
        ('          <div class="legend"><span><i style="background:var(--s1)"></i>Current</span><span><i style="background:var(--s2)"></i>31–60 days</span><span><i style="background:var(--s3)"></i>60+ days</span></div>\n          <div class="chart" id="ch-ap-aging"></div>\n        </div>',
         '          <div class="legend" id="lg-ap-aging"></div>\n          <div class="chart" id="ch-ap-aging"></div>\n          <div class="note" id="ap-aging-note"></div>\n        </div>'),
        ('<h3>Payments — by vendor</h3>',
         '<h3>Accounts payable — by vendor</h3>'),
        ('<div class="legend"><span><i style="background:var(--s3)"></i>Payments made</span></div>\n          <div class="chart" id="ch-ap-vendor"></div>',
         '<div class="legend"><span><i style="background:var(--s3)"></i>Outstanding balance</span></div>\n          <div class="chart" id="ch-ap-vendor"></div>'),
    ] + [
        ("    svg.appendChild(txt(sx(b.amount)+8,y+bh/2+4,fM(b.amount)+(b.share!=null?'  ·  '+fP(b.share,0):''),{class:'lbl'}));",
         "    /* A negative balance (an unapplied credit) has no bar, and its value\n       would otherwise be drawn left of the axis on top of the category\n       label. Park those labels just inside the axis instead. */\n    svg.appendChild(txt(Math.max(sx(b.amount),padL)+8,y+bh/2+4,fM(b.amount)+(b.share!=null?'  ·  '+fP(b.share,0):''),{class:'lbl'}));"),
    ]
    for a, b in subs:
        if html.count(a) != 1:
            sys.exit(f"expected exactly one of:\n  {a[:90]}\nfound {html.count(a)}")
        html = html.replace(a, b)

    # fill the static labels, and say what the comparison is, on every render
    old = "function render(){\n"
    assert html.count(old) == 1
    html = html.replace(old, """function render(){
  document.querySelectorAll('.cmp-label').forEach(n=>{ n.textContent = CMP; });
""")

    # a line under the P&L table naming the comparison, since it is not always budget
    old = """  $('#pl-note').textContent = M.opex_split"""
    assert html.count(old) == 1
    html = html.replace(old, """  const cmpNote = M.comparator && M.comparator!=='Budget'
    ? `This year is closed, so there is no budget to hold it against: the comparison column is `
      + `${CMP} actual, month for month. `
    : '';
  $('#pl-note').textContent = cmpNote + (M.opex_split""")
    old = """    : 'Operating expenses are shown as a single line: the financial statements give one operating-expenses figure per business unit, with no split into BU salaries, Noon HQ and other. Supply that split on the Manual Inputs tab of the source workbook and the three lines appear here.';"""
    assert html.count(old) == 1
    html = html.replace(old, """    : 'Operating expenses are shown as a single line: the financial statements give one operating-expenses figure per business unit, with no split into BU salaries, Noon HQ and other. Supply that split on the Manual Inputs tab of the source workbook and the three lines appear here.');""")

    # ── the four ageing panels follow the selected month ──────────────────
    old_block_start = "  $('#ar-aging-scope').textContent="
    old_block_end = "else awaiting('#ch-ap-vendor','No vendor-level payment schedule was supplied for the month.');"
    i = html.index(old_block_start)
    j = html.index(old_block_end) + len(old_block_end)
    html = html[:i] + """  /* Ageing and counterparty detail are reported month by month, so they
     follow the end of the selected period rather than one fixed date. A month
     with no schedule says so instead of showing a neighbouring month's. */
  const AGM = D.aging_by_month || {};
  const AG  = AGM[LBL[R.i1]] || null;
  const agAt = AG ? LBL[R.i1] : asAt;
  const arAg  = AG ? AG.ar_aging       : (D.ar_aging       || []);
  const arCon = AG ? AG.ar_by_contract : (D.ar_by_contract || []);
  const apAg  = AG ? AG.ap_aging       : (D.ap_aging       || []);
  const apVen = AG ? AG.ap_by_vendor   : (D.ap_by_vendor   || []);

  const arAgTot = sum(arAg.map(b=>b.amount));
  $('#ar-aging-scope').textContent = arAg.length ? `as at ${agAt} · total ${fM(arAgTot)}` : 'not yet supplied';
  if(arAg.length) hbarChart('#ch-ar-aging', arAg.map(b=>({label:b.bucket.replace(' — ',' · '),amount:b.amount,share:b.share})),
            {frame:'aging', padL:200, colors:['--s1','--s2','--s3'], total:arAgTot});
  else awaiting('#ch-ar-aging','No receivables aging was supplied for this month.');

  const arcTot = sum(arCon.map(c=>c.amount));
  $('#ar-contract-scope').textContent = arCon.length ? `as at ${agAt} · total ${fM(arcTot)}` : 'not yet supplied';
  if(arCon.length) hbarChart('#ch-ar-contract', arCon.map(c=>({label:c.contract,amount:c.amount,share:c.amount/arcTot})),
            {frame:'aging', padL:200, total:arcTot});
  else awaiting('#ch-ar-contract','No contract-level receivable schedule was supplied for this month.');

  /* Payables age in five buckets, not three, so the legend is built from the
     data rather than fixed in the markup. */
  const AP_COLS = ['--s1','--s4','--s2','--s5','--s3'];
  const apAgTot = sum(apAg.map(b=>b.amount));
  $('#lg-ap-aging').innerHTML = apAg.length
    ? apAg.map((b,i)=>`<span><i style="background:${css(AP_COLS[i%AP_COLS.length])}"></i>${esc(b.bucket)}</span>`).join('')
    : '';
  $('#ap-aging-scope').textContent = apAg.length ? `as at ${agAt} · subledger ${fM(apAgTot)}` : 'not yet supplied';
  if(apAg.length) hbarChart('#ch-ap-aging', apAg.map(b=>({label:b.bucket,amount:b.amount,share:b.share})),
            {frame:'aging', padL:200, colors:AP_COLS, total:apAgTot});
  else awaiting('#ch-ap-aging','No payables aging was supplied for this month.');
  $('#ap-aging-note').textContent = (AG && AG.ap_unreconciled)
    ? `The vendor subledger totals ${fM(apAgTot)} against ${fM(AG.ap_balance_sheet)} of payables on the `
      + `balance sheet. The ${fM(AG.ap_unreconciled)} difference is accruals and other payables that do `
      + `not pass through a vendor account, so the aging covers the subledger only.`
    : '';

  const venTot = sum(apVen.map(v=>v.amount));
  $('#ap-vendor-scope').textContent = apVen.length ? `as at ${agAt} · total ${fM(venTot)}` : 'not yet supplied';
  if(apVen.length) hbarChart('#ch-ap-vendor', apVen.map(v=>({label:v.vendor,amount:v.amount,share:v.amount/venTot})),
            {frame:'aging', padL:200, color:'--s3', total:venTot});
  else awaiting('#ch-ap-vendor','No vendor-level payable schedule was supplied for this month.');""" + html[j:]

    open(out, "w", encoding="utf-8").write(html)
    print(f"  Years   {', '.join(order)}  (opens on {order[-1]})")
    print(f"  Wrote   {out}  ({len(html):,} bytes)")


if __name__ == "__main__":
    main(*sys.argv[1:])
