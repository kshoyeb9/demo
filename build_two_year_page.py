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


def main(live, fy_json, out):
    html = open(live, encoding="utf-8").read()
    fy2526 = json.load(open(fy_json))

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

    open(out, "w", encoding="utf-8").write(html)
    print(f"  Years   {', '.join(order)}  (opens on {order[-1]})")
    print(f"  Wrote   {out}  ({len(html):,} bytes)")


if __name__ == "__main__":
    main(*sys.argv[1:])
