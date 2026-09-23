#!/usr/bin/env python3
"""
Re-applies this build's changes to the Project Dashboard's Monthly Projections tab:

  1. Budget Snapshot Analysis - compare any two snapshots budget code by budget code.
  2. Coverage honesty - the total row says how many jobs it actually covers, and the
     jobs with no snapshot for the month are named underneath instead of vanishing.

Why this exists as a script rather than a one-off edit: index.html is edited in
several places (here, the repo, other sessions), so a hand-merged copy gets
overwritten the next time a newer index.html comes down. Re-run this against the
newest index.html instead and the feature goes back on in one step.

    python add_snapshot_analysis.py index.html

Idempotent, and the two groups apply independently: whichever is already present is
left alone. Every edit is anchored to text that must exist exactly once; if an anchor
has moved, that group is skipped with a message rather than written half-applied.
"""
import io, sys, re

FUNCS = r"""/* ===================== Budget Snapshot Analysis ==========================
   One job, two snapshots, compared budget code by budget code - the same view
   Spark's Budget Projections screen gives. The rows are written per project by
   import_snapshot_lines.js, and only the selected job's file is fetched, so the
   page cost stays flat as jobs are added.
   A stored row is [revised budget, job-to-date cost, committed, EAC]. The
   comparison is on EAC: that is the number the month-end review turns on, and
   the one Spark puts side by side. */
const EAC_IDX=3;
const money2=n=>(n||n===0)?'$ '+Number(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}):'&mdash;';
/* Spark's sign convention, kept deliberately: Variance = Snapshot 1 - Snapshot 2.
   Positive means the estimate came DOWN between the two (money back); negative
   means it grew, and shows red in parentheses exactly as the source screen does. */
function vcell(n){ n=num(n); if(!Math.round(n*100)) return '<span class="muted">&mdash;</span>';
  return n>0 ? '<span class="funder">'+money2(n)+'</span>'
             : '<span class="fover">( '+money2(Math.abs(n))+' )</span>'; }
function snapLabel(s){
  if(!s) return '';
  if(String(s.id)==='current') return 'Current (live)';
  const base=s.label||monthName(s.month||'')||('Snapshot '+s.id);
  return base+(/approv/i.test(s.approval||'')?'':' — '+(s.approval||'unapproved'));
}
function renderCompare(){
  const jobSel=document.getElementById('cmpJob'); if(!jobSel) return;
  const meta=document.getElementById('cmpMeta'), out=document.getElementById('cmpList');
  const list=(SNAPIDX&&SNAPIDX.projects)||[];
  if(!list.length){
    jobSel.innerHTML=''; ['cmpA','cmpB'].forEach(id=>{const e=document.getElementById(id); if(e) e.innerHTML='';});
    if(meta) meta.textContent='';
    out.innerHTML='<div class="empty">No budget-code snapshot data yet. Run the <b>Budget snapshot LINES import</b> workflow &mdash; it writes <b>snapshot_lines/</b> into the same SharePoint folder &mdash; then reload this page.</div>';
    return; }
  if(!CMPJOB||!list.some(p=>p.job_no===CMPJOB)) CMPJOB=list[list.length-1].job_no;
  jobSel.innerHTML=list.map(p=>'<option value="'+esc(p.job_no)+'">'+esc(p.job_no)+' — '+esc(p.job_name||'')+'</option>').join('');
  jobSel.value=CMPJOB;
  cmpLoadJob();
}
async function cmpLoadJob(){
  const out=document.getElementById('cmpList'); if(!out) return;
  if(CMPFILE&&CMPFILE.job_no===CMPJOB){ cmpFill(); return; }
  if(CMPBUSY) return; CMPBUSY=true;
  out.innerHTML='<div class="loading">Loading budget codes for '+esc(CMPJOB)+'&hellip;</div>';
  try{
    const token=await getToken(); if(!token){ CMPBUSY=false; return; }
    const r=await fetchFile(token, SNAPLINES_DIR+'/'+CMPJOB+'.json');
    if(!r.ok) throw new Error('Graph HTTP '+r.status);
    CMPFILE=await r.json();
  }catch(e){ CMPFILE=null; CMPBUSY=false;
    out.innerHTML='<div class="empty">Could not load budget codes for '+esc(CMPJOB)+': '+esc((e&&e.message)||e)+'</div>'; return; }
  CMPBUSY=false; CMPA=''; CMPB=''; cmpFill();
}
/* Newest first, with the live budget at the top when the import captured one. */
function cmpSnaps(){ if(!CMPFILE) return [];
  const arr=(CMPFILE.snapshots||[]).slice().reverse();
  return CMPFILE.current ? [CMPFILE.current].concat(arr) : arr; }
function cmpFill(){
  const a=document.getElementById('cmpA'), b=document.getElementById('cmpB'); if(!a||!b) return;
  const snaps=cmpSnaps();
  if(!snaps.length){ a.innerHTML=''; b.innerHTML='';
    document.getElementById('cmpList').innerHTML='<div class="empty">No snapshots stored for this job yet.</div>'; return; }
  const opts=snaps.map(s=>'<option value="'+esc(String(s.id))+'">'+esc(snapLabel(s))+'</option>').join('');
  a.innerHTML=opts; b.innerHTML=opts;
  // Default to the most recent pair: Snapshot 2 = newest (Current when present), Snapshot 1 = the one before it.
  if(!CMPA||!snaps.some(s=>String(s.id)===CMPA)) CMPA=String((snaps[1]||snaps[0]).id);
  if(!CMPB||!snaps.some(s=>String(s.id)===CMPB)) CMPB=String(snaps[0].id);
  a.value=CMPA; b.value=CMPB;
  const cv=document.getElementById('cmpVar'); if(cv) cv.checked=CMPVARONLY;
  cmpTable();
}
function cmpRows(){
  if(!CMPFILE) return null;
  const snaps=cmpSnaps();
  const A=snaps.find(s=>String(s.id)===CMPA), B=snaps.find(s=>String(s.id)===CMPB);
  if(!A||!B) return null;
  const codes=CMPFILE.codes||{};
  const keys=[...new Set(Object.keys(A.rows||{}).concat(Object.keys(B.rows||{})))].sort();
  const rows=keys.map(k=>{
    const ra=(A.rows||{})[k], rb=(B.rows||{})[k];
    const va=ra?num(ra[EAC_IDX]):0, vb=rb?num(rb[EAC_IDX]):0;
    return { code:k, desc:(codes[k]&&codes[k].d)||'', type:(codes[k]&&codes[k].t)||'',
             a:va, b:vb, v:va-vb, onlyA:!rb, onlyB:!ra };
  });
  return {A:A,B:B,rows:rows};
}
function cmpTable(){
  const out=document.getElementById('cmpList'), meta=document.getElementById('cmpMeta'); if(!out) return;
  const d=cmpRows();
  if(!d){ out.innerHTML='<div class="empty">Pick two snapshots to compare.</div>'; if(meta) meta.textContent=''; return; }
  /* Totals are over EVERY code, not the filtered view - "show only variances" hides
     rows that did not move, it does not change what the job is projected to cost. */
  const totA=d.rows.reduce((t,r)=>t+num(r.a),0), totB=d.rows.reduce((t,r)=>t+num(r.b),0);
  const moved=d.rows.filter(r=>Math.round(r.v*100)!==0);
  const shown=CMPVARONLY?moved:d.rows;
  if(meta) meta.innerHTML=esc(CMPFILE.job_no||'')+' '+esc(CMPFILE.job_name||'')
    +'  ·  '+d.rows.length+' budget codes, '+moved.length+' moved'
    +(String(d.B.id)==='current'?'  ·  <span class="chip" style="background:#FFE9C7;color:#8a5a00;border:1px solid #f0cf8e" title="Snapshot 2 is the live budget view, not an approved month-end record.">Current is live, not reviewed</span>':'')
    +(CMPFILE.generated_at?('  ·  captured '+new Date(CMPFILE.generated_at).toLocaleString()):'');
  const th=(s,tot)=>'<th style="text-align:right">'+esc(snapLabel(s))+'<div style="font-weight:400;font-size:11px;color:#7a828f">Estimate at Completion</div><div style="color:#1F3864">'+money2(tot)+'</div></th>';
  const head='<tr><th class="l">Budget Code</th><th class="l">Type</th><th class="l">Budget Code Description</th>'
    +th(d.A,totA)+th(d.B,totB)
    +'<th style="text-align:right">Variance<div style="font-weight:400;font-size:11px;color:#7a828f">Snap 1 &minus; Snap 2</div><div>'+vcell(totA-totB)+'</div></th></tr>';
  const chip=t=>' <span class="chip" style="background:#eef3fb;color:#1F3864;border:1px solid #c3d4ee">'+t+'</span>';
  const body=shown.map(r=>
    '<tr><td class="l"><b>'+esc(r.code)+'</b>'+(r.onlyB?chip('new'):'')+(r.onlyA?chip('dropped'):'')+'</td>'
    +'<td class="l">'+esc(r.type)+'</td><td class="l">'+esc(r.desc)+'</td>'
    +'<td>'+(r.onlyB?'<span class="muted">0</span>':money2(r.a))+'</td>'
    +'<td>'+(r.onlyA?'<span class="muted">0</span>':money2(r.b))+'</td>'
    +'<td>'+vcell(r.v)+'</td></tr>').join('');
  const empty='<tr><td class="l" colspan="6"><span class="muted">Nothing moved between these two snapshots.</span></td></tr>';
  const tot='<tr style="font-weight:bold;background:#f7f9fc"><td class="l">Total</td><td></td><td></td><td>'+money2(totA)+'</td><td>'+money2(totB)+'</td><td>'+vcell(totA-totB)+'</td></tr>';
  out.innerHTML='<div class="rwrap"><table class="rtable"><thead>'+head+'</thead><tbody>'+(body||empty)+tot+'</tbody></table></div>';
}
function cmpCsv(){
  const d=cmpRows(); if(!d) return;
  const rows=CMPVARONLY?d.rows.filter(r=>Math.round(r.v*100)!==0):d.rows;
  const totA=d.rows.reduce((t,r)=>t+num(r.a),0), totB=d.rows.reduce((t,r)=>t+num(r.b),0);
  const q=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
  const lines=[['Budget Code','Type','Budget Code Description',
                'Snapshot 1 ('+snapLabel(d.A)+') EAC','Snapshot 2 ('+snapLabel(d.B)+') EAC','Variance'].map(q).join(',')];
  for(const r of rows) lines.push([r.code,r.type,r.desc,num(r.a).toFixed(2),num(r.b).toFixed(2),r.v.toFixed(2)].map(q).join(','));
  lines.push(['Total','','',totA.toFixed(2),totB.toFixed(2),(totA-totB).toFixed(2)].map(q).join(','));
  const blob=new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=(CMPFILE.job_no||'job').replace(/[^\w.-]+/g,'_')+'_snapshot_comparison.csv';
  document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },0);
}
"""

MARKUP = (
 "     +'<div style=\"border-top:2px solid #e3e8f0;margin:28px 0 0\"></div>'\n"
 "     +'<div class=\"fnote\" style=\"background:#fff7ef;border-color:#f0d9bf;margin-top:18px\">&#128200; <b>Budget Snapshot Analysis</b> &mdash; pick a job and any two snapshots to compare it budget code by budget code. <b>Variance = Snapshot 1 &minus; Snapshot 2</b>, so a positive number means the estimate came <i>down</i>; an increase shows red in parentheses. <b>Current</b> is the live budget view, not a reviewed month &mdash; it moves between refreshes.</div>'\n"
 "     +'<div class=\"controls\">'\n"
 "       +'<label>Job:</label> <select id=\"cmpJob\" style=\"padding:8px 12px;border:1px solid #c7cfdb;border-radius:8px;font-size:14px;min-width:260px\"></select>'\n"
 "       +'<label style=\"margin-left:10px\">Snapshot 1 (previous):</label> <select id=\"cmpA\" style=\"padding:8px 12px;border:1px solid #c7cfdb;border-radius:8px;font-size:14px\"></select>'\n"
 "       +'<label style=\"margin-left:10px\">Snapshot 2 (latest):</label> <select id=\"cmpB\" style=\"padding:8px 12px;border:1px solid #c7cfdb;border-radius:8px;font-size:14px\"></select>'\n"
 "       +'<label style=\"margin-left:14px;cursor:pointer\"><input type=\"checkbox\" id=\"cmpVar\" checked style=\"vertical-align:-1px;margin-right:5px\">Show only variances</label>'\n"
 "       +'<button type=\"button\" id=\"cmpCsv\" style=\"margin-left:14px;padding:8px 14px;border:1px solid #c7cfdb;border-radius:8px;background:#fff;font-size:13px;font-weight:600;cursor:pointer\">&#11123; Export CSV</button>'\n"
 "     +'</div>'\n"
 "     +'<div id=\"cmpMeta\" style=\"color:#7a828f;font-size:12px;margin:-4px 0 10px\"></div>'\n"
 "     +'<div id=\"cmpList\"></div>'\n")

WIRING = (
 "  { const cj=document.getElementById('cmpJob'); if(cj) cj.addEventListener('change',function(){ CMPJOB=this.value; CMPA=''; CMPB=''; cmpLoadJob(); }); }\n"
 "  { const ca=document.getElementById('cmpA'); if(ca) ca.addEventListener('change',function(){ CMPA=this.value; cmpTable(); }); }\n"
 "  { const cb=document.getElementById('cmpB'); if(cb) cb.addEventListener('change',function(){ CMPB=this.value; cmpTable(); }); }\n"
 "  { const cv=document.getElementById('cmpVar'); if(cv) cv.addEventListener('change',function(){ CMPVARONLY=this.checked; cmpTable(); }); }\n"
 "  { const cc=document.getElementById('cmpCsv'); if(cc) cc.addEventListener('click',cmpCsv); }")

EDITS = [
 ("path constant",
  "const PROJSNAP_PATH='Compliance Dashboard/proj_snapshots.json';",
  "const PROJSNAP_PATH='Compliance Dashboard/proj_snapshots.json';\nconst SNAPLINES_DIR='Compliance Dashboard/snapshot_lines';          // per-project budget-code rows behind the snapshot comparison",
  True),
 ("state",
  "let MONTHLY=null, MPMONTH='';",
  "let MONTHLY=null, MPMONTH='';\nlet SNAPIDX=null, CMPFILE=null, CMPJOB='', CMPA='', CMPB='', CMPVARONLY=true, CMPBUSY=false;   // Budget Snapshot Analysis",
  True),
 ("index load",
  "try{ const mr=await fetchFile(token, PROJSNAP_PATH); if(mr.ok){ MONTHLY=await mr.json(); } }catch(e){ MONTHLY=null; }",
  "try{ const mr=await fetchFile(token, PROJSNAP_PATH); if(mr.ok){ MONTHLY=await mr.json(); } }catch(e){ MONTHLY=null; }\n    // Budget-code snapshot index (best-effort; the comparison shows an empty state until the lines import runs)\n    try{ const sx=await fetchFile(token, SNAPLINES_DIR+'/_index.json'); if(sx.ok){ SNAPIDX=await sx.json(); } }catch(e){ SNAPIDX=null; }",
  True),
 ("Monthly panel markup",
  "     +'<div id=\"mpList\"></div>'\n",
  "     +'<div id=\"mpList\"></div>'\n" + MARKUP,
  True),
 ("event wiring",
  "  { const mm=document.getElementById('mpMonth'); if(mm) mm.addEventListener('change',function(){ MPMONTH=this.value; renderMonthly(); }); }",
  "  { const mm=document.getElementById('mpMonth'); if(mm) mm.addEventListener('change',function(){ MPMONTH=this.value; renderMonthly(); }); }\n" + WIRING,
  True),
 ("render call",
  "  renderMonthly();\n",
  "  renderMonthly();\n  renderCompare();\n",
  True),
 ("functions",
  "// Procore-style cost projection card",
  FUNCS + "// Procore-style cost projection card",
  True),
]


COVERAGE_EDITS = [
 ("coverage counts",
  """  const notAppr=j=>!!(j.approval_status&&!/approv/i.test(j.approval_status));   // explicitly not-yet-approved
  const unappr=jobs.filter(notAppr);""",
  """  const notAppr=j=>!!(j.approval_status&&!/approv/i.test(j.approval_status));   // explicitly not-yet-approved
  const unappr=jobs.filter(notAppr);
  /* COVERAGE. A month only holds the jobs that had a Procore budget-view snapshot, so a
     job with no snapshot is absent from snap.jobs entirely rather than showing as zero.
     Summing what is present and calling it "Company total" reads as the whole company
     when it can be a small slice of it - and that is exactly the number someone reaches
     for when a projection is being disputed. So: count the jobs the dashboard knows
     about, under the same filters, and name what is missing instead of dropping it. */
  const captured=new Set(jobs.map(j=>String(j.job_no)));
  let known=ROWS.slice();
  if(SELECTED) known=known.filter(r=>SELECTED.has(r.procore_id));
  if(q) known=known.filter(r=>((r.project_number||'')+' '+(r.project||'')).toLowerCase().includes(q));
  const missing=known.filter(r=>!captured.has(String(r.project_number)))
                     .sort((a,b)=>num(b.revised_contract||b.contract_amount)-num(a.revised_contract||a.contract_amount));
  const totalKnown=known.length||jobs.length;"""),
 ("coverage chip",
  """'  \u00b7  '+jobs.length+' jobs'
    +(unappr.length?""",
  """'  \u00b7  '+jobs.length+' of '+totalKnown+' jobs captured'
    +(missing.length?' <span class="chip" style="background:#fde8e8;color:#8a1f1f;border:1px solid #f0bcbc" title="'+missing.length+' job(s) on the dashboard have no budget snapshot for this month, so they are in none of the totals below.">&#9888; '+missing.length+' not captured</span>':'')
    +(unappr.length?"""),
 ("total row label",
  """<td class="l">Company total</td><td>'+money(totC)+'</td>""",
  """<td class="l">Total &mdash; '+jobs.length+' of '+totalKnown+' jobs</td><td>'+money(totC)+'</td>"""),
 ("uncaptured list",
  """  listEl.innerHTML='<div class="rwrap"><table class="rtable"><thead>'+head+'</thead><tbody>'+body+tot+'</tbody></table></div>';
}
/* ===================== Budget Snapshot Analysis""",
  """  /* Name the jobs the month did not capture. Without this the tab under-reports in
     silence; with it, the gap reads as a to-do list of snapshots to take or approve. */
  let miss='';
  if(missing.length){
    const everSnap=new Set();
    Object.keys(snaps).forEach(m=>((snaps[m]&&snaps[m].jobs)||[]).forEach(j=>everSnap.add(String(j.job_no))));
    miss='<div class="fnote" style="background:#fdf2f2;border-color:#f0bcbc;margin-top:14px">'
      +'<b>&#9888; Not captured for '+esc(monthName(MPMONTH))+' &mdash; '+missing.length+' job'+(missing.length===1?'':'s')+', excluded from every total above.</b>'
      +'<div style="margin-top:8px;font-size:12px;line-height:1.7">'
      +missing.map(r=>'<div><b>'+esc(r.project_number||'')+'</b> '+esc(r.project||'')
          +' <span class="muted">&mdash; '+(everSnap.has(String(r.project_number))
              ?'has snapshots in other months, none for this one'
              :'no budget snapshot has ever been captured in Procore')+'</span>'
          +' <span class="muted">('+money(r.revised_contract||r.contract_amount)+' contract)</span></div>').join('')
      +'</div><div style="margin-top:8px;font-size:12px" class="muted">Take or approve a budget-view snapshot in Procore for these jobs to bring them into the month.</div></div>';
  }
  listEl.innerHTML='<div class="rwrap"><table class="rtable"><thead>'+head+'</thead><tbody>'+body+tot+'</tbody></table></div>'+miss;
}
/* ===================== Budget Snapshot Analysis"""),
]

GROUPS = [
 ("Budget Snapshot Analysis", "cmpJob", [(n,o,x) for n,o,x,_ in EDITS]),
 ("Monthly Projections coverage", "totalKnown", COVERAGE_EDITS),
]

def apply_group(s, name, sentinel, edits):
    """Returns (new_source, message). The group is all-or-nothing."""
    if sentinel in s:
        return s, '  %s: already present, unchanged.' % name
    work = s
    for anchor, old, new in edits:
        n = work.count(old)
        if n != 1:
            return s, ('  %s: SKIPPED - anchor "%s" found %d times, expected 1. '
                       'index.html has changed around it; nothing written for this group.'
                       % (name, anchor, n))
        work = work.replace(old, new)
    return work, '  %s: applied.' % name


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else 'index.html'
    s = io.open(path, encoding='utf-8').read()
    original = s
    msgs = []
    for name, sentinel, edits in GROUPS:
        s, m = apply_group(s, name, sentinel, edits)
        msgs.append(m)
    print(path)
    for m in msgs:
        print(m)
    if s != original:
        io.open(path, 'w', encoding='utf-8').write(s)
        print('Written.')
    else:
        print('No changes needed.')
    return 1 if any('SKIPPED' in m for m in msgs) else 0

sys.exit(main())
