import win32com.client as w, re, shutil, sys, io
SRC = r"D:\倍优跟单进度表（模板）_v9.xlsx"
OUT = r"D:\倍优跟单进度表（模板）_v10.xlsx"
CAP = 100000
LOG = open("u9_opt_log.txt", "w", encoding="utf-8")
def log(s):
    print(s); LOG.write(s+"\n")

re_whole = re.compile(r"(?<![A-Za-z0-9_$])(\$?[A-Z]{1,3}):(\$?[A-Z]{1,3})(?![0-9])")
re_today = re.compile(r"TODAY\s*\(")

app=w.Dispatch("Ket.Application"); app.Visible=False; app.DisplayAlerts=False
shutil.copyfile(SRC, OUT)
log("copied v9 -> v10")
wb=app.Workbooks.Open(OUT)
log("opened v10 read-only=%s" % wb.ReadOnly)

# choose TODAY anchor cell (must be empty); set LAST to avoid self-replace
ws0=wb.Worksheets("00使用说明")
anchor_addr=None
for col in ["C","Z","Y","X"]:
    if ws0.Range(col+"2").Value is None:
        anchor_addr="00使用说明!%s2"%col; anchor_cell=ws0.Range(col+"2"); break
if anchor_addr is None:
    anchor_addr="00使用说明!Z100"; anchor_cell=ws0.Range("Z100")
ANCHOR="'$%s'"%anchor_addr.replace("!","'!$")  # -> '$00使用说明'!$C2
log("TODAY anchor = %s" % ANCHOR)

tot_tok=0; tot_today=0
for i in range(1, wb.Worksheets.Count+1):
    ws=wb.Worksheets(i); name=ws.Name
    ur=ws.UsedRange
    try:
        fa=ur.Formula
    except Exception as e:
        log("%s: read err %s"%(name,e)); continue
    if not isinstance(fa,(list,tuple)): fa=[[fa]]
    elif not isinstance(fa[0],(list,tuple)): fa=[fa]
    toks=set(); tdays=set()
    nform=0
    for r in fa:
        for c in r:
            if isinstance(c,str) and c.startswith("="):
                nform+=1
                for m in re_whole.finditer(c): toks.add(m.group(0))
                for m in re_today.finditer(c): tdays.add(m.group(0))
    if not toks and not tdays:
        continue
    log("%-14s formulas=%d whole-tokens=%d today-variants=%d"%(name,nform,len(toks),len(tdays)))
    # replace whole-col tokens
    for t in sorted(toks):
        la=t.split(":")[0]; lb=t.split(":")[1]
        rep="%s2:%s%d"%(la,lb,CAP)
        try:
            ws.UsedRange.Replace(t, rep, 2, 1, False)  # xlPart, xlByRows, MatchCase=False
            tot_tok+=1
        except Exception as e:
            log("  !! replace FAIL %s -> %s : %s"%(t,rep,e))
    # replace TODAY variants
    for td in sorted(tdays):
        try:
            ws.UsedRange.Replace(td, ANCHOR, 2, 1, False)
            tot_today+=1
        except Exception as e:
            log("  !! TODAY replace FAIL %s : %s"%(td,e))
    # re-count remaining
    fa2=ur.Formula
    if not isinstance(fa2,(list,tuple)): fa2=[[fa2]]
    elif not isinstance(fa2[0],(list,tuple)): fa2=[fa2]
    rem=0
    for r in fa2:
        for c in r:
            if isinstance(c,str) and c.startswith("="): rem+=len(re_whole.findall(c))
    if rem: log("  !! %s still has %d whole-col refs"%(name,rem))

# set TODAY anchor LAST
anchor_cell.Formula="=TODAY()"
log("set anchor %s = TODAY()" % anchor_addr)
log("total tokens replaced=%d  total TODAY-variants replaced=%d"%(tot_tok,tot_today))
wb.Save()
log("saved v10")
wb.Close(False); app.Quit(); LOG.close(); print("DONE")
