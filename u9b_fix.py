import win32com.client as w, shutil
OUT = r"D:\倍优跟单进度表（模板）_v10.xlsx"
GOOD = "'00使用说明'!$C$2"
app=w.Dispatch("Ket.Application"); app.Visible=False; app.DisplayAlerts=False
wb=app.Workbooks.Open(OUT)
log=[]
for i in range(1, wb.Worksheets.Count+1):
    ws=wb.Worksheets(i); name=ws.Name
    # bound explicit full-column 1048576 refs
    try:
        ws.UsedRange.Replace("1048576", "100000", 2, 1, False)
    except Exception as e:
        log.append("%s 1048576 err %s"%(name,e))
    # centralize TODAY (skip the anchor sheet itself to avoid self-ref)
    if name=="00使用说明":
        continue
    try:
        ws.UsedRange.Replace("TODAY()", GOOD, 2, 1, False)
        log.append("%s: TODAY()->anchor ok"%name)
    except Exception as e:
        log.append("%s TODAY err %s"%(name,e))
# ensure anchor cell holds =TODAY()
ws0=wb.Worksheets("00使用说明")
for col in ["C","Z","Y","X"]:
    if ws0.Range(col+"2").Formula.startswith("=TODAY"):
        anchor_cell=ws0.Range(col+"2"); break
else:
    anchor_cell=ws0.Range("C2"); anchor_cell.Formula="=TODAY()"
# note next to anchor
note=ws0.Range("D2")
if note.Value is None:
    note.Value="（本行C2=今日日期，供全表TODAY()引用，请勿删除）"
wb.Save()
wb.Close(False); app.Quit()
print("\n".join(log)); print("DONE")
