import json, subprocess, sys, tempfile
from pathlib import Path
import pandas as pd

CS_NAMES={"cs","c/s","case","cases","case qty","case quantity"}

def run_marker(pdf: Path):
    outdir=Path(tempfile.mkdtemp())
    subprocess.run([
        sys.executable,
        "-m","marker.scripts.convert_single",
        str(pdf),
        "--output_format","json",
        "--output_dir",str(outdir)
    ],check=True)
    jf=next(outdir.glob("*.json"))
    return json.loads(jf.read_text(encoding="utf-8"))

def find_cs(headers):
    for i,h in enumerate(headers):
        n=h.lower().strip().replace(".","").replace("_"," ")
        if n in CS_NAMES:
            return i
    return None

def extract_tables(doc,pdf_name):
    rows_out=[]
    for page in doc.get("pages",[]):
        for block in page.get("blocks",[]):
            if block.get("type")!="table":
                continue
            rows=block.get("rows",[])
            if len(rows)<2:
                continue
            headers=[c.get("text","").strip() or f"Column_{i+1}" for i,c in enumerate(rows[0]["cells"])]
            cs_idx=find_cs(headers)
            if cs_idx is None:
                continue
            for row in rows[1:]:
                vals=[c.get("text","").strip() for c in row.get("cells",[])]
                vals.extend([""]*(len(headers)-len(vals)))
                rec=dict(zip(headers,vals))
                try:
                    cs=float(str(rec[headers[cs_idx]]).replace(",",""))
                except:
                    cs=0
                if cs>0:
                    rec["Source PDF"]=pdf_name
                    rows_out.append(rec)
    return rows_out

def process(pdfs):
    all_rows=[]
    for pdf in pdfs:
        print("Processing",pdf.name)
        all_rows.extend(extract_tables(run_marker(pdf),pdf.name))
    if not all_rows:
        print("No matching rows.")
        return
    df=pd.DataFrame(all_rows)
    cols=["Source PDF"]+[c for c in df.columns if c!="Source PDF"]
    df=df[cols]
    out=Path("output")
    out.mkdir(exist_ok=True)
    outfile=out/"Extracted.xlsx"
    df.to_excel(outfile,index=False)
    print("Saved:",outfile.resolve())

if __name__=="__main__":
    if len(sys.argv)<2:
        print("Drag one or more PDF files onto this script/executable.")
        input("Press Enter to exit...")
        sys.exit()
    process([Path(p) for p in sys.argv[1:]])
