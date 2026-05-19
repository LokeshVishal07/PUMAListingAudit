"""
PUMA Marketplace Listing Audit
================================
Focused app: What needs to be listed?

RULES:
- ZeCom Tracker = YES
- Launch Date <= today + 30 days
- Inventory Stock > 0
- Special Request: ignored

OUTPUT: One Excel per region, 4 tabs per marketplace:
  {MP} - Already Listed      | listed EANs + stock + MP Product ID
  {MP} - No Action Needed    | articles fully listed
  {MP} - Add Variant         | missing EANs with stock > 0
  {MP} - Full New Listing    | eligible articles with zero listings
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="PUMA Listing Audit", layout="wide", page_icon="📋")

st.markdown("""
<style>
.header{background:linear-gradient(135deg,#1a1a2e,#0f3460);
  padding:1.5rem;border-radius:10px;text-align:center;color:white;margin-bottom:1.5rem;}
.header h1{font-size:1.8rem;font-weight:700;margin:0;}
.header p{color:#a0aec0;margin:.3rem 0 0;font-size:.85rem;}
.metric-box{background:white;border-radius:8px;padding:.85rem;border:1px solid #e2e8f0;
  text-align:center;margin-bottom:.5rem;}
.mv{font-size:1.5rem;font-weight:800;}.ml{font-size:.72rem;color:#718096;margin-top:.1rem;}
.cg{color:#276221;}.cb{color:#1a5276;}.co{color:#7b5800;}.cp{color:#4a235a;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class='header'>
  <h1>PUMA Marketplace Listing Audit</h1>
  <p>What needs to be listed? Full New Listings and Add Variants — by region and marketplace</p>
</div>""", unsafe_allow_html=True)

REGIONS = ["PH","MY","SG"]
REGION_MARKETPLACES = {
    "PH":["Lazada","Shopee","Zalora"],
    "MY":["Lazada","Shopee","Zalora","TikTok"],
    "SG":["Lazada","Shopee","Zalora"],
}
TODAY         = pd.Timestamp.today().normalize()
FUTURE_WINDOW = TODAY + pd.Timedelta(days=30)
STOCK_BUFFER  = {("PH","Lazada"):1}
INV_COL       = {"PH":"Avail_Qty","MY":"QtyAvailable","SG":"QTY"}
MP_CFG = {
    "Lazada":{"ean":"SellerSKU","status":"status","stock":"Quantity","id":"Product ID"},
    "Shopee":{"ean":"SKU","status":"Status","stock":"Stock","id":"Product ID"},
    "Zalora":{"ean":"SellerSku","status":"Status","stock":"Quantity","id":""},
    "TikTok":{"ean":"Seller sku","status":"Status","stock":"Quantity","id":"Product ID"},
}
FILE_TYPES=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"]
TAB_COLORS={
    "Already Listed"  :{"hdr":"#1e6f39","bg":"#c6efce","fc":"#276221"},
    "No Action Needed":{"hdr":"#1a5276","bg":"#d6eaf8","fc":"#1a5276"},
    "Add Variant"     :{"hdr":"#7b5800","bg":"#ffeb9c","fc":"#7b5800"},
    "Full New Listing":{"hdr":"#4a235a","bg":"#e8d5f5","fc":"#4a235a"},
}

def _ean(v):
    s=str(v).strip().split(".")[0].strip()
    return s if s not in("nan","None","NaT","") else ""
def _s(v):
    s=str(v).strip()
    return s if s not in("nan","None","NaT","") else ""
def _i(v):
    try:
        if v is None or(isinstance(v,float)and np.isnan(v)):return 0
        return max(0,int(float(str(v).strip().replace(",",""))))
    except:return 0

def _resolve(file,wanted,engines):
    for eng in engines:
        try:
            file.seek(0)
            sheets=pd.ExcelFile(file,engine=eng).sheet_names
            if wanted in sheets:return wanted
            m=next((s for s in sheets if s.strip().lower()==wanted.strip().lower()),None)
            return m if m else sheets[0]
        except:continue
    return wanted

def read_file(file,sheet_name=0,header=0):
    name=getattr(file,"name","")or""
    ext=name.rsplit(".",1)[-1].lower()if"."in name else""
    if ext in("csv","tsv"):
        sep="\t"if ext=="tsv"else","
        for enc in("utf-8","utf-8-sig","latin-1"):
            try:
                file.seek(0)
                df=pd.read_csv(file,sep=sep,header=header,dtype=str,encoding=enc,low_memory=False)
                df.columns=[str(c).strip()for c in df.columns]
                return df.dropna(how="all").reset_index(drop=True)
            except UnicodeDecodeError:continue
        return pd.DataFrame()
    engs={"xlsx":["openpyxl","xlrd"],"xls":["xlrd","openpyxl"],
          "xlsb":["pyxlsb","openpyxl"],"ods":["odf","openpyxl"]}.get(ext,["openpyxl","xlrd"])
    if isinstance(sheet_name,str):sheet_name=_resolve(file,sheet_name,engs)
    for eng in engs:
        try:
            file.seek(0)
            df=pd.read_excel(file,sheet_name=sheet_name,header=header,engine=eng)
            df.columns=[str(c).strip()for c in df.columns]
            return df.dropna(how="all").reset_index(drop=True)
        except(ImportError,Exception):continue
    try:
        file.seek(0)
        df=pd.read_excel(file,sheet_name=sheet_name,header=header)
        df.columns=[str(c).strip()for c in df.columns]
        return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"Cannot read '{name}': {e}")
        return pd.DataFrame()

def nc(df,cands,new):
    cs={c.strip().lower()for c in cands}
    for col in df.columns:
        if col.strip().lower()in cs:
            return df.rename(columns={col:new})if col!=new else df
    return df

def load_content(file):
    df=read_file(file,sheet_name="content")
    if df.empty:
        file.seek(0);df=read_file(file)
    df=nc(df,["Color_No","ColorNo","Color No","Article No","ArticleNo","PIM Article#","Style Number"],"Article_No")
    df=nc(df,["EAN","ean","Barcode","GTIN","UPC","Child SKU"],"EAN")
    if "Article_No"not in df.columns or"EAN"not in df.columns:
        st.error("Content: Missing Article_No or EAN column.")
        return pd.DataFrame(columns=["Article_No","EAN","Size"])
    df["Article_No"]=df["Article_No"].apply(_s).astype(str)
    df["EAN"]=df["EAN"].apply(_ean).astype(str)
    sc=next((c for c in df.columns if"size"in c.lower()and c not in("Article_No","EAN")),None)
    df["Size"]=df[sc].apply(_s)if sc else""
    df=df[(df["Article_No"]!="")&(df["EAN"]!="")]
    return df[["Article_No","EAN","Size"]].drop_duplicates("EAN").reset_index(drop=True)

def load_zecom(file, region="PH"):
    """
    Region-aware ZeCom loader.
    PH  → sheet "PH",  header row 2, Article col = PIM Article#
    MY  → sheet "MY",  header row 3, Article col = Style#
    SG  → sheet "SG",  header row 3, Article col = STYLE#

    Tracker columns: fuzzy match on column name containing keyword
    (case-insensitive, ignores spaces/underscores).
    Keywords: lazada | shopee | zalora | tiktok
    """
    sheet_pref = {"PH":"PH","MY":"MY","SG":"SG"}.get(region,"PH")
    # MY/SG confirmed header row = 3; PH = 2
    hdr_row = 2 if region == "PH" else 3

    df = pd.DataFrame()
    for sh in [sheet_pref, region, "Sheet1", "Sheet", 0]:
        try:
            file.seek(0)
            tmp = read_file(file, sheet_name=sh, header=hdr_row)
            if len(tmp) > 5:
                df = tmp; break
        except:
            continue
    if df.empty:
        file.seek(0); df = read_file(file, header=hdr_row)

    # Article No column — STYLE# takes priority for MY/SG, then common fallbacks
    art_candidates = [
        "STYLE#","Style#","style#",
        "PIM Article#","PIM Article","Article No","ArticleNo",
        "Color_No","ColorNo","Color No","Style Number","StyleNo",
    ]
    df = nc(df, art_candidates, "Article_No")

    # Tracker columns — FUZZY: find col whose name contains the keyword
    MP_KEYWORDS = {
        "Lazada" : "lazada",
        "Shopee" : "shopee",
        "Zalora" : "zalora",
        "TikTok" : "tiktok",
    }
    already_mapped = set()
    for mp, kw in MP_KEYWORDS.items():
        target = f"Tracker_{mp}"
        # Find first column not already mapped whose name contains the keyword
        match = next(
            (c for c in df.columns
             if kw in c.strip().lower().replace(" ","").replace("_","")
             and c not in already_mapped
             and c != "Article_No"),
            None
        )
        if match:
            df = df.rename(columns={match: target})
            already_mapped.add(target)

    df = nc(df,["Launch Dates","Launch Date","LaunchDate","Go Live","live date"],"Launch_Date")

    for col in [f"Tracker_{m}" for m in ["Lazada","Shopee","Zalora","TikTok"]] + ["Launch_Date","Article_No"]:
        if col not in df.columns:
            df[col] = np.nan

    df["Article_No"]  = df["Article_No"].apply(_s).astype(str)
    df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    df = df[df["Article_No"].str.match(r'^\S+.*\S+$', na=False) &
            (df["Article_No"].str.len() > 2)]
    return df.drop_duplicates("Article_No").reset_index(drop=True)

# EAN column priority per region (first match wins)
INV_EAN_COLS = {
    "PH": ["EAN","ean","Barcode","barcode","SKU","sku","Material"],
    "MY": ["Sku","SKU","sku","EAN","ean","Barcode","barcode","Material"],
    "SG": ["PROD_CODE","prod_code","Prod_Code","EAN","ean","Barcode","barcode","SKU","sku"],
}

def load_inv(file, region):
    """
    Region-aware inventory loader.
    EAN column priority:
      PH : EAN / Barcode / SKU
      MY : Sku / SKU / sku / EAN / Barcode
      SG : PROD_CODE / EAN / Barcode / SKU
    Stock column:
      PH : Avail_Qty
      MY : QtyAvailable
      SG : QTY
    Header row:
      PH / MY : 0 (standard)
      SG      : 4 (confirmed from file inspection)
    """
    # SG inventory has 4 metadata rows before headers
    hdr = 4 if region == "SG" else 0
    df = read_file(file, header=hdr)
    all_cols = list(df.columns)

    # Region-specific EAN column candidates
    ean_cands = INV_EAN_COLS.get(region, INV_EAN_COLS["PH"])
    df = nc(df, ean_cands, "EAN")

    if "EAN" not in df.columns:
        st.warning(f"[{region}] Inventory: EAN column not found. "
                   f"Expected one of: {ean_cands[:4]}. Columns: {all_cols[:10]}")
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {"error":"EAN not found"}

    df["EAN"] = df["EAN"].apply(_ean).astype(str)
    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]

    # Stock column
    primary = INV_COL.get(region, "Avail_Qty")
    sc = (primary if primary in df.columns else
          next((c for c in df.columns if c.strip().lower() == primary.lower()), None))
    if not sc:
        fallbacks = ["avail_qty","qtyavailable","qty","stock per ean",
                     "available","on hand","soh","quantity","stock"]
        sc = next((c for c in df.columns if c.strip().lower() in fallbacks), None)
        if sc:
            st.warning(f"[{region}] Inventory: using '{sc}' as stock col (expected '{primary}')")
        else:
            st.error(f"[{region}] Inventory: stock column '{primary}' not found. Cols: {all_cols}")

    df["Inv_Stock"] = (pd.to_numeric(df[sc], errors="coerce")
                       .fillna(0).clip(lower=0).astype(int)) if sc else 0
    dbg = {
        "Region"   : region,
        "EAN col"  : "EAN (from candidates)",
        "Stock col": sc or "NOT FOUND",
        "EAN rows" : len(df),
        "Non-zero" : int((df["Inv_Stock"] > 0).sum()),
    }
    return df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True), dbg

def _mp(file,mp,sht=None):
    cfg=MP_CFG[mp]
    df=read_file(file,sheet_name=sht)if sht else read_file(file)
    df=nc(df,[cfg["ean"],"SellerSKU","SKU","Seller sku","SellerSku"],"EAN")
    df=nc(df,[cfg["status"],"Status","status","ItemStatus"],"MP_Status")
    df=nc(df,[cfg["stock"],"Stock","Quantity","Available","Qty"],"MP_Stock")
    if cfg["id"]:df=nc(df,[cfg["id"],"ItemId","item_id","Product ID","ProductId"],"MP_ID")
    if"MP_ID"not in df.columns:df["MP_ID"]=""
    for col in["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns:df[col]=np.nan
    df["EAN"]=df["EAN"].apply(_ean).astype(str)
    df["MP_Stock"]=df["MP_Stock"].apply(_i)
    df["MP_Status"]=df["MP_Status"].apply(_s)
    df["MP_ID"]=df["MP_ID"].apply(lambda v:_s(str(v).split(".")[0]))
    df["Marketplace"]=mp
    df=df[df["EAN"].str.match(r'^\d{8,}$',na=False)]
    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")

def load_lazada(f):return _mp(f,"Lazada","template")
def load_shopee(f):return _mp(f,"Shopee")
def load_tiktok(f):return _mp(f,"TikTok")

def load_zalora(sf,stf):
    ds=read_file(sf,sheet_name="ProductStatuses")
    dk=read_file(stf,sheet_name="Sheet")
    # EAN = SellerSku, PID = ShopSku
    ds=nc(ds,["SellerSku","SellerSKU","Seller SKU"],"EAN")
    ds=nc(ds,["ShopSku","shopsku","Shop SKU","ShopSKU"],"MP_ID")
    ds=nc(ds,["Status","status"],"MP_Status")
    dk=nc(dk,["SellerSku","SellerSKU","Seller SKU"],"EAN")
    dk=nc(dk,["Quantity","Stock","Available","Qty"],"MP_Stock")
    # Also grab ShopSku from stock file as backup
    dk=nc(dk,["ShopSku","shopsku","Shop SKU","ShopSKU"],"ShopSku_stk")
    for col in["EAN","MP_Status"]:
        if col not in ds.columns:ds[col]=np.nan
    if"MP_ID"not in ds.columns:ds["MP_ID"]=""
    if"EAN"not in dk.columns:dk["EAN"]=np.nan
    if"MP_Stock"not in dk.columns:dk["MP_Stock"]=0
    if"ShopSku_stk"not in dk.columns:dk["ShopSku_stk"]=""
    ds["EAN"]=ds["EAN"].apply(_ean).astype(str)
    dk["EAN"]=dk["EAN"].apply(_ean).astype(str)
    ds["MP_Status"]=ds["MP_Status"].apply(_s)
    ds["MP_ID"]=ds["MP_ID"].apply(_s)
    m=ds.merge(dk[["EAN","MP_Stock","ShopSku_stk"]].drop_duplicates("EAN"),on="EAN",how="left")
    m["MP_Stock"]=m["MP_Stock"].apply(_i)
    # Fill MP_ID from stock file if missing in status file
    m["MP_ID"]=m.apply(
        lambda r: r["MP_ID"] if r["MP_ID"] else _s(str(r.get("ShopSku_stk",""))),axis=1)
    m["Marketplace"]="Zalora"
    m["EAN"]=m["EAN"].astype(str)
    m=m[m["EAN"].str.match(r'^\d{8,}$',na=False)]
    return m[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")

def run_audit(mp_dfs,inv_df,zecom_df,content_df,region):
    art_eans={};ean_size={};ean_art={}
    for _,row in content_df.iterrows():
        art=row["Article_No"];ean=row["EAN"]
        if art and ean:
            art_eans.setdefault(art,[]).append(ean)
            ean_size[ean]=row.get("Size","")
    inv={}
    if not inv_df.empty:
        for _,row in inv_df.iterrows():inv[row["EAN"]]=int(row["Inv_Stock"])
    results={}
    for mp in REGION_MARKETPLACES.get(region,["Lazada","Shopee","Zalora"]):
        mp_df=mp_dfs.get(mp,pd.DataFrame())
        mpx={}
        if not mp_df.empty:
            for _,row in mp_df.iterrows():
                if row["EAN"]:mpx[row["EAN"]]=row
        buf=STOCK_BUFFER.get((region,mp),0)
        al=[];na=[];av=[];fn=[]
        for _,z in zecom_df.iterrows():
            art=z["Article_No"]
            tr=_s(z.get(f"Tracker_{mp}","")).upper()
            ld=z["Launch_Date"]
            if tr!="YES":continue
            if pd.notna(ld)and ld>FUTURE_WINDOW:continue
            eans=art_eans.get(art,[])
            if not eans:continue
            ld_str=ld.date().isoformat()if pd.notna(ld)else"-"
            in_stk=[e for e in eans if max(0,inv.get(e,0)-buf)>0]
            if not in_stk:continue
            listed=[e for e in in_stk if e in mpx]
            missing=[e for e in in_stk if e not in mpx]
            # Already listed — all EANs currently on MP (regardless of stock)
            for ean in[e for e in eans if e in mpx]:
                row=mpx[ean]
                al.append({
                    "Article No":art,"EAN":ean,"Size":ean_size.get(ean,""),
                    "MP Product ID":_s(row.get("MP_ID","")),
                    "MP Status":_s(row.get("MP_Status","")),
                    "Inventory Stock":max(0,inv.get(ean,0)-buf),
                    "MP Stock":int(row.get("MP_Stock",0)),
                    "Launch Date":ld_str,
                })
            if not missing:
                na.append({
                    "Article No":art,"Total EANs":len(eans),
                    "In-Stock EANs":len(in_stk),"Listed EANs":len(listed),
                    "Launch Date":ld_str,"Tracker":tr,
                })
            elif len(listed)==0:
                for ean in missing:
                    fn.append({
                        "Article No":art,"EAN":ean,"Size":ean_size.get(ean,""),
                        "Inventory Stock":max(0,inv.get(ean,0)-buf),
                        "Launch Date":ld_str,"Tracker":tr,
                    })
            else:
                for ean in missing:
                    av.append({
                        "Article No":art,"Missing EAN":ean,
                        "Size":ean_size.get(ean,""),
                        "Inventory Stock":max(0,inv.get(ean,0)-buf),
                        "Already Listed":len(listed),
                        "Total In-Stock":len(in_stk),
                        "Launch Date":ld_str,
                    })
        results[mp]={
            "Already Listed"  :pd.DataFrame(al),
            "No Action Needed":pd.DataFrame(na),
            "Add Variant"     :pd.DataFrame(av),
            "Full New Listing":pd.DataFrame(fn),
        }
    return results

def build_excel(rdata,region):
    import xlsxwriter
    out=BytesIO()
    mps=REGION_MARKETPLACES.get(region,["Lazada","Shopee","Zalora"])
    tabs=["Already Listed","No Action Needed","Add Variant","Full New Listing"]
    with pd.ExcelWriter(out,engine="xlsxwriter")as writer:
        wb=writer.book
        def F(**kw):
            d={"font_name":"Arial","font_size":9,"border":1,"valign":"vcenter","align":"left"}
            d.update(kw);return wb.add_format(d)
        norm=F();ttl=F(bold=True,font_size=13,font_color="#0f3460",border=0)
        sub=F(italic=True,font_size=8,font_color="#718096",border=0)
        def ch(bg):return F(bold=True,bg_color=bg,font_color="#ffffff",align="center",text_wrap=True)
        def cd(bg,fc):return F(bg_color=bg,font_color=fc)
        def sv(v):return""if(isinstance(v,float)and np.isnan(v))else str(v)
        # Summary
        ws=wb.add_worksheet("Summary");writer.sheets["Summary"]=ws
        ws.write(0,0,f"PUMA Listing Audit — {region}",ttl)
        ws.write(1,0,f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}",sub)
        ws.write(1,3,f"Eligible = Tracker YES + Launch <=today+30d + Stock>0",sub)
        for ci,h in enumerate(["Marketplace","Already Listed","No Action Needed","Add Variant","Full New Listing"]):
            ws.write(3,ci,h,ch("#0f3460"));ws.set_column(ci,ci,22)
        ws.set_row(3,28)
        r=4
        for mp in mps:
            mpd=rdata.get(mp,{})
            ws.write(r,0,mp,norm)
            for ci,cat in enumerate(tabs):
                cnt=len(mpd.get(cat,pd.DataFrame()))
                col=TAB_COLORS[cat]
                ws.write(r,ci+1,cnt,cd(col["bg"],col["fc"]))
            r+=1
        # ── Tracker Analysis sheet ───────────────────────────────────────────
        ws_ta = wb.add_worksheet("Tracker Analysis")
        writer.sheets["Tracker Analysis"] = ws_ta
        ws_ta.write(0, 0, f"Tracker Analysis | {region}", ttl)
        ws_ta.write(1, 0, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", sub)

        def grp_fmt(bg): return F(bold=True,bg_color=bg,font_color="#ffffff",align="center",text_wrap=True,border=1)
        grp_hdr_f  = grp_fmt("#0f3460")
        col_hdr_f  = grp_fmt("#1a3a5c")
        listed_f   = F(bg_color="#c6efce",font_color="#276221")
        notlist_f  = F(bg_color="#ffc7ce",font_color="#9c0006")
        yes_f      = F(bg_color="#d6eaf8",font_color="#1a5276")
        num_f      = F(align="center",border=1)

        # Collect all EANs from all categories across all MPs
        ean_data = {}
        for mp in mps:
            cats = rdata.get(mp,{})
            for cat_name in ["Already Listed","Add Variant","Full New Listing","No Action Needed"]:
                df = cats.get(cat_name, pd.DataFrame())
                if df.empty: continue
                for _,row in df.iterrows():
                    ean = str(row.get("Missing EAN","")) if cat_name=="Add Variant" else str(row.get("EAN",""))
                    if not ean or ean=="nan": continue
                    art = str(row.get("Article No",""))
                    ld  = str(row.get("Launch Date","-"))
                    inv = row.get("Inventory Stock",0)
                    if ean not in ean_data:
                        ean_data[ean] = {"Color No":art,"EAN":ean,"Launch Date":ld,"Inv_Stock":inv}
                        for m in mps:
                            ean_data[ean][f"Tracker_{m}"] = "-"
                            ean_data[ean][f"Status_{m}"]  = "Not Listed"
                            ean_data[ean][f"PID_{m}"]     = ""
                    ean_data[ean][f"Tracker_{mp}"] = "YES"
                    if cat_name == "Already Listed":
                        ean_data[ean][f"Status_{mp}"] = "Listed"
                        ean_data[ean][f"PID_{mp}"]    = sv(row.get("MP Product ID",""))
                    elif cat_name in ("Add Variant","Full New Listing"):
                        ean_data[ean][f"Status_{mp}"] = "Not Listed"
                    elif cat_name == "No Action Needed":
                        ean_data[ean][f"Status_{mp}"] = "Listed"

        ta_rows = sorted(ean_data.values(), key=lambda x:(x["Color No"],x["EAN"]))
        ws_ta.write(1,3,f"Total EANs: {len(ta_rows):,}",sub)

        # Group header row (row 2), column header row (row 3), data from row 4
        GRP=2; COL=3; DAT=4
        # Fixed cols
        ws_ta.write(GRP,0,"Color No",grp_hdr_f); ws_ta.set_column(0,0,16)
        ws_ta.write(GRP,1,"EAN",grp_hdr_f);      ws_ta.set_column(1,1,18)
        ws_ta.write(GRP,2,"Launch Date",grp_hdr_f); ws_ta.set_column(2,2,14)

        cc=3
        # ZeCom Tracker group
        ws_ta.merge_range(GRP,cc,GRP,cc+len(mps)-1,"ZeCom Tracker",grp_hdr_f)
        zc_cols={}
        for mp in mps:
            ws_ta.write(COL,cc,mp,col_hdr_f); ws_ta.set_column(cc,cc,10); zc_cols[mp]=cc; cc+=1

        # Inventory group
        ws_ta.merge_range(GRP,cc,GRP,cc,"Inventory",grp_hdr_f)
        ws_ta.write(COL,cc,"Quantity",col_hdr_f); ws_ta.set_column(cc,cc,12); inv_c=cc; cc+=1

        # Marketplace Status group
        ws_ta.merge_range(GRP,cc,GRP,cc+len(mps)-1,"Marketplace Status",grp_hdr_f)
        st_cols={}
        for mp in mps:
            ws_ta.write(COL,cc,f"{mp} Status",col_hdr_f); ws_ta.set_column(cc,cc,14); st_cols[mp]=cc; cc+=1

        # Marketplace PID group
        ws_ta.merge_range(GRP,cc,GRP,cc+len(mps)-1,"Marketplace PID",grp_hdr_f)
        pid_c={}
        for mp in mps:
            ws_ta.write(COL,cc,f"{mp} PID",col_hdr_f); ws_ta.set_column(cc,cc,26); pid_c[mp]=cc; cc+=1

        ws_ta.freeze_panes(DAT,3)
        ws_ta.set_row(GRP,20); ws_ta.set_row(COL,20)

        # Data
        for ri,rec in enumerate(ta_rows):
            r=DAT+ri
            ws_ta.write(r,0,rec["Color No"],norm)
            ws_ta.write(r,1,rec["EAN"],norm)
            ws_ta.write(r,2,rec["Launch Date"],norm)
            for mp in mps:
                tr=rec.get(f"Tracker_{mp}","-")
                ws_ta.write(r,zc_cols[mp],tr,yes_f if tr=="YES" else norm)
            ws_ta.write_number(r,inv_c,int(rec.get("Inv_Stock",0)),num_f)
            for mp in mps:
                st=rec.get(f"Status_{mp}","Not Listed")
                ws_ta.write(r,st_cols[mp],st,listed_f if st=="Listed" else notlist_f)
            for mp in mps:
                ws_ta.write(r,pid_c[mp],str(rec.get(f"PID_{mp}",""))or"",norm)

        # ── Detail sheets ────────────────────────────────────────────────────
        num_cols={"Inventory Stock","MP Stock","Total EANs","In-Stock EANs",
                  "Listed EANs","Already Listed","Total In-Stock"}
        for mp in mps:
            mpd=rdata.get(mp,{})
            for cat in tabs:
                df=mpd.get(cat,pd.DataFrame())
                sn=f"{mp} - {cat}"[:31]
                col=TAB_COLORS[cat]
                ws2=wb.add_worksheet(sn);writer.sheets[sn]=ws2
                ws2.write(0,0,f"{mp} — {cat} | {region}",ttl)
                ws2.write(1,0,f"Records: {len(df):,}",sub)
                if df.empty:ws2.write(2,0,"No records.",sub);continue
                hf=ch(col["hdr"]);rf=cd(col["bg"],col["fc"])
                df=df.reset_index(drop=True)
                for ci,c in enumerate(df.columns):
                    ws2.write(2,ci,c,hf)
                    try:w=max(len(str(c)),int(df[c].astype(str).str.len().max()))
                    except:w=len(str(c))
                    ws2.set_column(ci,ci,min(w+3,45))
                ws2.freeze_panes(3,0)
                for ri,(_,rec)in enumerate(df.iterrows()):
                    for ci,c in enumerate(df.columns):
                        v=rec[c];s=sv(v)
                        if c in num_cols:
                            try:ws2.write_number(ri+3,ci,int(float(s))if s else 0,rf)
                            except:ws2.write(ri+3,ci,s,rf)
                        else:ws2.write(ri+3,ci,s,norm)
    return out.getvalue()

# Session state
try:_=st.session_state["lr"]
except(KeyError,AttributeError):st.session_state["lr"]={}
try:_=st.session_state["rd"]
except(KeyError,AttributeError):st.session_state["rd"]=[]

tab_up,tab_res,tab_help=st.tabs(["Upload & Run","Results & Download","Help"])

with tab_help:
    st.markdown("""
## Rules
- **Tracker = YES** + **Launch Date <= today+30 days** + **Stock > 0**
- Special Request file: ignored in this app

| Tab | Meaning |
|---|---|
| Already Listed | All EANs on MP now — with Inventory Stock + MP Product ID |
| No Action Needed | All in-stock EANs already listed — nothing to do |
| Add Variant | Some in-stock EANs missing from MP |
| Full New Listing | Eligible article — zero EANs on MP at all |

**Output:** One Excel per region (PH/MY/SG), 4 tabs per marketplace.
PH Lazada buffer: Avail_Qty - 1.
    """)

with tab_up:
    st.markdown("### Step 1 — Select Regions")
    sel_r=st.multiselect("Regions:",REGIONS,default=["PH"])
    st.markdown("### Step 2 — Content Master *(all regions)*")
    cf=st.file_uploader("Content Master [required]",type=FILE_TYPES,key="cf")
    st.markdown("### Step 3 — Region Files")
    rf2={}
    for region in sel_r:
        with st.expander(f"{region}",expanded=True):
            rf2[region]={}
            rf2[region]["zecom"]=st.file_uploader(f"ZeCom Tracker ({region}) [required]",type=FILE_TYPES,key=f"z_{region}")
            rf2[region]["inv"]=st.file_uploader(f"Inventory ({region})",type=FILE_TYPES,key=f"i_{region}",help="PH:Avail_Qty MY:QtyAvailable SG:QTY")
            c1,c2=st.columns(2)
            with c1:
                rf2[region]["laz"]=st.file_uploader(f"Lazada ({region})",type=FILE_TYPES,key=f"l_{region}")
                rf2[region]["sho"]=st.file_uploader(f"Shopee ({region})",type=FILE_TYPES,key=f"s_{region}")
            with c2:
                rf2[region]["zst"]=st.file_uploader(f"Zalora Status ({region})",type=FILE_TYPES,key=f"zs_{region}")
                rf2[region]["zsk"]=st.file_uploader(f"Zalora Stock ({region})",type=FILE_TYPES,key=f"zk_{region}")
                rf2[region]["ttk"]=st.file_uploader(f"TikTok ({region}) MY only",type=FILE_TYPES,key=f"t_{region}")
    st.markdown("---")
    if st.button("Run Listing Audit",type="primary",use_container_width=True):
        errs=[]
        if not cf:errs.append("Content Master required.")
        if not sel_r:errs.append("Select a region.")
        for rg in sel_r:
            if not rf2.get(rg,{}).get("zecom"):errs.append(f"[{rg}] ZeCom required.")
        for e in errs:st.error(e)
        if not errs:
            prog=st.progress(0,text="Loading...")
            with st.spinner("Content..."):
                content_df=load_content(cf)
            st.success(f"Content: {len(content_df):,} EANs / {content_df['Article_No'].nunique():,} articles")
            prog.progress(10)
            all_res={};step=10;ssz=max(1,int(85/max(len(sel_r),1)))
            for region in sel_r:
                r=rf2.get(region,{})
                st.markdown(f"#### {region}")
                with st.spinner(f"[{region}] ZeCom..."):
                    zecom_df=load_zecom(r["zecom"], region)
                st.write(f"  ZeCom: {len(zecom_df):,} articles")
                inv_df=pd.DataFrame(columns=["EAN","Inv_Stock"])
                if r.get("inv"):
                    with st.spinner(f"[{region}] Inventory..."):
                        inv_df,dbg=load_inv(r["inv"],region)
                    st.write(f"  Inventory: {len(inv_df):,} EANs | col={dbg.get('Used col','?')} | non-zero={dbg.get('Non-zero',0):,}")
                else:
                    st.warning(f"[{region}] No inventory — stock=0")
                mp_dfs={}
                if r.get("laz"):
                    mp_dfs["Lazada"]=load_lazada(r["laz"])
                    st.write(f"  Lazada: {len(mp_dfs['Lazada']):,} EANs")
                if r.get("sho"):
                    mp_dfs["Shopee"]=load_shopee(r["sho"])
                    st.write(f"  Shopee: {len(mp_dfs['Shopee']):,} EANs")
                if r.get("zst")and r.get("zsk"):
                    mp_dfs["Zalora"]=load_zalora(r["zst"],r["zsk"])
                    st.write(f"  Zalora: {len(mp_dfs['Zalora']):,} EANs")
                elif r.get("zst"):
                    st.warning(f"[{region}] Zalora stock file missing")
                if r.get("ttk"):
                    mp_dfs["TikTok"]=load_tiktok(r["ttk"])
                    st.write(f"  TikTok: {len(mp_dfs['TikTok']):,} EANs")
                if not mp_dfs:
                    st.warning(f"[{region}] No MP files — skip.")
                    step+=ssz;prog.progress(min(step,95));continue
                prog.progress(min(step+ssz//2,95),text=f"[{region}] Auditing {len(zecom_df):,} articles...")
                res=run_audit(mp_dfs,inv_df,zecom_df,content_df,region)
                all_res[region]=res
                for mp,cats in res.items():
                    st.write(f"  {mp}: Full New={len(cats.get('Full New Listing',pd.DataFrame())):,} | "
                             f"Add Variant={len(cats.get('Add Variant',pd.DataFrame())):,} | "
                             f"Already Listed={len(cats.get('Already Listed',pd.DataFrame())):,}")
                step+=ssz;prog.progress(min(step,95))
            prog.progress(100,text="Done!")
            st.session_state["lr"]=all_res
            st.session_state["rd"]=sel_r
            if all_res:st.success("Done! Go to Results & Download tab.")

with tab_res:
    results=st.session_state["lr"]
    rr=st.session_state["rd"]
    if not results:
        st.info("Run the audit first.")
    else:
        fnl=sum(len(c.get("Full New Listing",pd.DataFrame()))for rv in results.values()for c in rv.values())
        av=sum(len(c.get("Add Variant",pd.DataFrame()))for rv in results.values()for c in rv.values())
        al=sum(len(c.get("Already Listed",pd.DataFrame()))for rv in results.values()for c in rv.values())
        k1,k2,k3=st.columns(3)
        k1.markdown(f"<div class='metric-box'><div class='mv cp'>{fnl:,}</div><div class='ml'>Full New Listing (EANs)</div></div>",unsafe_allow_html=True)
        k2.markdown(f"<div class='metric-box'><div class='mv co'>{av:,}</div><div class='ml'>Add Variant (EANs)</div></div>",unsafe_allow_html=True)
        k3.markdown(f"<div class='metric-box'><div class='mv cg'>{al:,}</div><div class='ml'>Already Listed (EANs)</div></div>",unsafe_allow_html=True)
        brows=[]
        for region in rr:
            for mp,cats in results.get(region,{}).items():
                brows.append({"Region":region,"Marketplace":mp,
                    "Already Listed":len(cats.get("Already Listed",pd.DataFrame())),
                    "No Action Needed":len(cats.get("No Action Needed",pd.DataFrame())),
                    "Add Variant":len(cats.get("Add Variant",pd.DataFrame())),
                    "Full New Listing":len(cats.get("Full New Listing",pd.DataFrame())),
                })
        st.dataframe(pd.DataFrame(brows),use_container_width=True)
        d1,d2,d3=st.columns(3)
        dr=d1.selectbox("Region",rr)
        dmp=d2.selectbox("Marketplace",list(results.get(dr,{}).keys()))
        dc=d3.selectbox("Category",["Already Listed","No Action Needed","Add Variant","Full New Listing"])
        dv=results.get(dr,{}).get(dmp,{}).get(dc,pd.DataFrame())
        st.caption(f"{len(dv):,} records")
        st.dataframe(dv,use_container_width=True,height=400)
        st.markdown("---")
        st.markdown("### Download Reports")
        cols=st.columns(max(len(rr),1))
        for i,region in enumerate(rr):
            rd=results.get(region,{})
            if rd:
                with cols[i]:
                    x=build_excel(rd,region)
                    fn=f"PUMA_Listing_Audit_{region}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button(f"Download {region}",data=x,file_name=fn,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,type="primary")
        st.caption("One Excel per region | 4 tabs per marketplace")
