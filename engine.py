"""
engine.py
IQ + SIRFIS + Overrides 데이터를 받아 JT~KE 계산을 수행한다.
입력:
    iq_df       : IQ Sched Status Rpt (A~JN)
    sirfis_df   : SIRFIS 시트
    overrides_df: Overrides 조정값 테이블
    use_override: bool (Frontend에서 전달 — C2 스위치)
출력:
    iq_df에 JT~KE 컬럼이 추가된 DataFrame
"""

import pandas as pd
import numpy as np


# ---------------------------------------------------------------------------
# SIRFIS 헬퍼
# ---------------------------------------------------------------------------

def _normalize_supplier(name: str) -> str:
    """Supplier 전체 이름을 IQ RPT의 단축 이름으로 변환."""
    name = name.upper()
    if "LAM" in name:
        return "Lam"
    if "APPLIED MATERIALS" in name or "AMAT" in name:
        return "AMAT"
    if "TOKYO ELECTRON" in name or "TEL" in name:
        return "TEL"
    if "ASM" in name:
        return "ASMI"
    if "SCREEN" in name:
        return "SCREEN"
    if "HITACHI" in name:
        return "Hitachi"
    if "KLA" in name:
        return "KLA"
    return name.title()


def _build_sirfis_lookup(sirfis_df: pd.DataFrame) -> dict:
    """
    Need ID(숫자) 기준으로 SIRFIS 데이터를 dict로 인덱싱.
    반환: {need_id: {"supplier": ..., "type": ..., "mt_rdd": ..., "mt_sdd": ...,
                     "ck_rdd": ..., "ck_sdd": ...}}
    SIRFIS 컬럼 참조:
        C = Need ID, G = Supplier Name, J = Category(MT/CK),
        S = RDD, T = SDD
    """
    lookup = {}
    for _, row in sirfis_df.iterrows():
        need_id = row.get("Need ID")
        if need_id is None:
            continue
        try:
            need_id = int(need_id)
        except (ValueError, TypeError):
            continue

        category = str(row.get("Category", "")).strip()
        supplier = _normalize_supplier(str(row.get("Supplier Name", "")).strip())
        rdd = row.get("RDD")
        sdd = row.get("SDD")

        if need_id not in lookup:
            lookup[need_id] = {"supplier": supplier, "category": category,  "mt_rdd": None, "mt_sdd": None,
                               "ck_rdd": None, "ck_sdd": None}

        if category == "MT":
            # mirrors Excel XLOOKUP: use first match only
            if lookup[need_id]["mt_rdd"] is None:
                lookup[need_id]["mt_rdd"] = rdd
            if lookup[need_id]["mt_sdd"] is None:
                lookup[need_id]["mt_sdd"] = sdd
            if not lookup[need_id]["supplier"]:
                lookup[need_id]["supplier"] = supplier
            lookup[need_id]["category"] = category
        elif category == "CK":
            if lookup[need_id]["ck_rdd"] is None:
                lookup[need_id]["ck_rdd"] = rdd
            if lookup[need_id]["ck_sdd"] is None:
                lookup[need_id]["ck_sdd"] = sdd
            lookup[need_id]["category"] = category

    return lookup


# ---------------------------------------------------------------------------
# computing columns JT to KE
# ---------------------------------------------------------------------------

def compute_jt_ke(selected_supplier, iq_df, sirfis_df, overrides_df,
                  use_override, reference_date, 
                  match_sdd_rdd):

    # start with IQ Dataframe
    df = iq_df.copy()

    # print("iq_df---------------------", len(df))
    # a = input("Press Enter to continue...")


    # Need ID = A컬럼 우측 6자리 숫자
    df["_need_id"] = df["Event ID"].apply(lambda x: int(str(int(x))[-6:]) if pd.notna(x) else None)
    sirfis = _build_sirfis_lookup(sirfis_df)

    # print("sirfis", sirfis)
    # a = input()

    # Override adjustments lookup --- later need to fix how to get override    
    # could be later Json file
    override_map = {}
    for _, row in overrides_df.iterrows():
        entity = row.get("Event-Life")
        if pd.isna(entity) or entity is None:
            continue
        override_map[str(entity).strip()] = {"sdd":row.get("SDD"), "mrcl_delta": row.get("MRCL Finish Delta"), "wstie": row.get("WSTie")}

    # print("override_map", override_map)
    # a = input()    

    results = []

    for _ , row in df.iterrows():
        # column A event id - > last 6 digit number
        need_id   = row["_need_id"]        
        # column G Entity code- life
        entity    = str(row.get("Entity Code - Life", "")).strip()
        # fullfillment type
        ftype     = str(row.get("Fulfillment Type", "")).strip()
        # event type
        evt_type  = str(row.get("Event Type", "")).strip()
        # ws_tie
        ws_tie    = row.get("WS Tie")       # M 컬럼 원본 WSTie (엑셀 KE 아님)
        # mrcl -- column JD
        mrcl      = row.get("MRCL Finish")  # JD 컬럼
        # sdd
        p3_sdd    = row.get("SDD")          # GL 컬럼 (P3 MT SDD)
        # ck_sdd
        p3_ck_sdd = row.get("Conv Kit SDD")  # GQ 컬럼 (P3 CK SDD)

        # Rdd
        col_GJ  = row.get("RDD")          # GJ 컬럼 (RDD)
        # ck_Rdd
        col_GO = row.get("Conv Kit RDD")  # GO 컬럼 (CK RDD)

        sirfis_lookup = sirfis.get(need_id, {}) if need_id else {}
       
        overide_lookup = override_map.get(entity, {}) if use_override else {}      



        # new column generation

        # JT: Supplier (IQ RPT 수동 입력값 직접 사용) --- right now we are using this, but later need to make a mapping
        # col_jt = str(row.get("Supplier", "")).strip()
        col_jt = sirfis_lookup.get("supplier", "") or str(row.get("Supplier", "")).strip() or str(row.get("_excel_jt", "")).strip()
      

        # JU: Set Start < SDD? (x if condition met)

        # read get start, if empty it returns None: column GT
        set_start = row.get("Set Start")


        # print("sirfis_lookup", sirfis_lookup)
        # print("overide_lookup", overide_lookup)
        # print("p3_sdd", p3_sdd)
        # a = input()

        # temporay ka_value
        column_ka_sdd = overide_lookup.get("sdd") or sirfis_lookup.get("mt_sdd") or p3_sdd 
      
        if "Conv" in evt_type:
            conv_start = row.get("Conversion Start")
            ck_sdd = sirfis_lookup.get("ck_sdd")            
            col_ju = "x" if (conv_start and ck_sdd and conv_start < ck_sdd) else ""

        else:
            col_ju = "x" if (set_start and column_ka_sdd and set_start < column_ka_sdd) else ""

        # JV: SIRFIS MT RDD
        if not sirfis_lookup:
            col_jv = "not found"

        elif sirfis_lookup.get("category") == "MT":
            mt_rdd = sirfis_lookup.get("mt_rdd")
            col_jv = mt_rdd if mt_rdd is not None else "no SIRFIS RDD"
        else:
            col_jv = ""

        # JW: SIRFIS MT SDD
        if not sirfis_lookup:
            col_jw = "not found"
        elif sirfis_lookup.get("category") == "MT":
            mt_sdd = sirfis_lookup.get("mt_sdd")
            col_jw = mt_sdd if mt_sdd is not None else "no SIRFIS SDD"
        else:
            col_jw = ""

        # JX: SIRFIS CK RDD
        if not sirfis_lookup:
            col_jx = "not found"
        elif sirfis_lookup.get("category") == "CK":
            ck_rdd = sirfis_lookup.get("ck_rdd")
            col_jx = ck_rdd if ck_rdd is not None else "no SIRFIS RDD"
        else:
            col_jx = ""

        # JY: SIRFIS CK SDD
        if not sirfis_lookup:
            col_jy = "not found"
        elif sirfis_lookup.get("category") == "CK":
            ck_sdd = sirfis_lookup.get("ck_sdd")
            col_jy = ck_sdd if ck_sdd is not None else "no SIRFIS SDD"
        else:
            col_jy = ""
       

        # JZ: Override SDD
        col_jz = overide_lookup.get("sdd") if use_override else None


        # KA: SDD to be used
        jw_num = col_jw if (col_jw and not isinstance(col_jw, str)) else None
        jy_num = col_jy if (col_jy and not isinstance(col_jy, str)) else None
        max_sirfis = max(filter(lambda x: x is not None, [jw_num, jy_num]), default=None)

        if match_sdd_rdd == "yes":                                             # H2 = "yes"
            col_ka = col_GJ if "Install" in evt_type else col_GO

        elif use_override and col_jz and pd.notna(col_jz):           # C2 + JZ
            col_ka = col_jz

        elif col_jw == "no SIRFIS SDD" or col_jy == "no SIRFIS SDD":
            col_ka = "no SIRFIS SDD"

        elif max_sirfis:
            col_ka = max_sirfis

        else:
            # mirrors Excel SWITCH(J2, "Install", ..., "*Conv*", ..., "reuse")
            # SWITCH has no wildcard support → "Conv" hits the default "reuse"
            if evt_type == "Install":
                if ftype != "Buy":
                    col_ka = "reuse"
                elif p3_sdd:
                    col_ka = p3_sdd
                elif set_start and set_start <= reference_date:
                    col_ka = "docked"
                else:
                    col_ka = "no forecast?"
            else:
                col_ka = "reuse"

        

        # KB: MRCL Adjustment (Days) from Overrides
        col_kb = overide_lookup.get("mrcl_delta") if use_override else None

        # KC: MRCL to be used
        if use_override and col_kb is not None:
            try:
                col_kc = mrcl + pd.Timedelta(days=int(col_kb)) if pd.notna(mrcl) else mrcl
            except Exception:
                col_kc = mrcl
        else:
            col_kc = mrcl

        # KD: Override WSTie
        col_kd = overide_lookup.get("wstie") if use_override else None

        # KE: WSTie to be used
        if use_override and col_kd is not None and pd.notna(col_kd):
            col_ke = col_kd
        else:
            col_ke = ws_tie

        results.append({
            "Supplier": col_jt, "Set Start < SDD?": col_ju,
            "SIRFIS MT RDD": col_jv, "SIRFIS MT SDD": col_jw,
            "SIRFIS CK RDD": col_jx, "SIRFIS CK SDD": col_jy,
            "Override SDD": col_jz,
            "SDD to be used (Override->SIRFIS->P3)": col_ka,
            "IQ/MRCL Adjustment (Days)": col_kb,
            "MRCL to be used (includes IQ/MRCL Adjustments)": col_kc,
            "Override WSTie": col_kd,
            "WSTie to be used": col_ke,
        })

    df_cal = pd.DataFrame(results)
    # pandas infers datetime64, converting sentinel strings to NaT
    # fix: convert to object dtype first, then restore strings cell-by-cell
    sentinel_cols = ["SDD to be used (Override->SIRFIS->P3)", "SIRFIS MT RDD", "SIRFIS MT SDD", "SIRFIS CK RDD", "SIRFIS CK SDD"]
    for col in sentinel_cols:
        if col in df_cal.columns and any(isinstance(r.get(col), str) for r in results):
            df_cal[col] = df_cal[col].astype(object)
            for i, r in enumerate(results):
                v = r.get(col)
                if isinstance(v, str):
                    df_cal.at[i, col] = v
    df = df.drop(columns=["_need_id"])
    new_df = pd.concat([df, df_cal], axis=1)

    # print("new_df---------------------", len(new_df))
    # a = input("Press Enter to continue...")

    # final_df = new_df[new_df["Supplier"] == selected_supplier].copy() 

    return new_df


