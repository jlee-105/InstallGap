"""
dgs.py
JT~KE가 계산된 IQ DataFrame을 받아 Delivery Gap Summary를 생성한다.
필터: Supplier=Lam, WSTie 7000~17000(숫자), Event Type != Qual
"""

import pandas as pd
from datetime import datetime
import numpy as np
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)


def compute_dgs(selected_supplier, we_tie, iq_extended, report_date):


    """
    Parameters:
        supplier
        iq_extended : compute_jt_ke() 결과 DataFrame (JT~KE 포함)
        report_date : 보고 기준일 (e.g. datetime(2026, 7, 24))       

    Returns: DGS DataFrame (Entity별 행)
    """
    df = iq_extended.copy()

    # mirrors Excel: JT=$B$1, LEFT(J,4)<>"Qual", KE>=$B$2, KE<17001, ISNUMBER(KE), SORTBY(...,KE,1)
    ke = pd.to_numeric(df["WSTie to be used"], errors="coerce")
    df = df[
        (df["Supplier"] == selected_supplier) &
        (~df["Event Type"].str.startswith("Qual", na=False)) &
        (ke >= we_tie) &
        (ke < 17001) &
        ke.notna()
    ].sort_values("WSTie to be used")
   
    dgs = df.copy()
    dgs = dgs.reset_index(drop=True)


    # 2. 컬럼 계산
    results = []
    for _, row in dgs.iterrows():
        entity    = row.get("Entity Code - Life")
        ceid      = row.get("CEID")
        evt_type  = str(row.get("Event Type", ""))
        wstie     = row.get("WSTie to be used")
        cnd       = row.get("CND")
        ka        = row.get("SDD to be used (Override->SIRFIS->P3)")   # G: MT/CK SDD
        kc        = row.get("MRCL to be used (includes IQ/MRCL Adjustments)") or "no MRCL date"  # I: MRCL
        need_id   = str(int(row["Event ID"]))[-6:] if pd.notna(row.get("Event ID")) else ""

        # F: MT/CK RDD — mirrors Excel GJ(RDD) for Install, GO(Conv Kit RDD) otherwise
        if evt_type == "Install":
            f_rdd = row.get("RDD")
        else:
            f_rdd = row.get("Conv Kit RDD")

        # H: Set Start or Conv Start — mirrors Excel GT(Set Start) for Install, IM(Conv Start) otherwise
        if evt_type == "Install":
            h_start = row.get("Set Start") or "no Set start"
        else:
            h_start = row.get("Conversion Start") or "no Conv Start"

        # L: MRCL if Set matched to SDD — mirrors Excel: IF(AND(ISNUMBER(G),H>=report_date),I-(H-G),I)
        ka_dt = pd.to_datetime(ka) if (not isinstance(ka, str) and not pd.isna(ka)) else None
        h_dt  = pd.to_datetime(h_start) if (not isinstance(h_start, str) and not pd.isna(h_start)) else None
        kc_dt = pd.to_datetime(kc) if (not isinstance(kc, str) and not pd.isna(kc)) else None
        if ka_dt is not None and h_dt is not None and kc_dt is not None and h_dt >= report_date:
            l_mrcl = kc_dt - (h_dt - ka_dt)
        else:
            l_mrcl = kc_dt

        # M: Open Delivery? — mirrors Excel IFS(G=reuse/docked→no, G=no forecast?→yes/no forecast?, G<report→no, G>F→late, G<=F→on-time)
        ka_dt2 = pd.to_datetime(ka) if ka and not isinstance(ka, str) else None
        f_rdd_dt = pd.to_datetime(f_rdd) if f_rdd and not isinstance(f_rdd, str) else None

        if ka is None or ka == "reuse":
            m_open = "no"
        elif ka == "docked":
            m_open = "no"
        elif ka == "no forecast?":
            m_open = "yes/no forecast?"
        elif ka_dt2 is not None and ka_dt2 < report_date:
            m_open = "no"
        elif ka_dt2 is not None and f_rdd_dt is not None and ka_dt2 > f_rdd_dt:
            m_open = "yes/late"
        elif ka_dt2 is not None and f_rdd_dt is not None and ka_dt2 <= f_rdd_dt:
            m_open = "yes/on-time"
        else:
            m_open = "yes/late"

        # N: CK needed?
        n_ck = "yes" if "Conv" in evt_type else "no"

        results.append({
            "Entity":                        entity,       # A
            "CEID":                          ceid,         # B
            "Event Type":                    evt_type,     # C
            "WSTie":                         wstie,        # D
            "CND":                           cnd,          # E
            "MT/CK RDD":                     f_rdd,        # F
            "MT/CK SDD":                     ka,           # G
            "MT Set Start or CK Conv Start": mt_or_ck_start,      # H
            "MRCL Finish":                   kc,           # I
            "MRCL if Set matched to SDD":    l_mrcl,       # L
            "Open Delivery?":                m_open,       # M
            "CK needed?":                    n_ck,         # N
            "Need ID":                       need_id,      # O
        })

    result_df = pd.DataFrame(results)
    # restore string sentinel values lost to datetime inference (e.g. "no SIRFIS SDD" → NaT)
    for col in ["MT/CK RDD", "MT/CK SDD", "MT Set Start or CK Conv Start", "MRCL Finish", "MRCL if Set matched to SDD"]:
        original = [r.get(col) for r in results]
        if any(isinstance(v, str) for v in original):
            result_df[col] = np.array(original, dtype=object)
    return result_df


def compute_wstie_summary(entity_df: pd.DataFrame, lrp_df: pd.DataFrame) -> pd.DataFrame:
    """
    Entity별 DGS DataFrame을 받아 WSTie별 집계(U~AC 섹션)를 계산한다.

    Parameters:
        entity_df : compute_dgs() 결과
        lrp_df    : LRP 테이블 (date, wspw 컬럼)
    Returns:
        WSTie별 집계 DataFrame
    """
    # U: WSTie 고유값 정렬
    wstie_vals = sorted(entity_df["WSTie"].dropna().unique())

    results = []
    cumulative_max_mrcl = None    # AA용 (Y 기반)
    cumulative_max_mrcl_am = None  # AM용 (AK 기반)

    for ws in wstie_vals:
        grp = entity_df[entity_df["WSTie"] == ws]

        # W: LRP Capital Need Date (정확히 매칭 or 보간)
        if lrp_df is not None and len(lrp_df) > 0:
            lrp_sorted = lrp_df.sort_values("wspw").reset_index(drop=True)
            lrp_match = lrp_sorted[lrp_sorted["wspw"] == ws]
            if len(lrp_match) > 0:
                w_date = lrp_match.iloc[0]["date"]
            else:
                below = lrp_sorted[lrp_sorted["wspw"] < ws]
                above = lrp_sorted[lrp_sorted["wspw"] > ws]
                if len(below) > 0 and len(above) > 0:
                    x0 = below.iloc[-1]["wspw"]; x1 = above.iloc[0]["wspw"]
                    t0 = below.iloc[-1]["date"].timestamp(); t1 = above.iloc[0]["date"].timestamp()
                    t_interp = t0 + (t1 - t0) * (ws - x0) / (x1 - x0)
                    w_date = pd.Timestamp.fromtimestamp(round(t_interp))
                elif len(below) > 0:
                    w_date = below.iloc[-1]["date"]
                else:
                    w_date = None
        else:
            w_date = None

        # Y: late 장비의 L 최대값
        late = grp[grp["Open Delivery?"] == "yes/late"]
        l_vals = pd.to_datetime(late["MRCL if Set matched to SDD"], errors="coerce").dropna()
        if len(l_vals) > 0:
            y_date = l_vals.max()
        else:
            y_date = w_date

        # Z: Gap (Y vs W)
        if y_date is None or w_date is None:
            z_gap = "no gap"
        elif pd.Timestamp(y_date) <= pd.Timestamp(w_date):
            z_gap = "no gap"
        else:
            delta_months = (pd.Timestamp(y_date).year - pd.Timestamp(w_date).year) * 12 + \
                           (pd.Timestamp(y_date).month - pd.Timestamp(w_date).month)
            if delta_months == 0:
                z_gap = "<1 month"
            else:
                z_gap = f"{delta_months} month{'s' if delta_months > 1 else ''}"

        # AA: Cumulative latest MRCL
        if y_date is not None:
            if cumulative_max_mrcl is None or pd.Timestamp(y_date) > pd.Timestamp(cumulative_max_mrcl):
                cumulative_max_mrcl = y_date
        aa_date = cumulative_max_mrcl

        # AB: Open Deliveries (On Time/Total)
        open_rows = grp[grp["Open Delivery?"].str.startswith("yes", na=False)]
        total_open = len(open_rows)
        on_time = len(grp[grp["Open Delivery?"] == "yes/on-time"])
        if total_open == 0:
            ab_str = "none"
        else:
            pct = round(on_time / total_open * 100)
            ab_str = f"{pct}% ({on_time} / {total_open})"

        # AC: Late Tool Deliveries (RDD-SDD 차이 큰 순)
        if ab_str == "none" or ab_str.startswith("100%"):
            ac_str = "none"
        else:
            late_open = grp[grp["Open Delivery?"] == "yes/late"].copy()
            late_open["rdd_dt"] = pd.to_datetime(late_open["MT/CK RDD"], errors="coerce")
            late_open["sdd_dt"] = pd.to_datetime(late_open["MT/CK SDD"], errors="coerce")
            late_open["gap_days"] = (late_open["sdd_dt"] - late_open["rdd_dt"]).dt.days
            late_open = late_open.sort_values("gap_days", ascending=False)
            names = []
            for _, r in late_open.iterrows():
                name = str(r["Entity"])[:6]
                suffix = " (CK)" if r["CK needed?"] == "yes" else ""
                names.append(name + suffix)
            ac_str = ", ".join(names) if names else "none"

        # AK: 전체 장비(late+on-time) L 최대값
        all_l = pd.to_datetime(grp["MRCL if Set matched to SDD"], errors="coerce").dropna()
        ak_date = all_l.max() if len(all_l) > 0 else w_date

        # AL: Gap (AK vs W)
        if ak_date is None or w_date is None:
            al_gap = "no gap"
        elif pd.Timestamp(ak_date) <= pd.Timestamp(w_date):
            al_gap = "no gap"
        else:
            delta_months = (pd.Timestamp(ak_date).year - pd.Timestamp(w_date).year) * 12 + \
                           (pd.Timestamp(ak_date).month - pd.Timestamp(w_date).month)
            al_gap = "<1 month" if delta_months == 0 else f"{delta_months} month{'s' if delta_months > 1 else ''}"

        # AN: IQ Status (L≤W 비율)
        total_all = len(grp)
        on_time_lw = len(grp[pd.to_datetime(grp["MRCL if Set matched to SDD"], errors="coerce") <= pd.Timestamp(w_date)]) if w_date else 0
        if total_all == 0:
            an_str = "none"
        else:
            pct2 = round(on_time_lw / total_all * 100)
            an_str = f"{pct2}% ({on_time_lw} / {total_all})"

        # AO: Late IQ (L>W 장비, MRCL-W 차이 큰 순)
        if an_str.startswith("100%") or an_str == "none":
            ao_str = "none"
        else:
            late_iq = grp.copy()
            late_iq["l_dt"] = pd.to_datetime(late_iq["MRCL if Set matched to SDD"], errors="coerce")
            late_iq = late_iq[late_iq["l_dt"] > pd.Timestamp(w_date)] if w_date else late_iq
            late_iq = late_iq.sort_values("l_dt", ascending=False)
            ao_str = ", ".join(str(r["Entity"])[:6] for _, r in late_iq.iterrows()) if len(late_iq) > 0 else "none"

        # AM: Cumulative latest MRCL (AK 기반 — AA와 별도 변수)
        if ak_date is not None:
            if cumulative_max_mrcl_am is None or pd.Timestamp(ak_date) > pd.Timestamp(cumulative_max_mrcl_am):
                cumulative_max_mrcl_am = ak_date
        am_date = cumulative_max_mrcl_am

        results.append({
            "WSTie":                    ws,
            "WSPW":                     f"{ws/1000:.3g}K",
            "Capital Need (LRP)":       w_date,
            "Last MRCL w/delivery limiters only": y_date,
            "Gap":                      z_gap,
            "Cumulative latest MRCL":   aa_date,
            "Open Deliveries (On Time/Total)": ab_str,
            "Late Tool Deliveries":     ac_str,
            "Last MRCL w/delivery+all IQ": ak_date,
            "Gap (all IQ)":             al_gap,
            "Cumulative latest MRCL (all IQ)": am_date,
            "IQ Status (On Time/Total)": an_str,
            "Late IQ":                  ao_str,
        })

    return pd.DataFrame(results)


def compute_ramp_profile(wstie_summary: pd.DataFrame, lrp_df: pd.DataFrame) -> pd.DataFrame:
    """AW~BD: 월별 Ramp 프로필. AY/AZ=late only, BC/BD=all IQ."""
    wstie_vals = sorted(wstie_summary["WSTie"].dropna().unique())
    results = []

    for _, lrp_row in lrp_df.iterrows():
        aw_date = lrp_row["date"]
        ax_wspw = lrp_row["wspw"]

        # AY: AA <= AW인 WSTie 최대 (late only)
        eligible = wstie_summary[
            pd.to_datetime(wstie_summary["Cumulative latest MRCL"], errors="coerce") <= aw_date
        ]
        ay = int(eligible["WSTie"].max()) if len(eligible) > 0 else min(wstie_vals)
        az = min(ax_wspw, ay)

        # BC: AM <= AW인 WSTie 최대 (all IQ)
        eligible2 = wstie_summary[
            pd.to_datetime(wstie_summary["Cumulative latest MRCL (all IQ)"], errors="coerce") <= aw_date
        ]
        bc = int(eligible2["WSTie"].max()) if len(eligible2) > 0 else min(wstie_vals)
        bd = min(ax_wspw, bc)

        results.append({
            "Date (AW)": aw_date, "LRP WSTie (AX)": ax_wspw,
            "Unlocked (AY)": ay, "Ramp late only (AZ)": az,
            "Unlocked all IQ (BC)": bc, "Ramp all IQ (BD)": bd,
        })

    return pd.DataFrame(results)

