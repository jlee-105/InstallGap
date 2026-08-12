
"""
This project is for a tool developement to understand the gap 
between Intel plan of ramping versus current state

Made by Jaejin Lee 2026/08/05
"""


# calling modules
from datetime import datetime
from app.core import data_source as data
from app.core import engine
from app.core import dgs
import pandas as pd

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# getting input from user


supplier_name_library = ['Adventest', 'Air Products And Chemicals, Inc.', 'AMAT', 'ASMI',
                        'ASML', 'Brooks', 'Canon', 'Disco', 'E&R',
                        'Ebara', 'ESI', 'exclude', 'Gudeng', 'Hitachi',
                      'HORIBA Instruments Incorporated', 'HPSP','ITW',
                      'ITW EAE, A DIVISION OF ILLINOIS TOO', 
                      'JT TOKE', 'KLA', 'Kokusai', 'Lam', 'Lasertec',
                      'Lintec Of America, Inc.', 'Micro Engineering, Inc.', 'Mitsubishi Chemical Advanced Materi',
                      'Nikon', 'not found', 'Nova', 'Onto', 'Pacific Integrated Handling Inc',
                      'Pro-Fab', 'PRO-FAB, INC.', 'PSK', 'Rigaku', 'SCREEN', 'TEL', 'Veeco', 'Versum Materials US, LLC', 
                      'YES']

# select supplier
selected_supplier = 'Lam'

# selected WSTie 
selected_WETie = 7000


# year
year = 2026

# selected WW_WEEK
selected_WW = 30.5

# year
year = 2026

week_number = int(selected_WW)   # 30
decimal = round(selected_WW % 1, 1)  # 0.1 or 0.5

# Only 2 cases
if decimal == 0.1:
    day_of_week = 1   # Monday
else:                 # 0.5
    day_of_week = 5   # Friday


# Convert to date using ISO week (%V and %G)
selected_date = datetime.strptime(f"{year}-W{week_number:02d}-{day_of_week}", "%G-W%V-%u")

# getting all data (now from excel file -- > later from DB)
df_source = data.load_all()

# later we should directly read from excel file or db
IQ_df = df_source["iq"]

# Supplier ~ WSTie to be used --- this is because original EXCEL should not have these columns, but in the current EXCEL, these columns are included. So we need to remove these columns before calculation.
_cols = IQ_df.columns.tolist()
_start = next((i for i, c in enumerate(_cols) if c == 'Supplier'), None)
_end   = next((i for i, c in enumerate(_cols) if c and 'WSTie to be used' in str(c)), None)
if _start is not None and _end is not None:
    # preserve Excel-precomputed JT Supplier before dropping so engine can use it as fallback
    IQ_df["_excel_jt"] = IQ_df[_cols[_start]]
    IQ_df = IQ_df.drop(columns=_cols[_start:_end + 1])


SIRFIS_df = df_source['sirfis']
OVERRIDE_df = df_source['overrides']
MA_df = df_source['ma']

match_sdd_rdd = 'no'

# calculation of the other colums
IQ_calculated_output= engine.compute_jt_ke(selected_supplier=selected_supplier,
                                           iq_df=IQ_df, 
                                           sirfis_df=SIRFIS_df,
                                           overrides_df=OVERRIDE_df,  
                                           use_override=True, 
                                           reference_date=selected_date,
                                           match_sdd_rdd=match_sdd_rdd)




dgs_output_1st_part = dgs.compute_dgs(selected_supplier=selected_supplier, we_tie= selected_WETie, 
                             iq_extended=IQ_calculated_output, report_date=selected_date)


print("==============================================================")

print("len(dgs_output_1st_part)", len(dgs_output_1st_part))

print("dgs_output_1st_part---------------------", dgs_output_1st_part[dgs_output_1st_part["Entity"]=="CVD263-2"])

print("============================================================")


WSTie = sorted({int(x) for x in dgs_output_1st_part["WSTie"].unique()})
print("WSTie", len(WSTie))
a = input()

WSTie_K = [f"{x/1000:g}K" for x in WSTie] 

print("WSTie_K", WSTie_K)
a =input()



# dgs_output_ist_part = dgs_output_1st_part[dgs_output_1st_part["Supplier"]==selected_supplier]

# print("dgs_output_1st_part---------------------", len(dgs_output_1st_part))
# a = input()


# print("dgs_output---------------------", dgs_output_1st_part.columns)
# a = input()




# done until this ------------------------------------------------------------------

   
# lrp_df = pd.DataFrame([
#     {"date": pd.Timestamp("2026-06-01"), "wspw": 7000},
#     {"date": pd.Timestamp("2026-07-01"), "wspw": 7500},
#     {"date": pd.Timestamp("2026-08-01"), "wspw": 8000},
#     {"date": pd.Timestamp("2026-09-01"), "wspw": 8500},
#     {"date": pd.Timestamp("2026-10-01"), "wspw": 9000},
#     {"date": pd.Timestamp("2026-11-01"), "wspw": 9750},
#     {"date": pd.Timestamp("2026-12-01"), "wspw": 9750},
#     {"date": pd.Timestamp("2027-01-01"), "wspw": 10500},
#     {"date": pd.Timestamp("2027-02-01"), "wspw": 11250},
#     {"date": pd.Timestamp("2027-03-01"), "wspw": 11250},
#     {"date": pd.Timestamp("2027-04-01"), "wspw": 12000},
#     {"date": pd.Timestamp("2027-05-01"), "wspw": 12500},
#     {"date": pd.Timestamp("2027-06-01"), "wspw": 13250},
#     {"date": pd.Timestamp("2027-07-01"), "wspw": 13250},
#     {"date": pd.Timestamp("2027-08-01"), "wspw": 14000},
#     {"date": pd.Timestamp("2027-09-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2027-10-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2027-11-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2027-12-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2028-01-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2028-02-01"), "wspw": 14500},
#     {"date": pd.Timestamp("2028-03-01"), "wspw": 15000},
#     {"date": pd.Timestamp("2028-04-01"), "wspw": 15500},
#     {"date": pd.Timestamp("2028-05-01"), "wspw": 16000},
#     {"date": pd.Timestamp("2028-06-01"), "wspw": 16500},
#     {"date": pd.Timestamp("2028-07-01"), "wspw": 17000},
# ])

# print(datetime(2026, 7, 24))
# print(selected_date==datetime(2026, 7, 24))
# a = input()



# dgs_summary = dgs.compute_wstie_summary(dgs, lrp_df)

# print("=== U: WSTie unique values ===")
# print(summary[["WSTie"]].to_string())
# print()
# print("=== Table 1: U, W ===")
# print(summary[["WSTie", "Capital Need (LRP)"]].to_string())
# print()
# print("=== Table 2: Y, Z, AA, AB, AC ===")
# print(summary[["WSTie", "Last MRCL w/delivery limiters only", "Gap", "Cumulative latest MRCL", "Open Deliveries (On Time/Total)", "Late Tool Deliveries"]].to_string())
# print()
# print("=== Table 3: AK, AL, AM, AN, AO ===")
# print(summary[["WSTie", "Last MRCL w/delivery+all IQ", "Gap (all IQ)", "Cumulative latest MRCL (all IQ)", "IQ Status (On Time/Total)", "Late IQ"]].to_string())
# print()
# print("=== Table 4: AW, AX, AY, AZ, BC, BD ===")
# from dgs import compute_ramp_profile
# ramp = compute_ramp_profile(summary, lrp_df)
# print(ramp.to_string())











