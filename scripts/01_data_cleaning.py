"""
VJP205 - Food App Business
Buoc 1: Tien xu ly & lam sach du lieu (Data Pre-processing & Cleaning)

Chay:  python scripts/01_data_cleaning.py
Input: 05_FoodAppBusiness.xlsx (sheet "Data")
Output (thu muc data/processed/):
  - FoodApp_Clean.csv          -> nap vao Power BI / mo hinh
  - FoodApp_Clean.xlsx         -> Data_Clean, Data_Dictionary, Cleaning_Log,
                                  Removed_Rows, Quality_Before_After
"""
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "05_FoodAppBusiness.xlsx"
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 0. Doc du lieu & chuan hoa ten cot
# ---------------------------------------------------------------------------
raw = pd.read_excel(RAW, sheet_name="Data")
raw.columns = [c.replace("\n", " ").strip() for c in raw.columns]

RENAME = {
    "Monthly Income": "Income",
    "Active Since ... Days": "Tenure_Days",
    "Age": "Age",
    "Graduate": "Graduate",
    "Married": "Married",
    "Single": "Single",
    "No Of Children": "Kids",
    "No Of Teenager": "Teens",
    "No Of Days Since Last Purchase": "Recency",
    "Amount Spend OnWines": "Spend_Wines",
    "Amount Spent On Fruits": "Spend_Fruits",
    "Amount Spent On Meat": "Spend_Meat",
    "Amount Spent On Fish": "Spend_Fish",
    "Amount Spent On Sweet": "Spend_Sweets",
    "No Of Deals With Discount": "Deals_Purchases",
    "No Of Web Purchase": "Web_Purchases",
    "No Of Catalog Purchase": "Catalog_Purchases",
    "No Of Store Purchase": "Store_Purchases",
    "No Of Web Visits/ Month": "Web_Visits_Month",
    "Purchased In 1st Campaign": "Cmp1",
    "Purchased In 2nd Campaign": "Cmp2",
    "Purchased In 3rd Campaign": "Cmp3",
    "Purchased In 4th Campaign": "Cmp4",
    "Purchased In 5th Campaign": "Cmp5",
    "Total No Of Campaign Accepted": "Total_Campaigns",
    "CustomerComplain": "Complain",
}
df = raw.rename(columns=RENAME).copy()
df.insert(0, "Raw_Row", np.arange(2, len(df) + 2))  # so dong trong Excel goc

SPEND = ["Spend_Wines", "Spend_Fruits", "Spend_Meat", "Spend_Fish", "Spend_Sweets"]
CHANNEL = ["Web_Purchases", "Catalog_Purchases", "Store_Purchases"]
CMP = ["Cmp1", "Cmp2", "Cmp3", "Cmp4", "Cmp5"]
BINARY = ["Graduate", "Married", "Single", "Complain"] + CMP


def profile(d: pd.DataFrame, label: str) -> pd.DataFrame:
    num = d.drop(columns=["Raw_Row"], errors="ignore").select_dtypes("number")
    p = num.describe().T[["count", "mean", "std", "min", "50%", "max"]]
    p["missing"] = num.isna().sum()
    p.columns = [f"{label}_{c}" for c in p.columns]
    return p


log = []
removed = []


def drop_rows(d: pd.DataFrame, mask: pd.Series, rule: str, reason: str) -> pd.DataFrame:
    n = int(mask.sum())
    if n:
        r = d[mask].copy()
        r.insert(1, "Rule", rule)
        r.insert(2, "Reason", reason)
        removed.append(r)
    log.append({"Step": rule, "Action": "Loai bo", "Rows_Affected": n,
                "Reason": reason, "Rows_After": len(d) - n})
    return d[~mask].copy()


def note(rule: str, action: str, n: int, reason: str, d: pd.DataFrame):
    log.append({"Step": rule, "Action": action, "Rows_Affected": n,
                "Reason": reason, "Rows_After": len(d)})


before = profile(df, "Before")
note("R0", "Kiem tra", len(df), "Du lieu goc: 2.205 dong x 26 cot, tat ca kieu so nguyen", df)

# ---------------------------------------------------------------------------
# 1. Kiem tra cau truc: missing, kieu du lieu, mien gia tri
# ---------------------------------------------------------------------------
n_missing = int(df.isna().sum().sum())
note("R1", "Kiem tra", n_missing, "Gia tri thieu (missing) - khong co, khong can impute", df)

bad_binary = int((~df[BINARY].isin([0, 1])).sum().sum())
note("R1", "Kiem tra", bad_binary, "Bien nhi phan ngoai {0,1}", df)

bad_marital = int(((df["Married"] + df["Single"]) != 1).sum())
note("R1", "Kiem tra", bad_marital, "Married + Single != 1 (mau thuan tinh trang hon nhan)", df)

bad_total = int((df[CMP].sum(axis=1) != df["Total_Campaigns"]).sum())
note("R1", "Kiem tra", bad_total, "Total_Campaigns != tong Cmp1..Cmp5", df)

neg = int((df.drop(columns="Raw_Row") < 0).sum().sum())
note("R1", "Kiem tra", neg, "Gia tri am", df)

# ---------------------------------------------------------------------------
# 2. Ban ghi trung lap (khong co CustomerID -> trung toan bo 26 cot)
# ---------------------------------------------------------------------------
dup_mask = df.drop(columns="Raw_Row").duplicated(keep="first")
df = drop_rows(df, dup_mask, "R2",
               "Trung lap hoan toan 26/26 cot (ca thu nhap, so ngay hoat dong, chi tieu...) "
               "-> xac suat trung tu nhien gan bang 0; giu ban ghi dau tien")

# ---------------------------------------------------------------------------
# 3. Ban ghi bat hop ly ve logic kinh doanh
# ---------------------------------------------------------------------------
df["_TS"] = df[SPEND].sum(axis=1)
df["_TP"] = df[CHANNEL].sum(axis=1)

m = df["_TP"] == 0
df = drop_rows(df, m, "R3",
               "Tong so lan mua = 0 nhung van co chi tieu > 0 -> mau thuan")

m = df["Deals_Purchases"] > df["_TP"]
df = drop_rows(df, m, "R4",
               "So lan mua co giam gia lon hon tong so lan mua -> mau thuan")

m = df["_TS"] > 0.5 * df["Income"]
df = drop_rows(df, m, "R5",
               "Tong chi tieu > 50% thu nhap (ban ghi lien ke cao nhat chi ~3%) -> "
               "nghi sai so nhap lieu thu nhap")

m = (df["Web_Purchases"] > 20) & (df["Web_Visits_Month"] <= 1)
df = drop_rows(df, m, "R6",
               "Tren 20 lan mua qua web nhung <=1 luot truy cap/thang va chi tieu rat thap "
               "-> hanh vi bat thuong (bot/nhap lieu sai)")

df = df.drop(columns=["_TS", "_TP"])

# ---------------------------------------------------------------------------
# 4. Ngoai lai (outlier) - kiem tra nhung GIU LAI
# ---------------------------------------------------------------------------
outlier_rows = []
for c in ["Income", "Age", "Recency"] + SPEND + CHANNEL + ["Deals_Purchases", "Web_Visits_Month"]:
    q1, q3 = df[c].quantile([0.25, 0.75])
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    n = int(((df[c] < lo) | (df[c] > hi)).sum())
    outlier_rows.append({"Variable": c, "Q1": q1, "Q3": q3, "Lower": lo, "Upper": hi,
                         "IQR_Outliers": n, "Pct": round(n / len(df) * 100, 2),
                         "Skewness": round(df[c].skew(), 2)})
outliers = pd.DataFrame(outlier_rows)
spend_out = df[SPEND].gt(df[SPEND].quantile(0.75) + 1.5 * (df[SPEND].quantile(0.75) - df[SPEND].quantile(0.25)))
note("R7", "Giu lai", int(spend_out.any(axis=1).sum()),
     "So khach co it nhat 1 nhom chi tieu vuot nguong IQR. Ngoai lai IQR o chi tieu/kenh mua la khach hang chi tieu cao THAT (phan phoi lech phai) "
     "-> khong xoa, khong cap; dung log-transform khi chay mo hinh", df)

# ---------------------------------------------------------------------------
# 5. Bien du thua
# ---------------------------------------------------------------------------
df = df.drop(columns=["Single"])
note("R8", "Loai cot", 0, "Bo cot Single (= 1 - Married) tranh da cong tuyen / dummy trap", df)

# ---------------------------------------------------------------------------
# 6. Feature engineering
# ---------------------------------------------------------------------------
df.insert(0, "CustomerID", [f"C{i:04d}" for i in range(1, len(df) + 1)])

df["Tenure_Years"] = (df["Tenure_Days"] / 365).round(2)
df["Age_Group"] = pd.cut(df["Age"], [0, 34, 44, 54, 64, 200],
                         labels=["24-34", "35-44", "45-54", "55-64", "65+"])
df["Income_Group"] = pd.cut(df["Income"], [0, 30000, 50000, 70000, 90000, np.inf],
                            labels=["<30K", "30-50K", "50-70K", "70-90K", "90K+"], right=False)
df["Marital_Status"] = np.where(df["Married"] == 1, "Married/Partner", "Single")
df["Education"] = np.where(df["Graduate"] == 1, "Graduate", "Non-graduate")
df["Total_Children"] = df["Kids"] + df["Teens"]
df["Has_Children"] = (df["Total_Children"] > 0).astype(int)
df["Family_Size"] = 1 + df["Married"] + df["Total_Children"]

df["Total_Spend"] = df[SPEND].sum(axis=1)
for c in SPEND:
    df[f"Share_{c.split('_')[1]}"] = (df[c] / df["Total_Spend"]).round(4)

df["Total_Purchases"] = df[CHANNEL].sum(axis=1)
df["AOV"] = (df["Total_Spend"] / df["Total_Purchases"]).round(2)
df["Deal_Ratio"] = (df["Deals_Purchases"] / df["Total_Purchases"]).round(4)
for c in CHANNEL:
    df[f"Share_{c.split('_')[0]}"] = (df[c] / df["Total_Purchases"]).round(4)
df["Preferred_Channel"] = (df[CHANNEL].idxmax(axis=1)
                           .str.replace("_Purchases", "", regex=False))
df["Spend_to_Income_Pct"] = (df["Total_Spend"] / df["Income"] * 100).round(3)

df["Campaign_Responder"] = (df["Total_Campaigns"] > 0).astype(int)
df["Recency_Group"] = pd.cut(df["Recency"], [-1, 30, 60, 100],
                             labels=["0-30 ngay", "31-60 ngay", "61-99 ngay"])

# RFM (diem 1-5 theo ngu phan vi; R dao chieu: mua cang gan diem cang cao)
df["R_Score"] = pd.qcut(df["Recency"].rank(method="first"), 5, labels=[5, 4, 3, 2, 1]).astype(int)
df["F_Score"] = pd.qcut(df["Total_Purchases"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
df["M_Score"] = pd.qcut(df["Total_Spend"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
df["FM_Score"] = ((df["F_Score"] + df["M_Score"]) / 2).round(1)
df["RFM_Score"] = df["R_Score"] + df["F_Score"] + df["M_Score"]


def rfm_segment(r: int, fm: float) -> str:
    if r >= 4 and fm >= 4:
        return "Champions"
    if r >= 3 and fm >= 3:
        return "Loyal Customers"
    if r >= 4 and fm < 3:
        return "New / Promising"
    if r <= 2 and fm >= 3:
        return "At Risk"
    if r == 3 and fm < 3:
        return "Need Attention"
    return "Hibernating"


df["RFM_Segment"] = [rfm_segment(r, fm) for r, fm in zip(df["R_Score"], df["FM_Score"])]
df["Value_Tier"] = pd.qcut(df["Total_Spend"], [0, .5, .8, 1.0], labels=["Low", "Mid", "High"])

# Bien bien doi log cho mo hinh (giam lech phai)
df["Log_Income"] = np.log1p(df["Income"]).round(4)
df["Log_Total_Spend"] = np.log1p(df["Total_Spend"]).round(4)

note("R9", "Tao bien", len(df), "Feature engineering: 33 bien moi (nhan khau hoc, chi tieu, kenh, RFM, log)", df)

# ---------------------------------------------------------------------------
# 7. Kiem tra sau lam sach
# ---------------------------------------------------------------------------
assert df.isna().sum().sum() == 0
assert df.drop(columns=["CustomerID", "Raw_Row"]).duplicated().sum() == 0
assert (df["Total_Purchases"] > 0).all()
assert (df["Deals_Purchases"] <= df["Total_Purchases"]).all()
note("R10", "Kiem tra", len(df), "Du lieu sach: 0 missing, 0 trung lap, 0 mau thuan logic", df)

after = profile(df, "After")
quality = before.join(after, how="outer")

# ---------------------------------------------------------------------------
# 8. Data dictionary
# ---------------------------------------------------------------------------
DICT = [
    ("CustomerID", "Ma khach hang (tao moi)", "ID", "Tao moi"),
    ("Raw_Row", "So dong trong file Excel goc (truy vet)", "ID", "Tao moi"),
    ("Income", "Thu nhap (cot goc 'Monthly Income')", "So", "Goc"),
    ("Tenure_Days", "So ngay tu khi dang ky", "So", "Goc"),
    ("Age", "Tuoi", "So", "Goc"),
    ("Graduate", "1 = tot nghiep dai hoc", "Nhi phan", "Goc"),
    ("Married", "1 = da ket hon/song chung", "Nhi phan", "Goc"),
    ("Kids", "So tre nho trong gia dinh", "So", "Goc"),
    ("Teens", "So thanh thieu nien trong gia dinh", "So", "Goc"),
    ("Recency", "So ngay tu lan mua gan nhat", "So", "Goc"),
    ("Spend_Wines", "Chi tieu ruou vang", "So", "Goc"),
    ("Spend_Fruits", "Chi tieu trai cay", "So", "Goc"),
    ("Spend_Meat", "Chi tieu thit", "So", "Goc"),
    ("Spend_Fish", "Chi tieu ca", "So", "Goc"),
    ("Spend_Sweets", "Chi tieu do ngot", "So", "Goc"),
    ("Deals_Purchases", "So lan mua co giam gia", "So", "Goc"),
    ("Web_Purchases", "So lan mua qua website", "So", "Goc"),
    ("Catalog_Purchases", "So lan mua qua catalog", "So", "Goc"),
    ("Store_Purchases", "So lan mua tai cua hang", "So", "Goc"),
    ("Web_Visits_Month", "So luot truy cap web/thang", "So", "Goc"),
    ("Cmp1..Cmp5", "1 = chap nhan uu dai chien dich 1..5", "Nhi phan", "Goc"),
    ("Total_Campaigns", "Tong so chien dich da chap nhan (0-5)", "So", "Goc"),
    ("Complain", "1 = da khieu nai", "Nhi phan", "Goc"),
    ("Tenure_Years", "Tenure_Days / 365", "So", "Tao moi"),
    ("Age_Group", "Nhom tuoi: 24-34/35-44/45-54/55-64/65+", "Phan loai", "Tao moi"),
    ("Income_Group", "Nhom thu nhap: <30K/30-50K/50-70K/70-90K/90K+", "Phan loai", "Tao moi"),
    ("Marital_Status", "Married/Partner hoac Single", "Phan loai", "Tao moi"),
    ("Education", "Graduate hoac Non-graduate", "Phan loai", "Tao moi"),
    ("Total_Children", "Kids + Teens", "So", "Tao moi"),
    ("Has_Children", "1 = co con", "Nhi phan", "Tao moi"),
    ("Family_Size", "1 + Married + Total_Children", "So", "Tao moi"),
    ("Total_Spend", "Tong chi tieu 5 nhom hang", "So", "Tao moi"),
    ("Share_Wines..Share_Sweets", "Ty trong chi tieu tung nhom hang", "Ty le", "Tao moi"),
    ("Total_Purchases", "Web + Catalog + Store", "So", "Tao moi"),
    ("AOV", "Gia tri trung binh/don = Total_Spend / Total_Purchases", "So", "Tao moi"),
    ("Deal_Ratio", "Deals_Purchases / Total_Purchases (do nhay cam gia)", "Ty le", "Tao moi"),
    ("Share_Web/Catalog/Store", "Ty trong so lan mua theo kenh", "Ty le", "Tao moi"),
    ("Preferred_Channel", "Kenh co so lan mua nhieu nhat", "Phan loai", "Tao moi"),
    ("Spend_to_Income_Pct", "Total_Spend / Income x 100", "Ty le", "Tao moi"),
    ("Campaign_Responder", "1 = chap nhan it nhat 1 chien dich", "Nhi phan", "Tao moi"),
    ("Recency_Group", "0-30 / 31-60 / 61-99 ngay", "Phan loai", "Tao moi"),
    ("R_Score, F_Score, M_Score", "Diem RFM 1-5 theo ngu phan vi", "So", "Tao moi"),
    ("FM_Score", "(F + M) / 2", "So", "Tao moi"),
    ("RFM_Score", "R + F + M (3-15)", "So", "Tao moi"),
    ("RFM_Segment", "Champions / Loyal / New-Promising / Need Attention / At Risk / Hibernating", "Phan loai", "Tao moi"),
    ("Value_Tier", "Low (50% duoi) / Mid (30%) / High (top 20%) theo Total_Spend", "Phan loai", "Tao moi"),
    ("Log_Income, Log_Total_Spend", "log(1+x) dung cho mo hinh", "So", "Tao moi"),
]
dictionary = pd.DataFrame(DICT, columns=["Column", "Description_VI", "Type", "Source"])

# ---------------------------------------------------------------------------
# 9. Xuat file
# ---------------------------------------------------------------------------
df.to_csv(OUT / "FoodApp_Clean.csv", index=False, encoding="utf-8-sig")
removed_df = pd.concat(removed, ignore_index=True) if removed else pd.DataFrame()
with pd.ExcelWriter(OUT / "FoodApp_Clean.xlsx", engine="openpyxl") as xw:
    df.to_excel(xw, sheet_name="Data_Clean", index=False)
    dictionary.to_excel(xw, sheet_name="Data_Dictionary", index=False)
    pd.DataFrame(log).to_excel(xw, sheet_name="Cleaning_Log", index=False)
    removed_df.to_excel(xw, sheet_name="Removed_Rows", index=False)
    outliers.to_excel(xw, sheet_name="Outlier_Check", index=False)
    quality.to_excel(xw, sheet_name="Quality_Before_After")

print(pd.DataFrame(log).to_string(index=False))
print(f"\nClean: {df.shape[0]} dong x {df.shape[1]} cot -> {OUT}")
