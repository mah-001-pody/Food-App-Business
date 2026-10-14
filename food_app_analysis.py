# -*- coding: utf-8 -*-
"""Food App Business - mã tái lập toàn bộ kết quả mô hình trong báo cáo.

Chạy:  python food_app_analysis.py  (cần pandas, numpy, scipy, statsmodels, scikit-learn, openpyxl)
Đầu vào: 05_FoodAppBusiness.xlsx (sheet đầu tiên). Đầu ra: in kết quả ra màn hình.

Thứ tự: đọc Excel -> làm sạch (R1-R8) -> tạo biến -> OLS -> logistic -> chia train/test
-> chọn ngưỡng bằng CV trên train -> đánh giá một lần trên test -> CV và lift OOF trên toàn mẫu
-> cây quyết định -> độ nhạy (giữ 4 dòng nghi vấn; giữ 204 dòng trùng).
Mọi phép ngẫu nhiên dùng random_state = 42.
"""
import warnings
import numpy as np
import pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from scipy import stats
from sklearn.metrics import roc_auc_score, f1_score, precision_score, recall_score, confusion_matrix, brier_score_loss
from sklearn.model_selection import StratifiedKFold, train_test_split
from sklearn.tree import DecisionTreeClassifier

warnings.filterwarnings("ignore")
SEED = 42
RAW = "05_FoodAppBusiness.xlsx"
COLS = ["Income", "Tenure_Days", "Age", "Graduate", "Married", "Single", "Kids", "Teens", "Recency",
        "Spend_Wines", "Spend_Fruits", "Spend_Meat", "Spend_Fish", "Spend_Sweets", "Deals_Purchases",
        "Web_Purchases", "Catalog_Purchases", "Store_Purchases", "Web_Visits_Month",
        "Cmp1", "Cmp2", "Cmp3", "Cmp4", "Cmp5", "Total_Campaigns", "Complain"]


# ---------------------------------------------------------------- 1. Đọc và làm sạch
def load_raw(path=RAW):
    r = pd.read_excel(path, sheet_name=0)
    r.columns = COLS
    r["Raw_Row"] = r.index + 2                       # số dòng Excel (dòng 1 là tiêu đề)
    r["Total_Spend"] = r[["Spend_Wines", "Spend_Fruits", "Spend_Meat", "Spend_Fish", "Spend_Sweets"]].sum(axis=1)
    r["Total_Purchases"] = r[["Web_Purchases", "Catalog_Purchases", "Store_Purchases"]].sum(axis=1)
    return r


def clean(r, keep_duplicates=False, keep_suspect=False):
    """R2-R6 theo đúng thứ tự trong Bảng 3.1; mỗi bước áp dụng trên dữ liệu còn lại."""
    d = r.copy()
    if not keep_duplicates:                                         # R2: trùng toàn bộ 26 cột, giữ bản đầu
        d = d[~d[COLS].duplicated(keep="first")]
    d = d[~((d.Total_Spend > 0) & (d.Total_Purchases == 0))]        # R3
    d = d[~(d.Deals_Purchases > d.Total_Purchases)]                 # R4
    if not keep_suspect:
        d = d[d.Total_Spend / d.Income <= 0.5]                      # R5 (dòng 22)
        d = d[~((d.Web_Purchases > 20) & (d.Web_Visits_Month <= 1))]  # R6 (dòng 1782, 1874, 1951)
    return d.drop(columns="Single")                                 # R8


# ---------------------------------------------------------------- 2. Tạo biến
def quintile_score(s, order):
    """Ngũ phân vị; giá trị bằng nhau được xếp theo khóa ổn định Raw_Row (rank method='first')."""
    r = s.loc[order].rank(method="first")
    return pd.qcut(r, 5, labels=[1, 2, 3, 4, 5]).astype(int).reindex(s.index)


def features(d):
    d = d.copy()
    # Làm tròn như bảng Customers trong mon.pbix (năm 2 chữ số, AOV 2 chữ số, tỷ lệ 4 chữ số)
    d["Tenure_Years"] = (d.Tenure_Days / 365).round(2)
    d["AOV"] = (d.Total_Spend / d.Total_Purchases).round(2)
    d["Deal_Ratio"] = (d.Deals_Purchases / d.Total_Purchases).round(4)
    # Kênh có nhiều đơn nhất; khi bằng nhau ưu tiên Web -> Catalog -> Store (thứ tự cột trong idxmax)
    ch = d[["Web_Purchases", "Catalog_Purchases", "Store_Purchases"]]
    d["Preferred_Channel"] = ch.idxmax(axis=1).str.replace("_Purchases", "")
    d["Campaign_Responder"] = (d.Total_Campaigns > 0).astype(int)
    order = d.sort_values("Raw_Row").index
    d["R_Score"] = 6 - quintile_score(d.Recency, order)
    d["F_Score"] = quintile_score(d.Total_Purchases, order)
    d["M_Score"] = quintile_score(d.Total_Spend, order)
    d["FM_Score"] = (d.F_Score + d.M_Score) / 2
    R, FM = d.R_Score, d.FM_Score
    d["RFM_Segment"] = np.select(
        [(R >= 4) & (FM >= 4), (R >= 3) & (FM >= 3), (R >= 4), (R == 3), (R <= 2) & (FM >= 3)],
        ["Champions", "Loyal Customers", "Promising", "Need Attention", "At Risk"], "Hibernating")
    d["Income10K"] = d.Income / 10000
    d["Recency10"] = d.Recency / 10
    d["Web_Pref"] = (d.Preferred_Channel == "Web").astype(float)
    d["Catalog_Pref"] = (d.Preferred_Channel == "Catalog").astype(float)
    return d


OLS_F = ("np.log(Total_Spend) ~ np.log(Income) + Tenure_Years + Married + Teens + Kids + Deal_Ratio"
         " + Web_Visits_Month + Recency + Age + Graduate")
LOGIT_X = ["Income10K", "Kids", "Teens", "Age", "Graduate", "Married", "Tenure_Years", "Recency10",
           "Deal_Ratio", "Web_Visits_Month", "Catalog_Pref", "Web_Pref"]
TREE_X = ["Income", "Kids", "Teens", "Age", "Recency", "Deal_Ratio", "Web_Visits_Month", "Total_Spend",
          "Catalog_Purchases", "Tenure_Days"]


def logit_fit(X, y):
    return sm.Logit(y.astype(float), sm.add_constant(X.astype(float))).fit(disp=0)


def logit_pred(m, X):
    return m.predict(sm.add_constant(X.astype(float), has_constant="add"))


def oof_predict(X, y, folds):
    p = pd.Series(index=X.index, dtype=float)
    for a, b in folds.split(X, y):
        p.iloc[b] = logit_pred(logit_fit(X.iloc[a], y.iloc[a]), X.iloc[b]).values
    return p


def main():
    raw = load_raw()
    c = features(clean(raw))
    print(f"Sau làm sạch: {len(c)} khách")
    print(c.RFM_Segment.value_counts().to_string(), "\n", c.Preferred_Channel.value_counts().to_string())

    # ---- OLS ln(chi tiêu), HC3
    ols = smf.ols(OLS_F, c).fit(cov_type="HC3")
    print("\n== OLS ln(Total_Spend) ==\n", ols.summary2().tables[1].round(4), "\nR2", round(ols.rsquared, 4))
    aov = smf.ols("np.log(AOV) ~ np.log(Income) + Kids + Teens + Deal_Ratio + Web_Pref + Catalog_Pref", c).fit(cov_type="HC3")
    print("\n== OLS ln(AOV) ==\n", aov.params.round(3).to_string(), "\nR2", round(aov.rsquared, 3))

    # ---- Logistic trên toàn mẫu
    y, X = c.Campaign_Responder, c[LOGIT_X]
    lg = logit_fit(X, y)
    print("\n== Logistic (toàn mẫu) ==\n", pd.DataFrame({"b": lg.params, "OR": np.exp(lg.params), "p": lg.pvalues}).round(4))
    print("pseudo R2", round(lg.prsquared, 3), "AIC", round(lg.aic, 1))

    # ---- Chia train/test 70/30 theo tầng
    Xtr, Xte, ytr, yte = train_test_split(X, y, test_size=0.3, stratify=y, random_state=SEED)
    skf = StratifiedKFold(5, shuffle=True, random_state=SEED)
    # (a) Chọn ngưỡng CHỈ trên train: CV 5 lần trong train, dự báo OOF, xem F1 theo ngưỡng
    oof_tr = oof_predict(Xtr, ytr, skf)
    print("\n== Chọn ngưỡng trên train (OOF, n =", len(ytr), ") AUC", round(roc_auc_score(ytr, oof_tr), 3))
    for t in [0.2, 0.25, 0.3, 0.35, 0.4, 0.5]:
        pr = oof_tr >= t
        print(f"  ngưỡng {t:.2f}: F1 {f1_score(ytr, pr):.3f} recall {recall_score(ytr, pr):.3f} "
              f"precision {precision_score(ytr, pr):.3f} chọn {pr.mean():.3f}")
    # (b) Huấn luyện trên toàn bộ train, đánh giá một lần trên test
    m = logit_fit(Xtr, ytr)
    p = logit_pred(m, Xte)
    print("\n== Tập test (n =", len(yte), ", dương tính =", int(yte.sum()), ") AUC", round(roc_auc_score(yte, p), 3),
          "Brier", round(brier_score_loss(yte, p), 3), "(mô hình hằng số:", round(brier_score_loss(yte, np.full(len(yte), ytr.mean())), 3), ")")
    for t in [0.5, 0.3]:
        pr = (p >= t).astype(int)
        tn, fp, fn, tp = confusion_matrix(yte, pr).ravel()
        print(f"  ngưỡng {t}: TP {tp} FP {fp} FN {fn} TN {tn} | F1 {f1_score(yte, pr):.3f} "
              f"recall {recall_score(yte, pr):.3f} precision {precision_score(yte, pr):.3f} chọn {pr.mean():.3f}")
    cal = pd.DataFrame({"p": p, "y": yte}).assign(bin=lambda t: pd.qcut(t.p, 5, labels=False))
    print("  Hiệu chỉnh theo 5 nhóm điểm (dự báo TB / thực tế):\n", cal.groupby("bin")[["p", "y"]].mean().round(3).to_string())

    # (c) CV 5 lần trên TOÀN MẪU: AUC từng fold và lift dựa trên dự báo OOF
    aucs = []
    oof = pd.Series(index=X.index, dtype=float)
    for a, b in skf.split(X, y):
        pp = logit_pred(logit_fit(X.iloc[a], y.iloc[a]), X.iloc[b])
        oof.iloc[b] = pp.values
        aucs.append(roc_auc_score(y.iloc[b], pp))
    print("\n== CV toàn mẫu: AUC TB", round(np.mean(aucs), 3), "độ lệch chuẩn giữa fold", round(np.std(aucs), 3))
    s = oof.sort_values(ascending=False)
    for q in [0.2, 0.3]:
        k = round(q * len(y)); cap = y[s.index[:k]].sum() / y.sum()
        print(f"  top {int(q*100)}%: chứa {cap:.3f} số khách đã chấp nhận, lift {cap/q:.2f}")

    # ---- Cây quyết định (cùng chia train/test)
    tree = DecisionTreeClassifier(max_depth=3, min_samples_leaf=30, random_state=SEED)
    tree.fit(c.loc[Xtr.index, TREE_X], ytr)
    print("\n== Cây quyết định: AUC test", round(roc_auc_score(yte, tree.predict_proba(c.loc[Xte.index, TREE_X])[:, 1]), 3))

    # ---- Độ nhạy của các bước làm sạch
    for name, kw in [("giữ 4 dòng nghi vấn R5-R6", dict(keep_suspect=True)), ("giữ 204 dòng trùng", dict(keep_duplicates=True))]:
        d = features(clean(raw, **kw))
        o = smf.ols(OLS_F, d).fit(cov_type="HC3")
        l = logit_fit(d[LOGIT_X], d.Campaign_Responder)
        print(f"\n== Độ nhạy - {name}: n = {len(d)}; OLS R2 {o.rsquared:.3f}, ln(TN) {o.params['np.log(Income)']:.3f}, "
              f"con nhỏ {o.params['Kids']:.3f}; logistic OR thu nhập {np.exp(l.params['Income10K']):.2f}, "
              f"OR con nhỏ {np.exp(l.params['Kids']):.2f} (p {l.pvalues['Kids']:.3f})")


if __name__ == "__main__":
    main()
