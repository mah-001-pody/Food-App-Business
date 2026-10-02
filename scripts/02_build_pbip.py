"""
VJP205 - Food App Business
Buoc 2: Sinh du an Power BI (.pbip) - Semantic model (TMDL) + Report (PBIR)

Chay:   python scripts/02_build_pbip.py
Output: powerbi/FoodApp_Dashboard.pbip
        powerbi/FoodApp_Dashboard.SemanticModel/   (mo hinh du lieu, Power Query, DAX)
        powerbi/FoodApp_Dashboard.Report/          (7 trang bao cao)

Mo bang Power BI Desktop -> sua tham so DataFile -> Refresh -> Save As .pbix
(xem powerbi/HUONG_DAN.md)
"""
import hashlib
import json
import re
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "powerbi"
NAME = "FoodApp_Dashboard"
SM = OUT / f"{NAME}.SemanticModel"
RP = OUT / f"{NAME}.Report"
THEME_SRC = ROOT / "scripts" / "pbip_assets"

DEFAULT_PATH = r"C:\FoodApp\FoodApp_Clean.xlsx"
MEASURES = "_Measures"


def lt(seed: str) -> str:
    """lineageTag / logicalId co dinh (chay lai script khong doi ID)."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"foodapp/{seed}"))


def hid(seed: str) -> str:
    return hashlib.md5(seed.encode("utf-8")).hexdigest()[:20]


def q(name: str) -> str:
    """Quote ten trong TMDL khi can."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        return name
    return "'" + name.replace("'", "''") + "'"


# =============================================================================
# 1. SEMANTIC MODEL SPEC
# =============================================================================
# (ten cot, kieu TMDL, kieu M, formatString, summarizeBy, sortByColumn, hidden)
I, D, S = ("int64", "Int64.Type"), ("double", "type number"), ("string", "type text")


def c(name, t, fmt=None, summ="none", sort=None, hidden=False):
    return dict(name=name, dt=t[0], mt=t[1], fmt=fmt, summ=summ, sort=sort, hidden=hidden)


CUSTOMER_SOURCE = [
    c("CustomerID", S), c("Raw_Row", I, "0", hidden=True),
    c("Income", I, "#,##0", "average"), c("Tenure_Days", I, "#,##0", "average"),
    c("Tenure_Years", D, "0.0", "average"), c("Age", I, "0", "average"),
    c("Graduate", I, "0"), c("Married", I, "0"), c("Kids", I, "0", "sum"), c("Teens", I, "0", "sum"),
    c("Recency", I, "0", "average"),
    c("Spend_Wines", I, "#,##0", "sum"), c("Spend_Fruits", I, "#,##0", "sum"),
    c("Spend_Meat", I, "#,##0", "sum"), c("Spend_Fish", I, "#,##0", "sum"),
    c("Spend_Sweets", I, "#,##0", "sum"),
    c("Deals_Purchases", I, "#,##0", "sum"), c("Web_Purchases", I, "#,##0", "sum"),
    c("Catalog_Purchases", I, "#,##0", "sum"), c("Store_Purchases", I, "#,##0", "sum"),
    c("Web_Visits_Month", I, "0", "average"),
    c("Cmp1", I, "0", "sum"), c("Cmp2", I, "0", "sum"), c("Cmp3", I, "0", "sum"),
    c("Cmp4", I, "0", "sum"), c("Cmp5", I, "0", "sum"),
    c("Total_Campaigns", I, "0"), c("Complain", I, "0", "sum"),
    c("Age_Group", S), c("Income_Group", S), c("Marital_Status", S), c("Education", S),
    c("Total_Children", I, "0"), c("Has_Children", I, "0"), c("Family_Size", I, "0"),
    c("Total_Spend", I, "#,##0", "sum"),
    c("Share_Wines", D, "0.0%", "average"), c("Share_Fruits", D, "0.0%", "average"),
    c("Share_Meat", D, "0.0%", "average"), c("Share_Fish", D, "0.0%", "average"),
    c("Share_Sweets", D, "0.0%", "average"),
    c("Total_Purchases", I, "#,##0", "sum"), c("AOV", D, "#,##0.0", "average"),
    c("Deal_Ratio", D, "0.0%", "average"),
    c("Share_Web", D, "0.0%", "average"), c("Share_Catalog", D, "0.0%", "average"),
    c("Share_Store", D, "0.0%", "average"),
    c("Preferred_Channel", S), c("Spend_to_Income_Pct", D, "0.00", "average"),
    c("Campaign_Responder", I, "0", "sum"), c("Recency_Group", S),
    c("R_Score", I, "0"), c("F_Score", I, "0"), c("M_Score", I, "0"),
    c("FM_Score", D, "0.0"), c("RFM_Score", I, "0"),
    c("RFM_Segment", S), c("Value_Tier", S, sort="Value_Tier_Sort"),
]
# Cot tao them trong Power Query
CUSTOMER_ADDED = [
    c("Children_Label", S, sort="Total_Children"),
    c("Web_Visit_Group", S),
    c("Deal_Hunter", S),
    c("Campaigns_Label", S, sort="Total_Campaigns"),
    c("Value_Tier_Sort", I, "0", hidden=True),
    c("Spend_Rank", I, "0"),
    c("Spend_Decile", I, "0", hidden=True),
    c("Spend_Decile_Label", S, sort="Spend_Decile"),
]


def m_types(cols):
    return ", ".join(f'{{"{x["name"]}", {x["mt"]}}}' for x in cols)


CUSTOMERS_M = f"""let
    Source = Excel.Workbook(File.Contents(DataFile), null, true),
    Data_Clean = Source{{[Item="Data_Clean",Kind="Sheet"]}}[Data],
    Promoted = Table.PromoteHeaders(Data_Clean, [PromoteAllScalars=true]),
    Selected = Table.SelectColumns(Promoted, {{{", ".join(f'"{x["name"]}"' for x in CUSTOMER_SOURCE)}}}),
    Typed = Table.TransformColumnTypes(Selected, {{{m_types(CUSTOMER_SOURCE)}}}),
    #"Added Children_Label" = Table.AddColumn(Typed, "Children_Label", each Text.From([Total_Children]) & " con", type text),
    #"Added Web_Visit_Group" = Table.AddColumn(#"Added Children_Label", "Web_Visit_Group", each if [Web_Visits_Month] >= 7 then "Truy cập cao (≥7 lần/tháng)" else "Truy cập thấp (<7 lần/tháng)", type text),
    #"Added Deal_Hunter" = Table.AddColumn(#"Added Web_Visit_Group", "Deal_Hunter", each if [Deal_Ratio] >= 0.4 then "Săn giảm giá (≥40% đơn)" else "Mua giá thường (<40% đơn)", type text),
    #"Added Campaigns_Label" = Table.AddColumn(#"Added Deal_Hunter", "Campaigns_Label", each Text.From([Total_Campaigns]) & " chiến dịch", type text),
    #"Added Value_Tier_Sort" = Table.AddColumn(#"Added Campaigns_Label", "Value_Tier_Sort", each if [Value_Tier] = "Low" then 1 else if [Value_Tier] = "Mid" then 2 else 3, Int64.Type),
    #"Sorted by Spend" = Table.Sort(#"Added Value_Tier_Sort", {{{{"Total_Spend", Order.Descending}}, {{"CustomerID", Order.Ascending}}}}),
    #"Added Spend_Rank" = Table.AddIndexColumn(#"Sorted by Spend", "Spend_Rank", 1, 1, Int64.Type),
    RowCount = Table.RowCount(#"Added Spend_Rank"),
    #"Added Spend_Decile" = Table.AddColumn(#"Added Spend_Rank", "Spend_Decile", each Number.RoundUp([Spend_Rank] * 10 / RowCount), Int64.Type),
    Result = Table.AddColumn(#"Added Spend_Decile", "Spend_Decile_Label", each if [Spend_Decile] = 1 then "Top 10%" else Text.From(([Spend_Decile] - 1) * 10) & "-" & Text.From([Spend_Decile] * 10) & "%", type text)
in
    Result"""

SPEND_M = """let
    Source = Table.SelectColumns(Customers, {"CustomerID", "Spend_Wines", "Spend_Meat", "Spend_Fish", "Spend_Fruits", "Spend_Sweets"}),
    Unpivoted = Table.UnpivotOtherColumns(Source, {"CustomerID"}, "Category_Code", "Spend"),
    Map = [Spend_Wines = {"Rượu vang", 1}, Spend_Meat = {"Thịt", 2}, Spend_Fish = {"Cá", 3}, Spend_Fruits = {"Trái cây", 4}, Spend_Sweets = {"Đồ ngọt", 5}],
    #"Added Category" = Table.AddColumn(Unpivoted, "Category", each Record.Field(Map, [Category_Code]){0}, type text),
    #"Added Category_Sort" = Table.AddColumn(#"Added Category", "Category_Sort", each Record.Field(Map, [Category_Code]){1}, Int64.Type),
    Typed = Table.TransformColumnTypes(#"Added Category_Sort", {{"CustomerID", type text}, {"Category_Code", type text}, {"Spend", Int64.Type}})
in
    Typed"""

CHANNEL_M = """let
    Source = Table.SelectColumns(Customers, {"CustomerID", "Store_Purchases", "Web_Purchases", "Catalog_Purchases"}),
    Unpivoted = Table.UnpivotOtherColumns(Source, {"CustomerID"}, "Channel_Code", "Purchases"),
    Map = [Store_Purchases = {"Store", 1}, Web_Purchases = {"Web", 2}, Catalog_Purchases = {"Catalog", 3}],
    #"Added Channel" = Table.AddColumn(Unpivoted, "Channel", each Record.Field(Map, [Channel_Code]){0}, type text),
    #"Added Channel_Sort" = Table.AddColumn(#"Added Channel", "Channel_Sort", each Record.Field(Map, [Channel_Code]){1}, Int64.Type),
    Typed = Table.TransformColumnTypes(#"Added Channel_Sort", {{"CustomerID", type text}, {"Channel_Code", type text}, {"Purchases", Int64.Type}})
in
    Typed"""

CAMPAIGN_M = """let
    Source = Table.SelectColumns(Customers, {"CustomerID", "Cmp1", "Cmp2", "Cmp3", "Cmp4", "Cmp5"}),
    Unpivoted = Table.UnpivotOtherColumns(Source, {"CustomerID"}, "Campaign_Code", "Accepted"),
    #"Added Campaign_No" = Table.AddColumn(Unpivoted, "Campaign_No", each Number.From(Text.End([Campaign_Code], 1)), Int64.Type),
    #"Added Campaign" = Table.AddColumn(#"Added Campaign_No", "Campaign", each "Chiến dịch " & Text.From([Campaign_No]), type text),
    Typed = Table.TransformColumnTypes(#"Added Campaign", {{"CustomerID", type text}, {"Campaign_Code", type text}, {"Accepted", Int64.Type}})
in
    Typed"""

LOG_M = """let
    Source = Excel.Workbook(File.Contents(DataFile), null, true),
    Cleaning_Log = Source{[Item="Cleaning_Log",Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Cleaning_Log, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Step", type text}, {"Action", type text}, {"Rows_Affected", Int64.Type}, {"Reason", type text}, {"Rows_After", Int64.Type}}),
    Result = Table.AddIndexColumn(Typed, "Log_Order", 1, 1, Int64.Type)
in
    Result"""

FUNNEL_M = """let
    Log = Cleaning_Log,
    RawRows = Table.SelectRows(Log, each [Step] = "R0"){0}[Rows_After],
    AfterDedup = Table.SelectRows(Log, each [Step] = "R2"){0}[Rows_After],
    FinalRows = Table.SelectRows(Log, each [Step] = "R10"){0}[Rows_After],
    Funnel = #table(
        type table [Stage = text, Rows = Int64.Type, Stage_Sort = Int64.Type],
        {
            {"1. Dữ liệu gốc", RawRows, 1},
            {"2. Sau loại trùng lặp", AfterDedup, 2},
            {"3. Sau loại mâu thuẫn logic", FinalRows, 3}
        })
in
    Funnel"""


def static_m(cols_types, rows):
    t = ", ".join(f"{n} = {mt}" for n, mt in cols_types)

    def lit(v):
        return str(v) if isinstance(v, int) else '"' + v.replace('"', '""') + '"'
    body = ",\n            ".join("{" + ", ".join(lit(v) for v in r) + "}" for r in rows)
    return f"""let
    Source = #table(
        type table [{t}],
        {{
            {body}
        }})
in
    Source"""


SEGMENTS = [
    ("Champions", 1, "Mua gần đây, chi tiêu và tần suất cao nhất",
     "Chương trình VIP, ưu tiên catalog cao cấp; không cần giảm giá"),
    ("Loyal Customers", 2, "Mua đều đặn, giá trị khá",
     "Upsell / cross-sell để nâng lên Champions"),
    ("New / Promising", 3, "Mới mua gần đây nhưng giá trị còn thấp",
     "Nuôi dưỡng bằng ưu đãi kiểu Chiến dịch 3 (nhóm đại trà)"),
    ("Need Attention", 4, "Giá trị thấp, bắt đầu thưa dần",
     "Nhắc nhở tự động, mã giảm giá nhỏ"),
    ("At Risk", 5, "Từng chi tiêu cao nhưng đã lâu không quay lại",
     "WIN-BACK NGAY: catalog rượu vang & thịt cá nhân hoá, ưu đãi có hạn 30 ngày"),
    ("Hibernating", 6, "Giá trị thấp và lâu không mua",
     "Email tự động chi phí thấp; không dồn ngân sách"),
]

TABLES = {
    "Customers": dict(
        cols=CUSTOMER_SOURCE + CUSTOMER_ADDED, m=CUSTOMERS_M,
        desc="Bảng sự kiện chính: 1 dòng = 1 khách hàng (dữ liệu đã làm sạch)"),
    "Dim_IncomeGroup": dict(
        cols=[c("Income_Group", S, sort="Income_Sort"), c("Income_Sort", I, "0", hidden=True)],
        m=static_m([("Income_Group", "text"), ("Income_Sort", "Int64.Type")],
                   [("<30K", 1), ("30-50K", 2), ("50-70K", 3), ("70-90K", 4), ("90K+", 5)]),
        desc="Dimension nhóm thu nhập"),
    "Dim_AgeGroup": dict(
        cols=[c("Age_Group", S, sort="Age_Sort"), c("Age_Sort", I, "0", hidden=True)],
        m=static_m([("Age_Group", "text"), ("Age_Sort", "Int64.Type")],
                   [("24-34", 1), ("35-44", 2), ("45-54", 3), ("55-64", 4), ("65+", 5)]),
        desc="Dimension nhóm tuổi"),
    "Dim_Segment": dict(
        cols=[c("RFM_Segment", S, sort="Segment_Sort"), c("Segment_Sort", I, "0", hidden=True),
              c("Segment_Description", S), c("Recommended_Action", S)],
        m=static_m([("RFM_Segment", "text"), ("Segment_Sort", "Int64.Type"),
                    ("Segment_Description", "text"), ("Recommended_Action", "text")], SEGMENTS),
        desc="Dimension phân khúc RFM + hành động đề xuất"),
    "Spend_By_Category": dict(
        cols=[c("CustomerID", S), c("Category_Code", S, hidden=True),
              c("Spend", I, "#,##0", "sum"), c("Category", S, sort="Category_Sort"),
              c("Category_Sort", I, "0", hidden=True)],
        m=SPEND_M, desc="Unpivot chi tiêu 5 nhóm hàng"),
    "Purchases_By_Channel": dict(
        cols=[c("CustomerID", S), c("Channel_Code", S, hidden=True),
              c("Purchases", I, "#,##0", "sum"), c("Channel", S, sort="Channel_Sort"),
              c("Channel_Sort", I, "0", hidden=True)],
        m=CHANNEL_M, desc="Unpivot số lần mua theo 3 kênh"),
    "Campaign_Response": dict(
        cols=[c("CustomerID", S), c("Campaign_Code", S, hidden=True),
              c("Accepted", I, "0", "sum"), c("Campaign_No", I, "0", hidden=True),
              c("Campaign", S, sort="Campaign_No")],
        m=CAMPAIGN_M, desc="Unpivot phản hồi 5 chiến dịch"),
    "Cleaning_Log": dict(
        cols=[c("Step", S), c("Action", S), c("Rows_Affected", I, "#,##0", "sum"),
              c("Reason", S), c("Rows_After", I, "#,##0"), c("Log_Order", I, "0")],
        m=LOG_M, desc="Nhật ký làm sạch dữ liệu"),
    "Data_Funnel": dict(
        cols=[c("Stage", S, sort="Stage_Sort"), c("Rows", I, "#,##0", "sum"),
              c("Stage_Sort", I, "0", hidden=True)],
        m=FUNNEL_M, desc="Phễu dữ liệu trước/sau làm sạch"),
}

RELATIONSHIPS = [
    ("Customers", "Income_Group", "Dim_IncomeGroup", "Income_Group"),
    ("Customers", "Age_Group", "Dim_AgeGroup", "Age_Group"),
    ("Customers", "RFM_Segment", "Dim_Segment", "RFM_Segment"),
    ("Spend_By_Category", "CustomerID", "Customers", "CustomerID"),
    ("Purchases_By_Channel", "CustomerID", "Customers", "CustomerID"),
    ("Campaign_Response", "CustomerID", "Customers", "CustomerID"),
]

# (ten, DAX, format, folder)
F1, F2, F3, F4, F5, F6 = ("1. Tổng quan", "2. Khách hàng", "3. Sản phẩm & kênh",
                          "4. Chiến dịch", "5. RFM & giữ chân", "6. Chất lượng dữ liệu")
ACCEPTED = "FILTER(Campaign_Response, Campaign_Response[Accepted] = 1)"
MEASURE_LIST = [
    ("Số khách hàng", "COUNTROWS(Customers)", "#,##0", F1),
    ("Tổng chi tiêu", "SUM(Customers[Total_Spend])", "#,##0", F1),
    ("Chi tiêu TB/khách", "DIVIDE([Tổng chi tiêu], [Số khách hàng])", "#,##0", F1),
    ("Tổng số đơn", "SUM(Customers[Total_Purchases])", "#,##0", F1),
    ("AOV (giá trị TB/đơn)", "DIVIDE([Tổng chi tiêu], [Tổng số đơn])", "#,##0.0", F1),
    ("Recency TB (ngày)", "AVERAGE(Customers[Recency])", "0.0", F1),
    ("% khách hàng",
     "DIVIDE([Số khách hàng], CALCULATE([Số khách hàng], ALLSELECTED(Customers)))", "0.0%", F1),
    ("% doanh thu",
     "DIVIDE([Tổng chi tiêu], CALCULATE([Tổng chi tiêu], ALLSELECTED(Customers)))", "0.0%", F1),
    ("Top 20% khách chiếm % doanh thu", [
        "VAR _all = CALCULATETABLE(Customers, ALLSELECTED(Customers))",
        "VAR _n = ROUND(COUNTROWS(_all) * 0.2, 0)",
        "VAR _top = TOPN(_n, _all, Customers[Total_Spend], DESC)",
        "RETURN",
        "    DIVIDE(SUMX(_top, Customers[Total_Spend]), SUMX(_all, Customers[Total_Spend]))"],
     "0.0%", F1),
    ("% doanh thu theo thập phân vị", [
        "DIVIDE(",
        "    [Tổng chi tiêu],",
        "    CALCULATE([Tổng chi tiêu], REMOVEFILTERS(Customers[Spend_Decile_Label], Customers[Spend_Decile]))",
        ")"], "0.0%", F1),
    ("% doanh thu lũy kế", [
        "VAR _d = MAX(Customers[Spend_Decile])",
        "RETURN",
        "    DIVIDE(",
        "        CALCULATE([Tổng chi tiêu], REMOVEFILTERS(Customers[Spend_Decile_Label]), Customers[Spend_Decile] <= _d),",
        "        CALCULATE([Tổng chi tiêu], REMOVEFILTERS(Customers[Spend_Decile_Label], Customers[Spend_Decile]))",
        "    )"], "0.0%", F1),
    ("Thu nhập TB", "AVERAGE(Customers[Income])", "#,##0", F2),
    ("Số con TB", "AVERAGE(Customers[Total_Children])", "0.00", F2),
    ("Chi tiêu theo nhóm hàng", "SUM(Spend_By_Category[Spend])", "#,##0", F3),
    ("Tỷ trọng nhóm hàng",
     "DIVIDE([Chi tiêu theo nhóm hàng], CALCULATE([Chi tiêu theo nhóm hàng], ALLSELECTED(Spend_By_Category)))",
     "0.0%", F3),
    ("Số lần mua theo kênh", "SUM(Purchases_By_Channel[Purchases])", "#,##0", F3),
    ("Tỷ lệ mua giảm giá", "DIVIDE(SUM(Customers[Deals_Purchases]), [Tổng số đơn])", "0.0%", F3),
    ("Lượt truy cập web TB/tháng", "AVERAGE(Customers[Web_Visits_Month])", "0.0", F3),
    ("Số đơn web TB", "AVERAGE(Customers[Web_Purchases])", "0.0", F3),
    ("Khách phản hồi CD", "CALCULATE([Số khách hàng], Customers[Campaign_Responder] = 1)", "#,##0", F4),
    ("Tỷ lệ phản hồi CD", "DIVIDE([Khách phản hồi CD], [Số khách hàng])", "0.0%", F4),
    ("Khách chưa từng phản hồi", "CALCULATE([Số khách hàng], Customers[Campaign_Responder] = 0)", "#,##0", F4),
    ("% chưa từng phản hồi", "DIVIDE([Khách chưa từng phản hồi], [Số khách hàng])", "0.0%", F4),
    ("Lượt chấp nhận CD", "SUM(Campaign_Response[Accepted])", "#,##0", F4),
    ("Tỷ lệ chấp nhận CD", "DIVIDE([Lượt chấp nhận CD], [Số khách hàng])", "0.0%", F4),
    ("Thu nhập TB người nhận CD", f"AVERAGEX({ACCEPTED}, RELATED(Customers[Income]))", "#,##0", F4),
    ("Số con TB người nhận CD", f"AVERAGEX({ACCEPTED}, RELATED(Customers[Total_Children]))", "0.00", F4),
    ("Chi tiêu TB người nhận CD", f"AVERAGEX({ACCEPTED}, RELATED(Customers[Total_Spend]))", "#,##0", F4),
    ("Doanh thu rủi ro (At Risk)", 'CALCULATE([Tổng chi tiêu], Customers[RFM_Segment] = "At Risk")', "#,##0", F5),
    ("% doanh thu rủi ro", "DIVIDE([Doanh thu rủi ro (At Risk)], [Tổng chi tiêu])", "0.0%", F5),
    ("Khách At Risk", 'CALCULATE([Số khách hàng], Customers[RFM_Segment] = "At Risk")', "#,##0", F5),
    ("Khách VIP đang At Risk",
     'CALCULATE([Số khách hàng], Customers[RFM_Segment] = "At Risk", Customers[Value_Tier] = "High")', "#,##0", F5),
    ("Tỷ lệ win-back", "SELECTEDVALUE('Winback Rate'[Winback Rate], 0.1)", "0%", F5),
    ("Doanh thu giữ lại được", "[Doanh thu rủi ro (At Risk)] * [Tỷ lệ win-back]", "#,##0", F5),
    ("% tổng doanh thu giữ lại", "DIVIDE([Doanh thu giữ lại được], [Tổng chi tiêu])", "0.0%", F5),
    ("Dòng dữ liệu gốc", "CALCULATE(MAX(Data_Funnel[Rows]), Data_Funnel[Stage_Sort] = 1)", "#,##0", F6),
    ("Dòng dữ liệu sạch", "CALCULATE(MAX(Data_Funnel[Rows]), Data_Funnel[Stage_Sort] = 3)", "#,##0", F6),
    ("Dòng bị loại", "[Dòng dữ liệu gốc] - [Dòng dữ liệu sạch]", "#,##0", F6),
    ("% dữ liệu giữ lại", "DIVIDE([Dòng dữ liệu sạch], [Dòng dữ liệu gốc])", "0.0%", F6),
    ("Số dòng", "SUM(Data_Funnel[Rows])", "#,##0", F6),
]
MEASURE_NAMES = {m[0] for m in MEASURE_LIST}

COLS = {t: {x["name"] for x in spec["cols"]} for t, spec in TABLES.items()}
COLS["Winback Rate"] = {"Winback Rate"}


def check_dax():
    for name, expr, *_ in MEASURE_LIST:
        text = expr if isinstance(expr, str) else "\n".join(expr)
        for tbl, col in re.findall(r"('?[A-Za-z_][A-Za-z0-9_ ]*'?)\[([^\]]+)\]", text):
            tbl = tbl.strip("'")
            assert tbl in COLS and col in COLS[tbl], f"{name}: {tbl}[{col}]"
        bare = re.sub(r"'?[A-Za-z_][A-Za-z0-9_ ]*'?\[[^\]]+\]", "", text)
        for ref in re.findall(r"\[([^\]]+)\]", bare):
            assert ref in MEASURE_NAMES, f"{name}: [{ref}]"
    for spec in TABLES.values():
        for x in spec["cols"]:
            if x["sort"]:
                assert x["sort"] in {y["name"] for y in spec["cols"]}, x


# =============================================================================
# 2. TMDL WRITER
# =============================================================================
T = "\t"


def tmdl_column(tname, x):
    out = [f"{T}column {q(x['name'])}", f"{T*2}dataType: {x['dt']}"]
    if x["hidden"]:
        out.append(f"{T*2}isHidden")
    if x["fmt"]:
        out.append(f"{T*2}formatString: {x['fmt']}")
    out += [f"{T*2}lineageTag: {lt(tname + '/' + x['name'])}",
            f"{T*2}summarizeBy: {x['summ']}",
            f"{T*2}sourceColumn: {x['name']}"]
    if x["sort"]:
        out.append(f"{T*2}sortByColumn: {q(x['sort'])}")
    out += ["", f"{T*2}annotation SummarizationSetBy = Automatic", ""]
    return out


def tmdl_m_partition(pname, m):
    out = [f"{T}partition {q(pname)} = m", f"{T*2}mode: import", f"{T*2}source ="]
    out += [f"{T*4}{line}" if line else "" for line in m.splitlines()]
    return out + [""]


def tmdl_table(tname, spec):
    out = [f"/// {spec['desc']}", f"table {q(tname)}", f"{T}lineageTag: {lt(tname)}", ""]
    for x in spec["cols"]:
        out += tmdl_column(tname, x)
    out += tmdl_m_partition(tname, spec["m"])
    out += [f"{T}annotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def tmdl_measures():
    out = ["/// Bảng chứa toàn bộ measure DAX", f"table {MEASURES}", f"{T}lineageTag: {lt(MEASURES)}", ""]
    for name, expr, fmt, folder in MEASURE_LIST:
        if isinstance(expr, str):
            out.append(f"{T}measure {q(name)} = {expr}")
        else:
            out.append(f"{T}measure {q(name)} =")
            out += [f"{T*3}{line}" for line in expr]
        out += [f"{T*2}formatString: {fmt}", f"{T*2}displayFolder: {folder}",
                f"{T*2}lineageTag: {lt('m/' + name)}", ""]
    out += [f"{T}column Dummy", f"{T*2}dataType: int64", f"{T*2}isHidden", f"{T*2}formatString: 0",
            f"{T*2}lineageTag: {lt(MEASURES + '/Dummy')}", f"{T*2}summarizeBy: none",
            f"{T*2}sourceColumn: Dummy", "", f"{T*2}annotation SummarizationSetBy = Automatic", ""]
    out += tmdl_m_partition(MEASURES, "let\n    Source = #table(type table [Dummy = Int64.Type], {{1}})\nin\n    Source")
    out += [f"{T}annotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def tmdl_winback():
    return "\n".join([
        "/// Tham số What-if: tỷ lệ kéo khách At Risk quay lại",
        "table 'Winback Rate'", f"{T}lineageTag: {lt('Winback Rate')}", "",
        f"{T}column 'Winback Rate'", f"{T*2}formatString: 0%",
        f"{T*2}lineageTag: {lt('Winback Rate/col')}", f"{T*2}summarizeBy: none",
        f"{T*2}sourceColumn: [Value]", "",
        f"{T*2}extendedProperty ParameterMetadata =", f"{T*4}{{", f'{T*4}  "version": 0', f"{T*4}}}", "",
        f"{T*2}annotation SummarizationSetBy = User", "",
        f"{T}partition 'Winback Rate' = calculated", f"{T*2}mode: import",
        f"{T*2}source = GENERATESERIES(0, 0.5, 0.05)", "",
        f"{T}annotation PBI_Id = {hid('winback')}", ""])


def write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def wjson(path: Path, obj):
    write(path, json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def build_model():
    d = SM / "definition"
    wjson(SM / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": lt("semanticmodel")}})
    wjson(SM / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.0", "settings": {}})
    wjson(SM / ".pbi" / "editorSettings.json", {
        "version": "1.0", "autodetectRelationships": False, "parallelQueryLoading": True,
        "typeDetectionEnabled": True, "relationshipImportEnabled": False,
        "shouldNotifyUserOfNameConflictResolution": True})
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n")
    names = list(TABLES) + [MEASURES, "Winback Rate"]
    order = ["DataFile"] + list(TABLES) + [MEASURES]
    write(d / "model.tmdl", "\n".join([
        "model Model", f"{T}culture: en-US", f"{T}defaultPowerBIDataSourceVersion: powerBI_V3",
        f"{T}sourceQueryCulture: en-US", f"{T}dataAccessOptions", f"{T*2}legacyRedirects",
        f"{T*2}returnErrorValuesAsNull", "",
        f"annotation PBI_QueryOrder = {json.dumps(order, ensure_ascii=False)}", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", ""]
        + [f"ref table {q(n)}" for n in names]) + "\n")
    write(d / "expressions.tmdl", "\n".join([
        "/// Đường dẫn tới file FoodApp_Clean.xlsx trên máy bạn",
        f'expression DataFile = "{DEFAULT_PATH}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"{T}lineageTag: {lt('DataFile')}", "", f"{T}annotation PBI_ResultType = Text", ""]))
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {lt(f'rel/{ft}.{fc}->{tt}.{tc}')}",
                f"{T}fromColumn: {q(ft)}.{q(fc)}", f"{T}toColumn: {q(tt)}.{q(tc)}", ""]
    write(d / "relationships.tmdl", "\n".join(rel))
    for tname, spec in TABLES.items():
        write(d / "tables" / f"{tname}.tmdl", tmdl_table(tname, spec))
    write(d / "tables" / f"{MEASURES}.tmdl", tmdl_measures())
    write(d / "tables" / "Winback Rate.tmdl", tmdl_winback())


# =============================================================================
# 3. REPORT (PBIR) HELPERS
# =============================================================================
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
NAVY, ORANGE, GRAY, RED = "#1F4E79", "#E07B39", "#6B7280", "#C0392B"


def lit(v):
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return f"{v}L"
    if isinstance(v, float):
        return f"{v}D"
    return "'" + str(v).replace("'", "''") + "'"


def p(v):
    return {"expr": {"Literal": {"Value": lit(v)}}}


def color(hexv):
    return {"solid": {"color": p(hexv)}}


def fcol(table, col):
    assert col in COLS[table], f"{table}[{col}]"
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": col}}


def fmeas(name):
    assert name in MEASURE_NAMES, name
    return {"Measure": {"Expression": {"SourceRef": {"Entity": MEASURES}}, "Property": name}}


AGG = {"sum": (0, "Sum"), "avg": (1, "Avg")}


def field(spec):
    """'Table.Col' | 'm:Measure' | 'sum:Table.Col'"""
    if spec.startswith("m:"):
        n = spec[2:]
        return fmeas(n), f"{MEASURES}.{n}"
    if ":" in spec:
        fn, rest = spec.split(":", 1)
        t, col = rest.split(".", 1)
        code, lbl = AGG[fn]
        return {"Aggregation": {"Expression": fcol(t, col), "Function": code}}, f"{lbl}({t}.{col})"
    t, col = spec.split(".", 1)
    return fcol(t, col), f"{t}.{col}"


def projections(specs):
    out = []
    for s in specs:
        disp = None
        if isinstance(s, tuple):
            s, disp = s
        f, ref = field(s)
        pr = {"field": f, "queryRef": ref, "nativeQueryRef": ref.split(".", 1)[-1]}
        if disp:
            pr["displayName"] = disp
        out.append(pr)
    return out


class Page:
    def __init__(self, key, display, title, subtitle, sowhat=None, slicers=True):
        self.key, self.name, self.display = key, hid("page/" + key), display
        self.visuals, self.z = [], 1000
        self.extra = {}
        self.textbox("header", 20, 8, 620 if slicers else 1240, 62, [
            [(title, {"fontSize": "18pt", "fontWeight": "bold", "color": NAVY})],
            [(subtitle, {"fontSize": "10pt", "color": GRAY})]], frame=False)
        if slicers:
            for i, (fld, disp, grp) in enumerate([
                    ("Dim_IncomeGroup.Income_Group", "Nhóm thu nhập", "sync_income"),
                    ("Dim_AgeGroup.Age_Group", "Nhóm tuổi", "sync_age"),
                    ("Customers.Children_Label", "Số con", "sync_children"),
                    ("Dim_Segment.RFM_Segment", "Phân khúc RFM", "sync_segment")]):
                self.slicer(f"slicer_{grp}", 650 + i * 155, 10, 150, 58, fld, disp, grp)
        if sowhat:
            self.textbox("sowhat", 20, 642, 1240, 70, [
                [("SO WHAT?  ", {"fontSize": "10pt", "fontWeight": "bold", "color": ORANGE}),
                 (sowhat[0], {"fontSize": "10pt", "fontWeight": "bold", "color": "#1F2937"})],
                [(sowhat[1], {"fontSize": "9pt", "color": "#374151"})]],
                bg="#FFF4EC")

    def add(self, key, x, y, w, h, visual):
        self.z += 100
        self.visuals.append({
            "$schema": f"{SCHEMA}/visualContainer/2.0.0/schema.json",
            "name": hid(f"{self.key}/{key}"),
            "position": {"x": x, "y": y, "z": self.z, "height": h, "width": w, "tabOrder": self.z},
            "visual": visual})

    def chart(self, key, vtype, x, y, w, h, roles, title=None, sort=None, objects=None, labels=False):
        qs = {r: {"projections": projections(f)} for r, f in roles.items()}
        v = {"visualType": vtype, "query": {"queryState": qs}, "drillFilterOtherVisuals": True}
        if sort:
            fspec, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field(fspec)[0], "direction": direction}],
                                            "isDefaultSort": False}
        obj = dict(objects or {})
        if labels:
            obj.setdefault("labels", [{"properties": {"show": p(True)}}])
        if obj:
            v["objects"] = obj
        if title:
            v["visualContainerObjects"] = {"title": [{"properties": {"show": p(True), "text": p(title)}}]}
        self.add(key, x, y, w, h, v)

    def card(self, key, x, y, w, h, measure, label=None):
        self.chart(key, "card", x, y, w, h, {"Values": [(f"m:{measure}", label) if label else f"m:{measure}"]})

    def slicer(self, key, x, y, w, h, fld, disp, group, dropdown=True, single=False):
        obj = {}
        if dropdown:
            obj["data"] = [{"properties": {"mode": p("Dropdown")}}]
        if single:
            obj["selection"] = [{"properties": {"singleSelect": p(True)}}]
        v = {"visualType": "slicer",
             "query": {"queryState": {"Values": {"projections": projections([(fld, disp)])}}},
             "objects": obj, "drillFilterOtherVisuals": True}
        if group:
            v["syncGroup"] = {"groupName": group, "fieldChanges": True, "filterChanges": True}
        self.add(key, x, y, w, h, v)

    def textbox(self, key, x, y, w, h, paragraphs, frame=True, bg=None):
        paras = [{"textRuns": [{"value": t, "textStyle": {"fontFamily": "Segoe UI", **st}} for t, st in para]}
                 for para in paragraphs]
        v = {"visualType": "textbox",
             "objects": {"general": [{"properties": {"paragraphs": paras}}]},
             "drillFilterOtherVisuals": True}
        vco = {}
        if not frame:
            vco["background"] = [{"properties": {"show": p(False)}}]
            vco["border"] = [{"properties": {"show": p(False)}}]
        if bg:
            vco["background"] = [{"properties": {"show": p(True), "color": color(bg), "transparency": p(0.0)}}]
            vco["border"] = [{"properties": {"show": p(True), "color": color(ORANGE)}}]
        if vco:
            v["visualContainerObjects"] = vco
        self.add(key, x, y, w, h, v)

    def json(self):
        page = {"$schema": f"{SCHEMA}/page/1.4.0/schema.json", "name": self.name,
                "displayName": self.display, "displayOption": "FitToPage",
                "height": 720, "width": 1280}
        page.update(self.extra)
        return page


ASC, DESC = "Ascending", "Descending"


# =============================================================================
# 4. PAGES
# =============================================================================
def build_pages():
    pages = []

    # ---- Trang 1: Tong quan -------------------------------------------------
    pg = Page("overview", "1. Tổng quan",
              "Doanh thu nằm trong tay rất ít khách hàng",
              "Tổng quan hiệu quả kinh doanh Food App - 1.992 khách hàng sau làm sạch",
              ("20% khách hàng tạo ra ~53% doanh thu; rượu vang + thịt chiếm ~83% chi tiêu.",
               "Doanh nghiệp chịu rủi ro tập trung kép (ít khách VIP + ít nhóm hàng). "
               "Mất một phần nhỏ khách giá trị cao sẽ làm doanh thu giảm mạnh → ưu tiên giữ chân."))
    for i, m in enumerate(["Số khách hàng", "Tổng chi tiêu", "Chi tiêu TB/khách",
                           "AOV (giá trị TB/đơn)", "Tỷ lệ phản hồi CD", "Top 20% khách chiếm % doanh thu"]):
        pg.card(f"kpi{i}", 20 + i * 208, 78, 200, 82, m)
    pg.chart("pareto", "lineClusteredColumnComboChart", 20, 170, 610, 462,
             {"Category": [("Customers.Spend_Decile_Label", "Nhóm 10% khách (xếp theo chi tiêu)")],
              "Y": [("m:% doanh thu theo thập phân vị", "% doanh thu")],
              "Y2": [("m:% doanh thu lũy kế", "% lũy kế")]},
             "Pareto: mỗi 10% khách hàng đóng góp bao nhiêu doanh thu?",
             sort=("Customers.Spend_Decile_Label", ASC), labels=True)
    pg.chart("donut", "donutChart", 640, 170, 300, 462,
             {"Category": ["Spend_By_Category.Category"], "Y": [("m:Chi tiêu theo nhóm hàng", "Chi tiêu")]},
             "Cơ cấu chi tiêu theo nhóm hàng")
    pg.chart("treemap", "treemap", 950, 170, 310, 462,
             {"Group": ["Dim_Segment.RFM_Segment"], "Values": [("m:Tổng chi tiêu", "Chi tiêu")]},
             "Doanh thu theo phân khúc RFM")
    pages.append(pg)

    # ---- Trang 2: Chan dung khach hang --------------------------------------
    pg = Page("profile", "2. Chân dung khách hàng",
              "Thu nhập và con cái quyết định mức chi tiêu",
              "Ai chi tiêu nhiều, ai chi tiêu ít?",
              ("Thu nhập 90K+ chi gấp ~33 lần nhóm <30K (r = 0,83); khách không con chi gấp ~3 lần khách 2 con.",
               "Phân khúc theo thu nhập + cấu trúc gia đình, KHÔNG theo hôn nhân (chênh lệch nhỏ). "
               "Chân dung VIP: thu nhập ≥70K, ít/không con."))
    pg.chart("by_income", "clusteredColumnChart", 20, 78, 405, 270,
             {"Category": [("Dim_IncomeGroup.Income_Group", "Nhóm thu nhập")],
              "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo nhóm thu nhập",
             sort=("Dim_IncomeGroup.Income_Group", ASC), labels=True)
    pg.chart("by_children", "clusteredColumnChart", 435, 78, 405, 270,
             {"Category": [("Customers.Children_Label", "Số con")], "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo số con",
             sort=("Customers.Children_Label", ASC), labels=True)
    pg.chart("by_age", "clusteredBarChart", 850, 78, 410, 270,
             {"Category": [("Dim_AgeGroup.Age_Group", "Nhóm tuổi")], "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo nhóm tuổi",
             sort=("Dim_AgeGroup.Age_Group", ASC), labels=True)
    pg.chart("scatter", "scatterChart", 20, 358, 820, 274,
             {"Category": ["Customers.CustomerID"],
              "Series": [("Customers.Children_Label", "Số con")],
              "X": [("sum:Customers.Income", "Thu nhập")],
              "Y": [("sum:Customers.Total_Spend", "Tổng chi tiêu")]},
             "Thu nhập vs chi tiêu từng khách (màu = số con)")
    pg.chart("by_edu", "clusteredBarChart", 850, 358, 410, 132,
             {"Category": [("Customers.Education", "Học vấn")], "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo học vấn", labels=True)
    pg.chart("by_marital", "clusteredBarChart", 850, 500, 410, 132,
             {"Category": [("Customers.Marital_Status", "Hôn nhân")], "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo hôn nhân", labels=True)
    pages.append(pg)

    # ---- Trang 3: San pham & kenh -------------------------------------------
    pg = Page("channel", "3. Sản phẩm & kênh",
              "Catalog là kênh bị bỏ quên, web nhiều lượt xem nhưng ít đơn",
              "Khách mua gì, mua qua kênh nào?",
              ("Catalog: 7% khách nhưng 15% doanh thu, AOV 69 - cao hơn ~60% so với cửa hàng (44). "
               "Khách truy cập web ≥7 lần/tháng chỉ chi ~40% so với nhóm còn lại.",
               "Mở rộng catalog cho nhóm thu nhập ≥70K; với web, vấn đề là CHUYỂN ĐỔI chứ không phải traffic "
               "- nhóm truy cập nhiều chủ yếu săn giảm giá."))
    pg.chart("cat_income", "hundredPercentStackedColumnChart", 20, 78, 610, 270,
             {"Category": [("Dim_IncomeGroup.Income_Group", "Nhóm thu nhập")],
              "Series": [("Spend_By_Category.Category", "Nhóm hàng")],
              "Y": [("m:Chi tiêu theo nhóm hàng", "Chi tiêu")]},
             "Cơ cấu giỏ hàng theo nhóm thu nhập",
             sort=("Dim_IncomeGroup.Income_Group", ASC))
    pg.chart("aov_channel", "clusteredColumnChart", 640, 78, 300, 270,
             {"Category": [("Customers.Preferred_Channel", "Kênh ưa thích")],
              "Y": ["m:AOV (giá trị TB/đơn)"]},
             "AOV theo kênh mua ưa thích", sort=("m:AOV (giá trị TB/đơn)", DESC), labels=True)
    pg.chart("share_channel", "clusteredColumnChart", 950, 78, 310, 270,
             {"Category": [("Customers.Preferred_Channel", "Kênh ưa thích")],
              "Y": ["m:% khách hàng", "m:% doanh thu"]},
             "% khách vs % doanh thu theo kênh", sort=("m:% doanh thu", DESC), labels=True)
    pg.chart("channel_income", "hundredPercentStackedBarChart", 20, 358, 610, 274,
             {"Category": [("Dim_IncomeGroup.Income_Group", "Nhóm thu nhập")],
              "Series": [("Purchases_By_Channel.Channel", "Kênh")],
              "Y": [("m:Số lần mua theo kênh", "Số lần mua")]},
             "Cơ cấu số lần mua theo kênh và thu nhập",
             sort=("Dim_IncomeGroup.Income_Group", ASC))
    pg.chart("web_table", "tableEx", 640, 358, 620, 274,
             {"Values": [("Customers.Web_Visit_Group", "Nhóm truy cập web"), "m:Số khách hàng",
                         "m:Thu nhập TB", "m:Chi tiêu TB/khách", "m:Tỷ lệ mua giảm giá",
                         "m:Số đơn web TB", "m:Lượt truy cập web TB/tháng"]},
             "Truy cập nhiều ≠ mua nhiều: so sánh 2 nhóm truy cập web")
    pages.append(pg)

    # ---- Trang 4: Chien dich -------------------------------------------------
    pg = Page("campaign", "4. Hiệu quả chiến dịch",
              "Chiến dịch đang bỏ sót một nửa khách VIP",
              "Chiến dịch nào hiệu quả, với ai?",
              ("Chỉ 21% khách từng phản hồi; Chiến dịch 2 thất bại (1,3%); ~50% khách giá trị cao chưa từng phản hồi.",
               "Dừng/thiết kế lại CD2; nhân rộng cách tiếp cận của CD3 (chạm được nhóm thu nhập thấp hơn, nhiều con hơn); "
               "nhóm săn giảm giá cần mã giảm giá trực tiếp thay vì chiến dịch."))
    pg.chart("acc_rate", "clusteredColumnChart", 20, 78, 405, 270,
             {"Category": ["Campaign_Response.Campaign"], "Y": ["m:Tỷ lệ chấp nhận CD"]},
             "Tỷ lệ chấp nhận theo chiến dịch", sort=("Campaign_Response.Campaign", ASC), labels=True)
    pg.chart("acc_income", "clusteredColumnChart", 435, 78, 405, 270,
             {"Category": ["Campaign_Response.Campaign"],
              "Y": ["m:Thu nhập TB người nhận CD"]},
             "Thu nhập TB của người nhận từng chiến dịch",
             sort=("Campaign_Response.Campaign", ASC), labels=True)
    pg.chart("resp_income", "clusteredColumnChart", 850, 78, 410, 270,
             {"Category": [("Dim_IncomeGroup.Income_Group", "Nhóm thu nhập")], "Y": ["m:Tỷ lệ phản hồi CD"]},
             "Tỷ lệ phản hồi chiến dịch theo thu nhập",
             sort=("Dim_IncomeGroup.Income_Group", ASC), labels=True)
    pg.chart("spend_ncamp", "clusteredColumnChart", 20, 358, 405, 274,
             {"Category": [("Customers.Campaigns_Label", "Số chiến dịch đã nhận")], "Y": ["m:Chi tiêu TB/khách"]},
             "Chi tiêu TB/khách theo số chiến dịch đã nhận",
             sort=("Customers.Campaigns_Label", ASC), labels=True)
    pg.chart("tier_table", "tableEx", 435, 358, 405, 274,
             {"Values": [("Customers.Value_Tier", "Nhóm giá trị"), "m:Số khách hàng",
                         "m:Tỷ lệ phản hồi CD", "m:Khách chưa từng phản hồi", "m:% chưa từng phản hồi"]},
             "Phủ sóng chiến dịch theo nhóm giá trị khách")
    pg.chart("deal_resp", "clusteredColumnChart", 850, 358, 410, 274,
             {"Category": [("Customers.Deal_Hunter", "Hành vi giá")],
              "Y": ["m:Tỷ lệ phản hồi CD", "m:Tỷ lệ mua giảm giá"]},
             "Khách săn giảm giá ít phản hồi chiến dịch", labels=True)
    pages.append(pg)

    # ---- Trang 5: RFM & giu chan ----------------------------------------------
    pg = Page("rfm", "5. RFM & giữ chân",
              "Gần 40% doanh thu đang có nguy cơ mất",
              "Ai sắp rời đi và đáng bao nhiêu tiền?",
              ("480 khách At Risk (24%) tạo ra ~40% doanh thu, chi TB ~930 nhưng đã ~79 ngày không quay lại.",
               "Ưu tiên số 1: chiến dịch WIN-BACK cho At Risk (đặc biệt khách VIP). "
               "Kéo lại 10% doanh thu nhóm này ≈ +4% tổng doanh thu - rẻ hơn nhiều so với tìm khách mới."))
    for i, m in enumerate(["Doanh thu rủi ro (At Risk)", "% doanh thu rủi ro",
                           "Khách VIP đang At Risk", "Doanh thu giữ lại được"]):
        pg.card(f"kpi{i}", 20 + i * 240, 78, 232, 82, m)
    pg.slicer("winback", 980, 78, 280, 82, "Winback Rate.Winback Rate",
              "Giả định tỷ lệ win-back (mặc định 10%)", None, dropdown=True, single=True)
    pg.chart("bubble", "scatterChart", 20, 170, 610, 462,
             {"Category": ["Dim_Segment.RFM_Segment"],
              "X": ["m:Recency TB (ngày)"], "Y": ["m:Chi tiêu TB/khách"], "Size": ["m:Số khách hàng"]},
             "Bản đồ phân khúc: chi tiêu TB vs số ngày chưa quay lại (bóng = số khách)",
             objects={"categoryLabels": [{"properties": {"show": p(True)}}]})
    pg.chart("seg_table", "tableEx", 640, 170, 620, 250,
             {"Values": [("Dim_Segment.RFM_Segment", "Phân khúc"), "m:Số khách hàng", "m:% khách hàng",
                         "m:% doanh thu", "m:Chi tiêu TB/khách", "m:Recency TB (ngày)", "m:Tỷ lệ phản hồi CD"]},
             "So sánh 6 phân khúc RFM (chuột phải → Drill through để xem danh sách khách)")
    pg.chart("action_table", "tableEx", 640, 430, 620, 202,
             {"Values": [("Dim_Segment.RFM_Segment", "Phân khúc"),
                         ("Dim_Segment.Segment_Description", "Đặc điểm"),
                         ("Dim_Segment.Recommended_Action", "Hành động đề xuất")]},
             "Hành động đề xuất theo phân khúc")
    pages.append(pg)

    # ---- Trang 6: Chat luong du lieu -------------------------------------------
    pg = Page("quality", "6. Chất lượng dữ liệu",
              "Vì sao các con số này đáng tin",
              "Quy trình tiền xử lý & làm sạch dữ liệu (2.205 → 1.992 dòng)",
              ("Loại 204 bản ghi trùng lặp (9,3%) + 9 bản ghi mâu thuẫn logic; giữ lại 90,3% dữ liệu.",
               "Ngoại lai chi tiêu được GIỮ LẠI có chủ đích: đó là khách VIP thật (Champions/At Risk). "
               "Hạn chế: không có ngày giao dịch (dữ liệu dạng snapshot); cột Income nhiều khả năng là thu nhập năm."),
              slicers=False)
    for i, m in enumerate(["Dòng dữ liệu gốc", "Dòng bị loại", "Dòng dữ liệu sạch", "% dữ liệu giữ lại"]):
        pg.card(f"kpi{i}", 20 + i * 313, 78, 300, 82, m)
    pg.chart("funnel", "funnel", 20, 170, 500, 300,
             {"Category": [("Data_Funnel.Stage", "Giai đoạn")], "Y": [("m:Số dòng", "Số dòng")]},
             "Phễu dữ liệu", sort=("Data_Funnel.Stage", ASC), labels=True)
    pg.textbox("outlier_note", 20, 480, 500, 152, [
        [("Vì sao KHÔNG xoá ngoại lai?", {"fontSize": "11pt", "fontWeight": "bold", "color": NAVY})],
        [("• 488 khách có ít nhất 1 nhóm chi tiêu vượt ngưỡng IQR - đây là hành vi thật, không phải lỗi.",
          {"fontSize": "9pt", "color": "#374151"})],
        [("• Top 20% khách tạo ra 53% doanh thu: xoá ngoại lai = xoá nhóm khách quan trọng nhất.",
          {"fontSize": "9pt", "color": "#374151"})],
        [("• Chỉ loại bản ghi MÂU THUẪN LOGIC; khi mô hình hoá dùng log(1+x) để giảm lệch phải.",
          {"fontSize": "9pt", "color": "#374151"})]])
    pg.chart("log_table", "tableEx", 530, 170, 730, 462,
             {"Values": [("Cleaning_Log.Log_Order", "#"), ("Cleaning_Log.Step", "Bước"),
                         ("Cleaning_Log.Action", "Hành động"),
                         ("sum:Cleaning_Log.Rows_Affected", "Số dòng ảnh hưởng"),
                         ("Cleaning_Log.Reason", "Lý do"),
                         ("sum:Cleaning_Log.Rows_After", "Số dòng còn lại")]},
             "Nhật ký làm sạch dữ liệu (Cleaning Log)", sort=("Cleaning_Log.Log_Order", ASC))
    pages.append(pg)

    # ---- Trang 7: Drill-through chi tiet khach hang ----------------------------
    pg = Page("detail", "Chi tiết khách hàng",
              "Danh sách khách hàng theo phân khúc",
              "Trang drill-through: từ trang 5, chuột phải vào một phân khúc → Drill through → Chi tiết khách hàng",
              slicers=False)
    flt = "Filter_" + hid("dt_segment")
    pg.extra = {
        "filterConfig": {"filters": [{
            "name": flt, "field": fcol("Dim_Segment", "RFM_Segment"),
            "type": "Categorical", "howCreated": "Drillthrough"}]},
        "pageBinding": {
            "name": hid("binding_detail"), "type": "Drillthrough",
            "parameters": [{"name": "Param_" + flt, "boundFilter": flt,
                            "fieldExpr": fcol("Dim_Segment", "RFM_Segment")}]},
        "visibility": "HiddenInViewMode"}
    pg.chart("customer_table", "tableEx", 20, 78, 1240, 554,
             {"Values": [("Customers.CustomerID", "Mã KH"), ("Customers.RFM_Segment", "Phân khúc"),
                         ("sum:Customers.Income", "Thu nhập"), ("Customers.Age_Group", "Nhóm tuổi"),
                         ("Customers.Children_Label", "Số con"),
                         ("sum:Customers.Total_Spend", "Tổng chi tiêu"),
                         ("sum:Customers.Recency", "Recency (ngày)"),
                         ("sum:Customers.Total_Purchases", "Số đơn"),
                         ("Customers.Preferred_Channel", "Kênh ưa thích"),
                         ("sum:Customers.Total_Campaigns", "Số CD đã nhận"),
                         ("Customers.Value_Tier", "Nhóm giá trị")]},
             "Danh sách khách hàng (sắp xếp theo tổng chi tiêu)",
             sort=("sum:Customers.Total_Spend", DESC))
    pages.append(pg)
    return pages


# =============================================================================
# 5. REPORT WRITER
# =============================================================================
THEME_NAME = "FoodAppTheme.json"
BASE_THEME = "CY24SU10"


def build_report():
    d = RP / "definition"
    wjson(RP / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": lt("report")}})
    wjson(RP / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    wjson(d / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    wjson(d / "report.json", {
        "$schema": f"{SCHEMA}/report/1.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": BASE_THEME, "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": THEME_NAME, "reportVersionAtImport": "5.61", "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": THEME_NAME, "path": THEME_NAME, "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True}})
    for src, dst in [(f"{BASE_THEME}.json", f"SharedResources/BaseThemes/{BASE_THEME}.json"),
                     (THEME_NAME, f"RegisteredResources/{THEME_NAME}")]:
        target = RP / "StaticResources" / dst
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(THEME_SRC / src, target)

    pages = build_pages()
    wjson(d / "pages" / "pages.json", {
        "$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [pg.name for pg in pages], "activePageName": pages[0].name})
    for pg in pages:
        wjson(d / "pages" / pg.name / "page.json", pg.json())
        for v in pg.visuals:
            wjson(d / "pages" / pg.name / "visuals" / v["name"] / "visual.json", v)
    return pages


def main():
    check_dax()
    if OUT.exists():
        for sub in (SM, RP):
            shutil.rmtree(sub, ignore_errors=True)
    OUT.mkdir(parents=True, exist_ok=True)
    build_model()
    pages = build_report()
    wjson(OUT / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True}})
    write(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    n_vis = sum(len(pg.visuals) for pg in pages)
    print(f"OK: {len(TABLES) + 2} bảng, {len(MEASURE_LIST)} measure, {len(RELATIONSHIPS)} quan hệ, "
          f"{len(pages)} trang, {n_vis} visual -> {OUT}")


if __name__ == "__main__":
    main()
