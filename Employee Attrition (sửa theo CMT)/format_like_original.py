import copy, pandas as pd
from openpyxl import load_workbook
d = pd.read_pickle('/tmp/claude-0/q4/emp_final.pkl')
log = pd.read_pickle('/tmp/claude-0/q4/log.pkl')
wb = load_workbook('orig.xlsx')
ws, wd = wb['Data'], wb['Description']
H, Dt = ws['B1'], ws['B2']                       # mẫu định dạng tiêu đề / dữ liệu
hs = dict(font=copy.copy(H.font), alignment=copy.copy(H.alignment), number_format='General')
ds = dict(font=copy.copy(Dt.font), alignment=copy.copy(Dt.alignment), number_format='General')
orig_w = {ws.cell(1, j).value.replace('\n', ' ').strip() if ws.cell(1, j).value else '': ws.column_dimensions[ws.cell(1, j).column_letter].width for j in range(1, ws.max_column + 1)}
ws.delete_rows(1, ws.max_row)
for k in list(ws.column_dimensions): del ws.column_dimensions[k]
ws.row_dimensions[1].height = 39.75
cols = list(d.columns)
for j, c in enumerate(cols, 1):
    cell = ws.cell(1, j, c)
    for a, v in hs.items(): setattr(cell, a, v)
vals = d.astype(object).where(d.notna(), None).values.tolist()
for i, row in enumerate(vals, 2):
    ws.row_dimensions[i].height = 14.25
    for j, v in enumerate(row, 1):
        if isinstance(v, float) and v.is_integer(): v = int(v)
        cell = ws.cell(i, j, v)
        for a, val in ds.items(): setattr(cell, a, val)
from openpyxl.utils import get_column_letter
for j, c in enumerate(cols, 1):
    L = max(len(c) * 1.05 + 1.5, max(len(str(x)) for x in d[c].astype(str).head(300)) + 2)
    ws.column_dimensions[get_column_letter(j)].width = round(min(max(7.5, L), 36), 2)
for n in list(wb.defined_names): del wb.defined_names[n]
ws.defined_names.clear() if hasattr(ws, 'defined_names') else None

# ---- Description: cùng kiểu chữ với sheet gốc
T, Pp, V = wd['A1'], wd['A3'], wd['A6']
st = {k: dict(font=copy.copy(c.font), alignment=copy.copy(c.alignment), fill=copy.copy(c.fill)) for k, c in (('t', T), ('p', Pp), ('v', V))}
orig_desc = {str(wd.cell(r, 1).value).split(':')[0]: str(wd.cell(r, 1).value) for r in range(6, 40)}
wd.delete_rows(1, wd.max_row)
lines = [('t', 'About Dataset'), ('t', 'Context'),
 ('p', 'Final (cleaned) version of the IBM HR Analytics Employee Attrition & Performance dataset used in the final report. One row per employee (1,470 rows); this is the Employees table loaded into Power BI.'),
 ('t', 'Content'),
 ('p', 'The original 34 columns were checked (no missing values, no duplicate rows or Employee Numbers). Five columns were removed: Over18 and StandardHours (constant) and DailyRate, HourlyRate, MonthlyRate (|r| ≤ 0.035 with MonthlyIncome). Column names were standardised to PascalCase and Employee Number was renamed EmployeeID. 42 derived columns were added, giving 71 columns in total. Columns are listed below.')]
ren_back = {'EmployeeID': 'EmployeeNumber', 'YearsWithCurrManager': 'YearsWithCurrManager'}
derived = {
 'Attrition_Flag': 'Attrition coded Yes = 1, No = 0. (Numerical)', 'OverTime_Flag': 'OverTime coded Yes = 1, No = 0. (Numerical)',
 'Age_Group': 'Age band: 18-25, 26-35, 36-45, 46-55, 56-60. (Categorical)', 'Age_Sort': 'Sort order of Age_Group. (Numerical)',
 'Tenure_Group': 'YearsAtCompany band: 0-1, 2-5, 6-10, 11-20, 21+ years. (Categorical)', 'Tenure_Sort': 'Sort order of Tenure_Group. (Numerical)',
 'Is_New_Hire': '1 if YearsAtCompany ≤ 1 (less than 2 years of tenure). (Numerical)',
 'Income_Band': 'MonthlyIncome band in USD: <3K, 3-5K, 5-10K, 10-15K, 15K+. (Categorical)', 'Income_Band_Sort': 'Sort order of Income_Band. (Numerical)',
 'Annual_Income': 'MonthlyIncome × 12, used for replacement cost. (Numerical)', 'Log_MonthlyIncome': 'Natural log of MonthlyIncome, used in the logistic model. (Numerical)',
 'Income_vs_Level_Ratio': 'MonthlyIncome ÷ median MonthlyIncome of the same JobLevel. (Numerical)', 'Underpaid_vs_Level': '1 if Income_vs_Level_Ratio < 0.85. (Numerical)',
 'Low_Income_Level1': '1 if JobLevel = 1 and MonthlyIncome < 3,000 USD. (Numerical)',
 'EnvironmentSatisfaction_Label': 'Text label of EnvironmentSatisfaction 1-4 (Thấp, Trung bình, Cao, Rất cao). (Categorical)',
 'JobSatisfaction_Label': 'Text label of JobSatisfaction 1-4. (Categorical)', 'RelationshipSatisfaction_Label': 'Text label of RelationshipSatisfaction 1-4. (Categorical)',
 'WorkLifeBalance_Label': 'Text label of WorkLifeBalance 1-4. (Categorical)', 'JobInvolvement_Label': 'Text label of JobInvolvement 1-4. (Categorical)',
 'Education_Label': 'Text label of Education 1-5 (Dưới cao đẳng ... Tiến sĩ). (Categorical)', 'PerformanceRating_Label': 'Text label of PerformanceRating (only 3 and 4 occur). (Categorical)',
 'JobLevel_Label': 'Level 1-5. (Categorical)', 'StockOption_Label': 'Stock option level 0-3 as text. (Categorical)',
 'Satisfaction_Index': 'Mean of the four satisfaction scales (environment, job, relationship, work-life balance). (Numerical)',
 'Low_Satisfaction_Count': 'Number of the four satisfaction scales at level 1 (Low). (Numerical)', 'Low_Sat_Label': 'Low_Satisfaction_Count as 0, 1, 2, 3-4 scales. (Categorical)',
 'Any_Low_Satisfaction': '1 if at least one satisfaction scale is at level 1. (Numerical)', 'Has_Stock': '1 if StockOptionLevel > 0. (Numerical)',
 'Is_Single': '1 if MaritalStatus = Single. (Numerical)', 'Is_Level1': '1 if JobLevel = 1. (Numerical)', 'Travel_Frequent': '1 if BusinessTravel = Travel_Frequently. (Numerical)',
 'Distance_Group': 'DistanceFromHome band: ≤ 5, 6-15, > 15 miles. (Categorical)', 'No_Training': '1 if TrainingTimesLastYear = 0. (Numerical)',
 'Mgr_Change_Label': 'Employees with ≥ 2 years and YearsWithCurrManager = 0: recent manager change; employees with < 2 years shown separately; others: stable manager. (Categorical)',
 'Mgr_Change_Flag': '1 if Mgr_Change_Label is a recent manager change. (Numerical)', 'Promotion_Stagnation': '1 if YearsSinceLastPromotion ≥ 5. (Numerical)',
 'Years_Before_Company': 'TotalWorkingYears - YearsAtCompany. (Numerical)',
 'Flag_NumCompanies_Inconsistent': '1 if NumCompaniesWorked = 0 but TotalWorkingYears > YearsAtCompany (197 rows, kept and flagged). (Numerical)',
 'Risk_Flags': 'Count of 5 risk flags: overtime, Level 1, single, frequent travel, ≥ 1 low satisfaction scale. (Numerical)',
 'Risk_Group': 'Low (0-1 flags), Medium (2-3), High (4-5). (Categorical)', 'Risk_Group_Sort': 'Sort order of Risk_Group. (Numerical)',
 'Status': 'Đã nghỉ (left) / Đang làm (current). (Categorical)',
}
for c in cols:
    key = ren_back.get(c, c)
    if c in derived: lines.append(('v', f'{c}: {derived[c]} [Derived]'))
    elif key in orig_desc: lines.append(('v', c + orig_desc[key][len(key):]))
    else: lines.append(('v', f'{c}: (original column)'))
lines += [('t', 'Cleaning Log')] + [('v', f'{r.Step} - {r.Action} ({r.Rows_Affected}): {r.Reason}') for r in log.itertuples()]
for i, (k, txt) in enumerate(lines, 1):
    cell = wd.cell(i, 1, txt)
    for a, v in st[k].items(): setattr(cell, a, v)
    wd.row_dimensions[i].height = None
out = '/home/user/Food-App-Business/Employee Attrition (sửa theo CMT)/02_Employee Attrition_Clean.xlsx'
wb.save(out); print('saved', ws.max_row, ws.max_column, wd.max_row)
