import pandas as pd, numpy as np
raw = pd.read_excel('02_Employee Attrition.xlsx', sheet_name=0)
print(raw.shape, raw.columns.tolist()[:5])
d = raw.copy()
d.columns = [c.replace('\n', ' ').strip() for c in d.columns]
ren = {c: ''.join(w[:1].upper() + w[1:] for w in c.replace('/', ' ').split()) for c in d.columns}
ren['Employee Number'] = 'EmployeeID'; ren['Years With Current Manager'] = 'YearsWithCurrManager'
d = d.rename(columns=ren)
drop = ['Over18', 'StandardHours', 'DailyRate', 'HourlyRate', 'MonthlyRate']
log = [('R0', 'Kiểm tra', len(d), f'Dữ liệu gốc: {len(raw)} dòng × {raw.shape[1]} cột'),
       ('R1', 'Kiểm tra', int(d.isna().sum().sum()), 'Ô thiếu giá trị'),
       ('R1', 'Kiểm tra', int(d.duplicated().sum()), 'Dòng trùng hoàn toàn'),
       ('R1', 'Kiểm tra', int(d.EmployeeID.duplicated().sum()), 'Mã nhân viên trùng')]
d = d.drop(columns=drop)
log.append(('R2', 'Loại cột', 5, 'Bỏ Over18, StandardHours (hằng số) và DailyRate, HourlyRate, MonthlyRate (|r| ≤ 0,035 với MonthlyIncome)'))
n_base = d.shape[1]

lab4 = {1: 'Thấp', 2: 'Trung bình', 3: 'Cao', 4: 'Rất cao'}
y = (d.Attrition == 'Yes').astype(int)
d['Attrition_Flag'] = y
d['OverTime_Flag'] = (d.OverTime == 'Yes').astype(int)
d['Age_Group'] = pd.cut(d.Age, [17, 25, 35, 45, 55, 60], labels=['18-25', '26-35', '36-45', '46-55', '56-60']).astype(str)
d['Age_Sort'] = d.Age_Group.map({'18-25': 1, '26-35': 2, '36-45': 3, '46-55': 4, '56-60': 5})
d['Tenure_Group'] = pd.cut(d.YearsAtCompany, [-1, 1, 5, 10, 20, 100], labels=['0-1', '2-5', '6-10', '11-20', '21+']).astype(str)
d['Tenure_Sort'] = d.Tenure_Group.map({'0-1': 1, '2-5': 2, '6-10': 3, '11-20': 4, '21+': 5})
d['Is_New_Hire'] = (d.YearsAtCompany <= 1).astype(int)
d['Income_Band'] = pd.cut(d.MonthlyIncome, [0, 2999, 4999, 9999, 14999, 10**6], labels=['<3K', '3-5K', '5-10K', '10-15K', '15K+']).astype(str)
d['Income_Band_Sort'] = d.Income_Band.map({'<3K': 1, '3-5K': 2, '5-10K': 3, '10-15K': 4, '15K+': 5})
d['Annual_Income'] = d.MonthlyIncome * 12
d['Log_MonthlyIncome'] = np.log(d.MonthlyIncome).round(4)
med = d.groupby('JobLevel').MonthlyIncome.transform('median')
d['Income_vs_Level_Ratio'] = (d.MonthlyIncome / med).round(4)
d['Underpaid_vs_Level'] = (d.Income_vs_Level_Ratio < 0.85).astype(int)
d['Low_Income_Level1'] = ((d.JobLevel == 1) & (d.MonthlyIncome < 3000)).astype(int)
for c in ['EnvironmentSatisfaction', 'JobSatisfaction', 'RelationshipSatisfaction', 'WorkLifeBalance', 'JobInvolvement']:
    d[c + '_Label'] = d[c].map(lab4)
d['Education_Label'] = d.Education.map({1: 'Dưới cao đẳng', 2: 'Cao đẳng', 3: 'Cử nhân', 4: 'Thạc sĩ', 5: 'Tiến sĩ'})
d['PerformanceRating_Label'] = d.PerformanceRating.map({3: 'Tốt (3)', 4: 'Xuất sắc (4)'})
d['JobLevel_Label'] = 'Level ' + d.JobLevel.astype(str)
d['StockOption_Label'] = 'Mức ' + d.StockOptionLevel.astype(str)
sat4 = ['EnvironmentSatisfaction', 'JobSatisfaction', 'RelationshipSatisfaction', 'WorkLifeBalance']
d['Satisfaction_Index'] = d[sat4].mean(axis=1).round(2)
d['Low_Satisfaction_Count'] = (d[sat4] == 1).sum(axis=1)
d['Low_Sat_Label'] = d.Low_Satisfaction_Count.clip(upper=3).map({0: '0 thang', 1: '1 thang', 2: '2 thang', 3: '3-4 thang'})
d['Any_Low_Satisfaction'] = (d.Low_Satisfaction_Count >= 1).astype(int)
d['Has_Stock'] = (d.StockOptionLevel > 0).astype(int)
d['Is_Single'] = (d.MaritalStatus == 'Single').astype(int)
d['Is_Level1'] = (d.JobLevel == 1).astype(int)
d['Travel_Frequent'] = (d.BusinessTravel == 'Travel_Frequently').astype(int)
d['Distance_Group'] = pd.cut(d.DistanceFromHome, [0, 5, 15, 100], labels=['≤ 5 dặm', '6-15 dặm', '> 15 dặm']).astype(str)
d['No_Training'] = (d.TrainingTimesLastYear == 0).astype(int)
d['Mgr_Change_Label'] = np.select([d.YearsAtCompany < 2, d.YearsWithCurrManager == 0],
                                  ['Mới vào (< 2 năm)', 'Vừa đổi quản lý'], 'Quản lý ổn định')
d['Mgr_Change_Flag'] = (d.Mgr_Change_Label == 'Vừa đổi quản lý').astype(int)
d['Promotion_Stagnation'] = (d.YearsSinceLastPromotion >= 5).astype(int)
d['Years_Before_Company'] = d.TotalWorkingYears - d.YearsAtCompany
d['Flag_NumCompanies_Inconsistent'] = ((d.NumCompaniesWorked == 0) & (d.TotalWorkingYears > d.YearsAtCompany)).astype(int)
d['Risk_Flags'] = d[['OverTime_Flag', 'Is_Level1', 'Is_Single', 'Travel_Frequent', 'Any_Low_Satisfaction']].sum(axis=1)
d['Risk_Group'] = pd.cut(d.Risk_Flags, [-1, 1, 3, 5], labels=['Thấp (0-1 cờ)', 'Trung bình (2-3 cờ)', 'Cao (4-5 cờ)']).astype(str)
d['Risk_Group_Sort'] = d.Risk_Group.str[0].map({'T': 0}).fillna(0)
d['Risk_Group_Sort'] = d.Risk_Group.map({'Thấp (0-1 cờ)': 1, 'Trung bình (2-3 cờ)': 2, 'Cao (4-5 cờ)': 3})
d['Status'] = np.where(y == 1, 'Đã nghỉ', 'Đang làm')
n_new = d.shape[1] - n_base
log += [('R3', 'Gắn cờ', int(d.Flag_NumCompanies_Inconsistent.sum()), 'NumCompaniesWorked = 0 nhưng TotalWorkingYears > YearsAtCompany (chênh đúng 1 năm ở mọi trường hợp); giữ lại, gắn cờ'),
        ('R4', 'Chuẩn hóa', 0, 'Tên cột về PascalCase; Employee Number → EmployeeID'),
        ('R5', 'Tạo biến', n_new, f'{n_new} cột dẫn xuất (nhóm, nhãn, cờ rủi ro)'),
        ('R6', 'Kiểm tra', len(d), f'Dữ liệu sạch: {len(d)} dòng × {d.shape[1]} cột')]
print('base', n_base, 'new', n_new, 'total', d.shape)
d.to_pickle('./emp_final.pkl')
pd.DataFrame(log, columns=['Step', 'Action', 'Rows_Affected', 'Reason']).to_pickle('./log.pkl')
# ---- đối chiếu với báo cáo
r = lambda m: round(d[m].mean() * 100, 1) if m in d else None
g = lambda col: (d.groupby(col).Attrition_Flag.agg(['size', 'mean']) .assign(mean=lambda t: (t['mean'] * 100).round(1)))
print(g('Age_Group').loc['18-25'].tolist(), g('Tenure_Group').loc[['0-1', '11-20']].values.tolist())
print(g('Mgr_Change_Label').values.tolist(), g('Mgr_Change_Label').index.tolist())
print(g('Risk_Group').values.tolist(), g('Risk_Group').index.tolist(), d[d.Attrition_Flag == 0].groupby('Risk_Group').size().to_dict())
print('underpaid', g('Underpaid_vs_Level').values.tolist(), 'lowsat', g('Low_Sat_Label').values.tolist())
print('L1<3K', d[d.JobLevel == 1].groupby('Low_Income_Level1').Attrition_Flag.mean().round(3).tolist(), 'notrain', g('No_Training').values.tolist())
print('flag', d.Flag_NumCompanies_Inconsistent.sum(), 'annual income leavers', d.loc[y == 1, 'Annual_Income'].sum())
