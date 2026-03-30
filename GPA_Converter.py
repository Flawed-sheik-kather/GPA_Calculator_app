import pandas as pd

Excel_sheet = "gpa_test_input.xlsx"
Output_sheet = "GPA_Output.xlsx"

def Calculating_Scholarly_gpa(Score):
    gpa = []
    for i in Score:
        if i >= 70:
            gpa.append(4)
        elif i >= 65:
            gpa.append(3.7)
        elif i >= 60:
            gpa.append(3.3)
        elif i >= 50:
            gpa.append(3)
        elif i >= 45:
            gpa.append(2.3)
        elif i >= 40:
            gpa.append(2)
        elif i < 40:
            gpa.append(0)
    return gpa   

def Final_Scholaro_US_gpa(Course_GPA, Credit_hours_func1):
    a = 0
    b = []
    sum_credit = sum(Credit_hours_func1)
    
    for i in range(len(Course_GPA)):
        b.append(Course_GPA[i] * Credit_hours_func1[i])
    a = sum(b)
    
    USFinal_Score = round((a / sum_credit), 3)
    return USFinal_Score   

def Uk_Cumulative_Score(Percent, Credit_hours_func2):
    a = 0
    b = []
    sum_credit = sum(Credit_hours_func2)
    
    for i in range(len(Percent)):
        b.append(Percent[i] * Credit_hours_func2[i])
    a = sum(b)
    
    UKFinal_Score = round((a / sum_credit), 2)
    return UKFinal_Score

# ===================== Excel Workbook =====================

df = pd.read_excel(Excel_sheet, sheet_name="GPA_Calculator")

# ---- Sheet 1: Course GPA ----
df["Course GPA"] = Calculating_Scholarly_gpa(df["Percentage"].tolist())

course_gpa_df = df[[
    "Student ID",
    "Course Code",
    "Credit Hours",
    "Course GPA"
]]

# ---- Sheet 2: Final GPA ----
final_rows = []

for student_id in df["Student ID"].unique():
    student_df = df[df["Student ID"] == student_id]

    final_rows.append({
        "Student ID": student_id,
        "Final US GPA": Final_Scholaro_US_gpa(
            student_df["Course GPA"].tolist(),
            student_df["Credit Hours"].tolist()
        ),
        "UK Cumulative": Uk_Cumulative_Score(
            student_df["Percentage"].tolist(),
            student_df["Credit Hours"].tolist()
        )
    })

final_gpa_df = pd.DataFrame(final_rows)

# ---- Write Excel workbook ----
with pd.ExcelWriter(Output_sheet, engine="openpyxl") as writer:
    course_gpa_df.to_excel(writer, sheet_name="Course_GPA", index=False)
    final_gpa_df.to_excel(writer, sheet_name="Final_GPA", index=False)

print("Excel workbook created successfully.")