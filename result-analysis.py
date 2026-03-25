import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import warnings

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def generate_report():

    file_path = filedialog.askopenfilename(
        title="Select Result Sheet (Excel File)",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_path:
        return

    try:

        df = pd.read_excel(file_path, header=[3, 4])

        df.columns = [' '.join(col).strip() for col in df.columns.values]

        df.columns = [c.replace('Unnamed: 0_level_1', '')
                        .replace('Unnamed: 1_level_1', '')
                        .replace('Unnamed: 2_level_1', '')
                        .strip() for c in df.columns]

        df = df[df['USN'].notna()]

        
        for col in df.columns:
            if any(x in col for x in ["IA", "Ext", "TOT", "SGPA"]):
                df[col] = pd.to_numeric(df[col], errors='coerce')

        sgpa_col = [c for c in df.columns if "SGPA" in c][0]
        result_col = [c for c in df.columns if "Result" in c][0]
        name_col = [c for c in df.columns if "Name of the Student" in c][0]
        usn_col = [c for c in df.columns if "USN" in c][0]

        bins = [0, 4, 5, 6, 7, 8, 9, 10]
        labels = ['0-4', '4-5', '5-6', '6-7', '7-8', '8-9', '9+']

        df['SGPA_RANGE'] = pd.cut(df[sgpa_col], bins=bins, labels=labels)

        sgpa_df = df['SGPA_RANGE'].value_counts().sort_index().reset_index()
        sgpa_df.columns = ["SGPA Range", "Students"]

        pf_columns = [col for col in df.columns if "P/F" in col]

        subject_rows = []

        for col in pf_columns:
            subject = col.split()[0]

            passes = (df[col] == 'P').sum()
            fails = (df[col] == 'F').sum()

            pass_percent = round((passes / len(df)) * 100, 2)

            subject_rows.append([subject, passes, fails, pass_percent])

        subject_df = pd.DataFrame(
            subject_rows,
            columns=["Subject", "Pass Count", "Fail Count", "Pass %"]
        )

    
        difficulty_df = subject_df.copy()
        difficulty_df["Difficulty Index %"] = (
            difficulty_df["Fail Count"] / len(df) * 100
        ).round(2)

        difficulty_df = difficulty_df[
            ["Subject", "Fail Count", "Difficulty Index %"]
        ]

       
        failure_rows = []

        for col in pf_columns:

            subject = col.split()[0]

            failed = df[df[col] == 'F']

            for _, row in failed.iterrows():

                failure_rows.append([
                    subject,
                    row[usn_col],
                    row[name_col]
                ])

        failure_df = pd.DataFrame(
            failure_rows,
            columns=["Subject", "USN", "Student Name"]
        )

       
        subjects = sorted(set(col.split()[0]
                       for col in df.columns if "IA" in col))

        high_ia_rows = []

        for sub in subjects:

            ia_col = f"{sub} IA"
            ext_col = f"{sub} Ext"
            pf_col = f"{sub} P/F"

            filtered = df[(df[ia_col] >= 45) &
                          (df[pf_col] == 'F')]

            for _, row in filtered.iterrows():

                high_ia_rows.append([
                    sub,
                    row[usn_col],
                    row[name_col],
                    row[ia_col],
                    row[ext_col]
                ])

        high_ia_df = pd.DataFrame(
            high_ia_rows,
            columns=["Subject", "USN",
                     "Student Name", "IA Marks", "SEE Marks"]
        )

       
        grade_list = ['O', 'A+', 'A', 'B+', 'B',
                      'C', 'P', 'F', 'DX']

        gl_columns = [col for col in df.columns if "GL" in col]

        grade_rows = []

        for col in gl_columns:

            subject = col.split()[0]

            counts = df[col].value_counts()

            row = {"Subject": subject}

            for grade in grade_list:
                row[grade] = counts.get(grade, 0)

            grade_rows.append(row)

        grade_df = pd.DataFrame(grade_rows)
    

        cie_rows = []

        for sub in subjects:

            cie_rows.append([
                sub,
                round(df[f"{sub} IA"].mean(), 2),
                round(df[f"{sub} Ext"].mean(), 2)
            ])

        cie_df = pd.DataFrame(
            cie_rows,
            columns=["Subject", "Average CIE", "Average SEE"]
        )

        toppers_df = df[[usn_col,
                         name_col,
                         sgpa_col]].sort_values(
            by=sgpa_col,
            ascending=False
        ).head(10)

        toppers_df.columns = ["USN", "Name", "SGPA"]



        student_failure_summary = []

        for _, row in df.iterrows():

            failed_subjects = []

            for col in pf_columns:

                if row[col] == 'F':

                    subject_code = col.split()[0]
                    failed_subjects.append(subject_code)

            if failed_subjects:
                student_failure_summary.append([
                row[usn_col],
                row[name_col],
                len(failed_subjects),
                ", ".join(failed_subjects)
            ])

        student_failure_summary_df = pd.DataFrame(
            student_failure_summary,
            columns=[
                "USN",
                "Student Name",
                "Number of Failed Subjects",
                "Failed Course Codes"
            ]
        )


        output = "Result-Analysis-Report.xlsx"

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            sgpa_df.to_excel(writer, "SGPA Distribution", index=False)
            subject_df.to_excel(writer, "Subject Analysis", index=False)
            difficulty_df.to_excel(writer, "Subject Difficulty", index=False)
            failure_df.to_excel(writer, "Failure Student List", index=False)
            high_ia_df.to_excel(writer, "High IA Failed SEE", index=False)
            cie_df.to_excel(writer, "CIE vs SEE", index=False)
            grade_df.to_excel(writer, "Grade Distribution", index=False)
            toppers_df.to_excel(writer, "Top 10 Toppers", index=False)
            student_failure_summary_df.to_excel(writer, "Student Failure Summary", index=False)

            workbook = writer.book

          
            worksheet1 = writer.sheets["SGPA Distribution"]

            chart1 = workbook.add_chart({'type': 'column'})

            rows = len(sgpa_df) + 1

            chart1.add_series({
                'name': 'SGPA Distribution',
                'categories': f'=SGPA Distribution!$A$2:$A${rows}',
                'values': f'=SGPA Distribution!$B$2:$B${rows}',
            })

            chart1.set_title({'name': 'SGPA Distribution'})
            chart1.set_x_axis({'name': 'SGPA Range'})
            chart1.set_y_axis({'name': 'Students'})

            worksheet1.insert_chart('E2', chart1)


          
            worksheet2 = writer.sheets["Subject Analysis"]

            chart2 = workbook.add_chart({'type': 'column'})

            rows2 = len(subject_df) + 1

            chart2.add_series({
                'name': 'Pass %',
                'categories': f'=Subject Analysis!$A$2:$A${rows2}',
                'values': f'=Subject Analysis!$D$2:$D${rows2}',
            })

            chart2.set_title({'name': 'Subject Pass Percentage'})

            worksheet2.insert_chart('G2', chart2)


           
            worksheet3 = writer.sheets["CIE vs SEE"]

            chart3 = workbook.add_chart({'type': 'column'})

            rows3 = len(cie_df) + 1

            chart3.add_series({
                'name': 'Average CIE',
                'categories': f'=CIE vs SEE!$A$2:$A${rows3}',
                'values': f'=CIE vs SEE!$B$2:$B${rows3}',
            })

            chart3.add_series({
                'name': 'Average SEE',
                'categories': f'=CIE vs SEE!$A$2:$A${rows3}',
                'values': f'=CIE vs SEE!$C$2:$C${rows3}',
            })

            chart3.set_title({'name': 'CIE vs SEE Comparison'})
            chart3.set_x_axis({'name': 'Subjects'})
            chart3.set_y_axis({'name': 'Marks'})

            worksheet3.insert_chart('E2', chart3)

 
            result_summary = df[result_col].value_counts().reset_index()
            result_summary.columns = ["Result", "Students"]

            result_summary.to_excel(writer, "Overall Result", index=False)

            worksheet4 = writer.sheets["Overall Result"]

            chart4 = workbook.add_chart({'type': 'pie'})

            rows4 = len(result_summary) + 1

            chart4.add_series({
                'name': 'Overall Result',
                'categories': f'=Overall Result!$A$2:$A${rows4}',
                'values': f'=Overall Result!$B$2:$B${rows4}',
                'data_labels': {'percentage': True},
            })

            chart4.set_title({'name': 'Overall Result Distribution'})

            worksheet4.insert_chart('E2', chart4)

     
            worksheet5 = writer.sheets["Grade Distribution"]

            chart5 = workbook.add_chart({'type': 'column'})

            rows5 = len(grade_df) + 1
            grade_list = ['O','A+','A','B+','B','C','P','F','DX']

            for i, grade in enumerate(grade_list):
                col_letter = chr(66 + i)

                chart5.add_series({
                    'name': grade,
                    'categories': f'=Grade Distribution!$A$2:$A${rows5}',
                    'values': f'=Grade Distribution!${col_letter}$2:${col_letter}${rows5}',
                })

            chart5.set_title({'name': 'Grade Distribution by Subject'})
            chart5.set_x_axis({'name': 'Subjects'})
            chart5.set_y_axis({'name': 'Students'})

            worksheet5.insert_chart('L2', chart5)

        messagebox.showinfo(
            "Success",
            "Result Analysis Report Generated Successfully!"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Result Analysis Tool")
root.geometry("600x400")

try:
    root.iconbitmap(resource_path("foss-logo.ico"))
except:
    pass

# Load logo AFTER root created
logo = None

try:
    logo_img = Image.open(resource_path("foss-logo.png"))
    logo_img = logo_img.resize((170, 170))
    logo = ImageTk.PhotoImage(logo_img)
except:
    pass

title_label = tk.Label(
    root,
    text="Automatic Result Analysis Tool by FOSS Club, PESCE, Mandya",
    font=("Arial", 14)
)

title_label.pack(pady=20)

if logo:
    logo_label = tk.Label(root, image=logo)
    logo_label.pack()

root.logo = logo

btn = tk.Button(
    root,
    text="Select Result Sheet (Excel File)",
    command=generate_report,
    width=28,
    height=2
)

btn.pack(pady=25)

root.mainloop()
