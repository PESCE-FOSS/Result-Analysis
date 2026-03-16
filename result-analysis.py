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

logo_path = resource_path("foss-logo.png")
logo_img = Image.open(logo_path)


warnings.filterwarnings("ignore")

def generate_report():

    file_path = filedialog.askopenfilename(
        title="Select Result Sheet (Excel File)",
        filetypes=[("Excel Files","*.xlsx")]
    )

    if not file_path:
        return

    try:

        
        data_frame = pd.read_excel(file_path, header=[3,4])   #Read Excel File

        data_frame.columns = [' '.join(col).strip() for col in data_frame.columns.values]             #Pre-Processing

        data_frame.columns = [c.replace('Unnamed: 0_level_1','')
                       .replace('Unnamed: 1_level_1','')
                       .replace('Unnamed: 2_level_1','')
                       .strip() for c in data_frame.columns]

        data_frame = data_frame[data_frame['USN'].notna()]

        
        for col in data_frame.columns:
            if any(x in col for x in ["IA","Ext","TOT","SGPA"]):
                data_frame[col] = pd.to_numeric(data_frame[col], errors='coerce')

        sgpa_col = [c for c in data_frame.columns if "SGPA" in c][0]
        result_col = [c for c in data_frame.columns if "Result" in c][0]
        name_col = [c for c in data_frame.columns if "Name of the Student" in c][0]
        usn_col = [c for c in data_frame.columns if "USN" in c][0]

        

        bins=[0,4,5,6,7,8,9,10]                             # SGPA DISTRIBUTION
        labels=['0-4','4-5','5-6','6-7','7-8','8-9','9+']

        data_frame['SGPA_RANGE']=pd.cut(data_frame[sgpa_col],bins=bins,labels=labels)

        sgpa_df=data_frame['SGPA_RANGE'].value_counts().sort_index().reset_index()
        sgpa_df.columns=["SGPA Range","Students"]

        

        pf_columns=[col for col in data_frame.columns if "P/F" in col]                # Pass / Fail

        subject_rows=[]

        for col in pf_columns:

            subject=col.split()[0]

            passes=(data_frame[col]=='P').sum()
            fails=(data_frame[col]=='F').sum()

            pass_percent=round((passes/len(data_frame))*100,2)

            subject_rows.append([subject,passes,fails,pass_percent])

        subject_df=pd.DataFrame(
            subject_rows,
            columns=["Subject","Pass Count","Fail Count","Pass %"]
        )

        

        difficulty_rows=[]                            #for storing difficulty index

        for col in pf_columns:

            subject=col.split()[0]
            fails=(data_frame[col]=='F').sum()

            difficulty=round((fails/len(data_frame))*100,2)

            difficulty_rows.append([subject,fails,difficulty])

        difficulty_df=pd.DataFrame(
            difficulty_rows,
            columns=["Subject","Failure Count","Difficulty Index %"]
        )

        

        failure_rows=[]                          #Store Failure Student List

        for col in pf_columns:

            subject=col.split()[0]

            failed=data_frame[data_frame[col]=='F']

            for _,row in failed.iterrows():

                failure_rows.append([
                    subject,
                    row[usn_col],
                    row[name_col]
                ])

        failure_df=pd.DataFrame(
            failure_rows,
            columns=["Subject","USN","Student Name"]
        )

        

        student_fail_rows=[]
 
        for _,row in data_frame.iterrows():

            fails=[]

            for col in pf_columns:

                if row[col]=='F':
                    fails.append(col.split()[0])

            if len(fails)>0:

                student_fail_rows.append([
                    row[usn_col],
                    row[name_col],
                    len(fails),
                    ", ".join(fails)
                ])

        student_fail_df=pd.DataFrame(
            student_fail_rows,
            columns=[
                "USN",
                "Student Name",
                "Number of Failed Subjects",
                "Failed Subjects"
            ]
        )

        
        high_ia_rows=[]

        subjects=sorted(set(col.split()[0] for col in data_frame.columns if "IA" in col))

        for sub in subjects:

            ia_col=f"{sub} IA"
            ext_col=f"{sub} Ext"
            pf_col=f"{sub} P/F"

            if ia_col in data_frame.columns and pf_col in data_frame.columns:

                filtered=data_frame[(data_frame[ia_col]>=45) & (data_frame[pf_col]=='F')]

                for _,row in filtered.iterrows():

                    high_ia_rows.append([
                        sub,
                        row[usn_col],
                        row[name_col],
                        row[ia_col],
                        row[ext_col]
                    ])

        high_ia_df=pd.DataFrame(
            high_ia_rows,
            columns=["Subject","USN","Student Name","IA Marks","SEE Marks"]
        )

       

        grace_rows=[]

        for sub in subjects:

            ext_col=f"{sub} Ext"

            filtered=data_frame[(data_frame[ext_col]>=30) & (data_frame[ext_col]<35)]

            for _,row in filtered.iterrows():

                grace_rows.append([
                    sub,
                    row[usn_col],
                    row[name_col],
                    row[ext_col]
                ])

        grace_df=pd.DataFrame(
            grace_rows,
            columns=["Subject","USN","Student Name","SEE Marks"]
        )

       
        cie_rows=[]

        for sub in subjects:

            ia_col=f"{sub} IA"
            ext_col=f"{sub} Ext"

            cie_rows.append([
                sub,
                round(data_frame[ia_col].mean(),2),
                round(data_frame[ext_col].mean(),2)
            ])

        cie_df=pd.DataFrame(
            cie_rows,
            columns=["Subject","Average CIE","Average SEE"]
        )

       

        toppers_df=data_frame[[usn_col,name_col,sgpa_col]]\
            .sort_values(by=sgpa_col,ascending=False)\
            .head(10)

        toppers_df.columns=["USN","Name","SGPA"]


       
        grade_list = ['O','A+','A','B+','B','C','P','F','DX']

        gl_columns = [col for col in data_frame.columns if "GL" in col]

        grade_rows = []

        for col in gl_columns:

            subject = col.split()[0]

            grade_counts = data_frame[col].value_counts()

            row = {"Subject":subject}

            for grade in grade_list:
                row[grade] = grade_counts.get(grade,0)

            grade_rows.append(row)

        grade_df = pd.DataFrame(grade_rows)



      
        output="Result-Analysis-Report.xlsx"

        with pd.ExcelWriter(output,engine="xlsxwriter") as writer:

            sgpa_df.to_excel(writer,"SGPA Distribution",index=False)
            subject_df.to_excel(writer,"Subject Analysis",index=False)
            difficulty_df.to_excel(writer,"Subject Difficulty",index=False)
            failure_df.to_excel(writer,"Failure Student List",index=False)
            student_fail_df.to_excel(writer,"Student Failure Analysis",index=False)
            high_ia_df.to_excel(writer,"High IA Failed SEE",index=False)
            grace_df.to_excel(writer,"Grace Range Students",index=False)
            cie_df.to_excel(writer,"CIE vs SEE",index=False)
            toppers_df.to_excel(writer,"Top 10 Toppers",index=False)
            grade_df.to_excel(writer,"Grade Distribution",index=False)

            workbook = writer.book

            

            worksheet1 = writer.sheets["SGPA Distribution"]

            chart1 = workbook.add_chart({'type':'column'})

            rows = len(sgpa_df)+1

            chart1.add_series({
                'name':'SGPA Distribution',
                'categories':f'=SGPA Distribution!$A$2:$A${rows}',
                'values':f'=SGPA Distribution!$B$2:$B${rows}'
            })

            chart1.set_title({'name':'SGPA Distribution'})
            chart1.set_x_axis({'name':'SGPA Range'})
            chart1.set_y_axis({'name':'Students'})

            worksheet1.insert_chart('E2', chart1)

            

            worksheet2 = writer.sheets["Subject Analysis"]

            chart2 = workbook.add_chart({'type':'column'})

            rows2 = len(subject_df)+1

            chart2.add_series({
                'name':'Pass %',
                'categories':f'=Subject Analysis!$A$2:$A${rows2}',
                'values':f'=Subject Analysis!$D$2:$D${rows2}'
            })

            chart2.set_title({'name':'Subject Pass Percentage'})

            worksheet2.insert_chart('G2', chart2)

           
            worksheet3 = writer.sheets["CIE vs SEE"]

            chart3 = workbook.add_chart({'type':'column'})

            rows3 = len(cie_df)+1

            chart3.add_series({
                'name':'Average CIE',
                'categories':f'=CIE vs SEE!$A$2:$A${rows3}',
                'values':f'=CIE vs SEE!$B$2:$B${rows3}'
            })

            chart3.add_series({
                'name':'Average SEE',
                'categories':f'=CIE vs SEE!$A$2:$A${rows3}',
                'values':f'=CIE vs SEE!$C$2:$C${rows3}'
            })

            chart3.set_title({'name':'CIE vs SEE Comparison'})
            chart3.set_x_axis({'name':'Subjects'})
            chart3.set_y_axis({'name':'Marks'})

            worksheet3.insert_chart('E2', chart3)

           
            result_summary = data_frame[result_col].value_counts().reset_index()
            result_summary.columns=["Result","Students"]

            result_summary.to_excel(writer,"Overall Result",index=False)

            worksheet4 = writer.sheets["Overall Result"]

            chart4 = workbook.add_chart({'type':'pie'})

            rows4 = len(result_summary)+1

            chart4.add_series({
                'name':'Overall Result',
                'categories':f'=Overall Result!$A$2:$A${rows4}',
                'values':f'=Overall Result!$B$2:$B${rows4}',
                'data_labels':{'percentage':True}
            })

            chart4.set_title({'name':'Overall Result Distribution'})

            worksheet4.insert_chart('E2', chart4)


           
            worksheet5 = writer.sheets["Grade Distribution"]

            chart5 = workbook.add_chart({'type':'column'})

            rows5 = len(grade_df) + 1

            for i,grade in enumerate(grade_list):

                col_letter = chr(66 + i)   # B,C,D...

                chart5.add_series({
                    'name': grade,
                    'categories': f'=Grade Distribution!$A$2:$A${rows5}',
                    'values': f'=Grade Distribution!${col_letter}$2:${col_letter}${rows5}'
                })

            chart5.set_title({'name':'Grade Distribution by Subject'})
            chart5.set_x_axis({'name':'Subjects'})
            chart5.set_y_axis({'name':'Number of Students'})

            worksheet5.insert_chart('L2', chart5)

        messagebox.showinfo("Success","Result Analysis Report Generated!")

    except Exception as e:
        messagebox.showerror("Error",str(e))



root=tk.Tk()
root.title("Result Analysis")
root.iconbitmap(resource_path("foss-logo.ico"))
root.geometry("600x400")

label=tk.Label(root,text="Automatic Result Analysis Tool by FOSS Club, PESCE, Mandya",font=("Arial",14))
label.pack(pady=20)

#logo_img = Image.open(resource_path("foss-logo.png"))
logo_img = logo_img.resize((170,170))

logo = ImageTk.PhotoImage(logo_img)

logo_label = tk.Label(root,image=logo)
logo_label.pack(pady=10)


btn=tk.Button(root,text="Select Result Sheet (Excel File)",
              command=generate_report,
              width=25,height=2)

btn.pack()

root.mainloop()