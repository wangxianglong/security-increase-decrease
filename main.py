import tkinter as tk
from tkinter import filedialog,ttk,messagebox
import os
import re
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment,PatternFill,Border, Side,Font
from openpyxl.worksheet.dimensions import ColumnDimension,DimensionHolder
from openpyxl.utils import get_column_letter

root = tk.Tk()
root.geometry("580x350+50+50") # widthxheight+x+y
root.title("社保公积金增员减员生成器")
root.resizable(False,False)

xinyue_path = tk.StringVar() # 馨悦-供应链项目花名册
tianan_path = tk.StringVar() # 天安-供应链项目花名册


def select_file_xinyue():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    xinyue_path.set(selected_file_path)
    if len(xinyue_path.get()) > 0 and len(tianan_path.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("馨悦-供应链项目花名册是", f'{selected_file_path}\n\n如果文件不正确可重新选择')

def select_file_tianan():
    selected_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    tianan_path.set(selected_file_path)
    if len(xinyue_path.get()) > 0 and len(tianan_path.get()) > 0:
        button1.config(state=tk.ACTIVE)
    messagebox.showwarning("天安-供应链项目花名册是", f'{selected_file_path}\n\n如果文件不正确可重新选择')


wb_anhui_accident_insurance = None #安徽和众意外险Excel
wb_anhui_increase_decrease = None #安徽和众增减员申报Excel
wb_anhui_declaration = None #安徽和众社保公积金申报Excel

wb_xinyue_accident_insurance = None #馨悦意外险Excel
wb_xinyue_increase_decrease = None #馨悦增减员申报Excel
wb_xinyue_declaration = None #馨悦社保公积金申报Excel

wb_tianan_accident_insurance = None #天安意外险Excel
wb_tianan_increase_decrease = None #天安增减员申报Excel
wb_tianan_declaration = None #天安社保公积金申报Excel


def validate_date(date_string):
    date_pattern = r'\d{4}/\d{1,2}/\d{1,2}'
    if re.match(date_pattern, date_string):
        return True
    else:
        return False


# 生成调休加班明细表
def generate_excel():
    try:
        selected_year = year_combo.get()
        selected_month = month_combo.get().replace('月','')
        # print(get_days_of_month(int(selected_year), int(selected_month)))
      

        wb_roster_xinyue = load_workbook(filename = xinyue_path.get(),read_only = True,data_only=True) #读取馨悦-供应链项目花名册Excel
        sheet_inservice_xinyue = wb_roster_xinyue.get_sheet_by_name("花名册在职模板") #读取在职人员
        for row_index in range(1,sheet_inservice_xinyue.max_row + 1):
            join_date = sheet_inservice_xinyue.cell(row=row_index,column=10).value #入司日期
            if validate_date(join_date) is True:
                # 判断是否本年本月
                join_date_array = join_date.split("/")
                if int(selected_year) == int(join_date_array[0]) and int(selected_month) == int(join_date_array[1]):
                    # 需要买意外险
                    company_name = sheet_inservice_xinyue.cell(row=row_index,column=39).value # 签订合同主体单位名称
                    file_name = f"{company_name}{selected_year}年{selected_month}月 意外险.xlsx"
                    template_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), '意外险.xlsx')
                    if company_name == '广州馨悦商务服务有限公司':
                        if wb_xinyue_accident_insurance is None:
                            wb_xinyue_accident_insurance = load_workbook(template_path)
                            wb_xinyue_accident_insurance.save(file_name)
                            wb_xinyue_accident_insurance.close()
                            wb_xinyue_accident_insurance = load_workbook(file_name)
                        
                    elif company_name == '安徽和众企业服务有限公司':
                        if wb_anhui_accident_insurance is None:
                            wb_anhui_accident_insurance = load_workbook(template_path)
                            wb_anhui_accident_insurance.save(file_name)
                            wb_anhui_accident_insurance.close()
                            wb_anhui_accident_insurance = load_workbook(file_name) 

                    if int(join_date_array[2]) <= 15: # 社保增员
                        ...


            regularization_date = sheet_inservice_xinyue.cell(row=row_index,column=14).value #实际转正时间
            if validate_date(regularization_date) is True:
                # 判断是否本年本月
                regularization_date_array = regularization_date.split("/")
                if int(selected_year) == int(regularization_date_array[0]) and int(selected_month) == int(regularization_date_array[1]):
                    if int(regularization_date_array[2]) <= 15: # 公积金增员
                        ...


        global file_name
        # file_name = f"{content}.xlsx"
        # workbook.save(file_name)
        
        messagebox.showinfo("提示", f"生成文件【{file_name}】成功")
     
    except Exception as e:
        print(e)
        messagebox.showerror("错误", "生成文件失败，请检查选择的文件内容是否正确!原因：" + repr(e))

    finally:
        workbook.close()
       


if __name__ == '__main__':
   
    years = list(range(2024,2055))
    year_combo = ttk.Combobox(root, values=years,width=5)
    year_combo.current(0)  # 设置默认选择为"一月"
    year_combo.configure(state="readonly")
    year_combo.grid(column=0, row=0)

    months = ["1月", "2月", "3月", "4月", "5月", "6月",
          "7月", "8月", "9月", "10月", "11月", "12月"]
    month_combo = ttk.Combobox(root, values=months,width=5)
    month_combo.current(0)  # 设置默认选择为"一月"
    month_combo.configure(state="readonly")
    month_combo.grid(column=1, row=0)
    
   
    entry = tk.Entry(root, textvariable=xinyue_path,width=45)
    entry.grid(column=0, row=1)
    entry.configure(state="readonly")
    tk.Button(root, text="选择馨悦-供应链项目花名册", command=select_file_xinyue).grid(row=1, column=1)

    entry = tk.Entry(root, textvariable=tianan_path,width=45)
    entry.grid(column=0, row=2)
    entry.configure(state="readonly")
    tk.Button(root, text="选择天安-供应链项目花名册", command=select_file_tianan).grid(row=2, column=1)


    button1 = tk.Button(root, text="1、生成社保公积金增减员文件",command=generate_excel)
    button1.grid(row=5, column=1, pady=20)  # 使Button在row=1, column=1的位置
    button1.config(state=tk.DISABLED)
  

    root.mainloop()
 

