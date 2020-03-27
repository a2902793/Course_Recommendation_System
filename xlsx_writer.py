import openpyxl as xl
import os, re
from openpyxl.utils import get_column_letter

wbook = xl.Workbook()
wsheet = wbook.active
wsheet['A1'] = 'ID'
wsheet['B1'] = '課程名稱'            # Course_Title
wsheet['C1'] = '授課教師'            # Course_Instructor
wsheet['D1'] = '開課系級(中)'        # Course_Class(Chinese)
wsheet['E1'] = '開課系級(英)'        # Course_Class(English)
wsheet['F1'] = '開課資料'            # Course_Details
wsheet['G1'] = '目標類型'            # Objective_Methods
wsheet['H1'] = '核心能力'            # Core_Competences
wsheet['I1'] = '基本素養'            # Essential_Virtues 
wsheet['J1'] = '教學方法'            # Teaching Methods
wsheet['K1'] = '評量方式'            # Assessment
wsheet['L1'] = '出席率'              # Grading_Attendance
wsheet['M1'] = '平時評量'            # Grading_MarkofUsual 
wsheet['N1'] = '期中評量'            # Grading_MidtermExam 
wsheet['O1'] = '期末評量'            # Grading_FinalExam 
wsheet['P1'] = '其他'                # Grading_Other
wsheet['Q1'] = '其他<>'              # Grading_Other<>

directory = './testset/xlsx'

for filename in sorted(os.listdir(directory)):
    data = []
    pdfpath=os.path.join(directory, filename) #./testset/ + 0001.pdf
    rwb = xl.load_workbook(pdfpath, data_only=True)
    rsheet = rwb.worksheets[0]
    data.append(os.path.splitext(filename)[0])   # 添加 ID
    for cell in rsheet['1']:
        if cell.value == "課程名稱":
            coordinate = "%s1" %get_column_letter(cell.column+1)
            #print("科目 = %s" %sheet[coordinate].value)
            data.append(rsheet[coordinate].value)   # 添加課程名稱
        if cell.value == "授課\n教師":
            coordinate = "%s1" %get_column_letter(cell.column+1)
            #print("授課教師 = %s" %sheet[coordinate].value)
            data.append(rsheet[coordinate].value)   # 添加授課老師

    for cell in rsheet['3']:
        if cell.value == "開課系級":
            coordinate = "%s3" %get_column_letter(cell.column+1)
            #print("科目 = %s" %sheet[coordinate].value)
            data.append(rsheet[coordinate].value)   # 添加開課系級(中)
            coordinate = "%s4" %get_column_letter(cell.column+1)
            data.append(rsheet[coordinate].value)   # 添加開課系級(英)
        if cell.value == "開課\n資料":
            coordinate = "%s3" %get_column_letter(cell.column+1)
            #print("授課教師 = %s" %sheet[coordinate].value)
            data.append(rsheet[coordinate].value)   # 添加開課資料

    for row in rsheet.iter_rows():
        for entry in row:
            try:
                if '院、系(所)\n核心能力' in entry.value:
                    coordinate = '{row}{column}'.format(row=get_column_letter(entry.column-1), column=entry.row+1)
                    data.append(rsheet[coordinate].value)   # 添加目標類型
                    coordinate = '{row}{column}'.format(row=get_column_letter(entry.column), column=entry.row+1)
                    data.append(rsheet[coordinate].value)   # 添加核心能力
                    coordinate = '{row}{column}'.format(row=get_column_letter(entry.column+1), column=entry.row+1)
                    data.append(rsheet[coordinate].value)   # 添加基本素養
                    coordinate = '{row}{column}'.format(row=get_column_letter(entry.column+2), column=entry.row+1)
                    data.append(rsheet[coordinate].value)   # 添加教學方法
                    coordinate = '{row}{column}'.format(row=get_column_letter(entry.column+3), column=entry.row+1)
                    data.append(rsheet[coordinate].value)   # 添加評量方式
            except (AttributeError, TypeError):
                continue
            try:
                if '出席率' in entry.value:
                    attendance = re.search(r"(?<=出席率：).*?(?=%)", entry.value)
                    #print("出席率 = %s" %attendance.group(0).strip())
                    data.append(attendance.group(0).strip())   # 添加出席率
            except (AttributeError, TypeError):
                continue
            try:
                if '平時評量' in entry.value:
                    attendance = re.search(r"(?<=平時評量：).*?(?=%)", entry.value)
                    #print("平時評量 = %s" %attendance.group(0))
                    data.append(attendance.group(0))   # 添加平時評量
            except (AttributeError, TypeError):
                continue
            try:
                if '期中評量' in entry.value:
                    midterm = re.search(r"(?<=期中評量：).*?(?=%)", entry.value)
                    #print("期中評量 = %s" %midterm.group(0))
                    data.append(midterm.group(0))   # 添加期中評量
            except (AttributeError, TypeError):
                continue
            try:
                if '期末評量' in entry.value:
                    final = re.search(r"(?<=期末評量：).*?(?=%)", entry.value)
                    #print("期末評量 = %s" %final.group(0))
                    data.append(final.group(0))   # 添加期末評量
            except (AttributeError, TypeError):
                continue
            try:
                if '其他〈' in entry.value:
                    other = re.search(r"(?<=其他〈).*?(?=〉：)", entry.value)
                    #print("其他〈〉 = %s" %other.group(0))
                    data.append(other.group(0))   # 添加其他
            except (AttributeError, TypeError):
                continue
            try:
                if '〉：' in entry.value:
                    other_num = re.search(r"(?<=〉：).*?(?=%)", entry.value)
                    #print("其他〈...〉 = %s" %other_num.group(0))
                    data.append(other_num.group(0))   # 添加其他<>
            except (AttributeError, TypeError):
                continue
    
    wsheet.append(data)

wbook.save("sample.xlsx")