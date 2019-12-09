Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks
Dim n,a 
Dim b
b="2"

objExcel.Application.Visible = True
objExcel.Workbooks.Add
objExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 4).Value = "AGE"

Do
n = InputBox("ENTER NAME")
a = InputBox("ENTER AGE")
objExcel.Cells(b, 3).Value = n
objExcel.Cells(b, 4).Value = a
b=b+1
Loop until b>"1" and b="7"




objExcel.ActiveWorkbook.Saveas "C:\Users\TOUSIF\Desktop\vbs\aaa.xlsx"

objExcel.ActiveWorkbook.Close "C:\Users\TOUSIF\Desktop\vbs\aaa.xlsx"


objExcel.Application.Quit
WScript.Echo "Finished."
WScript.Quit