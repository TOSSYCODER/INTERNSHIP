Set obj = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks
Dim n,a 
Dim b
Dim c

b="2"
c = InputBox("ENTER NO. OF ENTRY")

objExcel.Application.Visible = True
objExcel.Workbooks.Add
objExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 4).Value = "AGE"




Do
n = InputBox("ENTER NAME")
a = InputBox("ENTER AGE")
if IsEmpty(n) Then
objExcel.ActiveWorkbook.Saveas "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
objExcel.ActiveWorkbook.Close "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
objExcel.Application.Quit
obj.DeleteFile "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
'WScript.Echo "Finished."
WScript.Quit

elseif IsEmpty(a) Then
objExcel.ActiveWorkbook.Saveas "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
objExcel.ActiveWorkbook.Close "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
objExcel.Application.Quit
obj.DeleteFile "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"
'WScript.Echo "Finished."
WScript.Quit
else
objExcel.Cells(b, 3).Value = n
objExcel.Cells(b, 4).Value = a
b=b+1
end if
Loop until b>"1" and b=c+2





objExcel.ActiveWorkbook.Saveas "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"

objExcel.ActiveWorkbook.Close "C:\Users\TOUSIF\Desktop\vbs\na.xlsx"


objExcel.Application.Quit
WScript.Echo "Finished."
WScript.Quit