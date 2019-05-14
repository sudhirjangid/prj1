Dim Choice
Choice = InputBox("Menu" & vbnewline & "1: Add" & vbnewline & "2: Delete")

        Select case Choice
            case 1
               call Add()
             
            case 2
               call sbDelete_Rows_Based_On_Criteria()
			   
			case else 
				MsgBox("Unknown Choice")
		End Select 
		

Function Add()
Dim obj, obj1, strExcelPath, NoOfEmployee, num
'strExcelPath="C:\Users\sujangid\Desktop\VB_MPT\Employee.xlsx"
Set obj = createobject("Excel.Application")  
obj.visible=True 
'obj.WorkBooks.Open strExcelPath
'Set obj1 = obj.ActiveWorkbook.Worksheets(1)
Set obj1 = obj.Workbooks.Open("C:\Users\sujangid\Desktop\VB_MPT\Employee.xlsx")
obj.Cells(1,1).Value= "Employee ID"   
obj.Cells(1,2).Value= "Employee Name" 
obj.Cells(1,3).Value= "Employee Email"
obj.Cells(1,4).Value= "Employee Salary"
NoOfEmployee = InputBox("Enter No. Of Employee")

For num = 2 To (Cint(NoOfEmployee)+1)

obj.Cells(num,1).Value= InputBox("Enter Employee ID")   
obj.Cells(num,2).Value= InputBox("Enter Employee Name") 
obj.Cells(num,3).Value= InputBox("Enter Employee Email")
obj.Cells(num,4).Value= InputBox("Enter Employee Salary")

If num <= (Cint(NoOfEmployee)) Then
MsgBox("Enter Next Employee Details")
End If

Next
obj1.Save

End Function


Function sbDelete_Rows_Based_On_Criteria()

Dim obj, obj1

'strExcelPath="C:\Users\sujangid\Desktop\VB_MPT\Employeee.xlsx"
Set obj = createobject("Excel.Application")  
obj.visible=false

'obj.WorkBooks.Open strExcelPath
'Set obj1 = obj.ActiveWorkbook.Worksheets(1)
Set obj1 = obj.Workbooks.Open("C:\Users\sujangid\Desktop\VB_MPT\Employee.xlsx")

Dim lRow 
Dim iCntr 
Dim Data
Data = InputBox("Enter Employee ID To Delete")
lRow = 10
For iCntr = lRow To 1 Step -1
    If obj.Cells(iCntr, 1) = Cint(data) Then
        obj.Rows(iCntr).Delete
    End If
	
Next
obj1.save
End Function



 
 