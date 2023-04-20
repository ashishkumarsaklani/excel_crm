Attribute VB_Name = "Module3"
Public Function addDaily()
Dim ws4 As Worksheet
Set ws4 = Worksheets("Daily")
Dim lrow As Long
Dim lab0 As Integer




      lrow = ws4.Cells(Rows.Count, 1).End(xlUp).Row
      
ws4.Range("A:A").NumberFormat = "dd/mm/yy"
ws4.Range("B:B").NumberFormat = "hh:mm"





If UserForm2.Controls("Label11").Caption = "" Then
 Else
'MsgBox (Range("A" & lrow).Offset(1, 1).Address)

lab0 = 1
Call putValue(lrow, lab0)
 End If
  

  
 
 If UserForm2.Controls("Label21").Caption = "" Then

  
 Else
lab0 = 2

Call putValue(lrow, lab0)
 End If
 


 If UserForm2.Controls("Label31").Caption = "" Then


  
 Else
lab0 = 3

Call putValue(lrow, lab0)
 End If
 
 
 If UserForm2.Controls("Label41").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
lab0 = 4

Call putValue(lrow, lab0)
 End If
  If UserForm2.Controls("Label51").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
lab0 = 5
Call putValue(lrow, lab0)
 End If
   If UserForm2.Controls("Label61").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
 
lab0 = 6

Call putValue(lrow, lab0)
 End If
 
 'MsgBox (ran2)
 
 
' ran2.value = UserForm2.Controls(lab1).Caption & "  " & UserForm2.Controls(lab2).Caption
 'ws4.Range(ran3).value = UserForm2.Controls(lab3).Caption
 'ws4.Range(ran4).value = UserForm2.Controls(lab4).Caption
 'ws4.Range(ran5).value = UserForm2.Controls(lab5).Caption
 'tot = CInt(Left(UserForm2.Controls("Label10").Caption, Len(UserForm2.Controls("Label10").Caption) - 2))
' ws3.Range("h35:h36").value = tot * (9 / 100)
 'ws3.Range("h37:h38").value = tot * (9 / 100)
 ' ws3.Range("g39:h41").value = (tot + tot * (18 / 100)) & "Rs"

End Function

Public Function putValue(lrow, lab0)
 Dim ws4 As Worksheet
Set ws4 = Worksheets("Daily")

lab1 = "Label" & lab0 & "1"

lab2 = "Label" & lab0 & "2"
lab3 = "Label" & lab0 & "3"
lab4 = "Label" & lab0 & "4"
lab5 = "Label" & lab0 & "5"
lab6 = "Label" & lab0 & "6"


 ws4.Range("A" & lrow).Offset(lab0, 0).value = Date
 ws4.Range("A" & lrow).Offset(lab0, 1).value = Now()
 'MsgBox (UserForm2.Controls(lab1).Caption)
 ws4.Range("A" & lrow).Offset(lab0, 2).value = UserForm2.Controls(lab1).Caption & "  " & UserForm2.Controls(lab2).Caption
  ws4.Range("A" & lrow).Offset(lab0, 3).value = UserForm2.Controls(lab3).Caption
 ws4.Range("A" & lrow).Offset(lab0, 4).value = UserForm2.Controls(lab4).Caption
  ws4.Range("A" & lrow).Offset(lab0, 5).value = UserForm2.Controls(lab5).Caption
 ws4.Range("A" & lrow).Offset(lab0, 6).value = UserForm2.Controls(lab6).Caption
  ws4.Range("A" & lrow).Offset(lab0, 7).value = UserForm2.Controls("TextBox3").value
  ws4.Range("A" & lrow).Offset(lab0, 8).value = UserForm2.Controls("TextBox4").value



End Function
