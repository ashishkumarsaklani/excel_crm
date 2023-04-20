Attribute VB_Name = "Module2"
Public Function printReady()



Dim ws3 As Worksheet
Set ws3 = Worksheets("Bill")

ws3.Range("a1:i47").Clear


ws3.Range("a1:b1").Interior.ColorIndex = 23
ws3.Range("c1:e1").Interior.ColorIndex = 8
ws3.Range("f1:i1").Interior.ColorIndex = 3
ws3.Range("a47:d47").Interior.ColorIndex = 3
ws3.Range("e47:g47").Interior.ColorIndex = 8
ws3.Range("h47:i47").Interior.ColorIndex = 23
ws3.Range("b3:h7").Interior.ColorIndex = 15
 ws3.Range("b3:h4,b5:h5,b6:d6,b7:d7,b9:c10,D9:H10,d12:h12,b12:C12,B15:B16,C15:E16,F15:F16,G15:G16,H15:H16,F35:G36,F37:G38,b44:h44,b45:h45,b18:b19,b21:b22,b24:b25,b27:b28,b30:b31,b33:b34,f39:f41,g39:h41,h35:h36,h37:h38").Merge
 ws3.Range("c18:e19,c21:e22,c24:e25,c27:e28,c30:e31,c33,e34,f18:f19,g18:g19,h18:h19,f21:f22,g21:g22,h21:h22,f24:f25,g24:g25,h24:h25,f27:f28,g27:g28,h27:h28,f30:f31,g30:g31,h30:h31,f33:f34,g33:g34,h33:h34").Merge
 


With ws3.Range("b3:h4")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(255, 0, 0)
    .Font.Size = 31
    .value = "Logical Trends"
End With

With ws3.Range("f35: h41")
    .Interior.ColorIndex = 15
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders().LineStyle = xlContinuous
End With
ws3.Range("f39:h41").Font.Size = 20

With ws3.Range("b5:h5")
    .Interior.ColorIndex = 15
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .Font.Size = 12
    .value = "A-191 Gulab Bagh Uttam Nagar Delhi 59"

    
End With

With ws3.Range("b6:d6")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .Font.Size = 10
    .value = "email :trendslogical@gmail.com"

End With
With ws3.Range("b7:d7")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .Font.Size = 10
    .value = "web :www.logicaltrends.com"
End With

With ws3.Range("g6:g6")
    .value = "Invoice No"
End With
With ws3.Range("g7:g7")
    .value = "Date"
End With


With ws3.Range("b9:C10")
    .value = "Name"
End With
With ws3.Range("D9:H10")
    .Font.Size = 14
End With

With ws3.Range("b12:C12")
    .value = "Address"
End With



With ws3.Range("b9:h10,b12:C12,D12:H12,B15:H16,B18:H19,B21:H22,B24:H25,B27:H28,B30:H31,B33:H34")
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Interior.ColorIndex = 15
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .Font.Size = 10
End With



 With ws3.Range("B15:b34")
   ' .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlRight).LineStyle = xlContinuous
 End With
 With ws3.Range("C15:E34")
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
 End With
  With ws3.Range("F15:g34")
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlRight).LineStyle = xlContinuous
 End With
 
   With ws3.Range("B15:B16")
    .value = "S. No."
 End With
    With ws3.Range("C15:E16")
    .value = "Perticulars"
 End With
     With ws3.Range("F15:F16")
    .value = "Qty."
 End With
 With ws3.Range("G15:G16")
    .value = "Rate"
 End With
With ws3.Range("H15:H16")
    .value = "Amount"
 End With
 With ws3.Range("F35:G36")
    .value = "CGST @ 9%"
 End With
  With ws3.Range("F37:G38")
    .value = "IGST @ 9%"
 End With
   With ws3.Range("F39:F41")
    .value = "Total"
 End With
 
    With ws3.Range("b44:h44")
    .value = "Terms Consumables are non refundable"
 End With
     With ws3.Range("b45:h45")
    .value = "Terms Refund Only available in next 10 Mins of purchase"
 End With
 
End Function

Public Function printBill()

 Dim ws3 As Worksheet
Dim tot As Integer
Dim ran1, ran2, ran3, ran4, ran5, lab1, lab2, lab3, lab4, lab5 As String
Set ws3 = Worksheets("Bill")

ws3.Range("h7").NumberFormat = "dd/mm/yyyy"
ws3.Range("h7").value = Date



If UserForm2.Controls("Label11").Caption = "" Then
 Else
 'MsgBox (UserForm2.Controls("Label11").Caption)
ran1 = "b18"
ran2 = "c18:e19"
ran3 = "F18:F19"
ran4 = "g18:g19"
ran5 = "h18:h19"
lab0 = "1"
lab1 = "Label11"
lab2 = "Label12"
lab3 = "Label13"
lab4 = "Label14"
lab5 = "Label16"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
  

  
 
 If UserForm2.Controls("Label21").Caption = "" Then
' MsgBox ("hi catagory is empty")

  
 Else
 ran1 = "b21"
ran2 = "c21:e22"
ran3 = "F21:F22"
ran4 = "g21:g22"
ran5 = "h21:h22"
lab0 = "2"
lab1 = "Label21"
lab2 = "Label22"
lab3 = "Label23"
lab4 = "Label24"
lab5 = "Label26"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
 


 If UserForm2.Controls("Label31").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
 ran1 = "b24"
ran2 = "c24:e25"
ran3 = "F24:F25"
ran4 = "g24:g25"
ran5 = "h24:h25"
lab0 = "3"
lab1 = "Label31"
lab2 = "Label32"
lab3 = "Label33"
lab4 = "Label34"
lab5 = "Label36"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
 
 
 If UserForm2.Controls("Label41").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
 ran1 = "b27"
ran2 = "c27:e28"
ran3 = "F27:F28"
ran4 = "g27:g28"
ran5 = "h27:h28"
lab0 = "4"
lab1 = "Label41"
lab2 = "Label42"
lab3 = "Label43"
lab4 = "Label44"
lab5 = "Label46"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
  If UserForm2.Controls("Label51").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
 ran1 = "b30"
ran2 = "c30:e31"
ran3 = "F30:F31"
ran4 = "g30:g31"
ran5 = "h30:h31"
lab0 = "5"
lab1 = "Label51"
lab2 = "Label52"
lab3 = "Label53"
lab4 = "Label54"
lab5 = "Label56"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
   If UserForm2.Controls("Label61").Caption = "" Then
 'MsgBox ("hi catagory is empty")

  
 Else
 ran1 = "b33"
ran2 = "c33:e34"
ran3 = "F33:F34"
ran4 = "g33:g34"
ran5 = "h33:h34"
lab0 = "6"
lab1 = "Label61"
lab2 = "Label62"
lab3 = "Label63"
lab4 = "Label64"
lab5 = "Label66"
Call putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)

 End If
 
  tot = CInt(Left(UserForm2.Controls("Label10").Caption, Len(UserForm2.Controls("Label10").Caption) - 2))
 ws3.Range("h35:h36").value = tot * (9 / 100)
 ws3.Range("h37:h38").value = tot * (9 / 100)
  ws3.Range("g39:h41").value = (tot + tot * (18 / 100)) & "Rs"
 
 
  ws3.Range("D9").value = UserForm2.Controls("TextBox3").value
  ws3.Range("D12").value = UserForm2.Controls("TextBox4").value
 ws3.Range("e47:g47").Merge
 ws3.Range("e47:g47").value = Time()

ws3.Range("a1:i47").PrintPreview


End Function

Public Function putValue(ran1, ran2, ran3, ran4, ran5, lab0, lab1, lab2, lab3, lab4, lab5)
 Dim ws3 As Worksheet
Set ws3 = Worksheets("Bill")

'MsgBox (ran2 & lab1 & lab2)

 ws3.Range(ran1).value = lab0
 ws3.Range(ran2).value = UserForm2.Controls(lab1).Caption & "  " & UserForm2.Controls(lab2).Caption
 ws3.Range(ran3).value = UserForm2.Controls(lab3).Caption
 ws3.Range(ran4).value = UserForm2.Controls(lab4).Caption
 ws3.Range(ran5).value = UserForm2.Controls(lab5).Caption
 
  



End Function

