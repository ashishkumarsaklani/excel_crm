Attribute VB_Name = "Module1"
Public iRaw As Integer


Public Function gethRange(x) As String
Dim rngC As Range
Dim lrow As Long
Dim r1 As String
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    'lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = x.Cells(1, Columns.Count).End(xlToLeft).Column
    'MsgBox ("1st one " & iCol)
    
    ' to get letter of last column
    lette = Split(Cells(1, lCol).Address, "$")(1)
     ' MsgBox ("2st one " & lette)
    r1 = "E1" & ":" & lette & "1"
    
    gethRange = r1
     
End Function

Public Function getvRange(sheet, x, y)
Dim rngC As Range
Dim lrow As Long
Dim r1 As String
Dim z As String

   lrow = sheet.Cells(Rows.Count, y).End(xlUp).Row
z = Right(Left(x, 2), 1)
    If (lrow < 2) Then
    lrow = 1
    End If
r1 = z & "2" & ":" & z & lrow
getvRange = r1

     ' MsgBox ("2st one " & r1)
End Function


Public Function findValueInRange(workSheetName, RangeN, ValuetoCheck, o) As String 'o as Col or Row
Dim ws As Worksheet
Set ws = Worksheets(workSheetName)

    With ws.Range(RangeN)
        Set c = .Find(ValuetoCheck, LookIn:=xlValues)
            If Not c Is Nothing Then
                If (o = "Col") Then
                    findValueInRange = c.Address & "," & c.Column
                ElseIf (o = "Col") Then
                   findValueInRange = c.Address & "," & c.Row
                ElseIf (o = "No") Then
                    findValueInRange = c.Address
                End If
                
            Else
                findValueInRange = "stop"
            End If
    End With
End Function

Public Function checkLastInven() As String
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
Dim lastRow As Long
lastRow = ws2.Range("B" & Rows.Count).End(xlUp).Row
checkLastInven = "B" & lastRow
End Function

Public Function generateCode(z) As String
Dim y As Long
y = 100000 + z
generateCode = "A" & y
End Function


Public Function actform(x)
Dim ws As Worksheet
Set ws = Worksheets("Admin")



Dim r1 As String
r1 = gethRange(ws)
     x.ComboBox1.Clear
     
     'MsgBox (r1)

            For Each rngC In ws.Range(r1)
              x.ComboBox1.AddItem rngC.value
  
            Next rngC
End Function


Public Function setTwo(x)
Dim ws As Worksheet
Set ws = Worksheets("Admin")
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
Dim r1 As String
Dim r2 As String
Dim add As String
Dim val As String
Dim larray() As String
Dim y As Integer

r1 = gethRange(ws)
val = x.ComboBox1.value
add = findValueInRange("Admin", r1, val, "Col")

    If Not (add = "stop") Then

        larray = Split(add, ",")
        y = larray(1)
              r2 = getvRange(ws, larray(0), y)
             x.ComboBox2.Clear
                  For Each rngD In ws.Range(r2)
                     x.ComboBox2.AddItem rngD.value
              
                  Next rngD
                               '      Do
                                   'c.value = 5
                                   'Set c = .FindNext(c)
                               '   Loop While Not c Is Nothing
    End If

End Function

