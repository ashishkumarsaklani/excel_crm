VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13230
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
Call setTwo(UserForm3)
    Me.ComboBox2.ListIndex = 1
End Sub



Private Sub CommandButton1_Click()
Dim ws As Worksheet
Set ws = Worksheets("Admin")
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
Dim r1 As String
Dim val1 As String
Dim val2 As String
r1 = gethRange()
Dim larray() As String
Dim y As Integer
Dim r3 As String
Dim Lastinv As Integer
r3 = checkLastInven()
Lastinv = Right(r3, 1)

If Not ((TextBox1.value = Empty) Or (TextBox2.value = Empty) Or (TextBox3.value = Empty)) Then

        val1 = findValueInRange("Admin", r1, ComboBox1.value, "Col")
            If (val1 = "stop") Then
        
            larray = Split(r1, ":")
            r1 = larray(1)
        
           ' MsgBox (r1)
            Range(r1).Offset(0, 1).value = ComboBox1.value
            Range(r1).Offset(1, 1).value = ComboBox2.value
            ws2.Range(r3).Offset(1, 0).value = ComboBox2.value
            ws2.Range(r3).Offset(1, 1).value = ComboBox1.value
            ws2.Range(r3).Offset(1, -1).value = generateCode(Lastinv)
            ws2.Range(r3).Offset(1, 2).value = TextBox1.value
            ws2.Range(r3).Offset(1, 3).value = TextBox2.value
            ws2.Range(r3).Offset(1, 4).value = TextBox3.value
            
                
            Else
        
             larray = Split(val1, ",")
                y = larray(1)
            r1 = getvRange(ws, larray(0), y)
            'MsgBox (r1)
        
             'area to check if value 2 is already presetnt
            val2 = findValueInRange("Admin", r1, ComboBox2.value, "Row")
 
         '  MsgBox (val2)
        
                If (val2 = "stop") Then
        
                larray = Split(r1, ":")
                r1 = larray(1)
                Range(r1).Offset(1, 0).value = ComboBox2.value
                ws2.Range(r3).Offset(1, 0).value = ComboBox2.value
                ws2.Range(r3).Offset(1, 1).value = ComboBox1.value
                ws2.Range(r3).Offset(1, -1).value = generateCode(Lastinv)
                ws2.Range(r3).Offset(1, 2).value = TextBox1.value
                ws2.Range(r3).Offset(1, 3).value = TextBox2.value
                ws2.Range(r3).Offset(1, 4).value = TextBox3.value
                End If
            End If
    Else
        MsgBox ("enter all values")
    End If
End Sub

Private Sub Label1_Click()
Dim ws As Worksheet
Set ws = Worksheets("Admin")
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
ws.Visible = xlSheetVisible
ws2.Visible = xlSheetVisible
Application.Visible = True


End Sub

Private Sub ComboBox2_Enter()
  ComboBox2.SetFocus
   SendKeys "%{Down}"
End Sub


Private Sub UserForm_Activate()
Call actform(UserForm3)
    Me.ComboBox1.ListIndex = 1
   
End Sub

Private Sub UserForm_Terminate()
ThisWorkbook.Saved = False


Application.Quit

End Sub
