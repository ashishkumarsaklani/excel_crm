VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Welcome to Logical Trends"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8340.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public role As String
Public value As Integer


Private Sub CommandButton2_Click()
'wait for some thing good


End Sub


Private Sub CommandButton1_Click()
 a = TextBox1.Text
 b = TextBox2.Text

'MsgBox (" " & a & vbCrLf & b)

Dim i As Integer




   If ((a = "ASHISH") And (b = "a1")) Then
        value = FindValue(a)
        'MsgBox ("welcome")
        
        'role = range("C" & value)
        
        UserForm1.Hide
        UserForm2.Show
        ActiveWorkbook.Unprotect ("admin")
            
    ElseIf (a = "ADMIN" And b = "admin") Then
        UserForm1.Hide
        UserForm3.Show
         ActiveWorkbook.Unprotect ("admin")
         Application.Visible = True
         

        
    Else
        MsgBox ("Incorrect Username Password")
    End If



End Sub
 
Public Function FindValue(valueToFind) As Integer
    Dim i As Integer
     For i = 1 To 5  ' Revise the 50 to include all of your values
    
        If (Cells(i, 1).value = valueToFind) Then
            'MsgBox ("Found value on row " & i)
           
           FindValue = i
           
            Exit Function
        End If
    Next i
 End Function
 




Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
TextBox1.Text = UCase(TextBox1.Text)
End Sub

Private Sub UserForm_Terminate()
ThisWorkbook.Saved = False
Application.Quit

End Sub
