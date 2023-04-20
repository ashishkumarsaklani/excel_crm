VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Bill"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13410
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Call setTwo(UserForm2)
ComboBox2.ListIndex = 1

End Sub



Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then ' for enter key

  ComboBox2.SetFocus
   SendKeys "%{Down}"


End If

End Sub

Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim a As String
Dim r As String
Dim lrow As Long
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
lrow = ws2.Range("B" & Rows.Count).End(xlUp).Row
r = "B2:B" & lrow
a = Me.ComboBox2.value


If Not (a = "") Then
            
             'MsgBox (r & "this")
            
            
            
                With ws2.Range(r)
                    Set c = .Find(a, LookIn:=xlValues)
                    If Not c Is Nothing Then
                            a = c.Address
                    End If
                End With
            
            'MsgBox (a)
            
                UserForm2.Controls("Label6").Caption = ws2.Range(a).Offset(0, 4).value & "Rs"
                UserForm2.Controls("TextBox1").value = "1"
                UserForm2.Controls("TextBox2").value = "0"
                UserForm2.Controls("Label8").Caption = CInt(ws2.Range(a).Offset(0, 4).value) - CInt(UserForm2.Controls("TextBox2").value) & "Rs"
 End If
    
    
End Sub

Private Sub CommandButton1_Click()
iRaw = iRaw + 10
If iRaw <= 60 Then

 UserForm2.Controls("Label" & (iRaw + 1)).Caption = UserForm2.Controls("ComboBox1").value
  UserForm2.Controls("Label" & (iRaw + 2)).Caption = UserForm2.Controls("ComboBox2").value
   UserForm2.Controls("Label" & (iRaw + 3)).Caption = UserForm2.Controls("TextBox1").value
    UserForm2.Controls("Label" & (iRaw + 4)).Caption = UserForm2.Controls("Label6").Caption
     UserForm2.Controls("Label" & (iRaw + 5)).Caption = UserForm2.Controls("TextBox1").value
      UserForm2.Controls("Label" & (iRaw + 6)).Caption = UserForm2.Controls("Label8").Caption

        UserForm2.Controls("Label10").Caption = (CInt(Left(UserForm2.Controls("Label8").Caption, Len(UserForm2.Controls("Label8").Caption) - 2)) + CInt(Left(UserForm2.Controls("Label10").Caption, Len(UserForm2.Controls("Label10").Caption) - 2))) & "Rs"
    Else
        MsgBox ("6 Items Only in one /n Bill Generate New Bill")
    End If
              Me.ComboBox1.SetFocus
    
    
     
End Sub

Private Sub Label1_Click()
UserForm2.Hide
UserForm3.Show

End Sub




Private Sub Label67_Click()
' print  and end
Call printReady
Call addDaily
Call printBill
Call Reset_C





End Sub

Private Sub Label9_Click()
Call addDaily
Call Reset_C
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim a As Integer
a = CInt(Left(UserForm2.Controls("Label6").Caption, Len(UserForm2.Controls("Label6").Caption) - 2))
a = a * CInt(UserForm2.Controls("TextBox1").value)
a = a - CInt(UserForm2.Controls("TextBox2").value)

UserForm2.Controls("Label8").Caption = CStr(a) & "Rs"
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

 a = CInt(Left(UserForm2.Controls("Label6").Caption, Len(UserForm2.Controls("Label6").Caption) - 2))
a = a * CInt(UserForm2.Controls("TextBox1").value)
a = a - CInt(UserForm2.Controls("TextBox2").value)

UserForm2.Controls("Label8").Caption = CStr(a) & "Rs"

End Sub

Private Sub ComboBox1_Enter()
  ComboBox1.SetFocus
   SendKeys "%{Down}"

End Sub

Private Sub ComboBox2_Enter()
  ComboBox2.SetFocus
   SendKeys "%{Down}"

End Sub



Private Sub UserForm_Activate()

Me.ComboBox1.Clear
Call actform(UserForm2)
            Me.ComboBox1.ListIndex = 1
              

   UserForm2.Controls("Label10").Caption = 0 & "Rs"

End Sub

Private Sub UserForm_Terminate()
Application.Quit

End Sub


Private Function Reset_C()

    Dim ctl As MSForms.Control

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl
    
    For iRaw = 10 To 60 Step 10
    
 UserForm2.Controls("Label" & (iRaw + 1)).Caption = ""
  UserForm2.Controls("Label" & (iRaw + 2)).Caption = ""
   UserForm2.Controls("Label" & (iRaw + 3)).Caption = ""
    UserForm2.Controls("Label" & (iRaw + 4)).Caption = ""
     UserForm2.Controls("Label" & (iRaw + 5)).Caption = ""
      UserForm2.Controls("Label" & (iRaw + 6)).Caption = ""

        UserForm2.Controls("Label10").Caption = 0 & "Rs"
    Next iRaw
    
   iRaw = 0

End Function

