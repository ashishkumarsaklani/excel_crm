VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Dim ws As Worksheet
Set ws = Worksheets("Admin")
Dim ws2 As Worksheet
Set ws2 = Worksheets("Inventory")
Dim ws3 As Worksheet
Set ws3 = Worksheets("Bill")
Dim ws4 As Worksheet
Set ws4 = Worksheets("Daily")

ws4.Range("a1:i1").value = Array("Date", "Time", "Product", "Quantity", "Rate", "Discount", "Price", "Name", "Address")

Application.Visible = False
UserForm1.Show


End Sub
