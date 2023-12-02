VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminHome 
   Caption         =   "Home - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmAdminHome.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAdminHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lblForecastTab_Click()
Unload Me
frmForecast.Show
End Sub

Private Sub Label3_Click()
Unload Me
frmInventory.Show
End Sub

Private Sub Label4_Click()
Unload Me
frmReport.Show
End Sub

Private Sub lblInventoryTab_Click()
Unload Me
frmInventory.Show

End Sub

Private Sub lblLogout2_Click()

Unload Me
IndexPage.Show

End Sub

Private Sub lblOrdersTab_Click()
Unload Me
frmOrder.Show
End Sub

Private Sub lblReportsTab_Click()
Unload Me
frmReport.Show
End Sub

Private Sub lblSettingsTab_Click()
Unload Me
frmSettings.Show
End Sub

Private Sub lblViewLatestOrder_Click()
Unload Me
frmOrder.Show
End Sub

Private Sub UserForm_Initialize()
Dim productidBS As String

Me.lblOrderid = Sheets("Order Shipping").Cells(Sheets("Order Shipping").Range("C1").End(xlDown).Row, Sheets("Order Shipping").Range("C1").End(xlDown).Column).Value
Me.lblTxtime = Sheets("Order Shipping").Cells(Sheets("Order Shipping").Range("D1").End(xlDown).Row, Sheets("Order Shipping").Range("D1").End(xlDown).Column).Value
Me.lblTotal = "$" & Sheets("Order Shipping").Cells(Sheets("Order Shipping").Range("I1").End(xlDown).Row, Sheets("Order Shipping").Range("I1").End(xlDown).Column).Value

Me.lblProductId = Sheets("Product").Cells(Sheets("Product").Range("A1").End(xlDown).Row, Sheets("Product").Range("A1").End(xlDown).Column).Value
Me.lblProductName = Sheets("Product").Cells(Sheets("Product").Range("B1").End(xlDown).Row, Sheets("Product").Range("B1").End(xlDown).Column).Value & " - " & Sheets("Product").Cells(Sheets("Product").Range("E1").End(xlDown).Row, Sheets("Product").Range("E1").End(xlDown).Column).Value

productidBS = Application.WorksheetFunction.XLookup(Application.WorksheetFunction.Max(Sheets("Sales Dashboard").Range("B60:B86")), Sheets("Sales Dashboard").Range("B60:B86"), Sheets("Sales Dashboard").Range("A60:A86"))
Me.lblProductidBS = productidBS
Me.lblQtyBS = Application.WorksheetFunction.Max(Sheets("Sales Dashboard").Range("B60:B86")) & " Pcs"
Me.lblProductnameBS = Application.WorksheetFunction.VLookup(productidBS, Sheets("Product").Range("A:B"), 2) & " - " & Application.WorksheetFunction.VLookup(productidBS, Sheets("Product").Range("A:E"), 5)




End Sub
