VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportSales 
   Caption         =   "Sales Reports - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmReportSales.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmReportSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label2_Click()
Unload Me
frmSalesData.Show


End Sub

Private Sub Label3_Click()
Unload Me
frmSalesDash.Show
End Sub

Private Sub lblForecastTab_Click()
Unload Me
frmForecast.Show
End Sub

Private Sub lblHomepageTab_Click()
Unload Me
frmAdminHome.Show
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

