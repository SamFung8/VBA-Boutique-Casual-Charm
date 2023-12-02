VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForecast 
   Caption         =   "Forecast - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmForecast.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
Unload Me
IndexPage.Show
End Sub

Private Sub lblHomepageTab_Click()
Unload Me
frmAdminHome.Show
End Sub

Private Sub lblInventoryTab_Click()
Unload Me
frmInventory.Show

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
