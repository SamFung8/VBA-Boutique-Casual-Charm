VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSalesDash 
   Caption         =   "Sales Dashboard - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmSalesDash.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSalesDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboSelect_Change()

Select Case cboSelect.Value
Case "Category Sales"
    UpdateChart1

Case "Sales by Size"
    UpdateChart2
Case "Sales by Gender"
    UpdateChart3
    
Case "Average Sales Amount"
UpdateChart4
Case "Payment Method Used"
UpdateChart5
Case Else
    
End Select

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

Private Sub UpdateChart1()
On Error Resume Next
Sheets("Sales Dashboard").Activate
Dim catName As String
Dim cat As Chart
    
    Set cat = Sheets("Sales Dashboard").ChartObjects("Chart 2").Chart

'   Save chart as GIF
    catName = Environ("temp") & "\cat.jpg"
    cat.Export Filename:=catName, filterName:="jpg"

'   Show the chart
    frmSalesDash.Image1.Picture = LoadPicture(catName)
     frmSalesDash.Repaint
    
    Kill catName

End Sub
Private Sub UpdateChart2()
On Error Resume Next
Sheets("Sales Dashboard").Activate
Dim sizeName As String
Dim size  As Chart

    Set size = Sheets("Sales Dashboard").ChartObjects("Chart 4").Chart

    sizeName = Environ("temp") & "\size.jpg"
    size.Export Filename:=sizeName, filterName:="jpg"

    frmSalesDash.Image1.Picture = LoadPicture(sizeName)
     frmSalesDash.Repaint
    Kill sizeName

End Sub
Private Sub UpdateChart3()
On Error Resume Next
Sheets("Sales Dashboard").Activate
Dim genderName As String
Dim gender   As Chart

    Set gender = Sheets("Sales Dashboard").ChartObjects("Chart 3").Chart

    genderName = Environ("temp") & "\gender.jpg"
    gender.Export Filename:=genderName, filterName:="jpg"

    frmSalesDash.Image1.Picture = LoadPicture(genderName)
     frmSalesDash.Repaint
    
    Kill genderName

End Sub
Private Sub UpdateChart4()
On Error Resume Next
Sheets("Sales Dashboard").Activate
Dim avgsalesName As String
Dim avgsales As Chart

    Set avgsales = Sheets("Sales Dashboard").ChartObjects("Chart 6").Chart

    avgsalesName = Environ("temp") & "\avgssale.jpg"
    avgsales.Export Filename:=avgsalesName, filterName:="jpg"

    frmSalesDash.Image1.Picture = LoadPicture(avgsalesName)
     frmSalesDash.Repaint
    Kill avgsalesName

End Sub
Private Sub UpdateChart5()
On Error Resume Next
Sheets("Sales Dashboard").Activate
Dim paymentName As String
Dim payment  As Chart

    Set payment = Sheets("Sales Dashboard").ChartObjects("Chart 5").Chart

    paymentName = Environ("temp") & "\payment.jpg"
    payment.Export Filename:=paymentName, filterName:="jpg"

    frmSalesDash.Image1.Picture = LoadPicture(paymentName)
    
    frmSalesDash.Repaint
    
    Kill paymentName

End Sub


Private Sub UserForm_Initialize()
With cboSelect
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    .AddItem ("Category Sales")
    .AddItem ("Sales by Size")
    .AddItem ("Sales by Gender")
    .AddItem ("Average Sales Amount")
    .AddItem ("Payment Method Used")

End With
UpdateChart1
End Sub
