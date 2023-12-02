VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSalesData 
   Caption         =   "Sales Data - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmSalesData.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSalesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboSelect_Change()

Select Case cboSelect.Value
Case "Years"
    UpdateChart1

Case "Months (2022)"
    UpdateChart2
Case "Months (2023)"
    UpdateChart3
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


Private Sub UserForm_Initialize()



With cboSelect
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    .AddItem ("Years")
    .AddItem ("Months (2022)")
    .AddItem ("Months (2023)")

End With

End Sub

Private Sub UpdateChart1()
On Error Resume Next
Dim PicFilename1 As String
Dim currentchart As Chart
    Sheets("Sales Data").Activate
    Set currentchart = Sheets("Sales Data").ChartObjects("Chart 1").Chart

'   Save chart as GIF
    PicFilename1 = Environ("temp") & "\temp1_" & Format(Now(), "yyyyMMddhhmmss") & ".jpg"
    currentchart.Export Filename:=PicFilename1, filterName:="jpg"

'   Show the chart
    frmSalesData.Image1.Picture = LoadPicture(vbNullString)
    frmSalesData.Image1.Picture = LoadPicture(PicFilename1)
    frmSalesData.Repaint
    
    Kill PicFilename1

End Sub
Private Sub UpdateChart2()
On Error Resume Next
Dim PicFilename2 As String
Dim currentchart As Chart
    Sheets("Sales Data").Activate
    Set currentchart = Sheets("Sales Data").ChartObjects("Chart 2").Chart

'   Save chart as GIF
    PicFilename2 = Environ("temp") & "\temp2_" & Format(Now(), "yyyyMMddhhmmss") & ".jpg"
    currentchart.Export Filename:=PicFilename2, filterName:="jpg"

'   Show the chart
frmSalesData.Image1.Picture = LoadPicture(vbNullString)
    frmSalesData.Image1.Picture = LoadPicture(PicFilename2)
    frmSalesData.Repaint
    
    Kill PicFilename2

End Sub
Private Sub UpdateChart3()
On Error Resume Next
Dim PicFilename3 As String
Dim currentchart As Chart
    Sheets("Sales Data").Activate
    Set currentchart = Sheets("Sales Data").ChartObjects("Chart 11").Chart

'   Save chart as GIF
    PicFilename3 = Environ("temp") & "\temp3_" & Format(Now(), "yyyyMMddhhmmss") & ".jpg"
    currentchart.Export Filename:=PicFilename3, filterName:="jpg"

'   Show the chart
frmSalesData.Image1.Picture = LoadPicture(vbNullString)
    frmSalesData.Image1.Picture = LoadPicture(PicFilename3)
    frmSalesData.Repaint
    
    Kill PicFilename3

End Sub


