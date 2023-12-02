VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportProfit 
   Caption         =   "Profit and Loss Report - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmReportProfit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblBackToPreviosPage_Click()
    Unload Me
    frmReport.Show
End Sub


Private Sub UserForm_Initialize()


Call LoadListViewAll
End Sub

Private Sub LoadListViewAll()
    Me.listViewProfit.ListItems.Clear
    With Me.listViewProfit
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
        .Width = 732
        
        With .ColumnHeaders
          .Clear
          .Add , , "Date", 100
          .Add , , "No. of Orders", 80
          .Add , , "Total Price", 80
          .Add , , "Total Cost", 80
          .Add , , "Revenue", 80
        End With

    End With

    Dim wksSource As Worksheet
    Dim rngData As Range
    Dim rngCell As Range
    Dim LstItem As ListItem
    Dim RowCount As Long
    Dim ColCount As Long
    Dim i As Long
    Dim j As Long
    
    Set wksSource = Worksheets("Profit Loss Report")
    
    Set rngData = wksSource.Range("A3").CurrentRegion
    
    RowCount = rngData.Rows.count
    
    ColCount = rngData.Columns.count
    
    For i = 2 To RowCount
    
        Set LstItem = Me.listViewProfit.ListItems.Add(Text:=rngData(i, 1).Value)
        
        For j = 2 To ColCount
            LstItem.ListSubItems.Add Text:=rngData(i, j).Value
        Next j
    Next i

End Sub
