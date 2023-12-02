VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportInventory 
   Caption         =   "Inventory Report - Admin"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmReportInventory.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmReportInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lblBackToPreviosPage_Click()
    Unload Me
    frmReport.Show
End Sub

Private Sub UserForm_Initialize()

    Me.cmbInventoryReportSelection.List = Array("List All", "Quantity Report", "Category Report")
    Me.cmbInventoryReportSelection.Value = "List All"
    Call LoadListViewAll
    
End Sub

Private Sub cmbInventoryReportSelection_Change()

    Select Case Me.cmbInventoryReportSelection
    Case "List All"
        Call LoadListViewAll
    Case "Quantity Report"
        Call LoadListViewQuantity
    Case "Category Report"
        Call LoadListViewCategory
    End Select
    
End Sub

Private Sub LoadListViewAll()
    Me.listViewInventory.ListItems.Clear
    With Me.listViewInventory
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
        .Width = 732
        
        With .ColumnHeaders
          .Clear
          .Add , , "Product ID", 80
          .Add , , "Product Name", 170
          .Add , , "Cost", 30
          .Add , , "Price", 30
          .Add , , "Color", 40
          .Add , , "Quantity S", 60
          .Add , , "Quantity M", 60
          .Add , , "Quantity L", 60
          .Add , , "Gender", 60
          .Add , , "Category", 70
          .Add , , "On Sale", 50
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
    
    Set wksSource = Worksheets("Product")
    
    Set rngData = wksSource.Range("A2").CurrentRegion
    
    RowCount = rngData.Rows.count
    
    ColCount = rngData.Columns.count
    
    For i = 2 To RowCount
    
        Set LstItem = Me.listViewInventory.ListItems.Add(Text:=rngData(i, 1).Value)
        
        For j = 2 To ColCount
            LstItem.ListSubItems.Add Text:=rngData(i, j).Value
        Next j
    Next i

End Sub

Private Sub LoadListViewQuantity()
    Me.listViewInventory.ListItems.Clear
    With Me.listViewInventory
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
        .Width = 732
        
        With .ColumnHeaders
          .Clear
          .Add , , "Product", 170
          .Add , , "Size S", 40
          .Add , , "Size M", 40
          .Add , , "Size L", 40
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
    
    Set wksSource = Worksheets("Inventory Report")
    
    Set rngData = wksSource.Range("A2").CurrentRegion
    
    RowCount = rngData.Rows.count
    
    ColCount = rngData.Columns.count
    
    For i = 3 To RowCount
    
        Set LstItem = Me.listViewInventory.ListItems.Add(Text:=rngData(i, 1).Value)
        
        For j = 2 To ColCount
            LstItem.ListSubItems.Add Text:=rngData(i, j).Value
        Next j
    Next i

End Sub

Private Sub LoadListViewCategory()
    Me.listViewInventory.ListItems.Clear
    With Me.listViewInventory
        .Gridlines = True
        .HideColumnHeaders = False
        .View = lvwReport
        .Width = 320
        
        With .ColumnHeaders
          .Clear
          .Add , , "Category", 170
          .Add , , "Sum of Cost", 60
          .Add , , "Sum of Price", 60
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
    
    Set wksSource = Worksheets("Inventory Report")
    
    Set rngData = wksSource.Range("I2").CurrentRegion
    
    RowCount = rngData.Rows.count
    
    ColCount = rngData.Columns.count
    
    For i = 3 To RowCount
    
        Set LstItem = Me.listViewInventory.ListItems.Add(Text:=rngData(i, 1).Value)
        
        For j = 2 To ColCount
            LstItem.ListSubItems.Add Text:=rngData(i, j).Value
        Next j
    Next i

    Call UpdateChart
    
End Sub

Private Sub UpdateChart()

    Dim sTempFile As String
    Dim oChart As Chart

    sTempFile = Environ("temp") & "\temp.gif"
 
    Set oChart = Worksheets("Inventory Report").ChartObjects("CategoryChart").Chart

    oChart.Export Filename:=sTempFile, filterName:="GIF"

    Me.imgChart.Picture = LoadPicture(sTempFile)

    Me.Repaint

    Kill sTempFile

End Sub
