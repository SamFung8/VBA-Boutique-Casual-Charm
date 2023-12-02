VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventory 
   Caption         =   "Inventory Management - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmInventory.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public InventoryEvent As New Collection

Private Sub lblLogout2_Click()
Unload Me
IndexPage.Show
End Sub

Private Sub UserForm_Initialize()
    Set InventoryEvent = New Collection
    Dim categoryRnage As Range
    Dim dictionaryWs As Worksheet
    
    Set dictionaryWs = Worksheets("Dictionary")
    Me.cmbFilterCategory.AddItem "All"
    For Each categoryRnage In dictionaryWs.Range("a2:a8")
        If categoryRnage.Value <> "" Then
            Me.cmbFilterCategory.AddItem categoryRnage.Value
        End If
    Next
    Me.cmbFilterCategory.Value = "All"
    Me.cmbFilterGender.List = Array("All", "Men", "Women")
    Me.cmbFilterGender.Value = "All"
    
    Call ResetTable(True, False, "", False, "", False, "", False, "")

End Sub


Function LoadPNGImg(ByVal Filename As String) As StdPicture
        With CreateObject("WIA.ImageFile")
            .LoadFile Filename
            Set LoadPNGImg = .FileData.Picture
        End With
End Function

Private Sub lblForecastTab_Click()
Unload Me
frmForecast.Show
End Sub

Private Sub lblHomepageTab_Click()
Unload Me
frmAdminHome.Show
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



Public Sub ResetTable(noFilter, filterId, filterIdVal, filterName, filterNameVal, filterGender, filterGenderVal, filterCategory, filterCategoryVal)
    Dim productWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer, i As Integer, j As Integer
    Dim countFilterRow As Integer
    countFilterRow = 0
    Me.frameInventoryTable.Controls.Clear
    
    Worksheets("Product").Activate
    Set productWs = ActiveSheet
    With productWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmInventory.frameInventoryTable.ScrollHeight = 0

        For i = 1 To lastRow
           
            Dim lblActionEdit As Control
            Dim lblActionDel As Control
            Dim actionEvent As InventoryEventHandler
            Dim lblProductId As MSForms.Label
            Dim lblProductName As MSForms.Label
            Dim lblProductGender As MSForms.Label
            Dim lblProductCategory As MSForms.Label
            Dim isId As Boolean, isName As Boolean, isGender As Boolean, isCategory As Boolean
                
            If i <> 1 Then
            
                isId = False
                isName = False
                isGender = False
                isCategory = False
                If filterId = True Then
                    If (productWs.Cells(i, 1).Value Like "*" & filterIdVal & "*") Then
                        isId = True
                    End If
                End If
                If filterName = True Then
                    If (productWs.Cells(i, 2).Value Like "*" & filterNameVal & "*") Then
                        isName = True
                    End If
                End If
                If filterGender = True Then
                    If (productWs.Cells(i, 9).Value = filterGenderVal) Then
                        isGender = True
                    End If
                End If
                If filterCategory = True Then
                    If (productWs.Cells(i, 10).Value = filterCategoryVal) Then
                        isCategory = True
                    End If
                End If
                
                
                If noFilter = True Or isId = True Or isName = True Or isGender = True Or isCategory = True Then
                    
                    countFilterRow = countFilterRow + 1
                    Set lblProductId = frameInventoryTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    Set lblProductName = frameInventoryTable.Controls.Add("Forms.Label.1", "lblProductName" & CStr(i))
                    Set lblProductGender = frameInventoryTable.Controls.Add("Forms.Label.1", "lblProductGender" & CStr(i))
                    Set lblProductCategory = frameInventoryTable.Controls.Add("Forms.Label.1", "lblProductCategory" & CStr(i))
                    
                    With lblProductId
                        .Caption = productWs.Cells(i, 1).Value
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .Left = 10
                        .Width = 60
                    End With
                            
                    With lblProductName
                        .Caption = productWs.Cells(i, 2).Value
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .Left = 75
                        .Width = 240
                    End With
                            
                    With lblProductGender
                        .Caption = productWs.Cells(i, 9).Value
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .Left = 110 * 3
                        .Width = 50
                    End With
                    
                    With lblProductCategory
                        .Caption = productWs.Cells(i, 10).Value
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .Left = 100 * 4
                        .Width = 90
                    End With
                
                    'Edit Button
                    Set lblActionEdit = frameInventoryTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                    With lblActionEdit
                        .Tag = productWs.Cells(i, 1).Value
                        .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                        .Caption = "edit"
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 15
                        .Width = 15
                        .Left = 114 * 5
                        .MousePointer = 99
                        .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
                    End With
                    Set actionEvent = New InventoryEventHandler
                    Set actionEvent.EditLabel = lblActionEdit
                    
                    ' Delete Button
                    Set lblActionDel = frameInventoryTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                    With lblActionDel
                        .Tag = productWs.Cells(i, 1).Value
                        .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                        .Caption = "del"
                        .Top = 37 * (countFilterRow)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 15
                        .Width = 15
                        .Left = 120 * 5
                        .MousePointer = 99
                        .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
                    End With
                    Set actionEvent.DelLabel = lblActionDel
                    InventoryEvent.Add actionEvent
                    
                    ' Row Divider
                    Dim lblRowDivider As MSForms.Label
                    Set lblRowDivider = frameInventoryTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                    With lblRowDivider
                        .Top = 37 * (countFilterRow) - 12
                        .BackColor = &H8000000F
                        .Height = 2
                        .Width = 700
                    End With
                    ' Set height
                With frmInventory.frameInventoryTable
                        .ScrollHeight = .ScrollHeight + 38
                    End With
                End If
                
            End If
            
        Next i
        
    End With
End Sub

Private Sub CheckFilterBox()
    Dim isAll As Boolean, isExistId As Boolean, isExistName As Boolean, isExistGender As Boolean, isExistCategory As Boolean
    Dim id As String, name As String, gender As String, category As String

    If Len(txtFilterProductId.Value) > 0 Then
        isExistId = True
        id = txtFilterProductId.Value
    Else
        isExistId = False
        id = ""
    End If
    If Len(txtFilterName.Value) > 0 Then
        isExistName = True
        name = txtFilterName.Value
    Else
        isExistName = False
        name = ""
    End If
    If cmbFilterGender.Value <> "All" Then
        isExistGender = True
        gender = cmbFilterGender.Value
    Else
        isExistGender = False
        gender = ""
    End If
    If cmbFilterCategory.Value <> "All" Then
        isExistCategory = True
        category = cmbFilterCategory.Value
    Else
        isExistCategory = False
        category = ""
    End If
    isAll = True
    If isExistId = True Or isExistName = True Or isExistGender = True Or isExistCategory = True Then
        isAll = False
    End If

    Call ResetTable(isAll, isExistId, id, isExistName, name, isExistGender, gender, isExistCategory, category)
    
End Sub

Private Sub txtFilterProductId_Change()
    Call CheckFilterBox
End Sub

Private Sub txtFilterName_Change()
    Call CheckFilterBox
End Sub

Private Sub cmbFilterGender_Change()
    Call CheckFilterBox
End Sub

Private Sub cmbFilterCategory_Change()
    Call CheckFilterBox
End Sub

Private Sub lblAddInventory_Click()
    Dim lastRow As Integer
    Dim productWs As Worksheet
    Dim tempId As String
    
    Worksheets("Product").Activate
    Set productWs = ActiveSheet
    
    frmInventoryDetails.Tag = "Add"
    frmInventoryDetails.lblFrmTitle.Caption = "Inventory Management - Add new product"
    
    frmInventoryDetails.txtProductId.Enabled = False
    frmInventoryDetails.cmbGender.List = Array("Men", "Women")
    frmInventoryDetails.lblOnsaleOff.Visible = False
    frmInventoryDetails.lblOnsaleOn.Visible = True
    
    lastRow = productWs.Cells(Rows.count, 1).End(xlUp).Row
    tempId = CStr(CInt(Left(productWs.Cells((lastRow), 1).Value, 4)) + 1) & "_"
    frmInventoryDetails.txtProductId.Value = tempId
    
    frmInventoryDetails.Show
End Sub

