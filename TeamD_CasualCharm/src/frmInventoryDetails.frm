VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventoryDetails 
   Caption         =   "Inventory Details - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmInventoryDetails.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmInventoryDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
frmReportSales.Show
End Sub

Private Sub lblSettingsTab_Click()
Unload Me
frmSettings.Show
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub


Private Sub btnSave_Click()
    Dim productWs As Worksheet
    Dim isGenderTrue As Boolean, isNameTrue As Boolean, isCategoryTrue As Boolean, isColorTrue As Boolean, isCostTrue As Boolean
    Dim isPriceTrue As Boolean, isQuantitySTrue As Boolean, isQuantityMTrue As Boolean, isQuantityLTrue As Boolean, isImgUrlTrue As Boolean
    Dim tempId As String, tempColor As String
    Dim lastRow As Integer
    
    isGenderTrue = True
    isNameTrue = True
    isCategoryTrue = True
    isColorTrue = True
    isCostTrue = True
    isPriceTrue = True
    isQuantitySTrue = True
    isQuantityMTrue = True
    isQuantityLTrue = True
    isImgUrlTrue = True
    
    Worksheets("Product").Activate
    Set productWs = ActiveSheet
    lastRow = productWs.Cells(Rows.count, 1).End(xlUp).Row
    
'check if enter value is right
' lblErrGender
If Len(cmbGender.Value) = 0 Then
    isGenderTrue = False
    lblErrGender.Visible = True
End If
If Len(txtProductName.Value) = 0 Then
    isNameTrue = False
    lblErrProductName.Visible = True
End If
If Len(cmbCategory.Value) = 0 Then
    isCategoryTrue = False
    lblErrCategory.Visible = True
End If
If Len(cmbColor.Value) = 0 Then
    isColorTrue = False
    lblErrColor.Visible = True
End If
If Len(txtCost.Value) = 0 Or IsInt(txtCost.Value) = False Then
    isCostTrue = False
    lblErrCost.Visible = True
End If
If Len(txtPrice.Value) = 0 Or IsInt(txtPrice.Value) = False Then
    isPriceTrue = False
    lblErrPrice.Visible = True
End If
If Len(txtQuantityS.Value) = 0 Or IsInt(txtQuantityS.Value) = False Then
    isQuantitySTrue = False
    lblErrQuantityS.Visible = True
End If
If Len(txtQuantityM.Value) = 0 Or IsInt(txtQuantityM.Value) = False Then
    isQuantityMTrue = False
    lblErrQuantityM.Visible = True
End If
If Len(txtQuantityL.Value) = 0 Or IsInt(txtQuantityL.Value) = False Then
    isQuantityLTrue = False
    lblErrQuantityL.Visible = True
End If
If Len(txtImgUrl.Value) = 0 Then
    isImgUrlTrue = False
    lblErrImgUrl.Visible = True
End If

Select Case Me.cmbColor.Value
        Case "White"
            tempColor = "WT"
        Case "Black"
            tempColor = "BK"
        Case "Blue"
            tempColor = "BE"
        Case Else
            tempColor = "WT"
        End Select
tempId = CStr(CInt(Left(productWs.Cells((lastRow), 1).Value, 4)) + 1) & "_" & tempColor

If frmInventoryDetails.Tag = "Edit" Then
    On Error Resume Next
        Call RequestDownload(Me.txtImgUrl.Value, Application.ActiveWorkbook.Path & "\product_img\" & Me.txtProductId.Value & ".jpeg")
        If Err Then
           lblErrImgUrl.Visible = True
           isImgUrlTrue = False
        Else
            lblErrImgUrl.Visible = False
            isImgUrlTrue = True
        End If
    On Error GoTo 0
Else
    On Error Resume Next
        Call RequestDownload(Me.txtImgUrl.Value, Application.ActiveWorkbook.Path & "\product_img\" & tempId & ".jpeg")
        If Err Then
           lblErrImgUrl.Visible = True
           isImgUrlTrue = False
        Else
            lblErrImgUrl.Visible = False
            isImgUrlTrue = True
        End If
    On Error GoTo 0
End If

' if ok
If isGenderTrue And isNameTrue And isCategoryTrue And isColorTrue And isCostTrue And isPriceTrue And isQuantitySTrue And isQuantityMTrue And isQuantityLTrue And isImgUrlTrue Then

    Dim i As Integer

    If frmInventoryDetails.Tag = "Edit" Then
        With productWs

            For i = 1 To lastRow
            
                If .Cells(i, 1).Value = Me.txtProductId.Value Then
                    
                    .Cells(i, 2).Value = Me.txtProductName.Value
                    .Cells(i, 3).Value = Me.txtCost.Value
                    .Cells(i, 4).Value = Me.txtPrice.Value
                    .Cells(i, 5).Value = Me.cmbColor.Value
                    .Cells(i, 6).Value = Me.txtQuantityS.Value
                    .Cells(i, 7).Value = Me.txtQuantityM.Value
                    .Cells(i, 8).Value = Me.txtQuantityL.Value
                    .Cells(i, 9).Value = Me.cmbGender.Value
                    .Cells(i, 10).Value = Me.cmbCategory.Value
                    If frmInventoryDetails.lblOnsaleOn.Visible = True Then
                        .Cells(i, 11).Value = "Y"
                    Else
                        .Cells(i, 11).Value = "N"
                    End If
                    
                    Me.imgProduct.Picture = LoadPicture(Application.ActiveWorkbook.Path & "\product_img\" & Me.txtProductId.Value & ".jpeg", 216, 198)
                    .Cells(i, 12).Value = Me.txtImgUrl.Value
                    Exit For
                End If
                
            Next i
        End With
        MsgBox "Save Successfully"
        Call frmInventory.ResetTable(True, False, "", False, "", False, "", False, "")
        Unload Me
        
    ElseIf frmInventoryDetails.Tag = "Add" Then
        Dim isExist As Boolean
        
        
        
        isExist = False
        With productWs
    
                .Cells((lastRow + 1), 1).Value = tempId
                .Cells((lastRow + 1), 2).Value = Me.txtProductName.Value
                .Cells((lastRow + 1), 3).Value = Me.txtCost.Value
                .Cells((lastRow + 1), 4).Value = Me.txtPrice.Value
                .Cells((lastRow + 1), 5).Value = Me.cmbColor.Value
                .Cells((lastRow + 1), 6).Value = Me.txtQuantityS.Value
                .Cells((lastRow + 1), 7).Value = Me.txtQuantityM.Value
                .Cells((lastRow + 1), 8).Value = Me.txtQuantityL.Value
                .Cells((lastRow + 1), 9).Value = Me.cmbGender.Value
                .Cells((lastRow + 1), 10).Value = Me.cmbCategory.Value
                If frmInventoryDetails.lblOnsaleOn.Visible = True Then
                    .Cells((lastRow + 1), 11).Value = "Y"
                Else
                    .Cells((lastRow + 1), 11).Value = "N"
                End If
                    
                Me.imgProduct.Picture = LoadPicture(Application.ActiveWorkbook.Path & "\product_img\" & tempId & ".jpeg", 216, 198)
                .Cells((lastRow + 1), 12).Value = Me.txtImgUrl.Value
                MsgBox "Add product successfully"
                Call frmInventory.ResetTable(True, False, "", False, "", False, "", False, "")
        End With
        
        Unload Me
        
    End If

End If



End Sub

Private Sub btnCheckImg_Click()
    Dim productWs As Worksheet
    Dim tempId As String, tempColor As String
    Dim lastRow As Integer, i As Integer
    Worksheets("Product").Activate
    Set productWs = ActiveSheet
    lastRow = productWs.Cells(Rows.count, 1).End(xlUp).Row
    Select Case Me.cmbColor.Value
    Case "White"
        tempColor = "WT"
    Case "Black"
        tempColor = "BK"
    Case "Blue"
        tempColor = "BE"
    Case Else
        tempColor = "WT"
    End Select
    tempId = CStr(CInt(Left(productWs.Cells((lastRow), 1).Value, 4)) + 1) & "_" & tempColor
    
    On Error Resume Next
    Call RequestDownload(Me.txtImgUrl.Value, Application.ActiveWorkbook.Path & "\product_img\" & tempId & ".jpeg")
    Me.imgProduct.Picture = LoadPicture(Application.ActiveWorkbook.Path & "\product_img\" & tempId & ".jpeg", 216, 198)
    If Err Then
       lblErrImgUrl.Visible = True
    Else
        lblErrImgUrl.Visible = False
    End If
    On Error GoTo 0
    
End Sub

Private Sub cmbCategory_Change()
    If Len(cmbCategory.Value) > 0 Then
        lblErrCategory.Visible = False
    End If
End Sub

Private Sub cmbColor_Change()
    Dim productWs As Worksheet
    Dim tempId As String, tempColor As String
    Dim lastRow As Integer, i As Integer
    Worksheets("Product").Activate
    Set productWs = ActiveSheet
    lastRow = productWs.Cells(Rows.count, 1).End(xlUp).Row
    Select Case Me.cmbColor.Value
    Case "White"
        tempColor = "WT"
    Case "Black"
        tempColor = "BK"
    Case "Blue"
        tempColor = "BE"
    Case Else
        tempColor = "WT"
    End Select
    tempId = CStr(CInt(Left(productWs.Cells((lastRow), 1).Value, 4)) + 1) & "_" & tempColor
    Me.txtProductId.Value = tempId
    If Len(cmbColor.Value) > 0 Then
        lblErrColor.Visible = False
    End If
    
End Sub

Private Sub cmbGender_Change()
    If Len(cmbGender.Value) > 0 Then
        lblErrGender.Visible = False
    End If
End Sub

Private Sub lblOnsaleOff_Click()
    lblOnsaleOff.Visible = False
    lblOnsaleOn.Visible = True
End Sub

Private Sub lblOnsaleOn_Click()
    lblOnsaleOff.Visible = True
    lblOnsaleOn.Visible = False
End Sub

Private Sub txtCost_Change()
    If Len(txtCost.Value) > 0 And IsInt(txtCost.Value) Then
        lblErrCost.Visible = False
    End If
End Sub

Private Sub txtImgUrl_Change()
    If Len(txtImgUrl.Value) > 0 Then
        lblErrImgUrl.Visible = False
    End If
End Sub

Private Sub txtPrice_Change()
    If Len(txtPrice.Value) > 0 And IsInt(txtCost.Value) Then
        lblErrPrice.Visible = False
    End If
End Sub

Private Sub txtProductName_Change()
    If Len(txtProductName.Value) > 0 Then
        lblErrProductName.Visible = False
    End If
End Sub

Private Sub txtQuantityS_Change()
    If Len(txtQuantityS.Value) > 0 And IsInt(txtQuantityS.Value) Then
        lblErrQuantityS.Visible = False
    End If
End Sub

Private Sub txtQuantityM_Change()
    If Len(txtQuantityM.Value) > 0 And IsInt(txtQuantityM.Value) Then
        lblErrQuantityM.Visible = False
    End If
End Sub

Private Sub txtQuantityL_Change()
    If Len(txtQuantityL.Value) > 0 And IsInt(txtQuantityL.Value) Then
        lblErrQuantityL.Visible = False
    End If
End Sub


