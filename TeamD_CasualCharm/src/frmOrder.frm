VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrder 
   Caption         =   "Order Management - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmOrder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OrderEvent As New Collection
Function LoadPNGImg(ByVal Filename As String) As StdPicture
        With CreateObject("WIA.ImageFile")
            .LoadFile Filename
            Set LoadPNGImg = .FileData.Picture
        End With
End Function

Private Sub cboFrom_Change()
On Error Resume Next
Me.frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#
    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            If cboFrom.Value <> "" And cboTo.Value <> "" And i > 1 Then
            
            If CDbl(orderWs.Cells(i, 4).Value) >= CDbl(CDate(cboFrom.Value)) And CDbl(orderWs.Cells(i, 4).Value) < CDbl(CDate(cboTo.Value)) Then
            count = count + 1
            For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
        End If
            
        ElseIf cboFrom.Value = "" Or cboTo.Value = "" Then
        
         For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
              Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
     
        End If
        
        Next i
        
    End With




End Sub


Private Sub cboTo_Change()
On Error Resume Next
Me.frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#

    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If cboFrom.Value <> "" And cboTo.Value <> "" And i > 1 Then
            
            If CDbl(orderWs.Cells(i, 4).Value) >= CDbl(CDate(cboFrom.Value)) And CDbl(orderWs.Cells(i, 4).Value) < CDbl(CDate(cboTo.Value)) Then

            count = count + 1
            For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            End If
            End If
            
        ElseIf cboFrom.Value = "" Or cboTo.Value = "" Then
       
        
         For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
       
        End If
        Next i
        
    End With
End Sub

Private Sub Label7_Click()
Unload Me
IndexPage.Show
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

Private Sub lblReportsTab_Click()
Unload Me
frmReport.Show
End Sub

Private Sub lblSettingsTab_Click()
Unload Me
frmSettings.Show
End Sub

Private Sub cboDelivery_Change()
On Error Resume Next

frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If orderWs.Cells(i, 6).Value = cboDelivery.Value Then
            count = count + 1
            For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            End If
            
        ElseIf cboDelivery = "" Then
        
         For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           
             Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            End If
       
        End If
        Next i
        
    End With



End Sub


Private Sub cboPayment_Change()
On Error Resume Next
frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If orderWs.Cells(i, 7).Value = cboPayment.Value Then
            count = count + 1
            For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            End If
            
        ElseIf cboPayment = "" Then
        
         For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
        
        End If
        Next i
        
    End With

End Sub

Private Sub txtCustname_Change()
On Error Resume Next
frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If Application.WorksheetFunction.VLookup(orderWs.Cells(i, 2).Value, Sheets("Order Customer").Range("A:B"), 2, False) = txtCustname.Value Then
            count = count + 1
            For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
            
        ElseIf txtCustname = "" Then
        
         For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            End If
       
        End If
        Next i
        
    End With

End Sub

Private Sub txtFrom_Change()
On Error Resume Next
frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim count As Integer 'count for the filtered items#

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 1 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If CDate(txtFrom) > orderWs.Cells(i, 4) Then
            count = count + 1
            For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (count)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (count)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (count) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
            
        ElseIf txtCustname = "" Then
        
         For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
        
        End If
        Next i
        
    End With

End Sub

Private Sub txtOrderid_Change()
On Error Resume Next

frameOrderTable.Controls.Clear

 Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 2 To lastRow
           Dim actionEvent As OrderEventHandler
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            If orderWs.Cells(i, 3).Value = txtOrderid.Value Then
            For j = 1 To (lastCol - 1)
               If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
                   
                
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
            
            End If
            
        ElseIf txtOrderid = "" Then
        
        
        For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
           
            
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
                End With
           Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
        
            
            End If
        
        
        End If
        Next i
        
    End With

End Sub



Private Sub UserForm_Initialize()
     On Error Resume Next
    Dim orderWs As Worksheet
    Dim lastRow As Integer, lastCol As Integer
    Dim arrDates() As Variant
    Dim startD, endD As Date
    Dim actionEvent As OrderEventHandler
    
    startD = CDate("1/1/2022")
    endD = Date
    arrDates = Application.Transpose(Evaluate("INDEX(TEXT(""" & startD & """+ROW(1:" & endD - startD + 1 & ")-1,""DD/MM/YYYY""),)"))
    
    With cboFrom
    .List = arrDates
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    End With
    
    With cboTo
    .List = arrDates
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    End With
    
    With cboPayment
    '.RowSource = "Dictionary!I2:I5"
    .List = Sheets("Dictionary").Range("I2:I5").Value
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    End With
    
    With cboDelivery
    '.RowSource = "Dictionary!G2:G5"
    .List = Sheets("Dictionary").Range("G2:G5").Value
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    .BackColor = RGB(218, 227, 243)
    End With
    
    With txtOrderid
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    End With
    
    With txtCustname
    .FontSize = 12
    .FontName = "Calibri"
    .ForeColor = RGB(77, 85, 122)
    End With
    

    
    Set orderWs = ThisWorkbook.Worksheets("Order Shipping")
    With orderWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row
        frmOrder.frameOrderTable.ScrollHeight = 0
        For i = 2 To lastRow
            
            lastCol = orderWs.Cells(i, Columns.count).End(xlToLeft).Column
            
            For j = 1 To (lastCol - 1)
                If i > 1 And (j = 2 Or j = 3 Or j = 4 Or j = 6 Or j = 7) Then
                    Set lblProductId = frameOrderTable.Controls.Add("Forms.Label.1", "lblProductId" & CStr(i))
                    
                    With lblProductId
                        .Caption = orderWs.Cells(i, j).Value
                        .Top = 37 * (i - 1)
                        .TextAlign = fmTextAlignCenter
                        .AutoSize = False
                        .WordWrap = False
                        .Height = 20
                        .FontSize = 11
                    End With
                    
                    If j = 2 Then
                        'lblProductId.Caption = orderWs.Cells(i, j).Value
                        lblProductId.Caption = orderWs.Cells(i, j).Value & " " & Application.WorksheetFunction.VLookup(orderWs.Cells(i, j).Value, Sheets("Order Customer").Range("A:B"), 2, False)
                        lblProductId.Left = 10
                        lblProductId.Width = 75
                        orderWs.Activate
                    ElseIf j = 3 Then
                        lblProductId.Left = 125
                        lblProductId.Width = 50
                    ElseIf j = 4 Then
                        lblProductId.Left = 200
                        lblProductId.Width = 100
                    ElseIf j = 6 Then
                        lblProductId.Left = 110 * 3
                        lblProductId.Width = 75
                    ElseIf j = 7 Then
                        lblProductId.Left = 100 * 4
                        lblProductId.Width = 90
                    End If
                   
                   
                End If
            Next j
            
        
            With frmOrder.frameOrderTable
                .ScrollHeight = .ScrollHeight + 37
            End With
             
            
            
            If i > 1 Then
            ' Edit button
            Set lblActionEdit = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionEdit" & CStr(i))
                With lblActionEdit
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_pencil_green.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    '.Picture = LoadPicture(editIconFilename)
                    .Caption = "edit"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 114 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            
            
            
            ' Delete Button
            
            Set lblActionDel = frameOrderTable.Controls.Add("Forms.Label.1", "lblActionDel" & CStr(i))
                With lblActionDel
                    .Picture = LoadPNGImg(Application.ActiveWorkbook.Path & "\assets\icon_delete.png")
                    .Tag = orderWs.Cells(i, 3).Value
                    '.Picture = LoadPicture(delIconFilename)
                    .Caption = "del"
                    .Top = 37 * (i - 1)
                    .TextAlign = fmTextAlignCenter
                    .AutoSize = False
                    .WordWrap = False
                    .Height = 15
                    .Width = 15
                    .Left = 120 * 5
                    .MousePointer = 99
                    .MouseIcon = LoadPicture(Application.ActiveWorkbook.Path & "\assets\cursor_pointer.ico")
            End With
            ' Row Divider
            Set lblRowDivider = frameOrderTable.Controls.Add("Forms.Label.1", "lblRowDivider" & CStr(i))
                With lblRowDivider
                    .Top = 37 * (i - 1) - 12
                    .BackColor = &H8000000F
                    .Height = 2
                    .Width = 700
            End With
                
            Set actionEvent = New OrderEventHandler
            Set actionEvent.EditLabel = lblActionEdit
            Set actionEvent.DelLabel = lblActionDel
            OrderEvent.Add actionEvent
                           
            
            End If
        Next i
        
    End With
    
    
End Sub

