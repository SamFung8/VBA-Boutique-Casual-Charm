VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderDetails 
   Caption         =   "Order Details - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmOrderDetails.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmOrderDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()
Unload frmOrderDetails
frmOrder.Show
End Sub

Private Sub btnSave_Click()


'check if enter value is right
If Me.cboProduct.Value = "" Or Me.cboProduct.Value = "Select Product" Then
    MsgBox "Please select a valid Product ID !"
    Exit Sub
ElseIf (Me.txtSize.Value <> "S" And Me.txtSize.Value <> "M" And Me.txtSize.Value <> "L") Or Me.txtSize.Value = "" Then
    MsgBox "Please enter valid Size (S / M / L) !"
    Exit Sub
ElseIf Me.txtSubtotal.Value = "" Or Me.txtSubtotal.Value < 0 Then
    MsgBox "Please enter correct number of subtotal amount !"
    Exit Sub
ElseIf Me.txtTotal.Value = "" Or Me.txtTotal.Value < 0 Then
    MsgBox "Please enter correct number of total amount !"
    Exit Sub
ElseIf Me.txtQuantity.Value = "" Or Me.txtQuantity.Value < 0 Then
    MsgBox "Please enter valid number for quantity !"
    Exit Sub
ElseIf Me.cboStatus.Value = "" Or (Me.cboStatus.Value <> "Preparing" And Me.cboStatus.Value <> "In Transit" And Me.cboStatus.Value <> "Shipped") Then
    MsgBox "Please enter valid status (Preparing / In Transit / Shipped)"
    Exit Sub
End If

' if ok
    
    Set orderproductWs = Sheets("Order Product")
    Set ordershippingWs = Sheets("Order Shipping")
    With orderproductWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row

        For i = 1 To lastRow
            If .Cells(i, 1).Value = Me.txtOrderid And .Cells(i, 3).Value = Me.cboProduct Then
                .Cells(i, 2).Value = frmOrderDetails.txtCustomerId
                .Cells(i, 4).Value = frmOrderDetails.txtSize
                .Cells(i, 6).Value = frmOrderDetails.txtSubtotal.Value
                .Cells(i, 5).Value = frmOrderDetails.txtQuantity
                Exit For
            End If
        Next i
    End With
    
    With ordershippingWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row

        For i = 1 To lastRow
            If .Cells(i, 3).Value = txtOrderid.Value Then
                .Cells(i, 6).Value = frmOrderDetails.cboStatus.Value
                Exit For
            End If
        Next i
    End With
    
MsgBox "Save Successfully"
Unload Me
End Sub

Private Sub cboProduct_Change()

Set orderproductWs = ThisWorkbook.Sheets("Order Product")
Set productWs = ThisWorkbook.Sheets("Product")

With orderproductWs
        lastRow = .Cells(Rows.count, 1).End(xlUp).Row

        For i = 1 To lastRow
            If .Cells(i, 1).Value = Me.txtOrderid.Value And Me.cboProduct.Value = .Cells(i, 3).Value Then
                Me.txtSize.Value = .Cells(i, 4).Value
                Me.txtQuantity.Value = .Cells(i, 5).Value
                Me.txtSubtotal.Value = .Cells(i, 6).Value
                Exit For
            End If
            
        Next i
    End With
    
With productWs
lastRow = .Cells(Rows.count, 1).End(xlUp).Row

        For i = 1 To lastRow
            If .Cells(i, 1).Value = Me.cboProduct.Value Then
                Me.lblPname = .Cells(i, 2).Value & " - " & .Cells(i, 5).Value
                Exit For
            End If
            
        Next i
End With

End Sub

