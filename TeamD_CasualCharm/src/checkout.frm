VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} checkout 
   Caption         =   "UserForm1"
   ClientHeight    =   11925
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   16212
   OleObjectBlob   =   "checkout.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "checkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub lblMaster_Click()
    If lblmaster.BorderStyle = fmBorderStyleSingle Then
        lblmaster.BorderStyle = fmBorderStyleNone
    Else
        lblmaster.BorderStyle = fmBorderStyleSingle
        lblrupay.BorderStyle = fmBorderStyleNone
        lblvisa.BorderStyle = fmBorderStyleNone
    End If
End Sub

Private Sub lblproductlist_Click()
    Me.Hide
    HomePage.Show
End Sub

Private Sub lblRupay_Click()
    If lblrupay.BorderStyle = fmBorderStyleSingle Then
        lblrupay.BorderStyle = fmBorderStyleNone
    Else
        lblrupay.BorderStyle = fmBorderStyleSingle
        lblmaster.BorderStyle = fmBorderStyleNone
        lblvisa.BorderStyle = fmBorderStyleNone
    End If
End Sub

Private Sub lblVisa_Click()
    If lblvisa.BorderStyle = fmBorderStyleSingle Then
        lblvisa.BorderStyle = fmBorderStyleNone
    Else
        lblvisa.BorderStyle = fmBorderStyleSingle
        lblmaster.BorderStyle = fmBorderStyleNone
        lblrupay.BorderStyle = fmBorderStyleNone
    End If
End Sub



Private Sub UserForm_Activate()
    lblsubtotal.Caption = "$" & CStr(currentBalance)
    If currentBalance = 0 Then
        lblshippingfee.Caption = "$0"
    Else
        lblshippingfee.Caption = "$100"
    End If
        
    
End Sub

Private Sub UserForm_Initialize()

Dim name As String
Dim card_number As String
Dim expirationDate As String
Dim cvv As String
Dim addressess_name As String
Dim email As String
Dim phone As String

Dim transitionDay As Date
Dim shippingDay As Date
Dim Address As String

name = txtCardName.Text
card_number = cardNumber.Text
expirationDate = txtexpiration.Text
cvv = txtcvv.Text
addressess_name = txtName.Text
email = txtEmail.Text
phone = txtTel.Text
    
    transitionDay = Now
    shippingDay = DateAdd("d", 8, transitionDay)
    shippingDay = DateValue(Format(shippingDay, "yyyy-mm-dd")) + TimeValue("13:00:00")
    
    
    
    ' Set the value of lblshipping label
    lblshipping.Caption = shippingDay
    lblshipping.Enabled = False
    
   
    lblsubtotal.Caption = "$" & CStr(currentBalance)
    
    
'''''

        
        
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub btnCheckout_Click()
    ' perform date validation
    
    If lblvisa.BorderStyle = fmBorderStyleNone And _
        lblmaster.BorderStyle = fmBorderStyleNone And _
        lblrupay.BorderStyle = fmBorderStyleNone Then
        MsgBox "Please select a payment method"
        
    ElseIf Len(txtCardName) = 0 Then
        MsgBox "Please enter name!"
        
    ElseIf Len(cardNumber) <> 16 Then
        MsgBox "Invalid card number!"
        
    ElseIf Len(txtexpiration) < 3 Then
        MsgBox "Invalid expiration date"
        
    ElseIf Len(txtcvv) <> 3 And Len(txtcvv) <> 4 Then
        MsgBox "Invalid cvv"
        
    ElseIf Len(txtName) = 0 Then
        MsgBox "Invalid name"
        
    ElseIf InStr(1, txtEmail.Text, "@") = 0 Then
        MsgBox "Invalid email address!"
        
    ElseIf Len(txtTel.Text) = 0 Then
        MsgBox "Invalid phone number"
  
    ElseIf Len(txtAdress) = 0 Then
        MsgBox "Invalid address"
        
    Else
        MsgBox "payment success!"
    
        ' now process detail on the excel!
        
        ' order customr sheet
        Dim NewEntry As Range
        Dim LastEntry As Range
        
        Set LastEntry = Sheets("Order Customer").Range("A1").End(xlDown)
        Set NewEntry = LastEntry.Offset(1, 0)
        
        Dim CID_number As Integer
        Dim CID As String

        CID_number = CInt(Mid(LastEntry, 2)) + 1
        CID = "C" & CID_number
        
        With NewEntry
        
        .Offset(0, 0).Value = CID
        .Offset(0, 1).Value = txtName
        .Offset(0, 2).Value = txtEmail
        .Offset(0, 3).Value = txtAdress
        .Offset(0, 4).Value = txtTel
        End With
        
        ' order product sheet
        Set LastEntry = Sheets("Order Product").Range("A1").End(xlDown)
        Set NewEntry = LastEntry.Offset(1, 0)
        
        Dim OID_number As Integer
        Dim OID As String

        OID_number = CInt(Mid(LastEntry, 2)) + 1
        OID = "O" & OID_number
        
        Dim index As Integer
        Dim target As Range
        For index = 0 To numShoppingCartItems - 1
        With NewEntry
        
        .Offset(index, 0).Value = OID
        .Offset(index, 1).Value = CID
        .Offset(index, 2).Value = shoppingCart(index).id
        .Offset(index, 3).Value = shoppingCart(index).selectedSize
        .Offset(index, 4).Value = shoppingCart(index).quantity
        .Offset(index, 5).Value = shoppingCart(index).totalprice
        .Offset(index, 6).Value = shoppingCart(index).totalcost
        
        ' product sheet
        ' decrease the quatity
      
        Set target = Sheets("Product").Range("A:A").Find(shoppingCart(index).id)
        
        If shoppingCart(index).selectedSize = "S" Then
            target.Offset(0, 5) = target.Offset(0, 5) - shoppingCart(index).quantity '  handling for S
        ElseIf shoppingCart(index).selectedSize = "M" Then
            target.Offset(0, 6) = target.Offset(0, 6) - shoppingCart(index).quantity '  handling for M
        ElseIf shoppingCart(index).selectedSize = "L" Then
            target.Offset(0, 7) = target.Offset(0, 7) - shoppingCart(index).quantity ' handling for L
        End If
     
        
        End With
        Next index
        
        
        
        'order shipping sheet!!
        Set LastEntry = Sheets("Order Shipping").Range("A1").End(xlDown)
        Set NewEntry = LastEntry.Offset(1, 0)
        
        Dim SID_number As Integer
        Dim SID As String

        SID_number = CInt(Mid(LastEntry, 2)) + 1
        SID = "S" & SID_number
        
        Dim transitionDay As Date
        Dim shippingDay As Date
        
        transitionDay = Now
        shippingDay = DateAdd("d", 8, transitionDay)
        shippingDay = DateValue(Format(shippingDay, "yyyy-mm-dd")) + TimeValue("13:00:00")

        Dim PaymentMethod As String
        
        If lblrupay.BorderStyle = fmBorderStyleSingle Then
            PaymentMethod = "RuPay"
        ElseIf lblvisa.BorderStyle = fmBorderStyleSingle Then
            PaymentMethod = "Visa"
        ElseIf lblmaster.BorderStyle = fmBorderStyleSingle Then
            PaymentMethod = "Mastercard"
        End If
        
        Dim CoverCard As String
        CoverCard = "XXXX-XXXX-XXXX-" & Right(cardNumber, 4)
        
        With NewEntry
        .Offset(0, 0).Value = SID
        .Offset(0, 1).Value = CID
        .Offset(0, 2).Value = OID
        .Offset(0, 3).Value = transitionDay
        .Offset(0, 4).Value = shippingDay
        .Offset(0, 5).Value = "Preparing"
        .Offset(0, 6).Value = PaymentMethod
        .Offset(0, 7).Value = CoverCard
        .Offset(0, 8).Value = currentBalance
        .Offset(0, 9).Value = currentCost
        .Offset(0, 10).Value = currentBalance - currentCost
        ' will implement the cost in lastest sheet
        ' will implement the revence in lastest databse
        
        End With
        
        ' handle receipt front end...
        receipt.lblReceipt.Caption = OID
        receipt.lblCustomerNo.Caption = SID
        receipt.lblcustomerName.Caption = txtName
        receipt.lblCustomerTel.Caption = txtTel
        receipt.lblCustomerEmail.Caption = txtEmail
        receipt.lblTranscationDate.Caption = transitionDay
        receipt.lblshippingTime.Caption = shippingDay
        receipt.lblAddress.Caption = txtAdress
        receipt.lblpaymentmethod.Caption = PaymentMethod
        receipt.lblcardNo.Caption = CoverCard
        receipt.lblsubtotal = currentBalance
        receipt.lblshipping = "$100"
        receipt.lbltotalprice = currentBalance + 100
        
        ' handle the product part in receipt
        
       
        Dim PIDS() As Variant
        Dim Names() As Variant
        Dim Quantitys() As Variant
        Dim Categorys() As Variant
        Dim Styles() As Variant
        Dim Colors() As Variant
        Dim Sizes() As Variant
        Dim Prices() As Variant
        
        PIDS = Array("PID1", "PID2", "PID3", "PID4")
        Names = Array("Name1", "Name2", "Name3", "Name4")
        Quantitys = Array("quantity1", "quantity2", "quantity3", "quantity4")
        Categorys = Array("category1", "category2", "category3", "category4")
        Styles = Array("style1", "style2", "style3", "style4")
        Colors = Array("color1", "color2", "color3", "color4")
        Sizes = Array("size1", "size2", "size3", "size4")
        Prices = Array("price1", "price2", "price3", "price4")
        
        
        
        
    
    
    
        If numShoppingCartItems > 0 Then
            For index = 0 To numShoppingCartItems - 1
                 receipt.Controls(PIDS(index)).Caption = shoppingCart(index).id
                 receipt.Controls(Names(index)).Caption = shoppingCart(index).name
                 receipt.Controls(Quantitys(index)).Caption = shoppingCart(index).quantity
                 receipt.Controls(Categorys(index)).Caption = Application.WorksheetFunction.VLookup(shoppingCart(index).id, Worksheets("Product").UsedRange, 10, False)
                 receipt.Controls(Styles(index)).Caption = Application.WorksheetFunction.VLookup(shoppingCart(index).id, Worksheets("Product").UsedRange, 9, False)
                 receipt.Controls(Colors(index)).Caption = shoppingCart(index).selectedColor
                 receipt.Controls(Sizes(index)).Caption = shoppingCart(index).selectedSize
                 receipt.Controls(Prices(index)).Caption = shoppingCart(index).totalprice
                 
                 
                 
                 
                 
              
              
            Next index
        End If
            
    
        
        
       
        Me.Hide
        receipt.Show
        
    End If
    
End Sub



