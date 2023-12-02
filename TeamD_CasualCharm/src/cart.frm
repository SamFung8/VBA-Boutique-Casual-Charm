VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cart 
   Caption         =   "UserForm1"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   18888
   OleObjectBlob   =   "cart.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function NewPID(NewColor As String, OldID As String) As String
    Dim ChangePID As String
    Dim ChangeColorID As String
    If NewColor = "Black" Then
        ChangeColorID = "BK"
    ElseIf NewColor = "White" Then
        ChangeColorID = "WT"
    ElseIf NewColor = "Blue" Then
        ChangeColorID = "BE"
    End If
    ChangePID = Left(OldID, 5)
    NewPID = ChangePID & ChangeColorID
End Function
Private Sub cartCheckout_Click()
 If numShoppingCartItems > 0 Then
      Me.Hide
      checkout.Show
  Else
      MsgBox "Don't have enough items to check out!", vbExclamation, "Alert"
  End If

End Sub

Private Sub ComboBoxColor1_Change()
  
    'update the photo //todo...
    shoppingCart(0).selectedColor = ComboBoxColor1.Value
    shoppingCart(0).id = NewPID(shoppingCart(0).selectedColor, shoppingCart(0).id)
End Sub
Private Sub ComboBoxSize1_Change()
    shoppingCart(0).selectedSize = ComboBoxSize1.Value
End Sub
Private Sub ComboBoxColor2_Change()
 
    shoppingCart(1).selectedColor = ComboBoxColor2.Value
    shoppingCart(1).id = NewPID(shoppingCart(1).selectedColor, shoppingCart(1).id)
End Sub
Private Sub ComboBoxSize2_Change()
    shoppingCart(1).selectedSize = ComboBoxSize2.Value
End Sub
Private Sub ComboBoxColor3_Change()
   
    shoppingCart(2).selectedColor = ComboBoxColor3.Value
    shoppingCart(2).id = NewPID(shoppingCart(2).selectedColor, shoppingCart(2).id)
End Sub
Private Sub ComboBoxSize3_Change()
    shoppingCart(2).selectedSize = ComboBoxSize3.Value
End Sub
Private Sub ComboBoxColor4_Change()
  
    shoppingCart(3).selectedColor = ComboBoxColor4.Value
    shoppingCart(3).id = NewPID(shoppingCart(3).selectedColor, shoppingCart(3).id)
End Sub
Private Sub ComboBoxSize4_Change()
    shoppingCart(3).selectedSize = ComboBoxSize4.Value
End Sub

Private Sub Label1_Click()

End Sub




Private Sub lblrubbish1_Click()
Dim temp(9999) As New shoppingCartProduct
Dim deleteItemIndex As Integer
Dim index As Integer
Dim gap As Integer

deleteItemIndex = 0
gap = 0

For index = 0 To numShoppingCartItems - 1
    If deleteItemIndex <> index Then
        temp(index - gap).id = shoppingCart(index).id
        temp(index - gap).name = shoppingCart(index).name
        temp(index - gap).selectedColor = shoppingCart(index).selectedColor
        temp(index - gap).selectedSize = shoppingCart(index).selectedSize
        temp(index - gap).quantity = shoppingCart(index).quantity
        temp(index - gap).totalprice = shoppingCart(index).totalprice
        temp(index - gap).totalcost = shoppingCart(index).totalcost
    Else
        gap = 1
    End If
Next index

numShoppingCartItems = numShoppingCartItems - 1
Erase shoppingCart

For index = 0 To numShoppingCartItems - 1
    shoppingCart(index).id = temp(index).id
    shoppingCart(index).name = temp(index).name
    shoppingCart(index).selectedColor = temp(index).selectedColor
    shoppingCart(index).selectedSize = temp(index).selectedSize
    shoppingCart(index).quantity = temp(index).quantity
    shoppingCart(index).totalprice = temp(index).totalprice
    shoppingCart(index).totalcost = temp(index).totalcost
Next index
 
Dim covers() As Variant
covers = Array("cover1", "cover2", "cover3", "cover4") ' Add more covers as needed
Me.Controls(covers(0)).Visible = True
Me.Controls(covers(1)).Visible = True
Me.Controls(covers(2)).Visible = True
Me.Controls(covers(3)).Visible = True
    
Call cart.UserForm_Activate

End Sub
Private Sub lblrubbish2_Click()
Dim temp(9999) As New shoppingCartProduct
Dim deleteItemIndex As Integer
Dim index As Integer
Dim gap As Integer

deleteItemIndex = 1
gap = 0

For index = 0 To numShoppingCartItems - 1
    If deleteItemIndex <> index Then
        temp(index - gap).id = shoppingCart(index).id
        temp(index - gap).name = shoppingCart(index).name
        temp(index - gap).selectedColor = shoppingCart(index).selectedColor
        temp(index - gap).selectedSize = shoppingCart(index).selectedSize
        temp(index - gap).quantity = shoppingCart(index).quantity
        temp(index - gap).totalprice = shoppingCart(index).totalprice
        temp(index - gap).totalcost = shoppingCart(index).totalcost
    Else
        gap = 1
    End If
Next index

numShoppingCartItems = numShoppingCartItems - 1
Erase shoppingCart

For index = 0 To numShoppingCartItems - 1
    shoppingCart(index).id = temp(index).id
    shoppingCart(index).name = temp(index).name
    shoppingCart(index).selectedColor = temp(index).selectedColor
    shoppingCart(index).selectedSize = temp(index).selectedSize
    shoppingCart(index).quantity = temp(index).quantity
    shoppingCart(index).totalprice = temp(index).totalprice
    shoppingCart(index).totalcost = temp(index).totalcost
Next index
 
Dim covers() As Variant
covers = Array("cover1", "cover2", "cover3", "cover4") ' Add more covers as needed
Me.Controls(covers(0)).Visible = True
Me.Controls(covers(1)).Visible = True
Me.Controls(covers(2)).Visible = True
Me.Controls(covers(3)).Visible = True
    
Call cart.UserForm_Activate

End Sub
Private Sub lblrubbish3_Click()
Dim temp(9999) As New shoppingCartProduct
Dim deleteItemIndex As Integer
Dim index As Integer
Dim gap As Integer

deleteItemIndex = 2
gap = 0

For index = 0 To numShoppingCartItems - 1
    If deleteItemIndex <> index Then
        temp(index - gap).id = shoppingCart(index).id
        temp(index - gap).name = shoppingCart(index).name
        temp(index - gap).selectedColor = shoppingCart(index).selectedColor
        temp(index - gap).selectedSize = shoppingCart(index).selectedSize
        temp(index - gap).quantity = shoppingCart(index).quantity
        temp(index - gap).totalprice = shoppingCart(index).totalprice
        temp(index - gap).totalcost = shoppingCart(index).totalcost
    Else
        gap = 1
    End If
Next index

numShoppingCartItems = numShoppingCartItems - 1
Erase shoppingCart

For index = 0 To numShoppingCartItems - 1
    shoppingCart(index).id = temp(index).id
    shoppingCart(index).name = temp(index).name
    shoppingCart(index).selectedColor = temp(index).selectedColor
    shoppingCart(index).selectedSize = temp(index).selectedSize
    shoppingCart(index).quantity = temp(index).quantity
    shoppingCart(index).totalprice = temp(index).totalprice
    shoppingCart(index).totalcost = temp(index).totalcost
Next index
 
Dim covers() As Variant
covers = Array("cover1", "cover2", "cover3", "cover4") ' Add more covers as needed
Me.Controls(covers(0)).Visible = True
Me.Controls(covers(1)).Visible = True
Me.Controls(covers(2)).Visible = True
Me.Controls(covers(3)).Visible = True
    
Call cart.UserForm_Activate

End Sub
Private Sub lblrubbish4_Click()
Dim temp(9999) As New shoppingCartProduct
Dim deleteItemIndex As Integer
Dim index As Integer
Dim gap As Integer

deleteItemIndex = 3
gap = 0

For index = 0 To numShoppingCartItems - 1
    If deleteItemIndex <> index Then
        temp(index - gap).id = shoppingCart(index).id
        temp(index - gap).name = shoppingCart(index).name
        temp(index - gap).selectedColor = shoppingCart(index).selectedColor
        temp(index - gap).selectedSize = shoppingCart(index).selectedSize
        temp(index - gap).quantity = shoppingCart(index).quantity
        temp(index - gap).totalprice = shoppingCart(index).totalprice
        temp(index - gap).totalcost = shoppingCart(index).totalcost
    Else
        gap = 1
    End If
Next index

numShoppingCartItems = numShoppingCartItems - 1
Erase shoppingCart

For index = 0 To numShoppingCartItems - 1
    shoppingCart(index).id = temp(index).id
    shoppingCart(index).name = temp(index).name
    shoppingCart(index).selectedColor = temp(index).selectedColor
    shoppingCart(index).selectedSize = temp(index).selectedSize
    shoppingCart(index).quantity = temp(index).quantity
    shoppingCart(index).totalprice = temp(index).totalprice
    shoppingCart(index).totalcost = temp(index).totalcost
Next index
 
Dim covers() As Variant
covers = Array("cover1", "cover2", "cover3", "cover4") ' Add more covers as needed
Me.Controls(covers(0)).Visible = True
Me.Controls(covers(1)).Visible = True
Me.Controls(covers(2)).Visible = True
Me.Controls(covers(3)).Visible = True
    
Call cart.UserForm_Activate

End Sub

Private Sub UserForm_Initialize()
' first  cloth
    ComboBoxColor1.AddItem "Black"
    ComboBoxColor1.AddItem "White"
    ComboBoxColor1.AddItem "Blue"
    
    ComboBoxSize1.AddItem "S"
    ComboBoxSize1.AddItem "M"
    ComboBoxSize1.AddItem "L"
 ' second cloth
    ComboBoxColor2.AddItem "Black"
    ComboBoxColor2.AddItem "White"
    ComboBoxColor2.AddItem "Blue"
    
    ComboBoxSize2.AddItem "S"
    ComboBoxSize2.AddItem "M"
    ComboBoxSize2.AddItem "L"
    
 ' third
    ComboBoxColor3.AddItem "Black"
    ComboBoxColor3.AddItem "White"
    ComboBoxColor3.AddItem "Blue"

    ComboBoxSize3.AddItem "S"
    ComboBoxSize3.AddItem "M"
    ComboBoxSize3.AddItem "L"

 ' fourth
    ComboBoxColor4.AddItem "Black"
    ComboBoxColor4.AddItem "White"
    ComboBoxColor4.AddItem "Blue"

    ComboBoxSize4.AddItem "S"
    ComboBoxSize4.AddItem "M"
    ComboBoxSize4.AddItem "L"
    
    Me.totalprice.Caption = "$ 0"
    
End Sub


Public Sub UserForm_Activate()

 
   
    Dim clothNames() As Variant
    clothNames = Array("clothName1", "clothName2", "clothName3", "clothName4") ' Add more cloth names as needed
    
    Dim comboBoxColors() As Variant
    comboBoxColors = Array("ComboBoxColor1", "ComboBoxColor2", "ComboBoxColor3", "ComboBoxColor4") ' Add more ComboBox names for colors as needed
    
    Dim comboBoxSizes() As Variant
    comboBoxSizes = Array("ComboBoxSize1", "ComboBoxSize2", "ComboBoxSize3", "ComboBoxSize4") ' Add more ComboBox names for sizes as needed
    
    Dim clothTotals() As Variant
    clothTotals = Array("clothTotal1", "clothTotal2", "clothTotal3", "clothTotal4") ' Add more cloth totals as needed
    
    Dim clothPrices() As Variant
    clothPrices = Array("clothPrice1", "clothPrice2", "clothPrice3", "clothPrice4") ' Add more cloth prices as needed
    
    Dim covers() As Variant
    covers = Array("cover1", "cover2", "cover3", "cover4") ' Add more covers as needed
    
    Dim totalprice As Double
    Dim totalcost As Double
    
    If numShoppingCartItems > 0 Then
        For index = 0 To numShoppingCartItems - 1
            Me.Controls(clothNames(index)).Caption = shoppingCart(index).name
            Me.Controls(comboBoxColors(index)).Value = shoppingCart(index).selectedColor
            Me.Controls(comboBoxSizes(index)).Value = shoppingCart(index).selectedSize
            Me.Controls(clothTotals(index)).Caption = shoppingCart(index).quantity
            Me.Controls(clothPrices(index)).Caption = shoppingCart(index).totalprice
            
            ' Hide the corresponding cover
            Me.Controls(covers(index)).Visible = False
            
            
            ' Calculate total price
            totalprice = totalprice + shoppingCart(index).totalprice
            totalcost = totalcost + shoppingCart(index).totalcost
        Next index
        
        ' Update the TotalPrice label
        Me.totalprice.Caption = "$" & CStr(totalprice)
        currentBalance = totalprice
        currentCost = totalcost
    Else
        totalprice = 0
        Me.totalprice.Caption = "$" & CStr(totalprice)
        currentBalance = totalprice
        currentCost = totalcost
    End If
End Sub
Private Sub BackToHome_Click()

    Me.Hide
    HomePage.Show
    
End Sub

