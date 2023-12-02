VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings - Admin"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17688
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSettings"
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

Private Sub UserForm_Initialize()
    Call resetForm
    
End Sub

Private Sub btnSave_Click()
    Dim isUsernameTrue As Boolean, isNameTrue As Boolean, isEmailTrue As Boolean, isPhoneTrue As Boolean, isOldPasswordTrue As Boolean, isNewPasswordTrue As Boolean
    
    isUsernameTrue = False
    isNameTrue = False
    isEmailTrue = False
    isPhoneTrue = False
    isOldPasswordTrue = False
    isNewPasswordTrue = False
    
    If Len(Me.txtUsername.Value) > 0 Then
        isUsernameTrue = True
    Else
        Me.lblErrorUsername.Visible = True
    End If
    If Len(Me.txtName.Value) > 0 Then
        isNameTrue = True
    Else
        Me.lblErrorName.Visible = True
    End If
    If Len(Me.txtEmail.Value) > 0 Then
        isEmailTrue = True
    Else
         Me.lblErrorEmail.Visible = True
    End If
    If Len(Me.txtPhone.Value) > 0 Then
        isPhoneTrue = True
    Else
        Me.lblErrorPhone.Visible = True
    End If
    If Len(Me.txtOldPassword.Value) > 0 Then
        isOldPasswordTrue = True
    Else
        Me.lblErrorOldPassword.Visible = True
    End If
    If Len(Me.txtNewPassword.Value) > 0 Then
        isNewPasswordTrue = True
    End If
    
    If isUsernameTrue And isNameTrue And isEmailTrue And isPhoneTrue And isOldPasswordTrue Then
        If Me.txtOldPassword.Value = Me.txtOldPassword.Tag Then
            With Sheets("Staff Info")
                .Cells(Me.lblStaffId.Tag, 2).Value = Me.txtName.Value
                .Cells(Me.lblStaffId.Tag, 4).Value = Me.txtPhone.Value
                .Cells(Me.lblStaffId.Tag, 5).Value = Me.txtEmail.Value
                .Cells(Me.lblStaffId.Tag, 6).Value = Me.txtUsername.Value
                
            End With
            
            If isNewPasswordTrue Then
                Sheets("Staff Info").Cells(Me.lblStaffId.Tag, 7).Value = Me.txtNewPassword.Value
            End If
            MsgBox "Settings Saved"
            Call resetForm
        Else
            isOldPasswordTrue = True
        End If
        
        
    End If
    
   
    
    
End Sub

Private Sub resetForm()
Dim TargetCell As Range
    Dim id As String

    If Len(loginStaffId) > 0 Then
        id = loginStaffId
    Else
        id = "S_1000"
    End If
    Set TargetCell = Range("Table27").Find(What:=id, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    
    Me.lblErrorUsername.Visible = False
    Me.lblErrorName.Visible = False
    Me.lblErrorEmail.Visible = False
    Me.lblErrorPhone.Visible = False
    Me.lblErrorOldPassword.Visible = False
    
    Me.lblStaffId.Caption = "Staff ID : " & id
    Me.lblStaffId.Tag = TargetCell.Row
    Me.lblStaffPost.Caption = Sheets("Staff Info").Cells(TargetCell.Row, 3)
    Me.txtUsername.Value = Sheets("Staff Info").Cells(TargetCell.Row, 6)
    Me.txtName.Value = Sheets("Staff Info").Cells(TargetCell.Row, 2)
    Me.txtEmail.Value = Sheets("Staff Info").Cells(TargetCell.Row, 5)
    Me.txtPhone.Value = Sheets("Staff Info").Cells(TargetCell.Row, 4)
    Me.txtOldPassword.Tag = Sheets("Staff Info").Cells(TargetCell.Row, 7)
    Me.txtOldPassword.Value = ""
    Me.txtNewPassword.Value = ""
    
End Sub
