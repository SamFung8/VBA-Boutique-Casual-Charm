VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminLogin 
   Caption         =   "Login - Admin"
   ClientHeight    =   10035
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17700
   OleObjectBlob   =   "frmAdminLogin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub lblBackToPreviosPage_Click()
    Unload Me
    IndexPage.Show
End Sub


Private Sub lblLogin_Click()

If Not IsError(Application.Match(txtUsername.Value, Sheets("Staff Info").Range("F2:F999"), 0)) Then
    If txtPw.Value = Application.WorksheetFunction.index(Sheets("Staff Info").Range("F2:G999"), Application.Match(txtUsername.Value, Sheets("Staff Info").Range("F2:F999"), 0), 2) Then
        Dim TargetCell As Range
        Set TargetCell = Range("Table27").Find(What:=txtUsername.Value, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)

        loginStaffId = Sheets("Staff Info").Cells(TargetCell.Row, 1)
        MsgBox "Login Successful !"
        
        Unload Me
        Dim Form As frmAdminHome
        Set Form = New frmAdminHome
        Form.Show
   
    Else
        Me.lblErrorMsg = "Invalid Username or Password"
        Me.txtPw.Value = ""
        Me.txtUsername.Value = ""
    End If
Else
   Me.lblErrorMsg = "Invalid Username or Password"
    Me.txtPw.Value = ""
    Me.txtUsername.Value = ""
End If

End Sub

