VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoanMenu 
   Caption         =   "Menu List TIP"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9990
   OleObjectBlob   =   "frmLoanMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoanMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNewClient_Click()
    switchAddEdit = False
    Unload Me
    frmAddClient.Show
End Sub

Private Sub cmdAddNewUser_Click()
    Unload Me
    frmAddUser.Show
End Sub

Private Sub cmdCompareLoan_Click()
    Unload Me
    frmLoanCompare.Show
End Sub

Private Sub cmdExistingClient_Click()
    Unload Me
    frmClientInfoLogIn.Show
End Sub

Private Sub cmdEstimateLoan_Click()
    Unload Me
    frmLoanCalculator.Show
End Sub

Private Sub cmdUpdateUser_Click()
    Unload Me
    frmEditUser.Show
End Sub

Private Sub cmdLogOffMenu_Click()
    If MsgBox(prompt:="You are about to log-off the system", Title:="Loging Off", Buttons:=vbOKCancel + vbInformation) = vbOK Then
        Unload Me
        
        Call frmUserLogOn.cmdLogOff_Click
        frmUserLogOn.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim ithUserID As Long
    ithUserID = Application.Match(CStr(Trim(frmUserLogOn.txtUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)
    With Me
        .lblLevel.Caption = UCase(frmUserLogOn.cboLevel.Value)
        .lblName.Caption = UCase(ThisWorkbook.Sheets("user").Range("D" & ithUserID).Value & " " & _
           ThisWorkbook.Sheets("user").Range("D" & ithUserID).Offset(0, -3).Value & " " & _
            ThisWorkbook.Sheets("user").Range("D" & ithUserID).Offset(0, -2).Value)
    End With
    
    'just for testing code...delete later
    'Me.cmdExistingClient.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
