VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddUser 
   Caption         =   "Add User"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   OleObjectBlob   =   "frmAddUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    
    If Len(txtUserID.Value) <> 0 Then
        If MsgBox(prompt:="User ID copied?", Title:="Copy user id", Buttons:=vbYesNo) = vbYes Then
            Unload Me
        End If
    ElseIf Len(Trim(txtFirstName.Value)) <> 0 Or Len(Trim(txtLastName.Value)) <> 0 Or Len(cboLevel.Value) <> 0 Then
        If MsgBox(prompt:="continue adding user?", Title:="Process of adding user", Buttons:=vbYesNo) = vbNo Then
            Unload Me
        End If
    Else
        If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
            Unload Me
            frmLoanMenu.Show
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    
    If Len(Trim(txtFirstName)) = 0 Or Len(Trim(txtLastName)) = 0 Or Len(Trim(cboLevel)) = 0 Then
        MsgBox prompt:="All User Information fields are required", Title:="missing information - verify"
        Exit Sub
    End If
    
    If Not IsNumeric(txtPasscode) Or Len(Trim(CStr(txtPasscode))) < 4 Then
        MsgBox prompt:="Passcode must be at least 4 NUMERIC characters", Title:="Verifying Passcode"
        txtPasscode.Value = vbNullString
        txtPasscodeConfirm.Value = vbNullString
        txtPasscode.SetFocus
        Exit Sub
    ElseIf txtPasscode <> txtPasscodeConfirm Then
        MsgBox prompt:="Passcode confirmation failed", Title:="Verifying Passcode"
        txtPasscodeConfirm.Value = vbNullString
        txtPasscodeConfirm.SetFocus
        Exit Sub
    End If
    
    Dim ithUser As Long
    ithUser = ThisWorkbook.Sheets("user").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    With ThisWorkbook.Sheets("user")
        .Range("A" & ithUser).Value = Trim(CStr(txtFirstName.Value))
        .Range("B" & ithUser).Value = Trim(CStr(txtLastName.Value))
        .Range("C" & ithUser).Value = cboLevel.Value
        .Range("D" & ithUser).Value = LCase(CStr(Mid(txtFirstName.Value, 1, 1) & _
            Mid(txtLastName.Value, 1, 3) & _
            Format(ithUser, "00#") _
            ))
        .Range("E" & ithUser).Value = txtPasscode.Value
        .Range("F" & ithUser).Value = Now()
        .Range("G" & ithUser).Value = "Yes"
        txtUserID.Value = .Range("D" & ithUser).Value
    End With
    
    MsgBox prompt:="User successfully added - Save User ID", Title:="Success Information"
    With Me
        .txtFirstName.Value = ""
        .txtLastName.Value = ""
        .cboLevel.Value = ""
        .txtPasscode.Value = ""
        .txtPasscodeConfirm.Value = ""
    End With

End Sub

Private Sub txtFirstName_AfterUpdate()
    Me.txtFirstName.Value = Trim(Me.txtFirstName.Value)
End Sub

Private Sub txtLastName_AfterUpdate()
    Me.txtLastName.Value = Trim(Me.txtLastName.Value)
End Sub

Private Sub UserForm_Initialize()
    With Me.cboLevel
        .AddItem "Supervisor"
        .AddItem "Representative"
        .AddItem "Analyst"
        .AddItem "Strategist"
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub
