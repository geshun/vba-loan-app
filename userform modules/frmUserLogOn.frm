VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserLogOn 
   Caption         =   "Log On to TIP Loan Management System"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15030
   OleObjectBlob   =   "frmUserLogOn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ithUserID As Long

Private Sub cboLevel_Change()
    If Me.cboLevel.Value = ThisWorkbook.Sheets("user").Range("C" & ithUserID).Value Then
        With Me
            .lblPasscode.Visible = True
            .txtPasscode.Visible = True
            .txtPasscode.SetFocus
        End With
    Else
        With Me
            If .cmdConfirm.Visible Then
                .cmdConfirm.Visible = False
            End If
            .txtPasscode.Value = vbNullString
            .lblPasscode.Visible = False
            .txtPasscode.Visible = False
        End With
    End If
End Sub

Private Sub cmdConfirm_Click()
    If Me.cboLevel.Value = "admin" Then
        Application.Visible = True
    End If
    
    Dim ithLogOn As Long
    ithLogOn = ThisWorkbook.Sheets("user_logon_logoff_TS").Range("A" & Rows.Count).End(xlUp).Row + 1
    With ThisWorkbook.Sheets("user_logon_logoff_TS")
        .Range("A" & ithLogOn).Value = UCase(Me.txtUserID.Value)
        .Range("B" & ithLogOn).Value = "LN"
        .Range("C" & ithLogOn).Value = "Log On"
        .Range("D" & ithLogOn).Value = Now
    End With
    ThisWorkbook.Save
    Me.Hide
    frmLoanMenu.Show
End Sub

Public Sub cmdLogOff_Click()
    Dim ithLogOff As Long
    ithLogOff = ThisWorkbook.Sheets("user_logon_logoff_TS").Range("A" & Rows.Count).End(xlUp).Row + 1
    With ThisWorkbook.Sheets("user_logon_logoff_TS")
        .Range("A" & ithLogOff).Value = UCase(Me.txtUserID.Value)
        .Range("B" & ithLogOff).Value = "LO"
        .Range("C" & ithLogOff).Value = "Log Off"
        .Range("D" & ithLogOff).Value = Now()
    End With
    ThisWorkbook.Save
    Call UserForm_Initialize
End Sub

Private Sub txtPasscode_Change()
    If Me.txtPasscode.Value = CStr(ThisWorkbook.Sheets("user").Range("E" & ithUserID).Value) Then
        With Me
            .cmdConfirm.Visible = True
            .cmdLogOff.Visible = True
            .cmdConfirm.SetFocus
        End With
    Else
        Me.cmdConfirm.Visible = False
        Me.cmdLogOff.Visible = False
    End If
End Sub

Private Sub txtUserID_AfterUpdate()
    If Me.txtUserID.Value <> "ge001" Then
        ThisWorkbook.Sheets("user_level").Activate
     
        Dim numOfLevels As Integer
        numOfLevels = ThisWorkbook.Sheets("user_level").Range("A" & Rows.Count).End(xlUp).Row
     
        Dim levelRange As Range
        Set levelRange = ThisWorkbook.Sheets("user_level").Range("A3:A" & numOfLevels)
        Me.cboLevel.RowSource = levelRange.Address
    End If
End Sub

Private Sub txtUserID_Change()
    
    If Not IsError(Application.Match(CStr(Trim(Me.txtUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)) Then
        
        ithUserID = Application.Match(CStr(Trim(Me.txtUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)
        
        If ThisWorkbook.Sheets("user").Range("G" & ithUserID).Value = "Yes" Then
            With Me
                .lblLevel.Visible = True
                .cboLevel.Visible = True
                .cboLevel.SetFocus
            End With
        End If
        
    Else
        With Me
            If .lblPasscode.Visible Then
                If .cmdConfirm.Visible Then
                    .cmdConfirm.Visible = False
                End If
                .txtPasscode.Value = vbNullString
                .lblPasscode.Visible = False
                .txtPasscode.Visible = False
            End If
            .cboLevel.Value = vbNullString
            .lblLevel.Visible = False
            .cboLevel.Visible = False
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
    
     ThisWorkbook.Sheets("user_level").Activate
     
     Dim numOfLevels As Integer
     numOfLevels = ThisWorkbook.Sheets("user_level").Range("A" & Rows.Count).End(xlUp).Row
     
     Dim levelRange As Range
     Set levelRange = ThisWorkbook.Sheets("user_level").Range("A2:A" & numOfLevels)
     Me.cboLevel.RowSource = levelRange.Address
     
     With Me
        
        .lblCurrentTime.Caption = CStr(Format(Now, "Mmm-dd-yyyy HH:mm:ss"))
        
        .lblUserID.Visible = True
        .txtUserID.Visible = True
        
        
        .lblLevel.Visible = False
        .cboLevel.Visible = False
        
        .lblPasscode.Visible = False
        .txtPasscode.Visible = False
        
        .cmdConfirm.Visible = False
        .txtUserID.Value = vbNullString
        
        .cmdLogOff.Visible = False
    End With
    
    'just for testing code...delete later
    Me.txtUserID.Value = "ge001"
    Me.cboLevel.Value = "admin"
    Me.txtPasscode.Value = 3030
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = 0 Then
'        ThisWorkbook.Save
'        ThisWorkbook.Close
'    End If
'End Sub


