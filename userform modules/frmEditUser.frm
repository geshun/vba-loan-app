VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditUser 
   Caption         =   "Edit User"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
   OleObjectBlob   =   "frmEditUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public numOfChanges As Integer
Dim arrayOne() As Variant

Private Sub txtFirstName_AfterUpdate()
    If Me.togUpdateVerified.Enabled = True Then
        numOfChanges = numOfChanges + 1
        Me.cmdUpdateChanges.Enabled = True
    End If
End Sub

Private Sub txtLastName_AfterUpdate()
    If Me.togUpdateVerified.Enabled = True Then
        numOfChanges = numOfChanges + 1
        Me.cmdUpdateChanges.Enabled = True
    End If
End Sub

Private Sub cboLevel_afterupdate()
    If Me.togUpdateVerified.Enabled = True Then
        numOfChanges = numOfChanges + 1
        Me.cmdUpdateChanges.Enabled = True
    End If
End Sub

Private Sub txtPasscode_AfterUpdate()
    If Me.togUpdateVerified.Enabled = True Then
        numOfChanges = numOfChanges + 1
        Me.cmdUpdateChanges.Enabled = True
    End If
End Sub

Private Sub chkBlackListUser_AfterUpdate()
    If Me.togUpdateVerified.Enabled = True Then
        numOfChanges = numOfChanges + 1
        Me.cmdUpdateChanges.Enabled = True
    End If
End Sub

Private Sub cboUserID_Change()
   
    ReDim arrayOne(4) As Variant
    
    Dim ithUserID As Long
    
    If Not IsError(Application.Match(CStr(Trim(Me.cboUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)) Then
        
        ithUserID = Application.Match(CStr(Trim(Me.cboUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)
    Else
        Exit Sub
    End If
    
    With ThisWorkbook.Sheets("user")
        Me.txtFirstName.Value = .Range("A" & ithUserID).Value
        Me.txtLastName.Value = .Range("B" & ithUserID).Value
        Me.cboLevel.Value = .Range("C" & ithUserID).Value
        Me.txtPasscode.Value = .Range("E" & ithUserID).Value
        
        arrayOne(0) = .Range("A" & ithUserID).Value
        arrayOne(1) = .Range("B" & ithUserID).Value
        arrayOne(2) = .Range("C" & ithUserID).Value
        arrayOne(3) = .Range("E" & ithUserID).Value
        arrayOne(4) = .Range("G" & ithUserID).Value
        
        If .Range("G" & ithUserID).Value = "No" Then
            Me.chkBlackListUser.Value = True
            MsgBox prompt:="This user is blacklisted" + vbNewLine + "User is not authorized to use this system", _
                    Title:="Unauthorized user - Warning", Buttons:=vbCritical
        Else
            Me.chkBlackListUser.Value = False
        End If
    End With
    
    If Len(Me.cboUserID.Value) <> 0 Then
        Me.togUpdateVerified.Enabled = True
    Else
        Me.togUpdateVerified.Value = False
        Me.togUpdateVerified.Enabled = False
    End If
    
End Sub

Private Sub cmdExit_Click()
    If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        frmLoanMenu.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdLogOffEditUser_Click()
    If MsgBox(prompt:="You are about to log-off the system", Title:="Loging Off", Buttons:=vbOKCancel) = vbOK Then
        Unload Me
        
        Call frmUserLogOn.cmdLogOff_Click
        frmUserLogOn.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdReset_Click()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            ctrl.Value = ""
        End If
    Next ctrl
End Sub

Private Sub cmdUpdateChanges_Click()
    Dim ithUserID As Long
    ithUserID = ThisWorkbook.Sheets("user").Range("D" & Rows.Count).End(xlUp).Row
    
    If Not IsError(Application.Match(CStr(Trim(Me.cboUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)) Then
        
        ithUserID = Application.Match(CStr(Trim(Me.cboUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)
    Else
        Exit Sub
    End If
    
    Dim arrayTwo(4) As Variant
    arrayTwo(0) = Me.txtFirstName.Value
    arrayTwo(1) = Me.txtLastName.Value
    arrayTwo(2) = Me.cboLevel.Value
    arrayTwo(3) = Me.txtPasscode.Value
    If Me.chkBlackListUser.Value = False Then
        arrayTwo(4) = "Yes"
    ElseIf Me.chkBlackListUser.Value = True Then
        arrayTwo(4) = "No"
    End If
    
    'ReDim Preserve arrayOne(5) 'this preserves the first 5 values before adding the 6th value
    'arrayOne(5) = "C"
    
    If Trim(Join(arrayOne, "")) <> Trim(Join(arrayTwo, "")) Then
        With ThisWorkbook.Sheets("user")
           .Range("A" & ithUserID).Value = Me.txtFirstName.Value
           .Range("B" & ithUserID).Value = Me.txtLastName.Value
           .Range("C" & ithUserID).Value = Me.cboLevel.Value
           .Range("E" & ithUserID).Value = Me.txtPasscode.Value
           If chkBlackListUser.Value = True Then
                .Range("G" & ithUserID).Value = "No"
            Else
                .Range("G" & ithUserID).Value = "Yes"
            End If
        End With
        MsgBox prompt:="Changes successfully made", Title:="Update Notification"
    Else
        MsgBox prompt:="No Changes Detected", Title:="Update Notification"
    End If
            
End Sub

Private Sub togUpdateVerified_Click()
    MsgBox numOfChanges
    If Me.togUpdateVerified.Value = True Then
        If MsgBox(prompt:="Entring Modifying Mode" + _
        vbNewLine + "Any changes made must be saved", Buttons:=vbOKCancel, Title:="User Editing Mode") = vbOK Then
        'enable editing
            With Me
                .txtFirstName.Enabled = True
                .txtLastName.Enabled = True
                .cboLevel.Enabled = True
                .txtPasscode.Enabled = True
                .txtPasscode.Enabled = True
                .chkBlackListUser.Enabled = True
                .cboUserID.Enabled = False
            End With
        Else
            Me.togUpdateVerified.Value = False
        End If
    Else
        With Me
            .txtFirstName.Enabled = False
            .txtLastName.Enabled = False
            .cboLevel.Enabled = False
            .txtPasscode.Enabled = False
            .txtPasscode.Enabled = False
            .chkBlackListUser.Enabled = False
            .cboUserID.Enabled = True
        End With
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Dim ithUserID As Long
    ithUserID = ThisWorkbook.Sheets("user").Range("D" & Rows.Count).End(xlUp).Row
    
    Dim user As Range
    Dim userRange As Range
    
    Set userRange = ThisWorkbook.Sheets("user").Range("D3:D" & ithUserID)
    
    For Each user In userRange
        With Me.cboUserID
            .AddItem user.Value
        End With
    Next user
    
    With Me.cboLevel
        .AddItem "Supervisor"
        .AddItem "Representative"
        .AddItem "Analyst"
        .AddItem "Strategist"
    End With
    
    ThisWorkbook.Sheets("user").Activate
    With Me.lstUsersonFile
        .ColumnCount = 7
        '.ColumnHeads = True
        .RowSource = ThisWorkbook.Worksheets("user").Range("A3:G" & ithUserID).Address
    End With
    
    Me.togUpdateVerified.Enabled = False
    Me.cmdUpdateChanges.Enabled = False
    
    'disabling editing
    With Me
        .txtFirstName.Enabled = False
        .txtLastName.Enabled = False
        .cboLevel.Enabled = False
        .txtPasscode.Enabled = False
        .txtPasscode.Enabled = False
        .chkBlackListUser.Enabled = False
    End With
    
    'initializing changes made
    numOfChanges = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub
