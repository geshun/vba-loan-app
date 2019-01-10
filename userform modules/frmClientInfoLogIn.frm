VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientInfoLogIn 
   Caption         =   "Searching Client"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14880
   OleObjectBlob   =   "frmClientInfoLogIn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClientInfoLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ithClientID As Long

Private Sub cmdFind_Click()
    If Len(Me.txtClientID.Value) = 0 Then
        MsgBox prompt:="Please Enter a Valid Client ID.", Title:="Client Searching Result."
        Me.txtClientID.SetFocus
        Exit Sub
        
    ElseIf Not IsError(Application.Match(CLng(Me.txtClientID.Value), _
        ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)) Then
 
        ithClientID = Application.Match(CLng(Me.txtClientID.Value), _
            ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
        
        Me.Hide
        frmClientProfile.Show
        
    Else
        MsgBox prompt:="No Match Found. Please Enter a Valid Client ID.", Title:="Client Searching Result."
        Me.txtClientID.SetFocus
        Exit Sub
    End If

End Sub

Private Sub cmdGo_Click()
    
    frmClientInfoLogIn.txtClientID.Value = Me.lstClientInfoExtra.Value
            
    Me.Hide
    frmClientProfile.Show
End Sub

Private Sub cmdLogOff_Click()

End Sub

Private Sub cmdMainMenu_Click()
    Unload Me
    frmLoanMenu.Show
End Sub

Private Sub cmdSearchwithPhone_Click()
    Dim ithPhoneNumber As Long
    Dim ithPhoneNumberCount As Long
    
    'On Error GoTo handler
    If Not IsError(Application.Match(CStr(Me.txtPhoneNumber.Value), _
        ThisWorkbook.Sheets("client_info_personal").Range("I:I"), 0)) Then
        
        ithPhoneNumberCount = WorksheetFunction.CountIf(ThisWorkbook.Sheets("client_info_personal").Range("I:I"), CStr(Me.txtPhoneNumber.Value))
        'MsgBox ithPhoneNumberCount
        
        Select Case ithPhoneNumberCount
            Case 1
                ithPhoneNumber = Application.Match(CStr(Me.txtPhoneNumber.Value), _
                    ThisWorkbook.Sheets("client_info_personal").Range("I:I"), 0)
                
                ithClientID = ithPhoneNumber
                Me.txtClientID.Value = ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Value
            
                Me.Hide
                frmClientProfile.Show
            Case Else
                ThisWorkbook.Sheets("client_info_personal").Activate
                ThisWorkbook.Sheets("client_info_personal").Range("I1").AutoFilter Field:=9, Criteria1:=Trim(CStr(Me.txtPhoneNumber))
                     
                frmClientInfoLoginExtra.Show
                Exit Sub
                    
        End Select
    Else
        MsgBox prompt:="No match found"
        Exit Sub
    End If
'handler:
'    MsgBox "failed"
'    Exit Sub
End Sub

Private Sub cmdSearch_Click()
    Dim ithFirstName As Long
    Dim ithLastName As Long
    
    Dim ithFirstNameCount As Long
    Dim ithLastNameCount As Long
    
    'On Error GoTo handler
    If Not IsError(Application.Match(Trim(CStr(Me.txtFirstName.Value)), _
        ThisWorkbook.Sheets("client_info_personal").Range("B:B"), 0)) And _
            Not IsError(Application.Match(Trim(CStr(Me.txtLastName.Value)), _
            ThisWorkbook.Sheets("client_info_personal").Range("D:D"), 0)) Then
        
        ithFirstNameCount = WorksheetFunction.CountIf(ThisWorkbook.Sheets("client_info_personal").Range("B:B"), Trim(CStr(Me.txtFirstName.Value)))
        ithLastNameCount = WorksheetFunction.CountIf(ThisWorkbook.Sheets("client_info_personal").Range("D:D"), Trim(CStr(Me.txtLastName.Value)))
        
        If ithFirstNameCount = 1 And ithLastNameCount = 1 Then
            ithFirstName = Application.Match(Trim(CStr(Me.txtFirstName.Value)), _
                ThisWorkbook.Sheets("client_info_personal").Range("B:B"), 0)
        
            ithLastName = Application.Match(Trim(CStr(Me.txtLastName.Value)), _
                ThisWorkbook.Sheets("client_info_personal").Range("D:D"), 0)
            
            If ithFirstName = ithLastName Then
                ithClientID = ithFirstName
                Me.txtClientID.Value = ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Value
            
                Me.Hide
                frmClientProfile.Show
            End If
            
        ElseIf ithFirstNameCount = 1 And ithLastNameCount > 1 Then
            ithFirstName = Application.Match(Trim(CStr(Me.txtFirstName.Value)), _
                ThisWorkbook.Sheets("client_info_personal").Range("B:B"), 0)
                
            If LCase(CStr(ThisWorkbook.Sheets("client_info_personal").Range("D" & ithFirstName).Value)) = LCase(Trim(CStr(Me.txtLastName.Value))) Then
                ithClientID = ithFirstName
                Me.txtClientID.Value = ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Value
            
                Me.Hide
                frmClientProfile.Show
            End If
            
        ElseIf ithFirstNameCount > 1 And ithLastNameCount = 1 Then
            ithLastName = Application.Match(Trim(CStr(Me.txtLastName.Value)), _
                ThisWorkbook.Sheets("client_info_personal").Range("D:D"), 0)
            
            If LCase(CStr(ThisWorkbook.Sheets("client_info_personal").Range("B" & ithLastName).Value)) = LCase(Trim(CStr(Me.txtFirstName.Value))) Then
                ithClientID = ithLastName
                Me.txtClientID.Value = ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Value
            
                Me.Hide
                frmClientProfile.Show
            End If
        
        Else
            Dim ithInfoPersonal As Long
            ithInfoPersonal = ThisWorkbook.Sheets("client_info_personal").Range("A" & Rows.Count).End(xlUp).Row
            
            With ThisWorkbook.Sheets("client_info_personal")
                .Range("B1").AutoFilter Field:=2, Criteria1:=Trim(CStr(Me.txtFirstName.Value))
                .Range("D1").AutoFilter Field:=4, Criteria1:=Trim(CStr(Me.txtLastName.Value))
   
                Dim infoPersonal As Range
                Dim infoPersonalRange As Range
                Set infoPersonalRange = .Range("B1:B" & ithInfoPersonal).Rows.SpecialCells(xlCellTypeVisible)
                
                If WorksheetFunction.CountA(infoPersonalRange) = 1 Then
                    .AutoFilterMode = False
                    MsgBox prompt:="No match found. Please modify search.", Title:="Searching Result."
                    Exit Sub
                End If
                
                frmClientInfoLoginExtra.Show
                Exit Sub
            End With
            
        End If
    Else
        MsgBox prompt:="No match found. Please modify search.", Title:="Searching Result."
        Exit Sub
    End If
    
'handler:
'    MsgBox prompt:="No match found. Please modify search.", Title:="Searching Result....."
    'Me.txtFirstName.SetFocus
 '   Exit Sub
End Sub

Private Sub txtClientID_AfterUpdate()
    On Error GoTo handler
    If Not IsError(CLng(Me.txtClientID.Value)) Then
        Exit Sub
    Else
        MsgBox prompt:="Invalid Entry. Please Client ID contains only numbers", Title:="Wrong Entry"
    End If
handler:
    MsgBox prompt:="Thre is an er"
End Sub

Private Sub txtClientID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox prompt:="Only Numbers are accepted", Title:="Wrong Key Notification"
    End If
End Sub

Private Sub txtPhoneNumber_AfterUpdate()
    Me.txtPhoneNumber.Value = Format(Me.txtPhoneNumber.Value, "000-000-0000")
End Sub

Private Sub txtPhoneNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii > 47 And KeyAscii < 58) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    'just for testing code...delete later
    Me.txtClientID.Value = 290010002
    Me.cmdFind.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
    End If
End Sub
