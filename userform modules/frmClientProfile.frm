VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientProfile 
   Caption         =   "Client Profile Verification - TIP"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15540
   OleObjectBlob   =   "frmClientProfile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClientProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreatePortfolio_Click()
    'if client has a good status and previous loan is fully paid then proceed to create loan else raise a warning and exit sub
    If Me.chkInactiveStatus.Value = False Then
        Unload Me
        frmCreateLoan.Show
    Else
        MsgBox prompt:="This Client is not in good standing." + vbNewLine + "Client cannot apply for any loan", Title:="Client Status Warning....", Buttons:=vbCritical
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If frmClientInfoLogIn.Visible = False Then
        frmClientInfoLogIn.Show
    End If
End Sub

Private Sub cmdMainMenu_Click()
    Unload Me
    Unload frmClientInfoLogIn
    frmLoanMenu.Show
End Sub

Private Sub cmdMakePayment_Click()
    If Not IsError(Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
        ThisWorkbook.Sheets("loan_list").Range("A:A"), 0)) Then
        Unload Me
        frmLoanRepayment.Show
    Else
        MsgBox prompt:="Client has not yet applied for loan." + vbNewLine + "Click on Create Portfolio to apply for Loan", Title:="Warning....", Buttons:=vbCritical
        Exit Sub
    End If
End Sub

Private Sub cmdUpdateClientProfile_Click()
    switchAddEdit = True
    Unload Me
    frmAddClient.Show
End Sub



Private Sub UserForm_Initialize()
    Dim ithClientID As Long
    ithClientID = Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
        ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
    With Me
        With ThisWorkbook.Sheets("client_info_personal")
            txtClientID.Value = .Range("A" & ithClientID).Value
            txtFirstName.Value = .Range("B" & ithClientID).Value
            txtLastName.Value = .Range("D" & ithClientID).Value
            txtClientSince.Value = .Range("K" & ithClientID).Value
        End With
    End With
    
    Me.chkInactiveStatus.Enabled = False
    If ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Offset(0, 9).Value <> "Active" Then
        Me.chkInactiveStatus.Value = True
    End If
    
    With Me.cboCategory
        .AddItem "Loans"
        .AddItem "Repayments"
    End With
    
    'put exception for filepath...if file exits
    'Dim picturePath As String
    'picturePath = ThisWorkbook.Path & "\LoanClientImage\" & CStr(txtClientID.Value) & ".jpg"
    'Me.imgClient.Picture = LoadPicture(picturePath)
    
    'just for testing code...delete later
    Me.cmdMakePayment.SetFocus
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

