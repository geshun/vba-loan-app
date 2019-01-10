VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoanCalculator 
   Caption         =   "Loan Estimator (GHS) - Truth Is Powerful"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14130
   OleObjectBlob   =   "frmLoanCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoanCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDuration_Change()
    With Me
        .txtTotalAmounttoPay.Value = ""
        .txtInteresttoPay.Value = ""
        .txtInteresttoPayPercent.Value = ""
        .txtAmountperSchedule.Value = ""
        .cboPaymentSchedule.Value = ""
        .cboEarlyPayment.Value = ""
        .txtAmountPaid.Value = ""
        .txtAmountSaved.Value = ""
    End With
End Sub

Private Sub cboEarlyPayment_Change()
    If Len(Me.txtTotalAmounttoPay.Value) > 0 Then
        If Me.cboEarlyPayment.Value = "" Then
            Me.txtAmountPaid.Value = ""
            Me.txtAmountSaved.Value = ""
        ElseIf Me.cboEarlyPayment.Value = "1 Week" Then
            Me.txtAmountPaid.Value = Me.txtTotalAmounttoPay.Value - WorksheetFunction.Round(Me.txtInteresttoPay.Value * (2 / 100), 0)
            Me.txtAmountSaved.Value = WorksheetFunction.Round(Me.txtInteresttoPay.Value * (2 / 100), 0)
        ElseIf Me.cboEarlyPayment.Value = "2 Weeks" Then
            Me.txtAmountPaid.Value = Me.txtTotalAmounttoPay.Value - WorksheetFunction.Round(Me.txtInteresttoPay.Value * (4 / 100), 0)
            Me.txtAmountSaved.Value = WorksheetFunction.Round(Me.txtInteresttoPay.Value * (4 / 100), 0)
        ElseIf Me.cboEarlyPayment.Value = "3 Weeks" Then
            Me.txtAmountPaid.Value = Me.txtTotalAmounttoPay.Value - WorksheetFunction.Round(Me.txtInteresttoPay.Value * (6 / 100), 0)
            Me.txtAmountSaved.Value = WorksheetFunction.Round(Me.txtInteresttoPay.Value * (6 / 100), 0)
        ElseIf Me.cboEarlyPayment.Value = "4 Weeks" Then
            Me.txtAmountPaid.Value = Me.txtTotalAmounttoPay.Value - WorksheetFunction.Round(Me.txtInteresttoPay.Value * (8 / 100), 0)
            Me.txtAmountSaved.Value = WorksheetFunction.Round(Me.txtInteresttoPay.Value * (8 / 100), 0)
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cboPaymentSchedule_Change()
    If Len(Me.txtTotalAmounttoPay.Value) > 0 Then
        If Me.cboPaymentSchedule.Value = "Daily (1 day)" Then
            Me.txtAmountperSchedule.Value = (Me.txtTotalAmounttoPay.Value / (Me.cboDuration / 28)) / 28
        ElseIf Me.cboPaymentSchedule.Value = "Weekly (7 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtTotalAmounttoPay.Value / (Me.cboDuration / 28)) / 4
        ElseIf Me.cboPaymentSchedule.Value = "Bi-Weekly (14 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtTotalAmounttoPay.Value / (Me.cboDuration / 28)) / 2
        ElseIf Me.cboPaymentSchedule.Value = "Monthly (28 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtTotalAmounttoPay.Value / (Me.cboDuration / 28)) / 1
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cboRate_Change()
    With Me
        .txtTotalAmounttoPay.Value = ""
        .txtInteresttoPay.Value = ""
        .txtInteresttoPayPercent.Value = ""
        .txtAmountperSchedule.Value = ""
        .cboPaymentSchedule.Value = ""
        .cboEarlyPayment.Value = ""
        .txtAmountPaid.Value = ""
        .txtAmountSaved.Value = ""
    End With
End Sub

Private Sub cmdExit_Click()
    If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        frmLoanMenu.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdGetCalculator_Click()
    frmBasicCalculator.Show
End Sub

Private Sub cmdGetSummary_Click()
    If Len(Me.txtPrincipal.Value) > 0 And Len(Me.cboRate.Value) > 0 And Len(Me.cboDuration.Value) > 0 Then
        If Len(Me.txtTotalAmounttoPay) > 0 Then
            Dim months As String
            If Me.cboDuration.Value = 28 Then
                months = " month"
            Else
                months = " months"
            End If
            MsgBox prompt:="A loan of GHS " & Me.txtPrincipal & " to be paid in " & Me.cboDuration & " days (" _
                 & Me.cboDuration / 28 & months & ")." + vbNewLine + _
                    "The interest for the whole term is GHS " & Me.txtInteresttoPay & " and payment is GHS " _
                    & Me.txtAmountperSchedule & " " & Me.cboPaymentSchedule & "." + vbNewLine + _
                    "The total amount to be paid as at the end of term is GHS " & Me.txtTotalAmounttoPay & " (loan amount plus interest).", Buttons:=vbInformation, _
                    Title:="Summary of Loan"
        Else
            Call cmdCompute_Click
            Dim month As String
            If Me.cboDuration.Value = 28 Then
                month = " month"
            Else
                month = " months"
            End If
            MsgBox prompt:="A loan of GHS " & Me.txtPrincipal & " to be paid in " & Me.cboDuration & " days (" _
                 & Me.cboDuration / 28 & month & ")." + vbNewLine + _
                    "The interest for the whole term is GHS " & Me.txtInteresttoPay & " and payment is GHS " _
                    & Me.txtAmountperSchedule & " " & Me.cboPaymentSchedule & "." + vbNewLine + _
                    "The total amount to be paid as at the end of term is GHS " & Me.txtTotalAmounttoPay & " (loan amount plus interest).", Buttons:=vbInformation, _
                    Title:="Summary of Loan"
        End If
    Else
        MsgBox prompt:="Either principal or rate or duration is missing", Title:="Loan Amount Information"
        Me.txtPrincipal.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAmountPaid_Change()
    Me.txtAmountPaid.Value = Format(Me.txtAmountPaid.Value, "#,##0.00")
End Sub

Private Sub txtAmountperSchedule_Change()
    txtAmountperSchedule.Value = Format(txtAmountperSchedule.Value, "#,##0.00")
End Sub

Private Sub txtAmountSaved_Change()
    Me.txtAmountSaved.Value = Format(Me.txtAmountSaved.Value, "#,##0.00")
End Sub

Private Sub txtPrincipal_Change()
     With Me
        .txtTotalAmounttoPay.Value = ""
        .txtInteresttoPay.Value = ""
        .txtInteresttoPayPercent.Value = ""
        .txtAmountperSchedule.Value = ""
        .cboPaymentSchedule.Value = ""
        .cboEarlyPayment.Value = ""
        .txtAmountPaid.Value = ""
        .txtAmountSaved.Value = ""
    End With
End Sub

Private Sub txtPrincipal_AfterUpdate()
    If IsNumeric(txtPrincipal.Value) And txtPrincipal.Value > 0 Then
        txtPrincipal.Value = Format(txtPrincipal.Value, "#,##0.00")
    Else
        txtPrincipal.Value = vbNullString
        txtPrincipal.SetFocus
    End If
End Sub

Private Sub cmdCompute_Click()
    If Len(Me.txtPrincipal.Value) > 0 And Len(Me.cboRate.Value) > 0 And Len(Me.cboDuration.Value) > 0 Then
        With Me
            .txtTotalAmounttoPay.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28) + .txtPrincipal.Value
            .txtTotalAmounttoPay.Value = Format(.txtTotalAmounttoPay.Value, "#,##0.00")
            
            .txtInteresttoPay.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28)
            .txtInteresttoPay.Value = Format(.txtInteresttoPay.Value, "#,##0.00")
            
            .txtInteresttoPayPercent.Value = CStr(Format((.txtInteresttoPay / .txtPrincipal) * 100, "##0.00")) & "%"
              
             If Len(Me.cboPaymentSchedule.Value) > 0 Then
                Call cboPaymentSchedule_Change
             Else
                .cboPaymentSchedule.Value = "Monthly (28 days)"
                .txtAmountperSchedule.Value = (.txtTotalAmounttoPay.Value / (.cboDuration / 28)) / 1
            End If
            
            If Len(Me.cboEarlyPayment.Value) > 0 Then
                Call cboEarlyPayment_Change
            End If
        End With
    Else
        If Len(Me.txtPrincipal.Value) <= 0 Then
            MsgBox prompt:="Enter value for principal", Title:="Loan Information - Principal"
            Me.txtPrincipal.SetFocus
        ElseIf Len(Me.cboRate.Value) <= 0 Then
            MsgBox prompt:="Select the interest rate", Title:="Loan Information - Interest Rate"
            Me.cboRate.SetFocus
        ElseIf Len(Me.cboDuration.Value) <= 0 Then
            MsgBox prompt:="Select loan duration", Title:="Loan Information - Duration of Loan"
            Me.cboDuration.SetFocus
        End If
        Exit Sub
    End If
End Sub

Private Sub cmdReset_Click()
    With Me
        .txtPrincipal.Value = ""
        .cboRate.Value = ""
        .cboDuration.Value = ""
        .txtTotalAmounttoPay.Value = ""
        .txtInteresttoPay.Value = ""
        .txtAmountperSchedule.Value = ""
        .cboPaymentSchedule.Value = ""
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .txtPrincipal.Value = 1000
        .txtPrincipal.Value = Format(.txtPrincipal.Value, "#,##0.00")
        
        With .cboDuration
            .AddItem 28
            .AddItem 56
            .AddItem 84
            .AddItem 112
            .AddItem 140
            .AddItem 168
            .AddItem 196
            .AddItem 224
            .AddItem 252
            .AddItem 280
            .AddItem 308
            .AddItem 336
        End With
        .cboDuration.Value = 112
        
        With .cboRate
            .AddItem 5#
            .AddItem 5.25
            .AddItem 5.5
            .AddItem 5.75
            .AddItem 6#
            .AddItem 6.25
            .AddItem 6.5
            .AddItem 6.75
            .AddItem 7#
            .AddItem 7.25
            .AddItem 7.5
            .AddItem 7.75
        End With
        .cboRate.Value = 6.5
        
        .txtTotalAmounttoPay.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28) + .txtPrincipal.Value
        .txtTotalAmounttoPay.Value = Format(.txtTotalAmounttoPay.Value, "#,##0.00")
        
        .txtInteresttoPay.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28)
        .txtInteresttoPay.Value = Format(.txtInteresttoPay.Value, "#,##0.00")
        
        .txtInteresttoPayPercent.Value = CStr(Format((.txtInteresttoPay / .txtPrincipal) * 100, "##0.00")) & "%"
        
        .txtAmountperSchedule.Value = .txtTotalAmounttoPay.Value / 4
        .txtAmountperSchedule.Value = Format(.txtAmountperSchedule.Value, "#,##0.00")
        
        With .cboPaymentSchedule
            .AddItem "Daily (1 day)"
            .AddItem "Weekly (7 days)"
            .AddItem "Bi-Weekly (14 days)"
            .AddItem "Monthly (28 days)"
        End With
        .cboPaymentSchedule.Value = "Monthly (28 days)"
        
        With .cboEarlyPayment
            .AddItem "1 Week"
            .AddItem "2 Weeks"
            .AddItem "3 Weeks"
            .AddItem "4 Weeks"
        End With
        
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub
