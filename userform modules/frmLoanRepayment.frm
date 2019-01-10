VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoanRepayment 
   Caption         =   "Loan Repayment - Collection"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16785
   OleObjectBlob   =   "frmLoanRepayment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoanRepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboLoanID_Change()
    'filling out the fields using selected loan ID
    Dim selectedLoanID As Long
    selectedLoanID = Application.Match(CStr(Me.cboLoanID.Value), _
        ThisWorkbook.Sheets("loan_list").Range("B:B"), 0)
        
    With ThisWorkbook.Sheets("loan_list")
        Me.txtPrincipal.Value = .Range("D" & selectedLoanID).Value
        Me.txtInterest.Value = .Range("H" & selectedLoanID).Value
        Me.txtDuration.Value = .Range("F" & selectedLoanID).Value
        Me.txtPaymentperSchedule.Value = .Range("M" & selectedLoanID).Value
        Me.txtPrincipalInterest.Value = .Range("G" & selectedLoanID).Value
        Me.txtSchedule.Value = .Range("L" & selectedLoanID).Value
        Me.txtStartDate.Value = .Range("O" & selectedLoanID).Value
        Me.txtEndDate.Value = .Range("P" & selectedLoanID).Value
        
        If CStr(Me.txtSchedule.Value) = "Daily (1 day)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 1
        ElseIf CStr(Me.txtSchedule.Value) = "Weekly (7 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 7
        ElseIf CStr(Me.txtSchedule.Value) = "Bi-Weekly (14 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 14
        ElseIf CStr(Me.txtSchedule.Value) = "Monthly (28 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 28
        End If
    End With
    
    'filling out the summary payment field using selected loan ID
    Dim loanIDonPayment As Long
    If Not IsError(Application.Match(CStr(Me.cboLoanID.Value), _
        ThisWorkbook.Sheets("loan_payment").Range("B:B"), 0)) Then
        
        With ThisWorkbook.Sheets("loan_payment")
            .Range("B1").AutoFilter Field:=2, Criteria1:=CStr(Me.cboLoanID.Value)
            Me.txtTotalPaid.Value = WorksheetFunction.Sum(.Range("E:E").Rows.SpecialCells(xlCellTypeVisible))
            Me.txtAmountRemain.Value = Me.txtPrincipalInterest.Value - Me.txtTotalPaid.Value
            Me.txtNumberofPayment.Value = WorksheetFunction.Count(.Range("E:E").Rows.SpecialCells(xlCellTypeVisible))
        End With
        
    Else
        Me.txtTotalPaid.Value = 0
        Me.txtAmountRemain.Value = Me.txtPrincipalInterest.Value
        Me.txtNumberofPayment.Value = 0
    End If
    
    If ThisWorkbook.Sheets("loan_payment").AutoFilterMode = True Then
        ThisWorkbook.Sheets("loan_payment").AutoFilterMode = False
    End If
    
    If CDbl(Me.txtTotalPaid.Value) < CDbl(Me.txtPrincipalInterest.Value) Then
        Me.lblLoanStatus.Caption = "Still Owing on this Loan"
    Else
        Me.lblLoanStatus.Caption = "Loan is Fully Paid Off"
    End If
    
    'PMT date to detect delinquency
    Select Case CStr(Me.txtSchedule.Value)
        Case "Daily (1 day)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 1
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 1
            End If
        Case "Weekly (7 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 7
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 7
            End If
        Case "Bi-Weekly (14 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 14
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 14
            End If
        Case "Monthly (28 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 28
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 28
            End If
    End Select
    
    'detect payment delinquency
    Dim pmtTrack As Long
    pmtTrack = Now() - CDate(Me.txtSchedulePMTDate.Value)
    If pmtTrack < 0 Then
        Me.txtDelinquencyStatus.Value = "Early PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    ElseIf pmtTrack = 0 Then
        Me.txtDelinquencyStatus.Value = "On-Time PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    ElseIf pmtTrack > 0 Then
        Me.txtDelinquencyStatus.Value = "Late PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    End If
    If Me.txtDelinquencyStatus.Value = "Late PMT" Then Me.txtDelinquencyStatus.BackColor = vbRed
    
    'setting payment type based on discount value
    If Me.txtDiscount.Value > 0 Then
        Me.cboPaymentType.Value = "Discounted Repayment"
    Else
        Me.cboPaymentType.Value = "Only Repayment"
    End If
    
    'lock payloan (unused method) based on the payment method
    If Me.cboPaymentMethod.Value = "Cash" Then
        Me.frameOtherPaymentMethod.Enabled = False
        Me.frameCashPayment.Enabled = True
    ElseIf Me.cboPaymentMethod.Value = "Mobile Money" Then
        Me.frameOtherPaymentMethod.Enabled = True
        Me.frameCashPayment.Enabled = False
    End If
    
    'lock all payloan if there is nothing to pay
    If Me.txtAmountRemain.Value = 0 Then
        Me.cmdPayLoan.Enabled = False
        Me.txtTotalAmount.Value = vbNullString
        
        Me.cmdMobileMoneyPayMM.Enabled = False
        Me.txtAmountMM.Value = vbNullString
        Me.txtAmountplusDiscountMM.Value = vbNullString
    Else
        Me.cmdPayLoan.Enabled = True
        Me.cmdMobileMoneyPayMM.Enabled = True
    End If
    
End Sub

Private Sub cboPaymentMethod_Change()
    If Me.cboPaymentMethod.Value = "Cash" Then
        Me.frameOtherPaymentMethod.Enabled = False
        Me.frameCashPayment.Enabled = True
    ElseIf Me.cboPaymentMethod.Value = "Mobile Money" Then
        Me.frameOtherPaymentMethod.Enabled = True
        Me.frameCashPayment.Enabled = False
    End If
End Sub

Private Sub cmdCompute_Click()
    Dim checkTotalAmount As Double
    checkTotalAmount = CDbl(Me.txtProduct1.Value) + _
                            CDbl(Me.txtProduct2.Value) + _
                            CDbl(Me.txtProduct5.Value) + _
                            CDbl(Me.txtProduct10.Value) + _
                            CDbl(Me.txtProduct20.Value) + _
                            CDbl(Me.txtProduct50.Value) + _
                            CDbl(Me.txtCoins.Value) + _
                            CDbl(Me.txtDiscount.Value)
    If checkTotalAmount = 0 Then
        Me.txtTotalAmount.Value = vbNullString
    Else
        Me.txtTotalAmount = checkTotalAmount
        Me.txtQuant1.Enabled = False
        Me.txtQuant2.Enabled = False
        Me.txtQuant5.Enabled = False
        Me.txtQuant10.Enabled = False
        Me.txtQuant20.Enabled = False
        Me.txtQuant50.Enabled = False
        Me.cmdCompute.Enabled = False
    End If
End Sub

Private Sub cmdCorrectPayment_Click()
    MsgBox "Select PMT ID to proceed"
    Me.cboPaymentID.SetFocus
End Sub

Private Sub cmdExit_Click()
    If MsgBox(prompt:="This will close the loan repayment form" + vbNewLine + _
        "Would you like to exit?", Title:="Exiting Loan Repayment", Buttons:=vbInformation + vbYesNo) = vbYes Then
        Unload Me
        frmClientProfile.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdPayLoan_Click()
    
    'verify amount paid
    If Len(Me.txtTotalAmount.Value) = 0 Or Me.txtTotalAmount.Value = 0 Then
        Me.txtTotalAmount.Value = vbNullString
        Exit Sub
    End If
    
    'verify amount of coins used
    If CDbl(Me.txtCoins.Value) > 2 Then
        If MsgBox(prompt:="Are you sure you have this much coins?" + vbNewLine + _
            Me.txtCoins.Value, Title:="Coin collection notification", Buttons:=vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'verify payment date
    If Me.txtLastPMTDate.Value = "None" And (CDate(Format(Me.dtpPaymentDate.Value, "dd-Mmm-yyyy")) - CDate(Me.txtStartDate.Value) < -1) Then
        MsgBox "Payment Date is before when the loan was granted" + vbNewLine + "Change date", _
            Title:="Error in dating", Buttons:=vbCritical
        Me.dtpPaymentDate.SetFocus
        Exit Sub
    ElseIf Me.txtLastPMTDate.Value <> "None" Then
        If (CDate(Format(Me.dtpPaymentDate.Value, "dd-Mmm-yyyy")) - CDate(Me.txtLastPMTDate.Value) < 0) Then
            MsgBox "Payment Date is before when the last payment was made" + vbNewLine + "Date should be after the last time client made payment", _
                Title:="Error in dating", Buttons:=vbCritical
            Me.dtpPaymentDate.SetFocus
            Exit Sub
        ElseIf (CDate(Format(Me.dtpPaymentDate.Value, "dd-Mmm-yyyy")) - CDate(Me.txtLastPMTDate.Value) = 0) Then
            If MsgBox(prompt:="There is a payment made on this date already." + vbNewLine + _
                "Is this an additional or new payment?", Title:="Same Date Payments", Buttons:=vbInformation + vbYesNo) = vbNo Then
                Me.dtpPaymentDate.SetFocus
                Exit Sub
            End If
        End If
    ElseIf (Me.dtpPaymentDate.Value - Now()) > 0 Then
        MsgBox prompt:="You cannot future date payment", Title:="Error in dating", Buttons:=vbCritical
        Me.dtpPaymentDate.SetFocus
        Exit Sub
    End If
    
    'verify excess payment
    If CDbl(Me.txtTotalAmount.Value) > CDbl(Me.txtAmountRemain.Value) Then
        MsgBox prompt:="Please you can't pay more than you owe", Title:="Payment warning", Buttons:=vbCritical
        Exit Sub
    End If
    
    'verify payment compared to scheduled payment
    If CDbl(Me.txtTotalAmount.Value) > CDbl(Me.txtPaymentperSchedule.Value) Then
        If MsgBox(prompt:="Client's payment of " + Me.txtTotalAmount.Value + _
            " is MORE than scheduled payment of " + Me.txtPaymentperSchedule.Value + vbNewLine + _
                "Do you want to proceed?", Title:="Payment Information", Buttons:=vbYesNo + vbInformation) = vbNo Then
            Exit Sub
        End If
    End If
    
    If CDbl(Me.txtTotalAmount.Value) < CDbl(Me.txtPaymentperSchedule.Value) Then
        If MsgBox(prompt:="Client's payment of " + Me.txtTotalAmount.Value + _
            " is LESS than scheduled payment of " + Me.txtPaymentperSchedule.Value + vbNewLine + _
                "Do you want to proceed?", Title:="Payment Information", Buttons:=vbYesNo + vbInformation) = vbNo Then
            Exit Sub
        End If
    End If
    
    If MsgBox(prompt:="Payment date is " + Format(Me.dtpPaymentDate.Value, "Long Date") + vbNewLine + _
        "Do you want to proceed?", Buttons:=vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(prompt:="Payment to be made is GHS " + Me.txtTotalAmount.Value + vbNewLine + _
        "Do you want to proceed?", Buttons:=vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim ithPayment As Long
    ithPayment = ThisWorkbook.Sheets("loan_payment").Range("A" & Rows.Count).End(xlUp).Row
    
    Dim ithClientLoan As Long
    
    With ThisWorkbook.Sheets("loan_list")
        .Range("A1").AutoFilter Field:=1, Criteria1:=CStr(frmClientInfoLogIn.txtClientID.Value)
        ithClientLoan = WorksheetFunction.CountA(.Range("B:B").Rows.SpecialCells(xlCellTypeVisible)) - 1
    End With
    ThisWorkbook.Sheets("loan_list").AutoFilterMode = False
    
    Dim ithLoanPay As Long
    
    With ThisWorkbook.Sheets("loan_payment")
        .Range("B1").AutoFilter Field:=2, Criteria1:=CStr(Me.cboLoanID.Value)
        ithLoanPay = WorksheetFunction.CountA(.Range("B:B").Rows.SpecialCells(xlCellTypeVisible))
    End With
    ThisWorkbook.Sheets("loan_payment").AutoFilterMode = False
    
    
    Me.cboPaymentID.Value = payment_id(Me.txtClientID.Value, Me.txtPrincipal.Value, ithClientLoan, ithLoanPay, Me.txtTotalAmount.Value)
    With ThisWorkbook.Sheets("loan_payment")
        .Range("A" & ithPayment + 1).Value = Me.txtClientID.Value
        .Range("B" & ithPayment + 1).Value = Me.cboLoanID.Value
        .Range("C" & ithPayment + 1).Value = Me.cboPaymentID.Value
        .Range("D" & ithPayment + 1).Value = Me.txtPrincipal.Value
        .Range("E" & ithPayment + 1).Value = Me.txtTotalAmount.Value 'field conditioned on the payment method used
        .Range("F" & ithPayment + 1).Value = 2222 'to be computed from previous
        .Range("G" & ithPayment + 1).Value = Me.cboPaymentType.Value
        .Range("H" & ithPayment + 1).Value = Me.cboPaymentBy.Value
        .Range("I" & ithPayment + 1).Value = Me.cboPaymentMethod.Value
        .Range("J" & ithPayment + 1).Value = Format(Me.dtpPaymentDate.Value, "dd-Mmm-yyyy") 'change to date timepicker
        .Range("K" & ithPayment + 1).Value = Me.lblUserID.Caption
    End With
    
    'save after making payment
    ThisWorkbook.Save
    
    'reset values
    Call cmdReset_Click
    
    'print recipt
    '--------------------------------------------
    
    'cleaning worksheet
    With ThisWorkbook.Sheets("pmt_receipt").Cells
        .Clear
        .Font.Name = "Calibri"
        .Font.Size = 9
    End With
    
    'setting tags
    With ThisWorkbook.Sheets("pmt_receipt")
        .Range("B1").Value = "STATEMENT OF REPAYMENT AS AT " & Format(Now(), "dd-Mmm-yyyy")
        .Range("B1:D1").Merge
        .Range("B1:D1").HorizontalAlignment = xlCenter
        .Range("B1:D1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B1:D1").Borders(xlEdgeBottom).ColorIndex = 46
        
        .Range("A2").Value = "Client ID:"
        .Range("A3").Value = "Loan ID:"
        
        .Range("C2").Value = "Principal + Interest:"
        .Range("C3").Value = "PMT Schedule:"
        
        .Range("E2").Value = "Start Date:"
        .Range("E3").Value = "End Date:"
        
        .Range("F4").Value = "GHS"
        .Range("F4").Font.Size = 8
        .Range("F4").HorizontalAlignment = xlRight
        .Range("F4").Font.Bold = True
        
        .Range("A5").Value = "PMT #"
        .Range("B5").Value = "PMT Date"
        .Range("C5").Value = "PMT Method"
        .Range("D5").Value = "PMT Type"
        .Range("E5").Value = "PMT By"
        .Range("F5").Value = "PMT Amount"
        
        .Range("A5:F5").Font.Bold = True
        .Range("A2:F3").HorizontalAlignment = xlLeft
    End With
    
    With ThisWorkbook.Sheets("pmt_receipt")
        .Range("B2").Value = Me.txtClientID.Value
        .Range("B3").Value = Me.cboLoanID.Value
        .Range("D2").Value = Me.txtPrincipalInterest.Value
        .Range("D3").Value = Me.txtSchedule.Value
        .Range("F2").Value = Me.txtStartDate.Value
        .Range("F3").Value = Me.txtEndDate.Value
        .Range("A3:F3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:F3").Borders(xlEdgeBottom).ColorIndex = 0
        .Range("A5:F5").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A5:F5").Borders(xlEdgeBottom).ColorIndex = 0
    End With
    
    'define variables and use sql ADO
    Dim sourceConn As String
    Dim destWks As Worksheet
    Dim destRng As Range
    Dim sqlStatement As String
    
    sourceConn = ActiveWorkbook.FullName
    Set destWks = ThisWorkbook.Sheets("pmt_receipt")
    Set destRng = destWks.Range("B6")
    
    Dim stringLoanID As String
    stringLoanID = CStr(Me.cboLoanID.Value)
    
    'sqlStatement = "SELECT [Payment Date],[Payment Method],[Payment Type],[Payment By],[Amount Paid] FROM [loan_payment$] WHERE [Loan ID] =""270010001.8.8.1K.2016822"";"
    sqlStatement = "SELECT [Payment Date],[Payment Method],[Payment Type],[Payment By],[Amount Paid] FROM [loan_payment$] WHERE [Loan ID] =" & """" & CStr(Me.cboLoanID.Value) & """" & ";"
    
    Call get_data_sqlExcel(sourceConn, destWks, destRng, sqlStatement)
    
    Dim ithPMTLast As Long
    ithPMTLast = ThisWorkbook.Sheets("pmt_receipt").Range("F" & Rows.Count).End(xlUp).Row
    
    Dim i As Long
    Dim pmtRange As Range
    Set pmtRange = ThisWorkbook.Sheets("pmt_receipt").Range("A6:A" & ithPMTLast)
    
    With ThisWorkbook.Sheets("pmt_receipt")
        pmtRange.Offset(0, 1).NumberFormat = "dd-Mmm-yyyy"
        pmtRange.Offset(0, 1).HorizontalAlignment = xlLeft
        pmtRange.Offset(0, 5).NumberFormat = "#,##0.00"
        
        For i = ithPMTLast To 6 Step -1
            .Range("A" & i).Value = i - 6 + 1
        Next i
        pmtRange.HorizontalAlignment = xlLeft
        
        .Range("A" & ithPMTLast & ":F" & ithPMTLast).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range("A" & ithPMTLast & ":F" & ithPMTLast).Borders(xlEdgeBottom).ColorIndex = 0
        .Range("E" & ithPMTLast + 1).Value = "Total Paid"
        .Range("F" & ithPMTLast + 1).Value = WorksheetFunction.Sum(.Range("F6:F" & ithPMTLast))
        .Range("F" & ithPMTLast + 1).NumberFormat = "#,##0.00"
        .Range("E" & ithPMTLast + 2).Value = "Amount Remain"
        .Range("F" & ithPMTLast + 2).Value = .Range("D2").Value - .Range("F" & ithPMTLast + 1).Value
        .Range("F" & ithPMTLast + 2).NumberFormat = "#,##0.00"
        .Columns.AutoFit
    End With
    
    'save as pdf
    On Error Resume Next
    ThisWorkbook.Sheets("pmt_receipt").ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path + "\" + CStr(Me.cboLoanID.Value) + ".pdf", Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    MsgBox prompt:="Thanks. Payment is successfully made and saved as pdf", Title:="Payment Success", Buttons:=vbInformation
End Sub

Private Sub cmdReset_Click()
    Me.txtQuant1.Value = ""
    Me.txtQuant1.Enabled = True
    Me.txtProduct1.Value = 0
    Me.txtQuant2.Value = ""
    Me.txtQuant2.Enabled = True
    Me.txtProduct2.Value = 0
    Me.txtQuant5.Value = ""
    Me.txtQuant5.Enabled = True
    Me.txtProduct5.Value = 0
    Me.txtQuant10.Value = ""
    Me.txtQuant10.Enabled = True
    Me.txtProduct10.Value = 0
    Me.txtQuant20.Value = ""
    Me.txtQuant20.Enabled = True
    Me.txtProduct20.Value = 0
    Me.txtQuant50.Value = ""
    Me.txtQuant50.Enabled = True
    Me.txtProduct50.Value = 0
    
    Me.txtCoins.Value = 0
    Me.txtTotalAmount.Value = ""
    
    Me.cmdCompute.Enabled = True
    
End Sub


Private Sub txtAmountMM_AfterUpdate()
    If IsNumeric(Me.txtAmountMM.Value) And Me.txtAmountMM.Value > 0 Then
        Me.txtAmountplusDiscountMM.Value = Format(CDbl(Me.txtDiscount.Value) + CDbl(Me.txtAmountMM.Value), "#,##0.00")
        Me.txtAmountMM.Value = Format(txtAmountMM.Value, "#,##0.00")
    Else
        Me.txtAmountMM.Value = vbNullString
        Me.txtAmountplusDiscountMM.Value = vbNullString
        Me.txtAmountMM.SetFocus
    End If
End Sub

Private Sub txtAmountRemain_Change()
    If Me.txtAmountRemain.Value > 0 Then
        Me.txtAmountRemain.Value = Format(Me.txtAmountRemain.Value, "#,##0.00")
    End If
End Sub

Private Sub txtDiscount_AfterUpdate()
    Me.txtDiscount.Value = Format(Me.txtDiscount.Value, "#,##0.00")
End Sub

Private Sub txtInterest_Change()
    Me.txtInterest.Value = Format(Me.txtInterest.Value, "#,##0.00")
End Sub

Private Sub txtNewAmount_AfterUpdate()
    If IsNumeric(Me.txtNewAmount.Value) And Me.txtNewAmount.Value > 0 Then
        Me.txtNewAmount.Value = Format(txtNewAmount.Value, "#,##0.00")
        Me.cmdNewPaymentApply.Enabled = True
    Else
        Me.txtNewAmount.Value = vbNullString
        Me.cmdNewPaymentApply.Enabled = False
        Me.txtNewAmount.SetFocus
    End If
End Sub

Private Sub txtPaymentperSchedule_Change()
    Me.txtPaymentperSchedule.Value = Format(Me.txtPaymentperSchedule.Value, "#,##0.00")
End Sub

Private Sub txtPrincipal_Change()
    Me.txtPrincipal.Value = Format(Me.txtPrincipal.Value, "#,##0.00")
End Sub

Private Sub txtPrincipalInterest_Change()
    Me.txtPrincipalInterest.Value = Format(Me.txtPrincipalInterest.Value, "#,##0.00")
End Sub

'coins
Private Sub txtCoins_AfterUpdate()
    On Error Resume Next
    If Len(Me.txtCoins.Value) = 0 Or (Len(Me.txtCoins.Value) - Len(Replace(Me.txtCoins.Value, ".", ""))) > 1 Or CLng(Me.txtCoins.Value) = 0 Then
        Me.txtCoins.Value = 0
        Me.txtCoins.SetFocus
        Exit Sub
    End If
    
    Me.txtCoins.Value = Me.txtCoins.Value * 1
    Me.txtCoins.Value = Format(Me.txtCoins.Value, "#,##0.00")
End Sub

Private Sub txtCoins_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'1 ghana cedis bill
Private Sub txtQuant1_AfterUpdate()
    If Len(Me.txtQuant1.Value) = 0 Or Me.txtQuant1.Value = 0 Then
        Me.txtQuant1.Value = vbNullString
        Me.txtProduct1.Value = 0
        Exit Sub
    Else
        Me.txtQuant1.Value = CLng(Me.txtQuant1.Value)
    End If
    
    Me.txtProduct1.Value = Me.txtQuant1.Value * 1
    Me.txtProduct1.Value = Format(Me.txtProduct1.Value, "#,##0.00")
End Sub

Private Sub txtQuant1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'2 ghana cedis bill
Private Sub txtQuant2_AfterUpdate()
    If Len(Me.txtQuant2.Value) = 0 Or Me.txtQuant2.Value = 0 Then
        Me.txtQuant2.Value = vbNullString
        Me.txtProduct2.Value = 0
        Exit Sub
    Else
        Me.txtQuant2.Value = CLng(Me.txtQuant2.Value)
    End If
    
    Me.txtProduct2.Value = Me.txtQuant2.Value * 2
    Me.txtProduct2.Value = Format(Me.txtProduct2.Value, "#,##0.00")
End Sub

Private Sub txtQuant2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'5 ghana cedis bill
Private Sub txtQuant5_AfterUpdate()
    If Len(Me.txtQuant5.Value) = 0 Or Me.txtQuant5.Value = 0 Then
        Me.txtQuant5.Value = vbNullString
        Me.txtProduct5.Value = 0
        Exit Sub
    Else
        Me.txtQuant5.Value = CLng(Me.txtQuant5.Value)
    End If
    
    Me.txtProduct5.Value = Me.txtQuant5.Value * 5
    Me.txtProduct5.Value = Format(Me.txtProduct5.Value, "#,##0.00")
End Sub

Private Sub txtQuant5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'10 ghana cedis bill
Private Sub txtQuant10_AfterUpdate()
    If Len(Me.txtQuant10.Value) = 0 Or Me.txtQuant10.Value = 0 Then
        Me.txtQuant10.Value = vbNullString
        Me.txtProduct10.Value = 0
        Exit Sub
    Else
        Me.txtQuant10.Value = CLng(Me.txtQuant10.Value)
    End If
    
    Me.txtProduct10.Value = Me.txtQuant10.Value * 10
    Me.txtProduct10.Value = Format(Me.txtProduct10.Value, "#,##0.00")
End Sub

Private Sub txtQuant10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'20 ghana cedis bill
Private Sub txtQuant20_AfterUpdate()
    If Len(Me.txtQuant20.Value) = 0 Or Me.txtQuant20.Value = 0 Then
        Me.txtQuant20.Value = vbNullString
        Me.txtProduct20.Value = 0
        Exit Sub
    Else
        Me.txtQuant20.Value = CLng(Me.txtQuant20.Value)
    End If
    
    Me.txtProduct20.Value = Me.txtQuant20.Value * 20
    Me.txtProduct20.Value = Format(Me.txtProduct20.Value, "#,##0.00")
End Sub

Private Sub txtQuant20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'50 ghana cedis bill
Private Sub txtQuant50_AfterUpdate()
    If Len(Me.txtQuant50.Value) = 0 Or Me.txtQuant50.Value = 0 Then
        Me.txtQuant50.Value = vbNullString
        Me.txtProduct50.Value = 0
        Exit Sub
    Else
        Me.txtQuant50.Value = CLng(Me.txtQuant50.Value)
    End If
    
    Me.txtProduct50.Value = Me.txtQuant50.Value * 50
    Me.txtProduct50.Value = Format(Me.txtProduct50.Value, "#,##0.00")
End Sub

Private Sub txtQuant50_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtTotalAmount_Change()
    If Len(Me.txtTotalAmount.Value) > 0 Then
        txtTotalAmount.Value = Format(txtTotalAmount.Value, "#,##0.00")
    Else
        Me.txtTotalAmount.Value = vbNullString
    End If
End Sub

Private Sub txtTotalPaid_Change()
    If Me.txtTotalPaid.Value > 0 Then
        Me.txtTotalPaid.Value = Format(Me.txtTotalPaid.Value, "#,##0.00")
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim ithUserID As Long
    ithUserID = Application.Match(CStr(Trim(frmUserLogOn.txtUserID.Value)), _
        ThisWorkbook.Sheets("user").Range("D:D"), 0)
    
    With Me
        .lblUserID.Caption = UCase(ThisWorkbook.Sheets("user").Range("D" & ithUserID).Value & " " & _
           ThisWorkbook.Sheets("user").Range("D" & ithUserID).Offset(0, -3).Value)
    End With
    
    Me.txtProduct1.Value = 0
    Me.txtProduct2.Value = 0
    Me.txtProduct5.Value = 0
    Me.txtProduct10.Value = 0
    Me.txtProduct20.Value = 0
    Me.txtProduct50.Value = 0
    Me.txtCoins.Value = 0
    
    'compute if there is a discount. note that discount can be applied only onces on the first payment
    'if payment count exceeds one then assign discount to 0
    Me.txtDiscount.Value = 0
                            
    Dim ithClientID As Long
    ithClientID = Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
        ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
        
    Me.txtClientID.Value = ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).Value
    
    Dim ithLoan As Long
    ithLoan = ThisWorkbook.Sheets("loan_list").Range("B" & Rows.Count).End(xlUp).Row
    
    With ThisWorkbook.Sheets("loan_list")
        .Range("A1").AutoFilter Field:=1, Criteria1:=CStr(frmClientInfoLogIn.txtClientID.Value)
        
        Dim loan As Range
        Dim loanRange As Range
        
        Set loanRange = .Range("B2:B" & ithLoan).Rows.SpecialCells(xlCellTypeVisible)
        
        For Each loan In loanRange
            Me.cboLoanID.AddItem loan.Value
        Next loan
    End With
    
    Me.cboLoanID.Text = Me.cboLoanID.List(Me.cboLoanID.ListCount - 1) 'minus 1 because .list(i) starts counting from 0. Hence setting value to the last item on the list
    'Me.cboLoanID.Text = Me.cboLoanID.List(WorksheetFunction.CountA(loanRange) - 1) 'minus 1 because .list(i) starts counting from 0
        
    If ThisWorkbook.Sheets("loan_list").AutoFilterMode = True Then
        ThisWorkbook.Sheets("loan_list").AutoFilterMode = False
    End If
    
    'formatting dateTime pickers
    Me.dtpPaymentDate.Value = Now()
    Me.dtpPaymentDate.Format = dtpCustom
    Me.dtpPaymentDate.CustomFormat = "dd-MMM-yyyy"
    Me.dtpDateMoneySentMM.Value = Now()
    Me.dtpDateMoneySentMM.Format = dtpCustom
    Me.dtpDateMoneySentMM.CustomFormat = "dd-MMM-yyyy"
    
    With Me.cboPaymentBy
        .AddItem "Client"
        .AddItem "Other"
    End With
    Me.cboPaymentBy.Value = "Client"
    
    With Me.cboPaymentMethod
        .AddItem "Cash"
        .AddItem "Mobile Money"
    End With
    Me.cboPaymentMethod.Value = Me.cboPaymentMethod.List(0)
    
    'filling out the fields using selected loan ID
    Dim selectedLoanID As Long
    selectedLoanID = Application.Match(CStr(Me.cboLoanID.Value), _
        ThisWorkbook.Sheets("loan_list").Range("B:B"), 0)
    With ThisWorkbook.Sheets("loan_list")
        Me.txtPrincipal.Value = .Range("D" & selectedLoanID).Value
        Me.txtInterest.Value = .Range("H" & selectedLoanID).Value
        Me.txtDuration.Value = .Range("F" & selectedLoanID).Value
        Me.txtPaymentperSchedule.Value = .Range("M" & selectedLoanID).Value
        Me.txtPrincipalInterest.Value = .Range("G" & selectedLoanID).Value
        Me.txtSchedule.Value = .Range("L" & selectedLoanID).Value
        Me.txtStartDate.Value = Format(.Range("O" & selectedLoanID).Value, "dd-Mmm-yyyy")
        Me.txtEndDate.Value = Format(.Range("P" & selectedLoanID).Value, "dd-Mmm-yyyy")
        
        If CStr(Me.txtSchedule.Value) = "Daily (1 day)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 1
        ElseIf CStr(Me.txtSchedule.Value) = "Weekly (7 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 7
        ElseIf CStr(Me.txtSchedule.Value) = "Bi-Weekly (14 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 14
        ElseIf CStr(Me.txtSchedule.Value) = "Monthly (28 days)" Then
            Me.txtTotalExpectedNumberofPayment.Value = CInt(Me.txtDuration.Value) / 28
        End If
        
    End With
    
    Dim loanIDonPayment As Long
    If Not IsError(Application.Match(CStr(Me.cboLoanID.Value), _
        ThisWorkbook.Sheets("loan_payment").Range("B:B"), 0)) Then
        
        With ThisWorkbook.Sheets("loan_payment")
            .Range("B1").AutoFilter Field:=2, Criteria1:=CStr(Me.cboLoanID.Value)
            Me.txtTotalPaid.Value = WorksheetFunction.Sum(.Range("E:E").Rows.SpecialCells(xlCellTypeVisible))
            Me.txtAmountRemain.Value = Me.txtPrincipalInterest.Value - Me.txtTotalPaid.Value
            Me.txtNumberofPayment.Value = WorksheetFunction.Count(.Range("E:E").Rows.SpecialCells(xlCellTypeVisible))
            Me.txtLastPMTDate.Value = Format(WorksheetFunction.Max(.Range("J:J").Rows.SpecialCells(xlCellTypeVisible)), "dd-Mmm-yyyy")
        End With
    Else
        Me.txtTotalPaid.Value = 0
        Me.txtAmountRemain.Value = Me.txtPrincipalInterest.Value
        Me.txtNumberofPayment.Value = 0
        Me.txtLastPMTDate.Value = "None"
    End If
    
    If ThisWorkbook.Sheets("loan_payment").AutoFilterMode = True Then
        ThisWorkbook.Sheets("loan_payment").AutoFilterMode = False
    End If
    
    If CDbl(Me.txtTotalPaid.Value) < CDbl(Me.txtPrincipalInterest.Value) Then
        Me.lblLoanStatus.Caption = "Still Owing on this Loan"
    Else
        Me.lblLoanStatus.Caption = "Loan is Fully Paid Off"
    End If
    
    'PMT date to detect delinquency
    
    Select Case CStr(Me.txtSchedule.Value)
        Case "Daily (1 day)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 1
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 1
            End If
        Case "Weekly (7 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 7
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 7
            End If
        Case "Bi-Weekly (14 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 14
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 14
            End If
        Case "Monthly (28 days)"
            If Me.txtAmountRemain.Value <= 0 Then
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value)) * 28
            Else
                Me.txtSchedulePMTDate.Value = CDate(Me.txtStartDate.Value) + _
                    (WorksheetFunction.Quotient(Me.txtTotalPaid.Value, Me.txtPaymentperSchedule.Value) + 1) * 28
            End If
    End Select
    Me.txtSchedulePMTDate.Value = Format(Me.txtSchedulePMTDate.Value, "dd-Mmm-yyyy")
    
    'detect payment delinquency
    Dim pmtTrack As Long
    pmtTrack = Now() - CDate(Me.txtSchedulePMTDate.Value)
    If pmtTrack < 0 Then
        Me.txtDelinquencyStatus.Value = "Early PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    ElseIf pmtTrack = 0 Then
        Me.txtDelinquencyStatus.Value = "On-Time PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    ElseIf pmtTrack > 0 Then
        Me.txtDelinquencyStatus.Value = "Late PMT"
        Me.txtPMTDaysOverdue.Value = pmtTrack
    End If
    If Me.txtDelinquencyStatus.Value = "Late PMT" Then Me.txtDelinquencyStatus.BackColor = vbRed
    
    'deteterming the discount allowed - based on just the previous loan
    'If in future client is allowed to take multiple loans then this part must be adjusted to a loop of all previous loans
    'not that care must be taking since multiple loan can complicate the calculation of discount
     
    If Me.txtTotalPaid.Value = 0 Then
        Dim previousLoanID As String
        Dim amountRemain As Double
        Dim previousInterest As Double
        Dim totalPaid As Double
        Dim principalPlusInterest As Double
        Dim previousLastPMTDate As Date
        Dim previousEndDate As Date
        
        If Me.cboLoanID.ListCount > 1 Then 'check if there are more than one loan applied by this client
            previousLoanID = Me.cboLoanID.List(Me.cboLoanID.ListCount - 2) 'set to the second list item from the bottom thus the previous loan
    
            Dim selectedPreviousLoanID As Long
            selectedPreviousLoanID = Application.Match(CStr(previousLoanID), _
                ThisWorkbook.Sheets("loan_list").Range("B:B"), 0)
            
            principalPlusInterest = ThisWorkbook.Sheets("loan_list").Range("G" & selectedPreviousLoanID).Value
            previousInterest = ThisWorkbook.Sheets("loan_list").Range("H" & selectedPreviousLoanID).Value
            previousEndDate = CDate(ThisWorkbook.Sheets("loan_list").Range("P" & selectedPreviousLoanID).Value)
            
            With ThisWorkbook.Sheets("loan_payment")
                .Range("B1").AutoFilter Field:=2, Criteria1:=CStr(previousLoanID)
                totalPaid = WorksheetFunction.Sum(.Range("E:E").Rows.SpecialCells(xlCellTypeVisible))
                previousLastPMTDate = CDate(WorksheetFunction.Max(.Range("J:J").Rows.SpecialCells(xlCellTypeVisible)))
                amountRemain = principalPlusInterest - totalPaid
            End With
            
            If ThisWorkbook.Sheets("loan_payment").AutoFilterMode = True Then
                ThisWorkbook.Sheets("loan_payment").AutoFilterMode = False
            End If
            
            If amountRemain = 0 Then
                Select Case CInt(previousEndDate - previousLastPMTDate)
                    Case 7 To 13
                        Me.txtDiscount.Value = Format(WorksheetFunction.Round(previousInterest * (2 / 100), 0), "#,#00.00")
                    Case 14 To 20
                        Me.txtDiscount.Value = Format(WorksheetFunction.Round(previousInterest * (4 / 100), 0), "#,#00.00")
                    Case 21 To 27
                        Me.txtDiscount.Value = Format(WorksheetFunction.Round(previousInterest * (6 / 100), 0), "#,#00.00")
                    Case Is >= 28
                        Me.txtDiscount.Value = Format(WorksheetFunction.Round(previousInterest * (8 / 100), 0), "#,#00.00")
                    Case Else
                        Me.txtDiscount.Value = 0
                End Select
            End If
            
        End If
    End If
    
    'setting payment type based on whether or not there is a discount
    If Me.txtDiscount.Value > 0 Then
        Me.cboPaymentType.Value = "Discounted Repayment"
    Else
        Me.cboPaymentType.Value = "Only Repayment"
    End If
    
    If Me.txtAmountRemain.Value = 0 Then
        Me.cmdPayLoan.Enabled = False
        Me.cmdCorrectPayment.Enabled = False
        Me.cmdMobileMoneyPayMM.Enabled = False
    End If
    
    'disenabling the new payment application
    Me.cmdNewPaymentApply.Enabled = False
End Sub
