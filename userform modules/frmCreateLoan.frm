VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateLoan 
   Caption         =   "Creating New Loan - Portfolio"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   OleObjectBlob   =   "frmCreateLoan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'you can't create loan for a client who is not done paying previous loan
'however, in future it will be allowed that an additional loan is created for a client who is in good standing even if the previous loan is not fully paid for
'but even if that should happen we have to require that client pays at least 60 percent of the previous loan
'FIND OUT HOW THE ABOVE CHANGES CAN AFFECT THE DISCOUNT CALCULATION

Private Sub cboDuration_Change()
    If Len(Me.txtPrincipal.Value) > 0 And Me.txtPrincipal.Value <> 0 Then
        cmdComputeLoan_Click
    End If
    
    'If Len(Me.txtPrincipalplusInterest.Value) > 0 Then
        Me.txtEndDate.Value = CDate(Me.txtStartDate.Value) + CInt(Me.cboDuration.Value)
    'End If
End Sub

Private Sub cboLoanPurpose_Change()
    If Me.cboLoanPurpose.Value = "Other (specify)" Then
        MsgBox prompt:="Please provide details of how this loan is going to be used", Title:="What is this loan for?"
        Me.txtPurposeSpecifics.SetFocus
    End If
End Sub

Private Sub cboPaymentSchedule_Change()
    If Len(Me.txtPrincipalplusInterest.Value) > 0 Then
        If Me.cboPaymentSchedule.Value = "Daily (1 day)" Then
            Me.txtAmountperSchedule.Value = (Me.txtPrincipalplusInterest.Value / (Me.cboDuration / 28)) / 28
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 1
        ElseIf Me.cboPaymentSchedule.Value = "Weekly (7 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtPrincipalplusInterest.Value / (Me.cboDuration / 28)) / 4
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 7
        ElseIf Me.cboPaymentSchedule.Value = "Bi-Weekly (14 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtPrincipalplusInterest.Value / (Me.cboDuration / 28)) / 2
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 14
        ElseIf Me.cboPaymentSchedule.Value = "Monthly (28 days)" Then
            Me.txtAmountperSchedule.Value = (Me.txtPrincipalplusInterest.Value / (Me.cboDuration / 28)) / 1
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 28
        End If
    Else
        If Me.cboPaymentSchedule.Value = "Daily (1 day)" Then
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 1
        ElseIf Me.cboPaymentSchedule.Value = "Weekly (7 days)" Then
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 7
        ElseIf Me.cboPaymentSchedule.Value = "Bi-Weekly (14 days)" Then
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 14
        ElseIf Me.cboPaymentSchedule.Value = "Monthly (28 days)" Then
            Me.txtNextDate.Value = CDate(Me.txtStartDate.Value) + 28
        End If
        'Exit Sub
    End If
End Sub

Private Sub cmdApproveLoan_Click()
    
    'verify if loan has already been created
    If Len(Me.txtLoanID.Value) > 0 Then
        MsgBox "Loan has already been created" + vbNewLine + "Cannot create another loan unless previous loan is paid off", Buttons:=vbCritical, Title:="Loan Information"
        Exit Sub
    End If
    
    'verify status of client
    If Me.txtClientStatus.Value <> "Active" Then
        MsgBox prompt:="This client is not in good standing." + vbNewLine + "You cannot create loan for this client", Title:="Client Status Information", Buttons:=vbCritical
        Exit Sub
    End If
    
    'verify terms and conditions
    If Me.chkTermsandCondition.Value = False Then
        MsgBox prompt:="Client must agree to terms and conditions." + vbNewLine + "Make sure conditions are reviewed with the client", _
            Title:="Terms and Conditions", Buttons:=vbCritical
        Exit Sub
    End If
    
    'verify principal
    If Me.txtPrincipal.Value = vbNullString Then
        MsgBox prompt:="Enter value for principal", Title:="Loan Information - Principal"
        Me.txtPrincipal.SetFocus
        Exit Sub
    End If
    
    'verify loan purpose
    If Me.cboLoanPurpose.Value = vbNullString Then
        MsgBox prompt:="Loan purpose must be specified", Title:="What is this loan for?"
        Me.cboLoanPurpose.SetFocus
        Exit Sub
    End If
    
    'verify loan purpose details
    If Me.cboLoanPurpose.Value = "Other (specify)" And Len(Trim(CStr(Me.txtPurposeSpecifics.Value))) < 16 Then
        MsgBox prompt:="Please provide details of how this loan is going to be used." + vbNewLine + _
            "Details must be at least 15 characters", Title:="What is this loan for?"
        Me.txtPurposeSpecifics.SetFocus
        Exit Sub
    End If
    
    'verify loan information
    Dim months As String
    If Me.cboDuration.Value = 28 Then
        months = " month"
    Else
        months = " months"
    End If
    If MsgBox(prompt:="Is this Loan Information Correct?" + vbNewLine + vbNewLine + _
            "A loan of GHS " & Me.txtPrincipal & " to be paid in " & Me.cboDuration & " days (" _
            & Me.cboDuration / 28 & months & ")." + vbNewLine + _
            "The interest for the whole term is GHS " & Me.txtInterest & " and payment is GHS " _
            & Me.txtAmountperSchedule & " " & Me.cboPaymentSchedule & "." + vbNewLine + _
            "The total amount to be paid as at the end of term is GHS " & Me.txtPrincipalplusInterest & " (loan amount plus interest).", Buttons:=vbInformation + vbYesNo, _
            Title:="Summary of Loan - Please verify") = vbNo Then
        Exit Sub
    End If
    
    'removing any filter on loan_list sheet if any
    If ThisWorkbook.Sheets("loan_list").AutoFilterMode = True Then
        ThisWorkbook.Sheets("loan_list").AutoFilterMode = False
    End If
    
    Dim ithLoan As Long
    ithLoan = ThisWorkbook.Sheets("loan_list").Range("A" & Rows.Count).End(xlUp).Row
    
    Dim ithClientLoan As Long
    With ThisWorkbook.Sheets("loan_list")
        .Range("A1").AutoFilter Field:=1, Criteria1:=CStr(frmClientInfoLogIn.txtClientID.Value)
        ithClientLoan = WorksheetFunction.CountA(.Range("B:B").Rows.SpecialCells(xlCellTypeVisible))
    End With
    ThisWorkbook.Sheets("loan_list").AutoFilterMode = False
    
    'send data to worksheet
    With ThisWorkbook.Sheets("loan_list")
        .Range("A" & ithLoan + 1).Value = Me.txtClientID.Value
        .Range("B" & ithLoan + 1).Value = loan_id(CLng(Me.txtClientID.Value), ithClientLoan, ithLoan, CDbl(Me.txtPrincipal.Value), Now()) 'change now to the datetime picker value
        .Range("C" & ithLoan + 1).Value = Me.lblUserID.Caption
        .Range("D" & ithLoan + 1).Value = Me.txtPrincipal.Value
        .Range("E" & ithLoan + 1).Value = Me.cboRate.Value
        .Range("F" & ithLoan + 1).Value = Me.cboDuration.Value
        .Range("G" & ithLoan + 1).Value = Me.txtPrincipalplusInterest.Value
        .Range("H" & ithLoan + 1).Value = Me.txtInterest.Value
        .Range("I" & ithLoan + 1).Value = 0 'setting to zero because I don't want to delete that field which would affect other functions and computations
        .Range("J" & ithLoan + 1).Value = 0 'setting to zero because I don't want to delete that field which would affect other functions and computations
        .Range("K" & ithLoan + 1).Value = Me.txtPrincipalplusInterest.Value - 0
        .Range("L" & ithLoan + 1).Value = Me.cboPaymentSchedule.Value
        .Range("M" & ithLoan + 1).Value = Me.txtAmountperSchedule.Value
        .Range("N" & ithLoan + 1).Value = Format(Now(), "dd-Mmm-yyyy")
        .Range("O" & ithLoan + 1).Value = Format(CDate(.Range("N" & ithLoan + 1).Value) + 1, "dd-Mmm-yyyy")
        .Range("P" & ithLoan + 1).Value = Format(CDate(.Range("O" & ithLoan + 1).Value) + CInt(Me.cboDuration.Value), "dd-Mmm-yyyy")
        .Range("Q" & ithLoan + 1).Value = Me.txtClientStatus.Value
        .Range("R" & ithLoan + 1).Value = Me.cboLoanPurpose.Value
        .Range("S" & ithLoan + 1).Value = Me.txtPurposeSpecifics.Value
    End With
    Me.txtLoanID.Value = loan_id(Me.txtClientID.Value, ithClientLoan, ithLoan, Me.txtPrincipal.Value, Now()) 'change now to the datetime picker value
    
    'save workbook
    ThisWorkbook.Save
    MsgBox "this will proceed to print contract"
    
    'create folder for loan
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim fPath As String
    fPath = ThisWorkbook.Path
    
    If Not fso.FolderExists(fPath + "\" + CStr(Me.txtClientID.Value) + "\" + CStr(Me.txtLoanID.Value)) Then
        fso.CreateFolder fPath + "\" + CStr(Me.txtClientID.Value) + "\" + CStr(Me.txtLoanID.Value)
    End If
    
    If Not fso.FolderExists(fPath + "\" + CStr(Me.txtClientID.Value) + "\" + CStr(Me.txtLoanID.Value) + "\" + "PMT") Then
        fso.CreateFolder fPath + "\" + CStr(Me.txtClientID.Value) + "\" + CStr(Me.txtLoanID.Value) + "\" + "PMT"
    End If
    
    Set fso = Nothing
    
    Me.cmdApproveLoan.Enabled = False
    Me.cmdViewSummary.Enabled = False
    
    'cleaning worksheet
    With ThisWorkbook.Sheets("pmt_receipt").Cells
        .Clear
        .Font.Name = "Calibri"
        .Font.Size = 9
    End With
    
    'setting tags
    With ThisWorkbook.Sheets("pmt_receipt")
        .Range("B1").Value = "REPAYMENT RECEIPT FOR " & UCase(Me.txtFirstName.Value) & " " & UCase(Me.txtLastName.Value)
        .Range("B1:D1").Merge
        .Range("B1:D1").HorizontalAlignment = xlCenter
        .Range("B1:D1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B1:D1").Borders(xlEdgeBottom).ColorIndex = 46
        
        .Range("A2").Value = "Client ID:"
        .Range("A3").Value = "Schd. Amt:"
        
        .Range("C2").Value = "Principal + Interest:"
        .Range("C3").Value = "PMT Schedule:"
        
        .Range("E2").Value = "Start Date:"
        .Range("E3").Value = "End Date:"
        
        .Range("F4").Value = "GHS"
        .Range("F4").Font.Size = 7
        .Range("F4").HorizontalAlignment = xlRight
        .Range("F4").Font.Bold = True
        
        .Range("A5").Value = "PMT #"
        .Range("B5").Value = "PMT Date"
        .Range("C5").Value = "PMT Method"
        .Range("D5").Value = "PMT Type"
        .Range("E5").Value = "PMT By"
        .Range("F5").Value = "PMT Amount"
        
        .Range("A5:F5").Font.Bold = True
        .Range("A5:F5").HorizontalAlignment = xlCenter
        .Range("A2:F3").HorizontalAlignment = xlLeft
    End With
    
    With ThisWorkbook.Sheets("pmt_receipt")
        .Range("B2").Value = Me.txtClientID.Value
        .Range("B3").Value = Me.txtAmountperSchedule.Value
        .Range("B3").NumberFormat = "#,##0.00"
        .Range("D2").Value = Me.txtPrincipalplusInterest.Value
        .Range("D3").Value = Me.cboPaymentSchedule.Value
        .Range("F2").Value = Format(Me.txtStartDate.Value, "dd-Mmm-yyyy")
        .Range("F3").Value = Format(Me.txtEndDate.Value, "dd-Mmm-yyyy")
        .Range("A3:F3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:F3").Borders(xlEdgeBottom).ColorIndex = 0
        .Range("A5:F5").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A5:F5").Borders(xlEdgeBottom).ColorIndex = 0
        .Columns.AutoFit
    End With
    
    Dim expectedNumberofPayment As Long
    If CStr(Me.cboPaymentSchedule.Value) = "Daily (1 day)" Then
            expectedNumberofPayment = CInt(Me.cboDuration.Value) / 1
        ElseIf CStr(Me.cboPaymentSchedule.Value) = "Weekly (7 days)" Then
            expectedNumberofPayment = CInt(Me.cboDuration.Value) / 7
        ElseIf CStr(Me.cboPaymentSchedule.Value) = "Bi-Weekly (14 days)" Then
            expectedNumberofPayment = CInt(Me.cboDuration.Value) / 14
        ElseIf CStr(Me.cboPaymentSchedule.Value) = "Monthly (28 days)" Then
            expectedNumberofPayment = CInt(Me.cboDuration.Value) / 28
    End If
    
    Dim i As Long
    For i = 1 To expectedNumberofPayment
        With ThisWorkbook.Sheets("pmt_receipt")
            .Range("A" & 5 + i).Value = i
            .Range("A" & 5 + i).HorizontalAlignment = xlLeft
            'making grid lines
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlEdgeBottom).ThemeColor = 1
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlEdgeBottom).Weight = xlThin
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlEdgeBottom).TintAndShade = -0.249946592608417
            
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlInsideVertical).ThemeColor = 1
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlInsideVertical).Weight = xlThin
            .Range("A" & 5 + i & ":" & "F" & 5 + i).Borders(xlInsideVertical).TintAndShade = -0.249946592608417
        End With
    Next i
    
    'reset form
    Call cmdReset_Click
    
    'save as pdf
    On Error Resume Next
    ThisWorkbook.Sheets("pmt_receipt").ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path + "\" + _
        CStr(Me.txtClientID.Value) + "\" + CStr(Me.txtLoanID.Value) + "\" + CStr(Me.txtLoanID.Value) + ".pdf", Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End Sub

Private Sub cmdComputeLoan_Click()
    If Len(Me.txtPrincipal.Value) > 0 And Len(Me.cboRate.Value) > 0 And Len(Me.cboDuration.Value) > 0 Then
        With Me
            .txtPrincipalplusInterest.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28) + .txtPrincipal.Value
            .txtPrincipalplusInterest.Value = Format(.txtPrincipalplusInterest.Value, "#,##0.00")
            
            .txtInterest.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28)
            .txtInterest.Value = Format(.txtInterest.Value, "#,##0.00")

             If Len(Me.cboPaymentSchedule.Value) > 0 Then
                Call cboPaymentSchedule_Change
             Else
                .cboPaymentSchedule.Value = "Bi-Weekly (14 days)"
                .txtAmountperSchedule.Value = (.txtPrincipalplusInterest.Value / (.cboDuration / 28)) / 2
            End If
        End With
    Else
        If Len(Me.txtPrincipal.Value) <= 0 Then
            MsgBox prompt:="Enter value for principal", Title:="Loan Information - Principal"
            Me.txtPrincipal.SetFocus
        End If
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
    If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        frmClientProfile.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdMainMenu_Click()
    If Me.txtLoanID.Value = vbNullString Then
        If MsgBox(prompt:="This loan is not approved. You may click ""Approve Loan"" to approve it." + vbNewLine + "Would you like to approve?", _
            Title:="Loan Approval Notice", Buttons:=vbYesNo) = vbYes Then
            Exit Sub
        End If
    End If
    
    If MsgBox(prompt:="Would you like to access main menu? Client ID will be required the next time.", Title:="Back to Main Menu", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        Unload frmClientInfoLogIn
        frmLoanMenu.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdReset_Click()
     With Me
        .txtPrincipal.Value = ""
        .cboRate.Value = 6
        .cboDuration.Value = 112
        .cboPaymentSchedule.Value = "Bi-Weekly (14 days)"
        .cboLoanPurpose.Value = ""
        .txtPurposeSpecifics.Value = ""
    End With
End Sub

Private Sub cmdViewSummary_Click()
    If Len(Me.txtPrincipal.Value) > 0 And Len(Me.cboRate.Value) > 0 And Len(Me.cboDuration.Value) > 0 Then
        If Len(Me.txtPrincipalplusInterest) > 0 Then
            Dim months As String
            If Me.cboDuration.Value = 28 Then
                months = " month"
            Else
                months = " months"
            End If
            MsgBox prompt:="A loan of GHS " & Me.txtPrincipal & " to be paid in " & Me.cboDuration & " days (" _
                 & Me.cboDuration / 28 & months & ")." + vbNewLine + _
                    "The interest for the whole term is GHS " & Me.txtInterest & " and payment is GHS " _
                    & Me.txtAmountperSchedule & " " & Me.cboPaymentSchedule & "." + vbNewLine + _
                    "The total amount to be paid as at the end of term is GHS " & Me.txtPrincipalplusInterest & " (loan amount plus interest).", Buttons:=vbInformation, _
                    Title:="Summary of Loan"
        Else
            Call cmdComputeLoan_Click
            Dim month As String
            If Me.cboDuration.Value = 28 Then
                month = " month"
            Else
                month = " months"
            End If
            MsgBox prompt:="A loan of GHS " & Me.txtPrincipal & " to be paid in " & Me.cboDuration & " days (" _
                 & Me.cboDuration / 28 & month & ")." + vbNewLine + _
                    "The interest for the whole term is GHS " & Me.txtInterest & " and payment is GHS " _
                    & Me.txtAmountperSchedule & " " & Me.cboPaymentSchedule & "." + vbNewLine + _
                    "The total amount to be paid as at the end of term is GHS " & Me.txtPrincipalplusInterest & " (loan amount plus interest).", Buttons:=vbInformation, _
                    Title:="Summary of Loan"
        End If
    Else
        MsgBox prompt:="Either principal or rate or duration is missing", Title:="Loan Amount Information"
        Me.txtPrincipal.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAmountperSchedule_Change()
    txtAmountperSchedule.Value = Format(txtAmountperSchedule.Value, "#,##0.00")
End Sub

Private Sub cboRate_Change()
    If Len(Me.txtPrincipal.Value) > 0 And Me.txtPrincipal.Value <> 0 Then
        cmdComputeLoan_Click
    End If
End Sub

Private Sub txtEndDate_Change()
    Me.txtEndDate.Value = Format(Me.txtEndDate, "dd-Mmm-yyyy hh:mm:ss")
End Sub

Private Sub txtNextDate_Change()
    Me.txtNextDate.Value = Format(Me.txtNextDate, "dd-Mmm-yyyy hh:mm:ss")
End Sub

Private Sub txtPrincipal_Change()
    With Me
        .txtPrincipalplusInterest.Value = vbNullString
        .txtInterest.Value = vbNullString
        .txtAmountperSchedule.Value = vbNullString
    End With
End Sub

Private Sub txtPrincipal_AfterUpdate()
    If Len(Me.txtPrincipal.Value) = 0 Or Me.txtPrincipal.Value = 0 Then
        Me.txtPrincipal.Value = vbNullString
        Exit Sub
    Else
        Me.txtPrincipal.Value = Format(Me.txtPrincipal.Value, "#,##0.00")
    End If
End Sub

Private Sub txtPrincipal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
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
   
    With Me
        .txtPrincipal.Value = 1000
        .txtPrincipal.Value = Format(.txtPrincipal.Value, "#,##0.00")
        
        .txtStartDate.Value = Format(Now(), "dd-Mmm-yyyy hh:mm:ss")
        .txtNextDate.Value = Now() + 1 + 14
        .txtEndDate.Value = Now() + 1 + 112
        
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
        .cboRate.Value = 6#
        
        .txtPrincipalplusInterest.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28) + .txtPrincipal.Value
        .txtPrincipalplusInterest.Value = Format(.txtPrincipalplusInterest.Value, "#,##0.00")
        
        .txtInterest.Value = (.txtPrincipal.Value * .cboRate.Value / 100) * (.cboDuration / 28)
        .txtInterest.Value = Format(.txtInterest.Value, "#,##0.00")
        
        .txtAmountperSchedule.Value = (.txtPrincipalplusInterest.Value / (.cboDuration / 28)) / 2
        .txtAmountperSchedule.Value = Format(.txtAmountperSchedule.Value, "#,##0.00")
        
        With .cboPaymentSchedule
            .AddItem "Daily (1 day)"
            .AddItem "Weekly (7 days)"
            .AddItem "Bi-Weekly (14 days)"
            .AddItem "Monthly (28 days)"
        End With
        .cboPaymentSchedule.Value = "Bi-Weekly (14 days)"
        
        With .cboLoanPurpose
            .AddItem "Infrastrature"
            .AddItem "More Goods"
            .AddItem "Other (specify)"
        End With
    End With
    
    Dim ithClientID As Long
    ithClientID = Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
        ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
    
    Dim ithLoan As Long
    ithLoan = ThisWorkbook.Sheets("loan_list").Range("A" & Rows.Count).End(xlUp).Row
 
    If Not IsError(Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
        ThisWorkbook.Sheets("loan_list").Range("A:A"), 0)) Then
        ThisWorkbook.Sheets("dirty_client_info").Cells.Clear
        
        With ThisWorkbook
            .Sheets("loan_list").Range("A1").AutoFilter Field:=1, Criteria1:=CStr(frmClientInfoLogIn.txtClientID.Value)
             
            .Sheets("loan_list").Range("A1:S" & ithLoan).Rows.SpecialCells(xlCellTypeVisible).Copy Destination:= _
                .Sheets("dirty_client_info").Range("A1")
               
            Dim loanInfo As Range
            Set loanInfo = ThisWorkbook.Sheets("dirty_client_info").UsedRange
            
            ThisWorkbook.Sheets("dirty_client_info").Select 'not useful when using .list property. but useful for .rowsource property
            With Me.lstLoanHistory
                .ColumnCount = 19
                .ColumnHeads = True 'not so useful for .list property
                '.List = loanInfo.Offset.Value
                .RowSource = loanInfo.Offset(1, 0).Address
            End With
                
        End With
    Else
        Me.lstLoanHistory.AddItem "No Loan History for this Client. You may apply for Loan"
    End If
    
    If ThisWorkbook.Sheets("loan_list").AutoFilterMode = True Then
        ThisWorkbook.Sheets("loan_list").AutoFilterMode = False
    End If
    
    With Me
        With ThisWorkbook.Sheets("client_info_personal")
            txtClientID.Value = .Range("A" & ithClientID).Value
            txtFirstName.Value = .Range("B" & ithClientID).Value
            txtLastName.Value = .Range("C" & ithClientID).Value
            txtClientStatus.Value = .Range("J" & ithClientID).Value
        End With
    End With
    
End Sub

