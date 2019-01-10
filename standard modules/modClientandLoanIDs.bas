Attribute VB_Name = "modClientandLoanIDs"
Option Explicit

Function client_id_age(birthDate As Date, Optional currentDate As Date = 0) As Integer
    ' Calculate a person's age, given the person's birth date and
    ' an optional "current" date.
    If currentDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        currentDate = Date
    End If
    client_id_age = DateDiff("yyyy", birthDate, currentDate) + _
        (currentDate < DateSerial(Year(currentDate), month(birthDate), Day(birthDate)))
End Function


Function loan_id(clientID As Long, nthLoan As Long, nthRow As Long, loanPrincipal As Double, loanDate As Date) As String
    
    'date the loan was made. change now to the value of the datetime picker
    Dim loanDateStr As String
    loanDateStr = CStr(Year(Now)) + _
        CStr(month(Now)) + CStr(Day(Now))
    
    'loan amount in millions "M" or thousands "K" or hundreds "H"
    Dim loanPrincipalStr As String
    If loanPrincipal >= 1000 And loanPrincipal < 1000000 Then
        loanPrincipal = Round(loanPrincipal / 1000, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "K"
    ElseIf loanPrincipal >= 1000000 Then
        loanPrincipal = Round(loanPrincipal / 1000000, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "M"
    ElseIf loanPrincipal >= 100 And loanPrincipal < 1000 Then
        loanPrincipal = Round(loanPrincipal / 100, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "H"
    End If
    
    'clientID to string
    Dim clintIDStr As String
    clintIDStr = Trim(CStr(clientID))
    
    'loan number per client to string
    Dim nthLoanStr As String
    nthLoanStr = Trim(CStr(nthLoan))
    
    'loan row to string
    Dim nthRowStr As String
    nthRowStr = Trim(CStr(nthRow))
    
    'return value of function
    loan_id = clintIDStr + "." + nthLoanStr + "." + nthRowStr + "." + loanPrincipalStr + "." + loanDateStr

End Function

Function payment_id(clientID As Long, loanPrincipal As Double, nthLoan As Long, nthPay As Long, nthAmount As Long) As String
    
    'loan amount in millions "M" or thousands "K" or hundreds "H"
    Dim loanPrincipalStr As String
    If loanPrincipal >= 1000 And loanPrincipal < 1000000 Then
        loanPrincipal = Round(loanPrincipal / 1000, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "K"
    ElseIf loanPrincipal >= 1000000 Then
        loanPrincipal = Round(loanPrincipal / 1000000, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "M"
    ElseIf loanPrincipal >= 100 And loanPrincipal < 1000 Then
        loanPrincipal = Round(loanPrincipal / 100, 0)
        loanPrincipalStr = CStr(loanPrincipal) + "H"
    End If
    
    'clientID to string
    Dim clintIDStr As String
    clintIDStr = Trim(CStr(clientID))
    
    'loan number per client to string
    Dim nthLoanStr As String
    nthLoanStr = Trim(CStr(nthLoan))
    
    'loan row to string
    Dim nthPayStr As String
    nthPayStr = Trim(CStr(nthPay))
    
    'loan row to string
    Dim nthAmountStr As String
    nthAmountStr = Trim(CStr(Round(nthAmount, 0))) + "PMT"
    
    'return value of function
    payment_id = clintIDStr + "." + loanPrincipalStr + "." + nthLoanStr + "." + nthPayStr + "." + nthAmountStr
End Function
