Attribute VB_Name = "modADOLibrary"
Option Explicit

Sub get_data()
    Dim bookString As String
    bookString = ActiveWorkbook.FullName
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fld As ADODB.Field
    'Dim objerror As ADODB.Error
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & bookString & ";" & _
        "Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    
    conn.Open
    
    'For Each objerror In conn.Errors
    '    Debug.Print objerror.Description
    'Next objerror
    
    With rs
        .ActiveConnection = conn
        .Source = "SELECT [Payment Date],[Payment Method],[Payment Type],[Payment By],[Amount Paid] FROM [loan_payment$] WHERE [Loan ID] = ""270010001.6.6.1K.2016815"";"
        '.source = "SELECT [Payment Date],[Payment Method],[Payment Type],[Payment By],[Amount Paid] FROM [loan_payment$] WHERE [Loan ID] = ""270010001.6.6.1K.2016815"";"
        '.source = "[detail$]"
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open
    End With
    
    Worksheets.Add
    Range("A2").CopyFromRecordset rs
    Range("A1").Select
    
    For Each fld In rs.Fields
        ActiveCell.Offset(0, 1).Select
    Next fld
    
    rs.Close
    If CBool(conn.State And adStateOpen) Then
        conn.Close
    End If
    
    Set rs = Nothing
    Set conn = Nothing
End Sub

Sub learn()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fld As ADODB.Field
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.ConnectionString = ""
    conn.Open
    
    With rs
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .Source = "SELECT * FROM []"
        .LockType = adLockReadOnly
        .Open
    End With
    
    ThisWorkbook.Sheets("trying").Range("A1").CopyFromRecordset rs
    
    rs.Close
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
End Sub

Sub get_data_sqlExcel(sourcePathConn As String, destPath As Worksheet, destRange As Range, sqlStatement As String)
    
    'sourcePathConn As String, destPath As Worksheet, destRange As Range, sqlStatement As String
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fld As ADODB.Field
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePathConn & ";" & _
        "Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    conn.Open
    
    With rs
        .ActiveConnection = conn
        .Source = sqlStatement
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open
    End With
    
    destRange.CopyFromRecordset rs
    
    rs.Close
    If CBool(conn.State And adStateOpen) Then
        conn.Close
    End If
    
    Set rs = Nothing
    Set conn = Nothing
End Sub


Sub trying()
    Dim sourceConn As String
    Dim destWks As Worksheet
    Dim destRng As Range
    Dim sqlStatement As String
    
    sourceConn = ActiveWorkbook.FullName
    Set destWks = ThisWorkbook.Sheets("pmt_receipt")
    Set destRng = destWks.Range("B6")
    
    sqlStatement = "SELECT [Payment Date],[Payment Method],[Payment Type],[Payment By],[Amount Paid] FROM [loan_payment$] WHERE [Loan ID] = ""270010001.7.7.1K.2016817"";"
    
    Call get_data_sqlExcel(sourceConn, destWks, destRng, sqlStatement)
End Sub
