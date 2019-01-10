VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientInfoLoginExtra 
   Caption         =   "Select Client to Log In"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10485
   OleObjectBlob   =   "frmClientInfoLoginExtra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClientInfoLoginExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    frmClientInfoLogIn.txtClientID.Value = Me.lstClientInfoExtra.Value
            
    Me.Hide
    frmClientProfile.Show
End Sub

Private Sub UserForm_Initialize()
    Dim ithInfoPersonal As Long
    ithInfoPersonal = ThisWorkbook.Sheets("client_info_personal").Range("B" & Rows.Count).End(xlUp).Row
             
    Dim infoPersonalRange As Range
        
    Set infoPersonalRange = ThisWorkbook.Sheets("client_info_personal").Range("A2:I" & ithInfoPersonal).Rows.SpecialCells(xlCellTypeVisible)
            
    With Me.lstClientInfoExtra
        .ColumnCount = 9
        '.ColumnHeads = True
        .RowSource = infoPersonalRange.Address
    End With
    
    If ThisWorkbook.Sheets("client_info_personal").AutoFilterMode = True Then
        ThisWorkbook.Sheets("client_info_personal").AutoFilterMode = False
    End If
End Sub
