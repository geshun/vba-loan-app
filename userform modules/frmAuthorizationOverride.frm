VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAuthorizationOverride 
   Caption         =   "Select Client to Log In"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12255
   OleObjectBlob   =   "frmAuthorizationOverride.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAuthorizationOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        frmLoanMenu.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo handler
    frmClientInfoLogIn.txtClientID.Value = Me.lstClientInfoExtra.Value
            
    Me.Hide
    frmClientProfile.Show
handler:
    Exit Sub
End Sub

Private Sub lstClientInfoExtra_Change()
    Me.lblClientID.Caption = "You have selected: " + CStr(Me.lstClientInfoExtra.Value)
End Sub

Private Sub UserForm_Initialize()
    
    
    ThisWorkbook.Sheets("dirty_client_info").Cells.Clear
    'ThisWorkbook.Sheets("dirty_client_info").Range("A1").Select
    
    Dim ithInfoPersonal As Long
    ithInfoPersonal = ThisWorkbook.Sheets("client_info_personal").Range("A" & Rows.Count).End(xlUp).Row
             
    Dim infoPersonalRange As Range
        
    ThisWorkbook.Sheets("client_info_personal").Range("A2:I" & ithInfoPersonal).Rows.SpecialCells(xlCellTypeVisible).Copy _
        Destination:=ThisWorkbook.Sheets("dirty_client_info").Range("A1")
        
    Set infoPersonalRange = ThisWorkbook.Sheets("dirty_client_info").Range("A1").CurrentRegion
    
    With Me.lstClientInfoExtra
        .ColumnCount = 9
        .List = infoPersonalRange.Value
    End With
    
    With ThisWorkbook.Sheets("client_info_personal")
        If .AutoFilterMode = True Then
            .AutoFilterMode = False
        End If
    End With
End Sub
