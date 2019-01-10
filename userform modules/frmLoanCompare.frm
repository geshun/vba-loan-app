VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoanCompare 
   Caption         =   "Loan Compare Scenarios"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14640
   OleObjectBlob   =   "frmLoanCompare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoanCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboExit_Click()
    If MsgBox(prompt:="Would you like to exit?", Title:="Exiting...", Buttons:=vbYesNo) = vbYes Then
        Unload Me
        frmLoanMenu.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub
