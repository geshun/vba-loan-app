VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddClient 
   Caption         =   "New Client"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13995
   OleObjectBlob   =   "frmAddClient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldInfoCollection As Collection

Private Sub cboBusinessOwner_Change()
    If Me.cboBusinessOwner.Value = "Self Owned" Then
        With Me
            .lblCoOwnerName.Visible = False
            .txtCoOwnerName.Value = vbNullString
            .txtCoOwnerName.Visible = False
            .lblCoOwnerRelationship.Visible = False
            .cboCoOwnerRelationship.Value = vbNullString
            .cboCoOwnerRelationship.Visible = False
        End With
    Else
        With Me
            .lblCoOwnerName.Visible = True
            .txtCoOwnerName.Visible = True
            .lblCoOwnerRelationship.Visible = True
            .cboCoOwnerRelationship.Visible = True
        End With
    End If
End Sub

Private Sub cboBusinessType_Change()
    Select Case switchAddEdit
        Case False
            If cboBusinessType.Value = "None" Then
                With Me
                    .txtBusinessName.Value = vbNullString
                    .txtBusinessAddress.Value = vbNullString
                    .txtYearsinBusiness.Value = vbNullString
                    .lblBusinessName.Caption = "Type of Work *"
                    .lblBusinessAddress.Caption = "Work Address *"
                    .lblYearsinBusiness.Caption = "Years Working *"
                    .lblBusinessOwner.Visible = False
                    .cboBusinessOwner.Value = vbNullString
                    .cboBusinessOwner.Visible = False
                    .lblCoOwnerName.Visible = False
                    .txtCoOwnerName.Value = vbNullString
                    .txtCoOwnerName.Visible = False
                    .lblCoOwnerRelationship.Visible = False
                    .cboCoOwnerRelationship.Value = vbNullString
                    .cboCoOwnerRelationship.Visible = False
                End With
            Else
                With Me
                    .txtBusinessName.Value = vbNullString
                    .txtBusinessAddress.Value = vbNullString
                    .txtYearsinBusiness.Value = vbNullString
                    .lblBusinessName.Caption = "Business Name *"
                    .lblBusinessAddress.Caption = "Business Address *"
                    .lblYearsinBusiness.Caption = "Years in Business *"
                    .lblBusinessOwner.Visible = True
                    .cboBusinessOwner.Visible = True
                    .lblCoOwnerName.Visible = True
                    .txtCoOwnerName.Visible = True
                    .lblCoOwnerRelationship.Visible = True
                    .cboCoOwnerRelationship.Visible = True
                End With
            End If
        Case True
        
    End Select
End Sub

Private Sub cmdExit_Click()
    
    'create a collection and gather the form controls as collection
    Dim newInfoCollectionExit As Collection
    Set newInfoCollectionExit = New Collection
                
    Dim ctrl As Control
                
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "CheckBox" Then
            newInfoCollectionExit.Add ctrl.Value
        End If
    Next ctrl
                
    Dim changeInfoIntExit As Integer
    Dim ithCollectionExit As Integer
                
    For ithCollectionExit = 1 To newInfoCollectionExit.Count
        If oldInfoCollection(ithCollectionExit) <> newInfoCollectionExit(ithCollectionExit) Then
            changeInfoIntExit = changeInfoIntExit + 1
            Exit For
        End If
    Next ithCollectionExit
                
    Select Case switchAddEdit
        Case False
            If Len(Me.txtClientID.Value) <> 0 Then 'to check if saved and client ID copied. If saved then there is a client ID
                If MsgBox(prompt:="Please make sure you have written down the client ID before exiting" _
                    + vbNewLine + "Have you copied the client's ID?", Title:="Copy Client ID", Buttons:=vbInformation + vbYesNo) = vbNo Then
                    Exit Sub
                Else
                    Unload Me
                    frmLoanMenu.Show
                End If
            Else
                If changeInfoIntExit <> 0 Then
                    If MsgBox(prompt:="You are in the process of creating new client" + vbNewLine + _
                        "Do you want to exit?", Title:="Process of Adding New Client", Buttons:=vbInformation + vbYesNo) = vbNo Then
                        Exit Sub
                    Else
                        Unload Me
                        frmLoanMenu.Show
                    End If
                Else
                    Unload Me
                    frmLoanMenu.Show
                End If
            End If
            
        Case Else
            If changeInfoIntExit <> 0 Then
                MsgBox prompt:="There are changes made to the client's information" + vbNewLine + _
                    "Update before exiting.", Title:="Exiting client profile update", Buttons:=vbInformation
                Exit Sub
                    
            Else
                Unload Me
                frmClientProfile.Show
            End If
    End Select
End Sub

Private Sub cmdReset_Click()
    Dim ctrl As Control
    Select Case switchAddEdit
        Case False
            If MsgBox(prompt:="This will erase all information entered" + vbNewLine + _
                "Would you like to proceed?", Title:="Resetting fields", Buttons:=vbInformation + vbYesNo) = vbNo Then
                Exit Sub
            Else
                For Each ctrl In Me.Controls
                    If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
                        ctrl.Value = vbNullString
                    End If
                    If TypeName(ctrl) = "CheckBox" Then
                        ctrl.Value = False
                    End If
                Next ctrl
            End If
        Case Else
        
            'create a collection and gather the form controls as collection
            Dim newInfoCollectionRemark As Collection
            Set newInfoCollectionRemark = New Collection
                
            For Each ctrl In Me.Controls
                If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "CheckBox" Then
                    newInfoCollectionRemark.Add ctrl.Value
                End If
            Next ctrl
                
            Dim changeInfoIntRemark As Integer
            Dim ithCollectionRemark As Integer
                
            For ithCollectionRemark = 1 To newInfoCollectionRemark.Count
                If oldInfoCollection(ithCollectionRemark) <> newInfoCollectionRemark(ithCollectionRemark) Then
                    changeInfoIntRemark = changeInfoIntRemark + 1
                    Exit For
                End If
            Next ithCollectionRemark
            
            If changeInfoIntRemark = 0 Then
                MsgBox prompt:="There are no changes made to the client's information" + vbNewLine + _
                    "No Need to Reset.", Title:="Exiting client profile update", Buttons:=vbInformation
                Exit Sub
            End If
        
            If MsgBox(prompt:="This will bring back original client information" + vbNewLine + _
                "Would you like to proceed?", Title:="Resetting fields", Buttons:=vbInformation + vbYesNo) = vbNo Then
                Exit Sub
            Else
                'i could have just looped through
                'but my controls where randomely created on the form and the collection chose the internal randomness even though there is a tab order
                Me.txtFirstName.Value = oldInfoCollection(1)
                Me.txtLastName.Value = oldInfoCollection(2)
                Me.cboGender.Value = oldInfoCollection(3)
                Me.cboAgeRange.Value = oldInfoCollection(4)
                Me.txtIDNumber.Value = oldInfoCollection(5)
                Me.cboSocialMedia.Value = oldInfoCollection(6)
                Me.txtHouseAddress.Value = oldInfoCollection(7)
                Me.txtPrimaryPhone.Value = oldInfoCollection(8)
                Me.txtSecondaryPhone.Value = oldInfoCollection(9)
                Me.cboIDType.Value = oldInfoCollection(10)
                Me.cboBusinessType.Value = oldInfoCollection(11)
                Me.txtPostalAddress.Value = oldInfoCollection(12)
                Me.txtEmail.Value = oldInfoCollection(13)
                Me.txtBusinessName.Value = oldInfoCollection(14)
                Me.txtBusinessAddress.Value = oldInfoCollection(15)
                Me.txtFamilySize.Value = oldInfoCollection(16)
                Me.cboMaritalStatus.Value = oldInfoCollection(17)
                Me.txtReferee.Value = oldInfoCollection(18)
                Me.txtRefereeContact.Value = oldInfoCollection(19)
                Me.cboBusinessOwner.Value = oldInfoCollection(20)
                Me.txtYoungestAge.Value = oldInfoCollection(21)
                Me.txtYearsLived.Value = oldInfoCollection(22)
                Me.txtYearsinBusiness.Value = oldInfoCollection(23)
                Me.chkFirstTime.Value = oldInfoCollection(24)
                Me.chkBelongstoGroup.Value = oldInfoCollection(25)
                'Me.txtClientID.Value = oldInfoCollection(26)
                Me.txtCoOwnerName.Value = oldInfoCollection(27)
                Me.cboCoOwnerRelationship.Value = oldInfoCollection(28)
                Me.cboClientStatus.Value = oldInfoCollection(29)
                Me.cboBelongstoGroup.Value = oldInfoCollection(30)
                Me.txtRemark.Value = oldInfoCollection(31)
                Me.txtMiddleName.Value = oldInfoCollection(32)
            End If
    End Select
End Sub

Private Sub cmdSave_Click()
    
    'verify required fields - personal information
    If Len(Trim(Me.txtFirstName.Value)) = 0 Or _
        Len(Trim(Me.txtLastName.Value)) = 0 Or _
        Len(Me.cboGender.Value) = 0 Or _
        Len(Me.cboAgeRange.Value) = 0 Or _
        Len(Me.cboIDType.Value) = 0 Or _
        Len(Me.txtIDNumber) = 0 Then
        MsgBox prompt:="All client's Personal Information marked as * are required", Title:="Missing information - verify"
        Me.txtFirstName.SetFocus
        Exit Sub
    End If
    
    'verify required fields - contact information
    If Len(Trim(Me.txtPrimaryPhone)) = 0 Or _
        Len(Trim(Me.txtHouseAddress)) = 0 Or _
        Len(Me.txtYearsLived.Value) = 0 Then
        MsgBox prompt:="All client's contact Information marked as * are required", Title:="Missing information - verify"
        Me.txtPrimaryPhone.SetFocus
        Exit Sub
    End If
    If Len(Me.txtPrimaryPhone) <> 12 Or Left(Me.txtPrimaryPhone, 1) <> 0 Or Mid(Me.txtPrimaryPhone, 2, 1) = 0 Then
        MsgBox prompt:="Client's phone number is not correct. Please check", Title:="Incomplete information - verify"
        Me.txtPrimaryPhone.SetFocus
        Exit Sub
    End If
    
    'verify required fields - business information
    If Len(Me.cboBusinessType) = 0 Or _
        Len(Trim(Me.txtBusinessName)) = 0 Or _
        Len(Trim(Me.txtBusinessAddress)) = 0 Or _
        Len(Me.txtYearsinBusiness) = 0 Or _
        (Me.cboBusinessType.Value <> "None" And Len(Me.cboBusinessOwner) = 0) Or _
        (Me.cboBusinessOwner <> "Self Owned" And Me.cboBusinessType.Value <> "None" And Len(Me.cboCoOwnerRelationship) = 0) Or _
        (Me.cboBusinessOwner <> "Self Owned" And Me.cboBusinessType.Value <> "None" And Len(Me.txtCoOwnerName) = 0) Then
        MsgBox prompt:="All client's Business Information marked as * are required", Title:="missing information - verify"
        Me.cboBusinessType.SetFocus
        Exit Sub
    End If
    'verify required fields - business information part two
    If Me.cboBusinessType.Value <> "None" And Len(Me.cboBusinessOwner) = 0 Then
        MsgBox prompt:="This is a required field"
        Me.cboBusinessOwner.SetFocus
    End If
    
    'verify required fields - family information
    If Len(cboMaritalStatus.Value) = 0 Or _
        Len(txtFamilySize.Value) = 0 Or _
        Len(txtReferee.Value) = 0 Or _
        Len(txtRefereeContact.Value) = 0 Then
        MsgBox prompt:="All client's Family Information marked as * are required", Title:="Missing information - verify"
        Me.cboMaritalStatus.SetFocus
        Exit Sub
    End If
    If Len(Me.txtRefereeContact.Value) <> 12 Or Left(Me.txtRefereeContact.Value, 1) <> 0 Or Mid(Me.txtRefereeContact.Value, 2, 1) = 0 Then
        MsgBox prompt:="Client's phone number is not correct. Please check", Title:="Incomplete information - verify"
        Me.txtRefereeContact.Value.SetFocus
        Exit Sub
    End If
    
    'verify required fields - other information
    If Len(cboClientStatus.Value) = 0 Or _
        Len(txtRemark.Value) = 0 Then
        MsgBox prompt:="All client's other Information marked as * are required", Title:="Missing information - verify"
        Me.cboClientStatus.SetFocus
        Exit Sub
    End If
    
    'deciding on what to do whether new client or existing client
    Select Case switchAddEdit
        Case False
            'verify client personal details compared with information on id
            If MsgBox(prompt:="Does name entered match client's name on " + Me.cboIDType + "?", Title:="Verifying Client Personal Detials", Buttons:=vbYesNo) = vbNo Then
                Exit Sub
            End If
            
            Dim branchNumber As Integer
            branchNumber = 1
            
            Dim ithClient As Long
            ithClient = ThisWorkbook.Sheets("client_info_personal").Range("A" & Rows.Count).End(xlUp).Row + 1
            
            With ThisWorkbook.Sheets("client_info_personal")
                .Range("A" & ithClient).Value = client_id_age(#9/23/1988#) & Format(branchNumber, "00#") & _
                        Format((ithClient - 1), "000#")
                .Range("B" & ithClient).Value = Me.txtFirstName.Value
                .Range("C" & ithClient).Value = Me.txtMiddleName.Value
                .Range("D" & ithClient).Value = Me.txtLastName.Value
                .Range("E" & ithClient).Value = Me.cboGender
                .Range("F" & ithClient).Value = Me.cboAgeRange
                .Range("G" & ithClient).Value = Me.cboIDType
                .Range("H" & ithClient).Value = Me.txtIDNumber
                .Range("I" & ithClient).Value = Me.txtPrimaryPhone.Value
                .Range("J" & ithClient).Value = Me.cboClientStatus.Value
                .Range("K" & ithClient).Value = Format(Now(), "dd-Mmm-yyyy")
                Me.txtClientID.Value = .Range("A" & ithClient).Value
            End With
            
            With ThisWorkbook.Sheets("client_info_contact")
                .Range("A" & ithClient).Value = client_id_age(#9/23/1988#) & Format(branchNumber, "00#") & _
                        Format((ithClient - 1), "000#")
                .Range("B" & ithClient).Value = Me.txtPrimaryPhone.Value
                .Range("C" & ithClient).Value = Me.txtSecondaryPhone.Value
                .Range("D" & ithClient).Value = Me.cboSocialMedia.Value
                .Range("E" & ithClient).Value = Me.txtHouseAddress.Value
                .Range("F" & ithClient).Value = Me.txtYearsLived.Value
                .Range("G" & ithClient).Value = Me.txtPostalAddress.Value
                .Range("H" & ithClient).Value = Me.txtEmail.Value
            End With
            
            With ThisWorkbook.Sheets("client_info_business")
                .Range("A" & ithClient).Value = client_id_age(#9/23/1988#) & Format(branchNumber, "00#") & _
                        Format((ithClient - 1), "000#")
                .Range("B" & ithClient).Value = Me.cboBusinessType.Value
                .Range("C" & ithClient).Value = Me.txtBusinessName.Value
                .Range("D" & ithClient).Value = Me.txtBusinessAddress.Value
                .Range("E" & ithClient).Value = Me.txtYearsinBusiness
                .Range("F" & ithClient).Value = Me.cboBusinessOwner
                .Range("G" & ithClient).Value = Me.txtCoOwnerName
                .Range("H" & ithClient).Value = Me.cboCoOwnerRelationship
            End With
            
            With ThisWorkbook.Sheets("client_info_family")
                .Range("A" & ithClient).Value = client_id_age(#9/23/1988#) & Format(branchNumber, "00#") & _
                        Format((ithClient - 1), "000#")
                .Range("B" & ithClient).Value = Me.cboMaritalStatus.Value
                .Range("C" & ithClient).Value = Me.txtFamilySize.Value
                .Range("D" & ithClient).Value = Me.txtYoungestAge
                .Range("E" & ithClient).Value = Me.txtReferee
                .Range("F" & ithClient).Value = Me.txtRefereeContact
            End With
            
            With ThisWorkbook.Sheets("client_info_other")
                .Range("A" & ithClient).Value = client_id_age(#9/23/1988#) & Format(branchNumber, "00#") & _
                        Format((ithClient - 1), "000#")
                .Range("B" & ithClient).Value = Me.cboClientStatus.Value
                .Range("C" & ithClient).Value = Me.txtRemark.Value
                .Range("D" & ithClient).Value = Me.chkFirstTime.Value
                .Range("E" & ithClient).Value = Me.chkBelongstoGroup
                .Range("F" & ithClient).Value = Me.cboBelongstoGroup
            End With
            ThisWorkbook.Save
            
            'create folder
            Dim fPath As String
            fPath = ThisWorkbook.Path
            
            Dim fso As Scripting.FileSystemObject
            Set fso = New Scripting.FileSystemObject
            
            If Not fso.FolderExists(fPath + "\" + CStr(Me.txtClientID.Value)) Then
                fso.CreateFolder fPath + "\" + CStr(Me.txtClientID.Value)
            End If
            
            Set fso = Nothing
            
            'notify of success
            MsgBox prompt:="Client successfully added to the system. Please WRITE DOWN the CLIENT'S ID." + vbNewLine + Me.txtClientID.Value, Title:="Success Notification"
            If MsgBox(prompt:="Have you copied Client ID and notified the client about it?", Title:="Copy Client ID", Buttons:=vbInformation + vbYesNo) = vbYes Then
                Unload Me
                frmLoanMenu.Show
            Else
                cmdSave.Enabled = False
                cmdReset.Enabled = False
                Exit Sub
            End If
            
        Case Else
            'create a collection and gather the form controls as collection
            Dim newInfoCollection As Collection
            Set newInfoCollection = New Collection
            
            Dim ctrl As Control
            
            For Each ctrl In Me.Controls
                If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "CheckBox" Then
                    newInfoCollection.Add ctrl.Value 'note that .Value does not show as one of the intellisence/properties of control
                End If
            Next ctrl
            
            Dim changeInfoInt As Integer
            Dim ithCollection As Integer
            
            For ithCollection = 1 To newInfoCollection.Count
                If oldInfoCollection(ithCollection) <> newInfoCollection(ithCollection) Then
                    changeInfoInt = changeInfoInt + 1
                    Exit For
                End If
            Next ithCollection
            
            If changeInfoInt = 0 Then
                MsgBox prompt:="No Changes Detected", Title:="Changes Made", Buttons:=vbInformation
                Exit Sub
            End If
            
            'notify if you want to save changes
            If MsgBox(prompt:="There are changes made to this client's profile." + vbNewLine + _
                "Do you want to update?", Title:="Changes Notification", Buttons:=vbInformation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            
            Dim ithClientUpdate As Long
            ithClientUpdate = Application.Match(CLng(Me.txtClientID.Value), _
                ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
            
            With ThisWorkbook.Sheets("client_info_personal")
                .Range("B" & ithClientUpdate).Value = Me.txtFirstName.Value
                .Range("C" & ithClientUpdate).Value = Me.txtMiddleName.Value
                .Range("D" & ithClientUpdate).Value = Me.txtLastName.Value
                .Range("E" & ithClientUpdate).Value = Me.cboGender
                .Range("F" & ithClientUpdate).Value = Me.cboAgeRange
                .Range("G" & ithClientUpdate).Value = Me.cboIDType
                .Range("H" & ithClientUpdate).Value = Me.txtIDNumber
                .Range("I" & ithClientUpdate).Value = Me.txtPrimaryPhone.Value
                .Range("J" & ithClientUpdate).Value = Me.cboClientStatus.Value
            End With
            
            With ThisWorkbook.Sheets("client_info_contact")
                .Range("B" & ithClientUpdate).Value = Me.txtPrimaryPhone.Value
                .Range("C" & ithClientUpdate).Value = Me.txtSecondaryPhone.Value
                .Range("D" & ithClientUpdate).Value = Me.cboSocialMedia.Value
                .Range("E" & ithClientUpdate).Value = Me.txtHouseAddress.Value
                .Range("F" & ithClientUpdate).Value = Me.txtYearsLived.Value
                .Range("G" & ithClientUpdate).Value = Me.txtPostalAddress.Value
                .Range("H" & ithClientUpdate).Value = Me.txtEmail.Value
            End With
            
            With ThisWorkbook.Sheets("client_info_business")
                .Range("B" & ithClientUpdate).Value = Me.cboBusinessType.Value
                .Range("C" & ithClientUpdate).Value = Me.txtBusinessName.Value
                .Range("D" & ithClientUpdate).Value = Me.txtBusinessAddress.Value
                .Range("E" & ithClientUpdate).Value = Me.txtYearsinBusiness
                .Range("F" & ithClientUpdate).Value = Me.cboBusinessOwner
                .Range("G" & ithClientUpdate).Value = Me.txtCoOwnerName
                .Range("H" & ithClientUpdate).Value = Me.cboCoOwnerRelationship
            End With
            
            With ThisWorkbook.Sheets("client_info_family")
                .Range("B" & ithClientUpdate).Value = Me.cboMaritalStatus.Value
                .Range("C" & ithClientUpdate).Value = Me.txtFamilySize.Value
                .Range("D" & ithClientUpdate).Value = Me.txtYoungestAge
                .Range("E" & ithClientUpdate).Value = Me.txtReferee
                .Range("F" & ithClientUpdate).Value = Me.txtRefereeContact
            End With
            
            With ThisWorkbook.Sheets("client_info_other")
                .Range("B" & ithClientUpdate).Value = Me.cboClientStatus.Value
                .Range("C" & ithClientUpdate).Value = Me.txtRemark.Value
                .Range("D" & ithClientUpdate).Value = Me.chkFirstTime.Value
                .Range("E" & ithClientUpdate).Value = Me.chkBelongstoGroup
                .Range("F" & ithClientUpdate).Value = Me.cboBelongstoGroup
            End With
            ThisWorkbook.Save
            
            MsgBox prompt:="Successfully updated", Title:="Update Notification", Buttons:=vbInformation
            
            Unload Me
            frmClientProfile.Show
    End Select
End Sub

Private Sub CommandButton1_Click()
    
    Dim ctl As Control
    Dim col As Collection
    Set col = New Collection
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "CheckBox" Or TypeName(ctl) = "ComboBox" Then
            MsgBox ctl.Name & " " & ctl.TabIndex & " " & ctl.TabStop & " " & ctl.Tag
            col.Add Item:=ctl.Value, Key:=CStr(ctl.TabIndex)
        End If
    Next
End Sub

Private Sub txtFamilySize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtFirstName_AfterUpdate()
    Select Case switchAddEdit
        Case False
            Me.txtFirstName.Value = Trim(Me.txtFirstName.Value)
        Case Else
            'Dim oldFirstName As String
            Me.txtFirstName.Value = Trim(Me.txtFirstName.Value)
    End Select
End Sub

Private Sub txtLastName_AfterUpdate()
    Me.txtLastName.Value = Trim(Me.txtLastName.Value)
End Sub

Private Sub txtMiddleName_AfterUpdate()
    Me.txtMiddleName.Value = Trim(Me.txtMiddleName.Value)
End Sub

Private Sub txtPrimaryPhone_AfterUpdate()
    Me.txtPrimaryPhone.Value = Format(Me.txtPrimaryPhone.Value, "000-000-0000")
End Sub

Private Sub txtPrimaryPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefereeContact_AfterUpdate()
    Me.txtRefereeContact.Value = Format(Me.txtRefereeContact.Value, "000-000-0000")
End Sub

Private Sub txtRefereeContact_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
    
    'If KeyAscii > 47 And KeyAscii < 58 Then
    '    KeyAscii = KeyAscii
    'Else
    '    KeyAscii = 0
    'End If
End Sub

Private Sub txtSecondaryPhone_AfterUpdate()
    Me.txtSecondaryPhone.Value = Format(Me.txtSecondaryPhone.Value, "000-000-0000")
End Sub

Private Sub txtSecondaryPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtYearsinBusiness_AfterUpdate()
    On Error Resume Next
    If Len(Me.txtYearsinBusiness.Value) = 0 Or (Len(Me.txtYearsinBusiness.Value) - Len(Replace(Me.txtYearsinBusiness.Value, ".", ""))) > 1 _
        Or CLng(Me.txtYearsinBusiness.Value) = 0 Then
        Me.txtYearsinBusiness.Value = vbNullString
        Me.txtYearsinBusiness.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtYearsinBusiness_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtYearsLived_AfterUpdate()
    On Error Resume Next
    If Len(Me.txtYearsLived.Value) = 0 Or (Len(Me.txtYearsLived.Value) - Len(Replace(Me.txtYearsLived.Value, ".", ""))) > 1 _
        Or CLng(Me.txtYearsLived.Value) = 0 Then
        Me.txtYearsLived.Value = vbNullString
        Me.txtYearsLived.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtYearsLived_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtYoungestAge_AfterUpdate()
    On Error Resume Next
    If Len(Me.txtYoungestAge.Value) = 0 Or (Len(Me.txtYoungestAge.Value) - Len(Replace(Me.txtYoungestAge.Value, ".", ""))) > 1 _
        Or CLng(Me.txtYoungestAge.Value) = 0 Then
        Me.txtYoungestAge.Value = vbNullString
        Me.txtYoungestAge.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtYoungestAge_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
           
            With Me.cboGender
                .AddItem "Male"
                .AddItem "Female"
            End With
            
            With Me.cboAgeRange
                .AddItem "21 - 30"
                .AddItem "31 - 40"
                .AddItem "41 - 50"
                .AddItem "51 - 60"
            End With
            
            With Me.cboIDType
                .AddItem "Drivers License"
                .AddItem "NHIS"
                .AddItem "Passport"
                .AddItem "Voters ID"
            End With
            
            With Me.cboSocialMedia
                .AddItem "Facebook"
                .AddItem "WhatsApp"
                .AddItem "Twitter"
                .AddItem "Instagram"
                .AddItem "Viber"
            End With
        
            With Me.cboMaritalStatus
                .AddItem "Single"
                .AddItem "Married"
                .AddItem "Divorce"
                .AddItem "Separated"
            End With
    
            With Me.cboBusinessType
                .AddItem "None"
                .AddItem "Hireway Trader"
                .AddItem "Provision Store"
                .AddItem "Educational"
                .AddItem "Other"
            End With
            
            With Me.cboBusinessOwner
                .AddItem "Self Owned"
                .AddItem "Co-Owned"
            End With
    
            With Me.cboCoOwnerRelationship
                .AddItem "Child"
                .AddItem "Extended Relative"
                .AddItem "Friend"
                .AddItem "Sibling"
                .AddItem "Spouse"
                .AddItem "Parent"
                .AddItem "Other"
            End With
        
            With Me
                .lblLevel.Caption = UCase(frmUserLogOn.cboLevel.Value)
                .lblUserID.Caption = UCase(frmUserLogOn.txtUserID.Value)
            End With
            
    Select Case switchAddEdit
        Case False
            With Me.cboClientStatus
                .AddItem "Active"
                .AddItem "Trial"
            End With
        
        Case Else
            Me.cmdSave.Caption = "Update"
            
            With Me.cboClientStatus
                .AddItem "Active"
                .AddItem "Delinquent"
                .AddItem "Trial"
                .AddItem "Misbehave"
            End With
            
            Dim ithClientID As Long
            ithClientID = Application.Match(CLng(frmClientInfoLogIn.txtClientID.Value), _
                ThisWorkbook.Sheets("client_info_personal").Range("A:A"), 0)
            
            ThisWorkbook.Sheets("dirty_client_info").Cells.ClearContents
            ThisWorkbook.Sheets("client_info_personal").Range("A" & ithClientID).EntireRow.Copy Destination:=ThisWorkbook.Sheets("dirty_client_info").Range("A1")
            
            With Me
                With ThisWorkbook.Sheets("client_info_personal")
                    txtClientID.Value = .Range("A" & ithClientID).Value
                    txtFirstName.Value = .Range("B" & ithClientID).Value
                    txtMiddleName.Value = .Range("C" & ithClientID).Value
                    txtLastName.Value = .Range("D" & ithClientID).Value
                    cboGender.Value = .Range("E" & ithClientID).Value
                    cboAgeRange.Value = .Range("F" & ithClientID).Value
                    cboIDType.Value = .Range("G" & ithClientID).Value
                    txtIDNumber.Value = .Range("H" & ithClientID).Value
                End With
            End With
            
            With Me
                With ThisWorkbook.Sheets("client_info_contact")
                    txtPrimaryPhone.Value = .Range("B" & ithClientID).Value
                    txtSecondaryPhone.Value = .Range("C" & ithClientID).Value
                    cboSocialMedia.Value = .Range("D" & ithClientID).Value
                    txtHouseAddress.Value = .Range("E" & ithClientID).Value
                    txtYearsLived.Value = .Range("F" & ithClientID).Value
                    txtPostalAddress.Value = .Range("G" & ithClientID).Value
                    txtEmail.Value = .Range("H" & ithClientID).Value
                End With
            End With
            
            With Me
                With ThisWorkbook.Sheets("client_info_business")
                    cboBusinessType.Value = .Range("B" & ithClientID).Value
                    txtBusinessName.Value = .Range("C" & ithClientID).Value
                    txtBusinessAddress.Value = .Range("D" & ithClientID).Value
                    txtYearsinBusiness.Value = .Range("E" & ithClientID).Value
                    cboBusinessOwner.Value = .Range("F" & ithClientID).Value
                    txtCoOwnerName.Value = .Range("G" & ithClientID).Value
                    cboCoOwnerRelationship.Value = .Range("H" & ithClientID).Value
                End With
            End With
        
            With Me
                With ThisWorkbook.Sheets("client_info_family")
                    cboMaritalStatus.Value = .Range("B" & ithClientID).Value
                    txtFamilySize.Value = .Range("C" & ithClientID).Value
                    txtYoungestAge.Value = .Range("D" & ithClientID).Value
                    txtReferee.Value = .Range("E" & ithClientID).Value
                    txtRefereeContact.Value = .Range("F" & ithClientID).Value
                End With
            End With
            
            With Me
                With ThisWorkbook.Sheets("client_info_other")
                    cboClientStatus.Value = .Range("B" & ithClientID).Value
                    txtRemark.Value = .Range("C" & ithClientID).Value
                    chkFirstTime.Value = .Range("D" & ithClientID).Value
                    chkBelongstoGroup.Value = .Range("E" & ithClientID).Value
                    cboBelongstoGroup.Value = .Range("F" & ithClientID).Value
                End With
            End With
    End Select
    
    Set oldInfoCollection = New Collection
    Dim ctlr As Control
            
    For Each ctlr In Me.Controls
        If TypeName(ctlr) = "TextBox" Or TypeName(ctlr) = "ComboBox" Or TypeName(ctlr) = "CheckBox" Then
            oldInfoCollection.Add ctlr.Value
        End If
    Next ctlr
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub
