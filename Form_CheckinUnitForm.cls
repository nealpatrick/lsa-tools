VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CheckinUnitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUnitSearch_Click()
    ' Assign User Entry to TempVars
    TempVars!CNumberSearch = Right(Trim(Form_CheckinUnitForm.txtCNumberEntry.Value), 5)
    
    ' Check if user entry exists in database. If it does, make linked report and fields visible. Otherwise, display an error.
    If ResourceExists("C", TempVars!CNumberSearch) Then
        DoCmd.Requery "SingleUnitByCNumber"
        Form_CheckinUnitForm.SingleUnitByCNumber.Visible = True
        Form_CheckinUnitForm.lblInstruction.Visible = True
        Form_CheckinUnitForm.chkDemobilize.Visible = True
        Form_CheckinUnitForm.lblDemobilize.Visible = True
        Form_CheckinUnitForm.btnSubmitCheckin.Visible = True
        Form_CheckinUnitForm.txtInjuryDamage.Visible = True
        Form_CheckinUnitForm.lblInjuryDamage.Visible = True
    Else
        msgNoUnit = MsgBox("That unit does not exist.", vbExclamation, "Incorrect Entry")
    End If
End Sub

Private Sub btnSubmitCheckin_Click()
    Dim OArray()
    Dim EArray()
    Dim strCNumber As String
    Dim strInjuryDamage As String
    Dim unitResources() As String
    strCNumber = Right(Trim(Form_CheckinUnitForm.txtCNumberEntry.Value), 5)
    
    unitResources = GetUnitResources(strCNumber)
    
    ' Assign Injury/Damage string
    If Not IsNull(Trim(Form_CheckinUnitForm.txtInjuryDamage.Value)) Then
        strInjuryDamage = Form_CheckinUnitForm.txtInjuryDamage.Value
    Else
        strInjuryDamage = "None"
    End If
    
    ' If the unit contains resources
    If Len(Join(unitResources)) > 0 Then
        Dim intOCount As Integer, intECount As Integer
        intOCount = 0
        intECount = 0
        
        ' Separate unitResources into EArray and OArray
        For Each resource In unitResources
            If Left(resource, 1) = "O" Then
                ReDim Preserve OArray(intOCount)
                OArray(intOCount) = resource
                intOCount = intOCount + 1
            ElseIf Left(resource, 1) = "E" Then
                ReDim Preserve EArray(intECount)
                EArray(intECount) = resource
                intECount = intECount + 1
            Else
                msgResourceError = MsgBox("There is a problem with resource #" & resource & ".", vbExclamation, "Resource Error")
            End If
        Next
        
        ' Checkin Unit Resources
        If intOCount > 0 Then Call CheckIn(OArray)
        If intECount > 0 Then Call CheckIn(, EArray)
        
        ' Dissolve unit
        Call DissolveUnit(strCNumber, strInjuryDamage)
        
        ' Demobilize each resource if requested by user
        If Form_CheckinUnitForm.chkDemobilize.Value = True Then
            Call Demobilize(OArray, EArray)
            msgConfirm = MsgBox("The resources have been checked in and demobilized. The unit has been dissolved.", vbInformation, "Check-In")
        Else
            msgConfirm = MsgBox("The resources have been checked in to staging. The unit has been dissolved.", vbInformation, "Check-In")
        End If
    Else
        ' Dissolve unit
        Call DissolveUnit(strCNumber, strInjuryDamage)
        ' If there are no resources, display message, take no action
        msgNoResource = MsgBox("There are no resources assigned to the specified unit. The unit has been dissolved.", vbInformation, "Get Unit Resources")
    End If
    
    'Close the form and open the main menu
    DoCmd.Close acForm, "CheckinUnitForm"
    DoCmd.OpenForm "StagingAreaMainMenu"
End Sub

Private Sub Form_Close()
    TempVars.RemoveAll
End Sub
