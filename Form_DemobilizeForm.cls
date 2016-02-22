VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DemobilizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Figure out what to do about MarkArrival.

Private Sub btnDemobilizeResources_Click()
    Dim strONumbers() As String
    Dim strENumbers() As String
    Dim strONumberEntry As String
    Dim strENumberEntry As String
    ' TO DO: If most recent StagingAction is NOT In Staging, make sure to Check the resource in.
    
    ' If form is blank, display error
    If IsNull(Form_DemobilizeForm.txtONumberEntry.Value) And IsNull(Form_DemobilizeForm.txtENumberEntry.Value) Then
        msgNoResource = MsgBox("You did not enter any resources. Keep trying!", vbInformation, "Resource Demobilization")
    ' If form is not blank:
    Else
        ' If Personnel are present
        If Not IsNull(Form_DemobilizeForm.txtONumberEntry.Value) Then
            strONumberEntry = Trim(CStr(Form_DemobilizeForm.txtONumberEntry.Value))
            strONumbers = Split(strONumberEntry, vbCrLf)
            Call Demobilize(strONumbers)
        End If
        ' If Equipment is present
        If Not IsNull(Form_DemobilizeForm.txtENumberEntry.Value) Then
            strENumberEntry = Trim(CStr(Form_DemobilizeForm.txtENumberEntry.Value))
            strENumbers = Split(strENumberEntry, vbCrLf)
            Call Demobilize(, strENumbers)
        End If
        msgConfirmation = MsgBox("The resources have been demobilized.", vbInformation, "Resource Demobilization")
        
        'Open StagingAreaMainMenu
        DoCmd.OpenForm "StagingAreaMainMenu"
        'Close MobilizeForm
        DoCmd.Close acForm, "DemobilizeForm"
    End If
    
End Sub

Private Sub btnDemobilizeUnit_Click()
    'CNumberEntry =
    'Get resources from specified unit, then:
    Dim OArray() As String
    Dim EArray() As String
    EString = Form_DemobilizeForm.txtENumberEntry.Value
    OString = Form_DemobilizeForm.txtONumberEntry.Value
    
    'If most recent StagingAction is NOT In Staging, make sure to Check the resource in.
    
    If IsNull(OString) And IsNull(EString) Then
        NoResourceMessage = MsgBox("You did not enter any resources.", vbInformation, "Resource Demobilization")
    Else
        If Not IsNull(OString) Then
            OArray = Split(OString, vbCrLf)
            Call Demobilize(OArray)
        End If
        If Not IsNull(EString) Then
            EArray = Split(EString, vbCrLf)
            Call Demobilize(, EArray)
        End If
        demobilizationConfirmation = MsgBox("The resources have been demobilized.", vbInformation, "Resource Demobilization")
    End If
    
    'Open StagingAreaMainMenu
    DoCmd.OpenForm "StagingAreaMainMenu"
    'Close MobilizeForm
    DoCmd.Close acForm, "DemobilizeForm"
End Sub
