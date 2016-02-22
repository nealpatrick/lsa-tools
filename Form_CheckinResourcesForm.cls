VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CheckinResourcesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' TO DO: Figure out what to do about MarkArrival.

Private Sub btnCheckInResources_Click()
    Dim strONumbers() As String
    Dim strENumbers() As String
    Dim strONumberEntry As String
    Dim strENumberEntry As String
        
    ' If form is blank, display error
    If IsNull(Form_CheckinResourcesForm.txtONumberEntry.Value) And IsNull(Form_CheckinResourcesForm.txtENumberEntry.Value) Then
        msgNoResource = MsgBox("You did not enter any resources. Keep trying!", vbInformation, "Resource Check-In")
    ' If form is not blank:
    Else
        ' If Personnel are present
        If Not IsNull(Form_CheckinResourcesForm.txtONumberEntry.Value) Then
            strONumberEntry = Trim(CStr(Form_CheckinResourcesForm.txtONumberEntry.Value))
            strONumbers = Split(strONumberEntry, vbCrLf)
            Call CheckIn(strONumbers)
        End If
        ' If Equipment is present
        If Not IsNull(Form_CheckinResourcesForm.txtENumberEntry.Value) Then
            strENumberEntry = Trim(CStr(Form_CheckinResourcesForm.txtENumberEntry.Value))
            strENumbers = Split(strENumberEntry, vbCrLf)
            Call CheckIn(, strENumbers)
        End If
        msgConfirmation = MsgBox("The resources have been checked in to staging.", vbInformation, "Resource Check-in")
        
        'Open StagingAreaMainMenu
        DoCmd.OpenForm "StagingAreaMainMenu"
        'Close MobilizeForm
        DoCmd.Close acForm, "CheckinResourcesForm"
    End If
End Sub
