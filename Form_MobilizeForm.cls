VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MobilizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnMobilize_Click()
    Dim strONumbers() As String
    Dim strENumbers() As String
    Dim strONumberEntry As String
    Dim strENumberEntry As String
    
    ' If form is blank, display error
    If IsNull(Form_MobilizeForm.txtONumberEntry.Value) And IsNull(Form_MobilizeForm.txtENumberEntry.Value) Then
        msgNoResource = MsgBox("You did not enter any resources. Keep trying!", vbInformation, "Resource Mobilization")
    ' If form is not blank...
    Else
        ' If Assign checkbox is checked...
        If Form_MobilizeForm.chkAssign.Value = True Then
            'Open DeployCreateNewUnitForm
            DoCmd.OpenForm "DeployCreateNewUnitForm"
            'Pass entry to DeployCreateNewUnitForm
            strONumberEntry = Trim(CStr(Form_MobilizeForm.txtONumberEntry.Value))
            strENumberEntry = Trim(CStr(Form_MobilizeForm.txtENumberEntry.Value))
            Form_DeployCreateNewUnitForm.txtONumberEntry.Value = strONumberEntry
            Form_DeployCreateNewUnitForm.txtENumberEntry.Value = strENumberEntry
            'Close MobilizeForm
            DoCmd.Close acForm, "MobilizeForm"
            
            msgPassResources = MsgBox("The resources have NOT YET been mobilized. They will be mobilized when you assign them to a unit in this next step.", vbExclamation, "Resource Mobilization")
        'If Assign checkbox is not checked:
        Else
            ' If Personnel are present...
            If Not IsNull(Form_MobilizeForm.txtONumberEntry.Value) Then
                strONumberEntry = Trim(CStr(Form_MobilizeForm.txtONumberEntry.Value))
                strONumbers = Split(strONumberEntry, vbCrLf)
                Call Mobilize(strONumbers)
            End If
            ' If Equipment is present...
            If Not IsNull(Form_MobilizeForm.txtENumberEntry.Value) Then
                strENumberEntry = Trim(CStr(Form_MobilizeForm.txtENumberEntry.Value))
                strENumbers = Split(strENumberEntry, vbCrLf)
                Call Mobilize(, strENumbers)
            End If

        ' Open StagingAreaMainMenu
        DoCmd.OpenForm "StagingAreaMainMenu"
        ' Close MobilizeForm
        DoCmd.Close acForm, "MobilizeForm"
        End If
    End If
End Sub
