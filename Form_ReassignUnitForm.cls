VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReassignUnitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub btnUnitSearch_Click()
    ' Assign User Entry to TempVars
    TempVars!CNumberSearch = Right(Trim(Form_ReassignUnitForm.txtCNumberEntry.Value), 5)
    
    ' Check if unit exists in database. If it does, make linked report and fields visible. Otherwise, display an error.
    If ResourceExists("C", TempVars!CNumberSearch) Then
        DoCmd.Requery "SingleUnitByCNumber"
        Form_ReassignUnitForm.SingleUnitByCNumber.Visible = True
        Form_ReassignUnitForm.lblInstruction.Visible = True
        Form_ReassignUnitForm.lblNewLocationEntry.Visible = True
        Form_ReassignUnitForm.txtNewLocationEntry.Visible = True
        Form_ReassignUnitForm.btnSubmitReassign.Visible = True
    Else
        ErrorMessage = MsgBox("That unit does not exist.", vbExclamation, "Incorrect Entry")
    End If
End Sub

Private Sub btnSubmitReassign_Click()
    Dim strCNumber As String
    Dim strNewLocation As String
    strCNumber = Right(Trim(Form_ReassignUnitForm.txtCNumberEntry.Value), 5)
    strNewLocation = Trim(Form_ReassignUnitForm.txtNewLocationEntry.Value)
    
    Call ReassignUnit(strCNumber, strNewLocation)
    
    ' Open StagingAreaMainMenu
    DoCmd.OpenForm "StagingAreaMainMenu"
    ' Close ReassignUnitForm
    DoCmd.Close acForm, "ReassignUnitForm"
End Sub

Private Sub Form_Close()
    TempVars.RemoveAll
End Sub
