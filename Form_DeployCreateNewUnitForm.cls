VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DeployCreateNewUnitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUnitEntryDone_Click()
    ' Declare variables and assign user entry to variables
    Dim strONumbers() As String
    Dim strENumbers() As String
    Dim strONumberEntry As String
    Dim strENumberEntry As String
    Dim strCNumberEntry As String
    Dim strCrewLeaderEntry As String
    
    ' If the form's C Number field or Crew Leader field is empty...
    If IsNull(Form_DeployCreateNewUnitForm.CNumber.Value) Or IsNull(Form_DeployCreateNewUnitForm.CrewLeader.Value) Then
        msgRequiredMissing = MsgBox("You have not entered either a unit number or crew leader. Keep trying!" & vbCrLf & "(If you are trying to reach the Main Menu, click the Home button.)", vbInformation, "Resource Deployment")
    ' If the Crew Leader field is incorrect
    Else
        ' Clean up fields and proceed
        strCNumberEntry = Right(Trim(Form_DeployCreateNewUnitForm.CNumber.Value), 5)
        Form_DeployCreateNewUnitForm.CNumber.Value = "C" + strCNumberEntry
        strCrewLeaderEntry = Right(Trim(Form_DeployCreateNewUnitForm.CrewLeader.Value), 5)
        Form_DeployCreateNewUnitForm.CrewLeader.Value = "O" + strCrewLeaderEntry
        ' If the crew leader does not exist in the database...
        If Not ResourceExists("O", strCrewLeaderEntry) Then
            msgCrewLeaderError = MsgBox("There is a problem with the Crew Leader field. You must enter a valid crew leader.", vbCritical, "Resource Deployment")
        ' If the crew leader exists in the database...
        Else
            ' If no resources are entered, deploy unit and take no further actions
            If IsNull(Form_DeployCreateNewUnitForm.txtONumberEntry.Value) And IsNull(Form_DeployCreateNewUnitForm.txtENumberEntry.Value) Then
                ' Save user entry for fields that are linked to table
                If Me.Dirty Then Me.Dirty = False
                Call ResourceCreated("C", strCNumberEntry)
                NoResourceMessage = MsgBox("The unit has been deployed. You have assigned no resources to this unit. Resources should be assigned later.", vbInformation, "Resource Deployment")
            ' If resources are entered, try to mobilize them, then assign them to this unit.
            Else
                ' Save user entry for fields that are linked to table
                If Me.Dirty Then Me.Dirty = False
                Call ResourceCreated("C", strCNumberEntry)
                ' If personnel numbers have been entered...
                If Form_DeployCreateNewUnitForm.txtONumberEntry.Value <> vbNullString Then
                    strONumberEntry = Trim(Form_DeployCreateNewUnitForm.txtONumberEntry.Value)
                    strONumbers = Split(strONumberEntry, vbCrLf)
                    Call Mobilize(strONumbers)
                    Call AssignToUnit(strCNumberEntry, strONumbers)
                End If
                ' If equipment numbers have been entered, mobilize and assign those units
                If Form_DeployCreateNewUnitForm.txtENumberEntry.Value <> vbNullString Then
                    strENumberEntry = Trim(Form_DeployCreateNewUnitForm.txtENumberEntry.Value)
                    strENumbers = Split(strENumberEntry, vbCrLf)
                    Call Mobilize(, strENumbers)
                    Call AssignToUnit(strCNumberEntry, , strENumbers)
                End If
                assignmentConfirmation = MsgBox("The unit has been deployed and the resources have been assigned.", vbInformation, "Resource Deployment")
            End If
            ' Open Main Menu and close This form.
            DoCmd.OpenForm "StagingAreaMainMenu"
            DoCmd.Close acForm, "DeployCreateNewUnitForm"
        End If
    End If
End Sub

Private Sub btnUnitEntryAnother_Click()
    ' Declare variables and assign user entry to variables
    Dim strONumbers() As String
    Dim strENumbers() As String
    Dim strONumberEntry As String
    Dim strENumberEntry As String
    Dim strCNumberEntry As String
    Dim strCrewLeaderEntry As String
    
    ' If the form's C Number field or Crew Leader field is empty...
    If IsNull(Form_DeployCreateNewUnitForm.CNumber.Value) Or IsNull(Form_DeployCreateNewUnitForm.CrewLeader.Value) Then
        msgRequiredMissing = MsgBox("You have not entered either a unit number or crew leader. Keep trying!" & vbCrLf & "(If you are trying to reach the Main Menu, click the Home button.)", vbInformation, "Resource Deployment")
    ' If the Crew Leader field is incorrect
    Else
        ' Clean up fields and proceed
        strCNumberEntry = Right(Trim(Form_DeployCreateNewUnitForm.CNumber.Value), 5)
        Form_DeployCreateNewUnitForm.CNumber.Value = "C" + strCNumberEntry
        strCrewLeaderEntry = Right(Trim(Form_DeployCreateNewUnitForm.CrewLeader.Value), 5)
        Form_DeployCreateNewUnitForm.CrewLeader.Value = "O" + strCrewLeaderEntry
        ' If the crew leader does not exist in the database...
        If Not ResourceExists("O", strCrewLeaderEntry) Then
            msgCrewLeaderError = MsgBox("There is a problem with the Crew Leader field. You must enter a valid crew leader.", vbCritical, "Resource Deployment")
        ' If the crew leader exists in the database...
        Else
            ' If no resources are entered, deploy unit and take no further actions
            If IsNull(Form_DeployCreateNewUnitForm.txtONumberEntry.Value) And IsNull(Form_DeployCreateNewUnitForm.txtENumberEntry.Value) Then
                ' Save user entry for fields that are linked to table
                If Me.Dirty Then Me.Dirty = False
                Call ResourceCreated("C", strCNumberEntry)
                NoResourceMessage = MsgBox("The unit has been deployed. You have assigned no resources to this unit. Resources should be assigned later.", vbInformation, "Resource Deployment")
            ' If resources are entered, try to mobilize them, then assign them to this unit.
            Else
                ' Save user entry for fields that are linked to table
                If Me.Dirty Then Me.Dirty = False
                Call ResourceCreated("C", strCNumberEntry)
                ' If personnel numbers have been entered...
                If Form_DeployCreateNewUnitForm.txtONumberEntry.Value <> vbNullString Then
                    strONumberEntry = Trim(Form_DeployCreateNewUnitForm.txtONumberEntry.Value)
                    strONumbers = Split(strONumberEntry, vbCrLf)
                    Call Mobilize(strONumbers)
                    Call AssignToUnit(strCNumberEntry, strONumbers)
                End If
                ' If equipment numbers have been entered, mobilize and assign those units
                If Form_DeployCreateNewUnitForm.txtENumberEntry.Value <> vbNullString Then
                    strENumberEntry = Trim(Form_DeployCreateNewUnitForm.txtENumberEntry.Value)
                    strENumbers = Split(strENumberEntry, vbCrLf)
                    Call Mobilize(, strENumbers)
                    Call AssignToUnit(strCNumberEntry, , strENumbers)
                End If
                assignmentConfirmation = MsgBox("The unit has been deployed and the resources have been assigned.", vbInformation, "Resource Deployment")
            End If
            ' Close and reopen the form
            DoCmd.Close acForm, "DeployCreateNewUnitForm"
            DoCmd.OpenForm "DeployCreateNewUnitForm"
        End If
    End If
End Sub
