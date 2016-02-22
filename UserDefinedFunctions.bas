Attribute VB_Name = "UserDefinedFunctions"
Option Compare Database

Function Mobilize(Optional OArray As Variant = Empty, Optional EArray As Variant = Empty)
    Dim strONumber As String
    Dim strENumber As String
    Dim intMobilized As Integer
    Dim intNotMobilized As Integer
    intNotMobilized = 0
    intMobilized = 0
        
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False
    
    ' If OArray is non-empty...
    If Not IsEmpty(OArray) Then
        For Each ONumber In OArray
            strONumber = Right(Trim(CStr(ONumber)), 5)
            ' If ONumber is valid...
            If ResourceExists("O", strONumber) Then
                ' And if the resource is not already mobilized
                If Not IsMobilized("O", strONumber) Then
                    ' Update Mobilized field, add Staging Actions
                    DoCmd.RunSQL ("UPDATE Personnel SET Mobilized = True WHERE Personnel.ONumber = 'O" + strONumber + "'")
                    DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Mobilized', 'O" + strONumber + "')")
                    DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('In Staging', 'O" + strONumber + "')")
                    intMobilized = intMobilized + 1
                ' But if the resource has already been mobilized...
                Else
                    msgAlreadyMobilized = MsgBox("Resource number O-" + strONumber + " is already mobilized.", vbInformation, "Resource Mobilization")
                    intNotMobilized = intNotMobilized + 1
                End If
            ' If ONumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number O-" + strONumber + " does not exist in the database. You should probably note this number and verify it later.", vbInformation, "Resource Entry Error")
                intNotMobilized = intNotMobilized + 1
            End If
        Next
        
        msgConfirmation = MsgBox("Of the PERSONNEL you entered, " + CStr(intMobilized) + " mobilized, and " + CStr(intNotMobilized) + " failed to mobilize.")
    End If
    
    ' If EArray is non-empty...
    If Not IsEmpty(EArray) Then
        For Each ENumber In EArray
            strENumber = Right(Trim(CStr(ENumber)), 5)
            ' If ENumber is valid...
            If ResourceExists("E", strENumber) Then
                ' And if the resource is not already mobilized...
                If Not IsMobilized("E", strENumber) Then
                    ' Update Mobilized field, add Staging Actions
                    DoCmd.RunSQL ("UPDATE Equipment SET Mobilized = True WHERE Equipment.ENumber = 'E" + strENumber + "'")
                    DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Mobilized', 'E" + strENumber + "')")
                    DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('In Staging', 'E" + strENumber + "')")
                    intMobilized = intMobilized + 1
                ' But if the resource has already been mobilized...
                Else
                    msgAlreadyMobilized = MsgBox("Resource number E-" + strENumber + " is already mobilized.", vbInformation, "Resource Mobilization")
                    intNotMobilized = intNotMobilized + 1
                End If
            ' If ENumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number E-" + strENumber + " does not exist in the database. You should probably note this number and verify it later.", vbInformation, "Resource Entry Error")
                intNotMobilized = intNotMobilized + 1
            End If
        Next
        
        msgConfirmation = MsgBox("Of the EQUIPMENT you entered, " + CStr(intMobilized) + " mobilized, and " + CStr(intNotMobilized) + " failed to mobilize.")
    End If
    DoCmd.SetWarnings True
    ' ---End SQL operations---
    
End Function

Function ResourceCreated(ResourceType As String, ResourceNumber As String)
   Dim strRNumber As String
   Dim strRType As String
   strRType = ResourceType
   strRNumber = Right(Trim(ResourceNumber), 5)
   
   ' ---Begin SQL operations
    DoCmd.SetWarnings False

    ' If RNumber is valid, add Staging Action
    If strRType = "O" Then
        DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit) VALUES ('Unit Created', 'O" + strRNumber + "')")
    End If
    If strRType = "E" Then
        DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit) VALUES ('Unit Created', 'E" + strRNumber + "')")
    End If
    If strRType = "C" Then
        DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit) VALUES ('Unit Created', 'C" + strRNumber + "')")
        DoCmd.RunSQL ("UPDATE Units SET Active = True WHERE CNumber = 'C" + strRNumber + "'")
    End If

    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function AssignToUnit(CNumber As String, Optional OArray As Variant = Empty, Optional EArray As Variant = Empty)
    Dim strCNumber As String
    Dim strONumber As String
    Dim strENumber As String
    strCNumber = Right(Trim(CNumber), 5)
    
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False
    
    ' Loop through each non-empty array
    ' If OArray is non-empty:
    If Not IsEmpty(OArray) Then
        For Each ONumber In OArray
            strONumber = Right(Trim(CStr(ONumber)), 5)
            ' If ONumber is valid, update CurrentUnit field, add Staging Action
            If ResourceExists("O", strONumber) Then
                DoCmd.RunSQL ("UPDATE Personnel SET Personnel.CurrentUnit = 'C" + strCNumber + "' WHERE Personnel.ONumber = 'O" + strONumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource, Unit) VALUES ('Deployed', 'O" + strONumber + "', 'C" + strCNumber + "')")
            ' If ONumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number O-" + strONumber + " does not exist in the database. You should probably note this number and verify it later.", vbInformation, "Resource Entry Error")
            End If
        Next
    End If
    If Not IsEmpty(EArray) Then
        For Each ENumber In EArray
            strENumber = Right(Trim(CStr(ENumber)), 5)
            ' If ENumber is valid, update CurrentUnit field, add Staging Action
            If ResourceExists("E", Trim(CStr(ENumber))) Then
                DoCmd.RunSQL ("UPDATE Equipment SET Equipment.CurrentUnit = 'C" + strCNumber + "' WHERE Equipment.ENumber = 'E" + strENumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource, Unit) VALUES ('Deployed', 'E" + strENumber + "', 'C" + strCNumber + "')")
            ' If ENumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number E-" + strENumber + " does not exist in the database. You should probably note this number and verify it later.", vbInformation, "Resource Entry Error")
            End If
        Next
    End If
    
     DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function ReassignUnit(CNumber As String, NewLocation As String)
   Dim strCNumber As String
   Dim strLocation As String
   strCNumber = Right(Trim(CNumber), 5)
   strLocation = NewLocation
    
    ' ---Begin SQL operations
    DoCmd.SetWarnings False

    ' If CNumber is valid, update Location field, add Staging Actions
    If ResourceExists("C", strCNumber) Then
        DoCmd.RunSQL ("UPDATE Units SET Units.Location = '" + strLocation + "' WHERE Units.CNumber = 'C" + strCNumber + "'")
        DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit, Location) VALUES ('Unit Relocated', 'C" + strCNumber + "', '" + strLocation + "')")
        DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit, Location) VALUES ('Deployed', 'C" + strCNumber + "', '" + strLocation + "')")
        ' TO DO: IMPLEMENT STAGING ACTION FOR ALL INCLUDED PERSONNEL AND EQUIPMENT AS WELL
        msgConfirmReassign = MsgBox("The unit's location has been updated.", vbInformation, "Unit Reassignment")
    ' If CNumber is NOT valid, display message, take no action on database
    Else
        msgNoResource = MsgBox("Resource number C-" + strCNumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
    End If

    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function CheckIn(Optional OArray As Variant = Empty, Optional EArray As Variant = Empty)
    Dim strONumber As String
    Dim strENumber As String
    
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False
    
    ' Loop through each non-empty array
    ' If OArray is non-empty:
    If Not IsEmpty(OArray) Then
        For Each ONumber In OArray
            strONumber = Right(Trim(CStr(ONumber)), 5)
            ' If ONumber is valid, add Staging Action
            If ResourceExists("O", strONumber) Then
                DoCmd.RunSQL ("UPDATE Personnel SET CurrentUnit = null WHERE Personnel.ONumber = 'O" + strONumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Checked In', 'O" + strONumber + "')")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('In Staging', 'O" + strONumber + "')")
            ' If ONumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number O-" + strONumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    If Not IsEmpty(EArray) Then
        For Each ENumber In EArray
            strENumber = Right(Trim(CStr(ENumber)), 5)
            ' If ENumber is valid, add Staging Action
            If ResourceExists("E", strENumber) Then
                DoCmd.RunSQL ("UPDATE Equipment SET CurrentUnit = null WHERE Equipment.ENumber = 'E" + strENumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Checked In', 'E" + strENumber + "')")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('In Staging', 'E" + strENumber + "')")
            ' If ENumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number E-" + strENumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    
    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function DissolveUnit(CNumber As String, Optional InjuryDamage As String = vbNullString)
    Dim strCNumber As String
    strCNumber = Right(CNumber, 5)
    
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False

    ' Update pertinent fields in designated Unit. Create Staging Action.
    DoCmd.RunSQL ("UPDATE Units SET Active = False WHERE CNumber = 'C" + strCNumber + "'")
    DoCmd.RunSQL ("UPDATE Units SET InjuryDamage = '" + InjuryDamage + "' WHERE CNumber = 'C" + strCNumber + "'")
    DoCmd.RunSQL ("Update Units SET Location = '' WHERE CNumber = 'C" + strCNumber + "'")
    DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Unit) VALUES ('Unit Dissolved', 'C" + strCNumber + "')")

    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function Demobilize(Optional OArray As Variant = Empty, Optional EArray As Variant = Empty)
    Dim strONumber As String
    Dim strENumber As String
    
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False
    
    ' Loop through each non-empty array
    ' If OArray is non-empty:
    If Not IsEmpty(OArray) Then
        For Each ONumber In OArray
            strONumber = Right(Trim(CStr(ONumber)), 5)
            ' If ONumber is valid, update mobilized field, add Staging Action
            If ResourceExists("O", strONumber) Then
                DoCmd.RunSQL ("UPDATE Personnel SET Personnel.Mobilized = False WHERE Personnel.ONumber = 'O" + strONumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Demobilized', 'O" + strONumber + "')")
            'If ONumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number O-" + strONumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    ' If EArray is non-empty:
    If Not IsEmpty(EArray) Then
        For Each ENumber In EArray
            strENumber = Right(Trim(CStr(ENumber)), 5)
            ' If ENumber is valid, update Mobilized field, add Staging Action
            If ResourceExists("E", strENumber) Then
                DoCmd.RunSQL ("UPDATE Equipment SET Equipment.Mobilized = False WHERE Equipment.ENumber = 'E" + strENumber + "'")
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Demobilized', 'E" + strENumber + "')")
            ' If ENumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number E-" + strENumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function MarkHomeArrival(Optional OArray As Variant = Empty, Optional EArray As Variant = Empty)
    Dim strONumber As String
    Dim strENumber As String
    
    ' ---Begin SQL operations---
    DoCmd.SetWarnings False
    
    ' Loop through each non-empty array
    ' If OArray is non-empty:
    If Not IsEmpty(OArray) Then
        For Each ONumber In OArray
            strONumber = Right(Trim(CStr(ONumber)), 5)
            ' If ONumber is valid, add Staging Action
            If ResourceExists("O", strONumber) Then
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Arrived Home', 'O" + strONumber + "')")
            ' If ONumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number O-" + strONumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    If Not IsEmpty(EArray) Then
        For Each ENumber In EArray
            strENumber = Right(Trim(CStr(ENumber)), 5)
            ' If ENumber is valid, add Staging Action
            If ResourceExists("E", strENumber) Then
                DoCmd.RunSQL ("INSERT INTO StagingActions (ActionType, Resource) VALUES ('Arrived Home', 'E" + strENumber + "')")
            ' If ENumber is NOT valid, display message, take no action on database
            Else
                msgNoResource = MsgBox("Resource number E-" + strENumber + " does not exist in the database. You might want to double check that.", vbInformation, "Unit Entry Error")
            End If
        Next
    End If
    
    DoCmd.SetWarnings True
    ' ---End SQL operations---
End Function

Function ResourceExists(ResourceType As String, ResourceNumber As String) As Boolean
    Dim strRType As String
    Dim strRNumber As String
    strRType = ResourceType
    strRNumber = Right(ResourceNumber, 5)

    ' Check table corresponding to resource type for value corresponding to ResourceNumber
    If strRType = "O" Then
        If Not IsNull(DLookup("[ONumber]", "Personnel", "ONumber = 'O" & strRNumber & "'")) Then
            ResourceExists = True
            Exit Function
        End If
    ElseIf strRType = "E" Then
        If Not IsNull(DLookup("[ENumber]", "Equipment", "ENumber = 'E" & strRNumber & "'")) Then
            ResourceExists = True
            Exit Function
        End If
    ElseIf strRType = "C" Then
        If Not IsNull(DLookup("[CNumber]", "Units", "[CNumber] = 'C" & strRNumber & "'")) Then
            ResourceExists = True
            Exit Function
        End If
    Else
        ResourceExists = False
    End If
End Function

Function GetUnitResources(CNumber As String) As String()
    Dim c As String
    c = Right(CNumber, 5)
    
    If ResourceExists("C", c) Then
        Dim strResults() As String
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim i As Integer
        i = 0
        
        Set db = CurrentDb
        
        ' Query Personnel table for ONumbers in Current Unit
        Set rs = db.OpenRecordset("SELECT ONumber FROM Personnel WHERE CurrentUnit = 'C" & c & "'")
        ' If the query did NOT return an empty recordset...
        If Not rs.BOF Then
            rs.MoveFirst
            Do Until rs.EOF
                ReDim Preserve strResults(i)
                strResults(i) = CStr(rs.Fields(0))
                rs.MoveNext
                i = i + 1
            Loop
        End If
        rs.Close
                
        ' Query Equipment table for ENumbers in Current Unit
        Set rs = db.OpenRecordset("SELECT ENumber FROM Equipment WHERE CurrentUnit = 'C" & c & "'")
        ' If the query did NOT return an empty recordset...
        If Not rs.BOF Then
            rs.MoveFirst
            Do Until rs.EOF
                ReDim Preserve strResults(i)
                strResults(i) = CStr(rs.Fields(0))
                rs.MoveNext
                i = i + 1
            Loop
        End If
        
        ' Close and trash recordset and database objects
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
        
        GetUnitResources = strResults
    Else
        msgNoUnit = MsgBox("That unit does not exist.", vbInformation, "Get Unit Resources")
    End If
End Function

Function IsMobilized(ResourceType As String, ResourceNumber As String) As Boolean
    Dim strRType As String
    Dim strRNumber As String
    strRType = ResourceType
    strRNumber = Right(ResourceNumber, 5)

    ' Check table corresponding to resource type for value corresponding to ResourceNumber
    If strRType = "O" Then
        If DLookup("[Mobilized]", "Personnel", "ONumber = 'O" & strRNumber & "'") Then
            IsMobilized = True
            Exit Function
        End If
    ElseIf strRType = "E" Then
        If DLookup("[Mobilized]", "Equipment", "ENumber = 'E" & strRNumber & "'") Then
            IsMobilized = True
            Exit Function
        End If
    Else
        IsMobilized = False
    End If
End Function
