## VBA | Test Score Entry Form
### SETTINGS AND HOME BUTTON
```
Option Compare Database

Private Sub HomeBtn_Click()
    DoCmd.Close
    DoCmd.OpenForm ("MainScreen")
End Sub
```
### CONDITIONAL DATA ENTRY
```
Private Sub GED_Took_Test_Change()
    If Me.GED_Took_Test = "Yes" Then
        Me.GED_Test_Score.Enabled = True
    Else
        Me.GED_Test_Score.Enabled = False
    End If
End Sub
```
### "CLEAR FIELDS" CLICK EVENTS
```
Private Sub TABE_L_ClearFields_Click()
    ClearData Me.ScoreEntrySheets
End Sub

Private Sub TABE_Clear_Fields_Click()
    ClearData Me.ScoreEntrySheets
End Sub

Private Sub GED_P_ClearFields_Click()
    ClearData Me.ScoreEntrySheets
End Sub

Private Sub GED_Clear_Fields_Click()
    ClearData Me.ScoreEntrySheets
End Sub

Private Sub OMJ_WorkKeys_Clear_Fields_Click()
    ClearData Me.ScoreEntrySheets
End Sub

Private Sub ACT_WorkKeys_Clear_Fields_Click()
    ClearData Me.ScoreEntrySheets
End Sub
```
### "CLEAR FIELDS" OPERATION
```
Public Sub ClearData(tabct As TabControl)
    Dim ct As Control
    Dim i As Integer
    i = 0
    For Each ct In tabct
        If TypeOf ct Is TextBox Or TypeOf ct Is CheckBox Or TypeOf ct Is ComboBox Then
            Debug.Print (i & " " & ct.name)
            ct = Null
            i = i + 1
        End If
    Next
End Sub
```
### "SAVE ENTRY" BUTTONS
**Official GED Practice Test**
```
Private Sub GED_P_Save_Entry_Click()
    Dim other_info As Variant
    otherMissingFields = check_AllControls(Me.ScoreEntrySheets.Pages("Official GED Practice Tests"))

    If Len(Join(otherMissingFields)) > 0 Then
        response = MsgBox("Please complete all fields before submitting.", 1, "Missing Fields")
        Exit Sub
    Else:
        Dim searchID As Variant
        searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![GED_P_StudentName]")
        Dim existingRec As Variant
        existingRec = DLookup("[GEDPracticeID]", "GEDPracticeT", "[ClientID] = " & searchID & _
                      " AND [GEDPracticeTest] = '" & GED_P_TestSelection & "' AND [GEDPracticeTestDay] = #" & _
                      GED_P_TestDate & "# AND [GEDPracticeScore] = " & GED_P_TestScore)

        ' If a record already exists, prompt user
        If Not IsNull(existingRec) Then
            response = MsgBox("Error: This record already exists!", 1, "Missing Fields")
            Exit Sub
        End If

        CurrentDb.Execute "INSERT INTO GEDPracticeT(ClientID, GEDPracticeTest, " & _
            "GEDPracticeTestDay, GEDPracticeScore)" & " VALUES('" & searchID & "','" & _
            GED_P_TestSelection & "','" & GED_P_TestDate & "'," & GED_P_TestScore & ");"

        MsgBox ("New official GED practice test scores added!")
        ClearData Me.ScoreEntrySheets
    End If
End Sub
```
**TABE Survey**
```
Private Sub SaveTABE_Click()
    ' All fields must be complete to submit TABE Survey data
    Dim otherMissingFields As Variant
    otherMissingFields = check_AllControls(Me.ScoreEntrySheets.Pages("TABE Survey"))
    If Len(Join(otherMissingFields)) > 0 Then
        response = MsgBox("Please complete all fields before submitting.", 1, "Missing Fields")
        Exit Sub
    End If

    ' If all fields are complete, insert new record into TabeT using data
    Dim searchID As Variant
    searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![TABE_StudentEntry]")

    CurrentDb.Execute "INSERT INTO TabeT(ClientID, TestDate," & _
        "TABEForm, ReadingTestLevel, ReadingSS, ReadingGE," & _
        "LanguageTestLevel, LanguageSS, LanguageGE, MathTestLevel," & _
        "CompMathSS, CompMathGE, AppliedMathSS, AppliedMathGE," & _
        "OverallMathSS, OverallMathGE)" & _
        " VALUES ('" & searchID & "','" & Me.TABEDateEntry & "','" & _
        Me.TABE_Form & "','" & Me.R_BookLevel & "'," & Me.R_SS & "," & _
        Me.R_GLE & ",'" & Me.L_BookLevel & "'," & Me.L_SS & "," & _
        Me.L_GLE & ",'" & Me.M_BookLevel & "'," & Me.CM_SS & "," & _
        Me.CM_GLE & "," & Me.AM_SS & "," & Me.AM_GLE & "," & _
        Me.CombMath_SS & "," & Me.CombM_GLE & ");"

    MsgBox ("New TABE scores added!")
    ClearData Me.ScoreEntrySheets
End Sub
```
**TABE Locator**
```
Private Sub TABE_L_Save_Entries_Click()
    Dim other_info As Variant
    otherMissingFields = check_AllControls(Me.ScoreEntrySheets.Pages("TABE Locator"))

    If Len(Join(otherMissingFields)) > 0 Then
        response = MsgBox("Please complete all fields before submitting.", 1, "Missing Fields")
        Exit Sub
    Else:
        Dim searchID As Variant
        searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![TABE_L_Student_Name]")
        CurrentDb.Execute "INSERT INTO TABELocatorT(ClientID, LocatorTestDay," & _
            "ReadingNC, ReadingBookLevel,LanguageArtsNC, LanguageArtsBookLevel," & _
            "AppliedMathNC, ComputationalMathNC, CombinedMathNC, MathBookLevel)" & _
            " VALUES ('" & searchID & "','" & Me.TABE_L_Test_Date & "'," & _
            Me.TABE_L_ReadingNC & ",'" & Me.TABE_L_ReadingLevel & "'," & _
            Me.TABE_L_LArtsNC & ",'" & Me.TABE_L_LArtsLevel & "'," & _
            Me.TABE_L_AppliedMath_NC & "," & Me.TABE_L_CompMath_NC & "," & _
            Me.TABE_L_CombinedM_NC & ",'" & Me.TABE_Math_BookLevel & "');"

        MsgBox ("New TABE Locator scores added!")
        ClearData Me.ScoreEntrySheets
    End If
End Sub
```
**Official GED Test**
```
Private Sub GED_Save_Entry_Click()
    Dim missingFields As Variant
    missingFields = check_AllControls(Me.ScoreEntrySheets.Pages("Official GED Tests"))
    Dim numMissingFields As Integer
    numMissingFields = Len(Join(missingFields))
    Dim response As Integer

    ' Ensure that the user has completed all required fields
    If numMissingFields > 0 Then
        If IsNull(Me.GED_Student_Name) Or IsNull(Me.GED_Test) Or IsNull(Me.GED_Test_Date) Then
            response = MsgBox("Please complete at minimum the 'Student Name,' 'Scheduled Test,' and 'Test Date' fields.", 0, "Missing Fields")
            Exit Sub
        End If
    End If

    ' Proceed with submission if all required fields are complete
    ' Retrieve client ID
    Dim searchID As Variant
    searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![GED_Student_Name]")

    ' If record does not exist, insert new record; otherwise, update existing record
    Dim existingRec As Variant
    existingRec = DLookup("[GEDTest_ID]", "GEDTestT", "[ClientID] = " & searchID & _
                  " AND [GEDTestDay] = #" & GED_Test_Date & "# AND [GEDTest] = '" & GED_Test & "'")

    Dim strSQL As String
    If IsNull(GED_Test_Score) Then
    score = 0
    End If

    If IsNull(existingRec) Then
           strSQL = "INSERT INTO GEDTestT(ClientID, GEDTestDay, GEDTest, [TookTest?], Score, Comments)" & _
                    " VALUES('" & searchID & "', '" & GED_Test_Date & "', '" & GED_Test & "', '" & GED_Took_Test & _
                    "', '" & score & "', '" & GED_Comments & "');"

    Else
        strSQL = "UPDATE GEDTestT " & _
                 "SET [TookTest?] = " & GED_Took_Test & ", [Score] = '" & score & "', [Comments] = '" & GED_Comments & "' " & _
                 "WHERE ([ClientID] = " & searchID & " AND [GEDTestDay] = #" & GED_Test_Date & "# AND [GEDTest] = '" & GED_Test & "');"
        Debug.Print (strSQL)

    End If

    CurrentDb.Execute strSQL

    ' Inform user that entries were added and then reset tab
    MsgBox ("New official GED test scores added!")
    ClearData Me.ScoreEntrySheets
    Me.GED_Test_Score.Enabled = False
End Sub
```
**Ohio Means Jobs (OMJ) WorkKeys Practice Test**
```
Private Sub OMJ_WorkKeys_Save_Entry_Click()
    Dim other_info As Variant
    otherMissingFields = check_AllControls(Me.ScoreEntrySheets.Pages("OhioMeansJobs WorkKeys Practice"))
    If Len(Join(otherMissingFields)) > 0 Then
        response = MsgBox("Please complete all fields before submitting.", 1, "Missing Fields")
        Exit Sub
    Else:
        Dim searchID As Variant
        searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![OMJ_WorkKeys_Student_Name]")
        CurrentDb.Execute "INSERT INTO OMJWorkKeysT(ClientID, TestDate, PracticeTest, TestScore, EstWorkKeysLevel)" & _
            " VALUES('" & searchID & "','" & OMJ_WorkKeys_Test_Date & "','" & OMJ_WorkKeys_Practice_Test & _
            "'," & OMJ_WorkKeys_Test_Score & ",'" & OMJ_WorkKeys_Level & "');"

        MsgBox ("New OhioMeansJobs practice WorkKeys scores added!")
        ClearData Me.ScoreEntrySheets
    End If
End Sub
```
**Official ACT WorkKeys Test**
```
Private Sub ACT_WorkKeys_Save_Entry_Click()
    Dim other_info As Variant
    otherMissingFields = check_AllControls(Me.ScoreEntrySheets.Pages("Official ACT WorkKeys"))
    If Len(Join(otherMissingFields)) > 0 Then
        response = MsgBox("Please complete all fields before submitting.", 1, "Missing Fields")
        Exit Sub
    Else:
        Dim searchID As Variant
        searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![ScoreEntry]![ACT_WorkKeys_Student_Name]")
        CurrentDb.Execute "INSERT INTO ACTWorkKeysT(ClientID, TestDate, WorkKeysTest, ScaleScore, WorkKeysLevel)" & _
            " VALUES('" & searchID & "','" & ACT_WorkKeys_Test_Date & "','" & ACT_WorkKeys_Test & _
            "'," & ACT_WorkKeys_Scale_Score & ",'" & ACT_WorkKeys_Level & "');"

        MsgBox ("New ACT WorkKeys scores added!")
        ClearData Me.ScoreEntrySheets
    End If
End Sub
```
**Check for Missing Fields Before Form Submission**
```
Public Function checkTABE_NeededValues() As Boolean
    If IsNull(Me.TABE_StudentEntry.Value) Or IsNull(Me.TABEDateEntry.Value) Or IsNull(Me.TABE_Form.Value) Then
        checkTABE_NeededValues = False
    ElseIf Me.TABE_StudentEntry.Value = "" Or Me.TABEDateEntry.Value = "" Or Me.TABE_Form.Value = "" Then
        checkTABE_NeededValues = False
    Else:
        checkTABE_NeededValues = True
    End If
    End Function

    Public Function check_AllControls(tabPage As Page) As Variant
    Debug.Print ("You are in this page: " & tabPage.name)
    Dim missing_Entries() As String
    ReDim missing_Entries(16)
    Dim i As Integer
    i = 0
    For Each ct In tabPage.Controls
        If TypeOf ct Is TextBox Or TypeOf ct Is CheckBox Or TypeOf ct Is ComboBox Then
            If IsNull(ct) Or ct = "" Then
                'Debug.Print (ct.name & " " & "is null.")
                missing_Entries(i) = ct.Controls.Item(0).name
                'Debug.Print (missing_Entries(i))
                i = i + 1
            End If
        End If
    Next
    ReDim Preserve missing_Entries(i)
    check_AllControls = missing_Entries
End Function
```
