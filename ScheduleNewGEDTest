## VBA | MODAL FORM FOR SCHEDULING A NEW GED TEST
```
Option Compare Database

Private Sub GEDTestEntryCancelBtn_Click()
    DoCmd.Close
End Sub

Private Sub TestScheduleEntryOK_Click()
    ' Check that all required fields are completed before submitting
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            'Debug.Print (ctl.name)
            If ctl = "" Or IsNull(ctl) Then
                'Debug.Print ("I am empty.")
                If ctl.name <> "CommentsEntry" Then
                    response = MsgBox("Please complete all required fields before submitting.", 1, "Missing Fields")
                    Exit Sub
                End If
            End If
        End If
    Next

    ' Enter reported information into GEDTestT
    ' Retrieve client ID
    Dim searchID As Variant
    searchID = DLookup("[ClientID]", "NameSearchQ", "[WholeName]=Forms![GEDTestScheduleEntryF]![ClientNameEntry]")

    ' If record does not exist, insert new record; otherwise, update existing record
    Dim existingRec As Variant
    existingRec = DLookup("[GEDTest_ID]", "GEDTestT", "[ClientID] = " & searchID & _
                  " AND [GEDTestDay] = #" & TestDateEntry & "# AND [GEDTest] = '" & TestSubjectEntry & "'")

    If IsNull(existingRec) Then
        CurrentDb.Execute "INSERT INTO GEDTestT(ClientID, GEDTestDay, " & _
            "GEDTestTime, GEDTest, GEDLocation, Transportation, Comments)" & " VALUES('" & searchID & "','" & _
             TestDateEntry & "','" & TestTimeEntry & "','" & TestSubjectEntry & "','" & _
             TestLocationEntry & "','" & TransportationEntry & "','" & CommentsEntry & "');"

    Else
        strSQL = "UPDATE GEDTestT " & _
            "SET [GEDTestTime] = #" & TestTimeEntry & "#, [Transportation] = '" & TransportationEntry & _
                "', [GEDLocation] = '" & TestLocationEntry & "', [Comments] = '" & CommentsEntry & "' " & _
            "WHERE ([ClientID] = " & searchID & " AND [GEDTestDay] = #" & TestDateEntry & _
                "# AND [GEDTest] = '" & TestSubjectEntry & "');"
        Debug.Print (strSQL)
        CurrentDb.Execute strSQL
    End If
    DoCmd.Close
End Sub
```
