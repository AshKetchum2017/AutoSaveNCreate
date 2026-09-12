Option Explicit

Private pUsers As Collection

Private Sub UserForm_Initialize()
    Dim settings As SNCSettingsStore
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo LoadFailed
    Set settings = New SNCSettingsStore
    Set pUsers = settings.LoadUsers()
    RefreshList
    Exit Sub
LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "UserSettingsMenu.Initialize", errorDescription
End Sub

Private Sub cmdAdd_Click()
    EditUser 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    If lbxUserLists.ListIndex < 0 Then Exit Sub
    EditUser lbxUserLists.ListIndex + 1
End Sub

Private Sub cmdRemove_Click()
    If lbxUserLists.ListIndex < 0 Then Exit Sub
    pUsers.Remove lbxUserLists.ListIndex + 1
    RefreshList
End Sub

Private Sub cmdSave_Click()
    Dim settings As SNCSettingsStore
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveUsers pUsers
    Unload Me
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan daftar user (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub lbxUserLists_Click()
    cmdModify.Enabled = (lbxUserLists.ListIndex >= 0)
    cmdRemove.Enabled = (lbxUserLists.ListIndex >= 0)
End Sub

Private Sub RefreshList(Optional ByVal selectedIndex As Long = 0)
    Dim i As Long
    Dim item As Variant
    lbxUserLists.Clear
    lbxUserLists.ColumnCount = 2
    For i = 1 To pUsers.Count
        item = pUsers(i)
        lbxUserLists.AddItem CStr(item(0))
        lbxUserLists.List(i - 1, 1) = CStr(item(1))
    Next i
    If selectedIndex > 0 And selectedIndex <= pUsers.Count Then lbxUserLists.ListIndex = selectedIndex - 1
    lbxUserLists_Click
End Sub

Private Sub EditUser(ByVal itemIndex As Long)
    Dim editor As Object
    Dim settings As SNCSettingsStore
    Dim item As Variant
    Dim newUsers As Collection
    Dim i As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim operation As String
    On Error GoTo EditFailed
    Set settings = New SNCSettingsStore
    operation = "Membuka AddUserSettings (termasuk txbUserName)"
    Set editor = UserForms.Add("AddUserSettings")
    If itemIndex > 0 Then
        item = pUsers(itemIndex)
        editor.BeginEdit CStr(item(0)), CStr(item(1))
    Else
        editor.BeginEdit vbNullString, vbNullString
    End If
    Do
        operation = "Menampilkan AddUserSettings"
        editor.Show vbModal
        If Not editor.Accepted Then Exit Do
        operation = "Memeriksa nama operator"
        If NameAvailable(editor.ResultName, itemIndex) Then
            item = Array(CStr(editor.ResultName), CStr(editor.ResultTokens))
            If itemIndex = 0 Then
                pUsers.Add item
                itemIndex = pUsers.Count
            Else
                Set newUsers = New Collection
                For i = 1 To pUsers.Count
                    If i = itemIndex Then
                        newUsers.Add item
                    Else
                        newUsers.Add pUsers(i)
                    End If
                Next i
                Set pUsers = newUsers
            End If
            RefreshList itemIndex
            Exit Do
        End If
    Loop
    Unload editor
    Exit Sub
EditFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not editor Is Nothing Then Unload editor
    On Error GoTo 0
    MsgBox operation & vbCrLf & "Error " & CStr(errorNumber) & ": " & errorDescription, vbExclamation, "AutoSaveNCreate"
End Sub

Private Function NameAvailable(ByVal userName As String, ByVal exceptIndex As Long) As Boolean
    Dim settings As SNCSettingsStore
    On Error GoTo NameFailed
    Set settings = New SNCSettingsStore
    settings.EnsureUniqueName pUsers, userName, exceptIndex
    NameAvailable = True
    Exit Function
NameFailed:
    MsgBox Err.Description, vbExclamation, "AutoSaveNCreate"
End Function
