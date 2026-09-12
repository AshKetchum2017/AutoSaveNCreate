Option Explicit

' Class Module (Name) = SNCSettingsStore.
Private Const REG_APP As String = "RinCorelMacros"
Private Const REG_SECTION As String = "AutoSaveNCreate"

' Satu nilai registry untuk seluruh daftar. Tab memisahkan nama/pola;
' LF memisahkan row. Karakter tersebut ditolak sebelum penyimpanan.
Public Function LoadUsers() As Collection
    Dim users As Collection
    Dim serialized As String
    Dim rows As Variant
    Dim fields As Variant
    Dim i As Long
    Set users = New Collection
    serialized = GetSetting(REG_APP, REG_SECTION, "UserSettings", "")
    If Len(serialized) > 0 Then
        rows = Split(serialized, vbLf)
        For i = 0 To UBound(rows)
            fields = Split(rows(i), vbTab)
            If UBound(fields) <> 1 Then Err.Raise 5, "SNCSettingsStore", "Data setting user di registry tidak valid pada row " & CStr(i + 1) & "."
            ValidateUser fields(0), fields(1)
            EnsureUniqueName users, CStr(fields(0))
            users.Add Array(CStr(fields(0)), CStr(fields(1)))
        Next i
    End If
    Set LoadUsers = users
End Function

Public Sub SaveUsers(ByVal users As Collection)
    Dim serialized As String
    Dim i As Long
    Dim item As Variant
    Dim checkedUsers As Collection
    Set checkedUsers = New Collection
    For i = 1 To users.Count
        item = users(i)
        ValidateUser CStr(item(0)), CStr(item(1))
        EnsureUniqueName checkedUsers, CStr(item(0))
        checkedUsers.Add item
        If i > 1 Then serialized = serialized & vbLf
        serialized = serialized & CStr(item(0)) & vbTab & CStr(item(1))
    Next i
    SaveSetting REG_APP, REG_SECTION, "UserSettings", serialized
End Sub

Public Sub ValidateUser(ByVal userName As String, ByVal tokenText As String)
    Dim parser As SNCFolderParser
    Dim i As Long
    Dim code As Long
    If Len(Trim$(userName)) = 0 Then Err.Raise 5, "SNCSettingsStore", "Nama operator belum diisi."
    For i = 1 To Len(userName)
        code = AscW(Mid$(userName, i, 1))
        If code >= 0 And code < 32 Then Err.Raise 5, "SNCSettingsStore", "Nama operator tidak boleh mengandung tab atau baris baru."
    Next i
    Set parser = New SNCFolderParser
    parser.ValidateTemplate tokenText
End Sub

Public Sub EnsureUniqueName(ByVal users As Collection, ByVal userName As String, Optional ByVal exceptIndex As Long = 0)
    Dim i As Long
    Dim item As Variant
    For i = 1 To users.Count
        If i <> exceptIndex Then
            item = users(i)
            If StrComp(Trim$(CStr(item(0))), Trim$(userName), vbTextCompare) = 0 Then
                Err.Raise 5, "SNCSettingsStore", "Nama operator sudah ada: " & userName
            End If
        End If
    Next i
End Sub

Public Function LoadDirectory(ByVal modeName As String) As String
    Dim fso As Object
    Dim savedPath As String
    savedPath = GetSetting(REG_APP, REG_SECTION, DirectoryKey(modeName), "")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(savedPath) Then LoadDirectory = savedPath
End Function

Public Sub SaveDirectory(ByVal modeName As String, ByVal folderPath As String)
    Dim fso As Object
    Dim key As String
    key = DirectoryKey(modeName)
    folderPath = Trim$(folderPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Err.Raise 5, "SNCSettingsStore", "Pilih directory tujuan yang sudah ada dan valid."
    SaveSetting REG_APP, REG_SECTION, key, folderPath
End Sub

Private Function DirectoryKey(ByVal modeName As String) As String
    Select Case modeName
        Case "DieA", "HiDie", "KissA", "HiKiss"
            DirectoryKey = "Directory" & modeName
        Case Else
            Err.Raise 5, "SNCSettingsStore", "Mode directory tidak dikenal: " & modeName
    End Select
End Function
