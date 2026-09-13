Option Explicit

' Code-behind UserForm (Name) = DirectorySettings.
Private pModeName As String
Private pSourceDocument As Document
Private pSession As SNCDirectorySession
Private Const FILE_DIALOG_SAVE_AS As Long = 1
Private Const SELECT_FOLDER_DUMMY_FILE As String = "Select this folder"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_USENEWUI As Long = &H50

Public Sub BeginEdit(ByVal modeName As String, Optional ByVal sourceDocument As Document = Nothing)
    Dim settings As SNCSettingsStore
    Set settings = New SNCSettingsStore
    txbDirectory.Text = settings.LoadDirectory(modeName)
    pModeName = modeName
    Set pSourceDocument = sourceDocument
    Set pSession = GetSNCDirectorySession()
    Me.Caption = "Root Directory - " & modeName
    RefreshDisposableList 0
End Sub

Private Sub cmdAdd_Click()
    Dim settings As SNCSettingsStore
    Dim parser As SNCFolderParser
    Dim detector As SNCModeDetector
    Dim users As Collection
    Dim fso As Object
    Dim folderPath As String
    Dim fileName As String
    Dim modeName As String
    Dim displayText As String
    Dim conflict As Boolean
    Dim index As Long
    On Error GoTo AddFailed
    folderPath = pSession.NormalizeDirectory(txbDirectory.Text)
    If Not pSourceDocument Is Nothing Then fileName = pSourceDocument.FileName
    Set detector = New SNCModeDetector
    modeName = detector.DetectDirectoryMode(folderPath, fileName, conflict)
    If conflict Then
        If MsgBox("Petunjuk kategori directory/nama CDR bertentangan." & vbCrLf & _
            folderPath & vbCrLf & "CDR: " & fileName & vbCrLf & vbCrLf & _
            "OK = lanjut menambahkan path saja tanpa kategori." & vbCrLf & _
            "Cancel = batalkan penambahan.", vbOKCancel Or vbExclamation Or vbDefaultButton2, "AutoSaveNCreate") <> vbOK Then Exit Sub
    End If
    Set settings = New SNCSettingsStore
    displayText = folderPath
    If Len(modeName) > 0 Then
        Set parser = New SNCFolderParser
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set users = settings.LoadUsers()
        displayText = parser.DisposableParentName(fso.GetFileName(folderPath), users) & _
            " | " & detector.ModeCaption(modeName) & " | " & folderPath
    End If
    index = pSession.AddDirectory(folderPath, displayText)
    txbDirectory.Text = settings.LoadDirectory(pModeName)
    RefreshDisposableList index
    Exit Sub
AddFailed:
    MsgBox "Gagal menambah directory sementara (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdBrowse_Click()
    Dim selectedPath As String
    On Error GoTo BrowseFailed
    selectedPath = SelectFolder(Trim$(txbDirectory.Text))
    If Len(selectedPath) > 0 Then txbDirectory.Text = selectedPath
    Exit Sub
BrowseFailed:
    MsgBox "Gagal memilih directory (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    On Error GoTo ClearFailed
    pSession.ClearAll
    RefreshDisposableList 0
    Exit Sub
ClearFailed:
    MsgBox "Gagal mengosongkan daftar (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdRemove_Click()
    Dim index As Long
    On Error GoTo RemoveFailed
    If lbxDisposableLists.ListIndex < 0 Then Exit Sub
    index = lbxDisposableLists.ListIndex + 1
    pSession.RemoveDirectory index
    If index > pSession.Count Then index = pSession.Count
    RefreshDisposableList index
    Exit Sub
RemoveFailed:
    MsgBox "Gagal menghapus item directory (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdSave_Click()
    Dim settings As SNCSettingsStore
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveDirectory pModeName, txbDirectory.Text
    pSession.ClearSelection
    Unload Me
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan directory (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo SelectFailed
    If lbxDisposableLists.ListIndex < 0 Then Exit Sub
    pSession.SelectDirectory pSourceDocument, lbxDisposableLists.ListIndex + 1
    Unload Me
    Exit Sub
SelectFailed:
    MsgBox "Gagal memilih directory sementara (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub lbxDisposableLists_Click()
    ' Memilih row tidak mengganti textbox root permanen.
    cmdSelect.Enabled = (lbxDisposableLists.ListIndex >= 0)
    cmdRemove.Enabled = (lbxDisposableLists.ListIndex >= 0)
End Sub

Private Sub RefreshDisposableList(ByVal selectedIndex As Long)
    Dim i As Long
    lbxDisposableLists.Clear
    lbxDisposableLists.ColumnCount = 1
    lbxDisposableLists.MultiSelect = fmMultiSelectSingle
    For i = 1 To pSession.Count
        lbxDisposableLists.AddItem pSession.DisplayAt(i)
    Next i
    If selectedIndex > 0 And selectedIndex <= pSession.Count Then lbxDisposableLists.ListIndex = selectedIndex - 1
    cmdClear.Enabled = (pSession.Count > 0)
    lbxDisposableLists_Click
End Sub

Private Sub txbDirectory_Change()
    ' Draft: cmdAdd menambah daftar sesi; cmdSave menyimpan root permanen.
End Sub

' Mengikuti pola ExportRelatedSettings: dialog Save As dummy, lalu fallback Shell.
Private Function SelectFolder(ByVal initialPath As String) As String
    Dim fso As Object
    Dim selectedPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(initialPath) Then initialPath = vbNullString
    If TrySelectFolderWithFileDialog(initialPath, selectedPath) Then
        SelectFolder = selectedPath
        Exit Function
    End If
    SelectFolder = SelectFolderWithBrowseForFolder(initialPath)
End Function

Private Function TrySelectFolderWithFileDialog(ByVal initialPath As String, ByRef selectedFolder As String) As Boolean
    Dim cst As Object
    Dim selectedItem As String
    On Error Resume Next
    Set cst = Application.CorelScriptTools
    If Err.Number <> 0 Or cst Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    selectedItem = cst.GetFileBox("All Files (*.*)|*.*", "Pilih Folder Tujuan", _
        FILE_DIALOG_SAVE_AS, SELECT_FOLDER_DUMMY_FILE, vbNullString, initialPath, "Select Folder")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    selectedFolder = ResolveFolderFromDialogPath(selectedItem)
    ' Cancel juga handled: jangan membuka fallback setelah user membatalkan.
    TrySelectFolderWithFileDialog = True
End Function

Private Function SelectFolderWithBrowseForFolder(ByVal initialPath As String) As String
    Dim shellApp As Object
    Dim folder As Object
    Dim rootFolder As Variant
    Dim fso As Object
    rootFolder = initialPath
    If Len(initialPath) = 0 Then rootFolder = 0
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, "Pilih Folder Tujuan", BIF_RETURNONLYFSDIRS Or BIF_USENEWUI, rootFolder)
    If folder Is Nothing Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folder.Self.Path) Then SelectFolderWithBrowseForFolder = folder.Self.Path
End Function

Private Function ResolveFolderFromDialogPath(ByVal selectedItem As String) As String
    Dim fso As Object
    Dim folderPath As String
    selectedItem = Trim$(selectedItem)
    If Len(selectedItem) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(selectedItem) Then
        ResolveFolderFromDialogPath = selectedItem
    Else
        folderPath = fso.GetParentFolderName(selectedItem)
        If fso.FolderExists(folderPath) Then ResolveFolderFromDialogPath = folderPath
    End If
End Function
