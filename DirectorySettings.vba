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
    If Not sourceDocument Is Nothing Then
        If Len(pSession.GetDirectory(sourceDocument, modeName)) > 0 Then
            txbDirectory.Text = pSession.GetDirectory(sourceDocument, modeName)
            Me.Caption = "Directory Sekali Pakai - " & modeName
        End If
    End If
End Sub

Private Sub cmdDisposableSave_Click()
    On Error GoTo SaveFailed
    If pSession Is Nothing Then Err.Raise 5, "DirectorySettings", "Konteks directory belum tersedia."
    pSession.SetDirectory pSourceDocument, pModeName, txbDirectory.Text
    Unload Me
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan directory sekali pakai (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
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

Private Sub cmdSave_Click()
    Dim settings As SNCSettingsStore
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.SaveDirectory pModeName, txbDirectory.Text
    Unload Me
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan directory (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub txbDirectory_Change()
    ' Perubahan tetap berupa draft sampai cmdSave ditekan.
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
