Option Explicit

' Tambahkan TextBox (Name) = txbUserName pada UserForm existing.
Private pAccepted As Boolean

Public Sub BeginEdit(ByVal userName As String, ByVal tokenText As String)
    pAccepted = False
    txbUserName.Text = userName
    txbUserTokens.Text = tokenText
End Sub

Public Property Get Accepted() As Boolean
    Accepted = pAccepted
End Property

Public Property Get ResultName() As String
    ResultName = Trim$(txbUserName.Text)
End Property

Public Property Get ResultTokens() As String
    ResultTokens = txbUserTokens.Text
End Property

Private Sub cmdCancel_Click()
    pAccepted = False
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    Dim settings As SNCSettingsStore
    On Error GoTo SaveFailed
    Set settings = New SNCSettingsStore
    settings.ValidateUser ResultName, ResultTokens
    pAccepted = True
    Me.Hide
    Exit Sub
SaveFailed:
    MsgBox "Setting user belum valid (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "AutoSaveNCreate"
End Sub

Private Sub cmdTokenHint_Click()
    Dim hints As Object
    On Error GoTo HintFailed
    Set hints = UserForms.Add("UserTokenHints")
    hints.Show vbModeless
    Exit Sub
HintFailed:
    MsgBox "Token: / = kategori folder, ~ = nama folder sumber lowercase." & vbCrLf & _
        "Gunakan tanda kutip untuk teks literal; spasi dan - mengikuti pola." & vbCrLf & _
        "Contoh: " & Chr$(34) & "r" & Chr$(34) & " - / ~" & vbCrLf & _
        "Token __ belum digunakan.", vbInformation, "AutoSaveNCreate - Token"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub

Private Sub txbUserTokens_Change()

End Sub
