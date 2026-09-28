Option Explicit

' Class Module (Name) = SNCBehaviorSession. One AutoSNC instance per queue step.
Private pMenu As Object
Private pDefaults As Object

Public Sub Start(ByVal script As String, ByVal observer As Object, ByVal token As String)
    Dim contract As New SNCBehaviorContract, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, existing As Object, value As Variant
    Dim operation As String, number As Long, source As String, description As String
    Dim closed As Boolean

    On Error GoTo Failed
    operation = "Memvalidasi kontrak AutoSNC"
    Set block = contract.Validate(script)

    operation = "Memeriksa AutoSNC yang sudah terbuka"
    For Each existing In VBA.UserForms
        If StrComp(TypeName(existing), "AutoSNC", vbTextCompare) = 0 Then _
            Err.Raise 5, "SNCBehaviorSession", "Tutup AutoSNC yang sudah terbuka sebelum menjalankan behavior."
    Next existing
    Set existing = Nothing

    operation = "Membuka AutoSNC"
    Set pMenu = VBA.UserForms.Add("AutoSNC")
    pMenu.MRBindRunner observer, token
    pMenu.MRSetBehaviorMode True
    pMenu.Show vbModeless

    operation = "Membaca nilai awal kontrol AutoSNC"
    SnapshotDefaults block

    For Each ins In block.Instructions
        operation = ins.FormName & " / " & ins.Target & " (posisi " & CStr(ins.Position) & ")"
        Select Case ins.Kind
            Case "stage", "skip"
                ' Omitted and skipped controls keep the form's initialized values.
            Case "action"
                Select Case LCase$(ins.Target)
                    Case "cmdprocess": pMenu.MRBehaviorProcess
                    Case "cmdclose"
                        pMenu.MRBehaviorClose
                        closed = True
                End Select
            Case Else
                value = ins.Value
                If ins.Kind = "default" Then value = pDefaults(LCase$(ins.Target))
                pMenu.MRBehaviorValue ins.Target, value
        End Select
    Next ins

    If Not closed Then pMenu.MRSetBehaviorMode False
    Set pMenu = Nothing
    Set pDefaults = Nothing
    Exit Sub

Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error Resume Next
    If Not pMenu Is Nothing Then
        pMenu.MRDetachRunner
        pMenu.MRSetBehaviorMode False
        If Not pMenu.Visible Then Unload pMenu
    End If
    Set pMenu = Nothing
    Set pDefaults = Nothing
    On Error GoTo 0
    Err.Raise number, "SNCBehaviorSession", operation & vbCrLf & _
        "Source asli: " & source & vbCrLf & description
End Sub

Private Sub SnapshotDefaults(ByVal block As MRBehaviorBlock)
    Dim ins As MRBehaviorInstruction, target As String
    Set pDefaults = CreateObject("Scripting.Dictionary")
    For Each ins In block.Instructions
        If ins.Kind = "default" Then
            target = LCase$(ins.Target)
            If Not pDefaults.Exists(target) Then _
                pDefaults.Add target, pMenu.MRBehaviorReadValue(ins.Target)
        End If
    Next ins
End Sub
