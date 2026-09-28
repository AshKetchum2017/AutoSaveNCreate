Option Explicit

' Class Module (Name) = SNCBehaviorContract. Pure semantic preflight.
Private pText As String
Private pAt As Long

Public Function Validate(ByVal script As String) As MRBehaviorBlock
    Dim parser As New MRBehaviorParser
    Dim blocks As Collection, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction
    Dim stageSeen As Boolean, processed As Boolean, closed As Boolean

    pText = script
    pAt = 1
    Set blocks = parser.Parse(script, Array("AutoSaveNCreate"))
    Set block = blocks(1)

    For Each ins In block.Instructions
        pAt = ins.Position
        If closed Then Fail "Tidak boleh ada instruksi setelah @cmdClose."
        If StrComp(ins.FormName, "AutoSNC", vbTextCompare) <> 0 Then _
            Fail "UserForm tidak terdaftar: " & ins.FormName

        Select Case ins.Kind
            Case "stage"
                If stageSeen Then Fail "AutoSNC hanya boleh mempunyai satu tahap."
                stageSeen = True
            Case "action"
                Select Case LCase$(ins.Target)
                    Case "cmdprocess"
                        If processed Then Fail "@cmdProcess hanya boleh dijalankan sekali."
                        processed = True
                    Case "cmdclose"
                        If Not processed Then Fail "Jalankan @cmdProcess sebelum @cmdClose."
                        closed = True
                    Case Else
                        Fail "Action tidak terdaftar: AutoSNC.@" & ins.Target
                End Select
            Case Else
                If processed Then Fail "Setelah @cmdProcess hanya @cmdClose yang boleh dijalankan."
                ValidateValue ins
        End Select
    Next ins

    pAt = Len(pText) + 1
    If Not processed Then Fail "Akhiri behavior AutoSNC dengan @cmdProcess (boleh diikuti @cmdClose)."

    Set Validate = block
End Function

Private Sub ValidateValue(ByVal ins As MRBehaviorInstruction)
    Dim expected As String
    Dim settings As SNCSettingsStore

    Select Case LCase$(ins.Target)
        Case "cmbuserselection", "cmbcorelversion": expected = "string"
        Case "chkembedcolorprofiles", "chkembedfonts", _
             "optdiea", "opthidie", "optkissa", "opthikiss": expected = "boolean"
        Case Else
            Fail "Target tidak terdaftar: AutoSNC." & ins.Target
    End Select

    If ins.Kind = "skip" Or ins.Kind = "default" Then Exit Sub
    If ins.Kind <> expected Then Fail ins.Target & " harus berupa " & expected & "."

    If expected = "string" Then
        If Len(Trim$(CStr(ins.Value))) = 0 Then Fail ins.Target & " tidak boleh kosong."
        If LCase$(ins.Target) = "cmbcorelversion" Then
            Set settings = New SNCSettingsStore
            If Not settings.IsCorelVersionLabel(CStr(ins.Value)) Then _
                Fail "Versi CDR tidak tersedia: " & CStr(ins.Value)
        End If
    End If
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long, lastBreak As Long
    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    lastBreak = InStrRev(prefix, vbLf)
    column = pAt - lastBreak
    Err.Raise 5, "SNCBehaviorContract", "Semantic AutoSNC baris " & CStr(line) & _
        ", kolom " & CStr(column) & ": " & message
End Sub
