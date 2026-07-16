Option Explicit

Public gValues As Object
Public gChanges As Collection
Public gTargetWorkbook As Workbook
Public gSelectedSheets As Object
Public gReturnToPrefill As Boolean

Public Const APP_NAME As String = "MR_Helper"
Public Const LEGACY_APP_NAME As String = "FormularProPredvyplneni"
Public Const APP_SECTION As String = "Settings"
Public Const PHRASES_KEY As String = "Phrases"
Public Const LEGACY_TERMS_KEY As String = "Terms"
Public Const VALUE_HISTORY_KEY As String = "ValueHistory"
Public Const TOOL_TITLE As String = "MR_Helper"
Public Const LOGO_FILE_NAME As String = "MR_Helper_logo.jpg"
Public Const ASSET_FOLDER_NAME As String = "MR_Helper_assets"

Public Sub ShowPrefillForm(Optional control As Object)
    If ActiveWorkbook Is Nothing Then Exit Sub

    If ActiveWorkbook Is ThisWorkbook Then
        MsgBox "Aktivujte sešit, který chcete předvyplnit.", _
               vbExclamation, _
               TOOL_TITLE
        Exit Sub
    End If

    Set gTargetWorkbook = ActiveWorkbook
    Set gValues = NewTextDictionary()
    Set gSelectedSheets = NewTextDictionary()

    frmPrefill.Show
End Sub

Public Function NewTextDictionary() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set NewTextDictionary = d
End Function

Public Sub ApplyBrandLogo(ByVal imageControl As Object)
    Dim logoPath As String
    On Error Resume Next
    logoPath = ThisWorkbook.Path & Application.PathSeparator & ASSET_FOLDER_NAME & _
               Application.PathSeparator & LOGO_FILE_NAME
    If Dir$(logoPath) = vbNullString Then
        logoPath = ThisWorkbook.Path & Application.PathSeparator & "assets" & _
                   Application.PathSeparator & LOGO_FILE_NAME
    End If
    If Dir$(logoPath) <> vbNullString Then
        Set imageControl.Picture = LoadPicture(logoPath)
        imageControl.PictureSizeMode = 3
        imageControl.BorderStyle = 0
        imageControl.SpecialEffect = 0
        imageControl.ZOrder 0
    End If
    On Error GoTo 0
End Sub

Public Function LoadPhrases() As Collection
    Dim result As New Collection
    Dim raw As String
    Dim parts() As String
    Dim i As Long
    Dim phrase As String

    raw = GetSetting( _
        APP_NAME, _
        APP_SECTION, _
        PHRASES_KEY, _
        vbNullString _
    )

    If Len(raw) = 0 Then
        raw = GetSetting( _
            APP_NAME, _
            APP_SECTION, _
            LEGACY_TERMS_KEY, _
            vbNullString _
        )
    End If

    If Len(raw) = 0 Then
        raw = GetSetting(LEGACY_APP_NAME, APP_SECTION, PHRASES_KEY, vbNullString)
        If Len(raw) = 0 Then
            raw = GetSetting(LEGACY_APP_NAME, APP_SECTION, LEGACY_TERMS_KEY, vbNullString)
        End If
    End If

    If Len(raw) > 0 Then
        parts = Split(raw, vbTab)

        For i = LBound(parts) To UBound(parts)
            phrase = Trim$(parts(i))

            If Len(phrase) > 0 Then
                AddUnique result, phrase
            End If
        Next i
    End If

    AddUnique result, "Vypracoval:"
    AddUnique result, "Ověřil a schválil:"
    AddUnique result, "Rozdělovník:"
    AddUnique result, "První vydání:"
    AddUnique result, "Číslo/datum revize:"
    AddUnique result, "Zákazník:"
    AddUnique result, "Číslo odlitku:"
    AddUnique result, "Název odlitku:"
    AddUnique result, "Operace:"
    AddUnique result, "Stroj:"
    AddUnique result, "Schválil:"
    AddUnique result, "Ověřil:"

    SavePhrases result

    Set LoadPhrases = result
End Function

Public Sub SavePhrases(ByVal phrases As Collection)
    Dim i As Long
    Dim raw As String

    For i = 1 To phrases.Count
        If InStr(CStr(phrases(i)), vbTab) = 0 Then
            If Len(raw) > 0 Then
                raw = raw & vbTab
            End If

            raw = raw & CStr(phrases(i))
        End If
    Next i

    SaveSetting _
        APP_NAME, _
        APP_SECTION, _
        PHRASES_KEY, _
        raw
End Sub

Public Sub AddUnique( _
    ByVal items As Collection, _
    ByVal value As String _
)
    Dim item As Variant

    For Each item In items
        If StrComp(CStr(item), value, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next item

    items.Add value
End Sub

Public Function IsRequiredPhrase(ByVal phrase As String) As Boolean
    Select Case LCase$(Trim$(phrase))
        Case LCase$("Vypracoval:"), _
             LCase$("Ověřil a schválil:"), _
             LCase$("Rozdělovník:"), _
             LCase$("První vydání:"), _
             LCase$("Číslo/datum revize:"), _
             LCase$("Zákazník:"), _
             LCase$("Číslo odlitku:"), _
             LCase$("Název odlitku:"), _
             LCase$("Operace:"), _
             LCase$("Stroj:")
            IsRequiredPhrase = True
    End Select
End Function

Public Function PhrasesInSelectedSheets( _
    ByVal allPhrases As Collection _
) As Collection

    Dim visible As New Collection
    Dim phrase As Variant

    For Each phrase In allPhrases
        If CountPhraseInSelectedSheets(CStr(phrase)) > 0 Then
            visible.Add CStr(phrase)
        End If
    Next phrase

    Set PhrasesInSelectedSheets = visible
End Function

Public Function CountPhraseInSelectedSheets( _
    ByVal phrase As String _
) As Long

    Dim ws As Worksheet
    Dim rng As Range
    Dim hit As Range
    Dim firstAddress As String

    If gTargetWorkbook Is Nothing Then Exit Function
    If gSelectedSheets Is Nothing Then Exit Function

    On Error GoTo SafeExit

    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If gSelectedSheets.Exists(ws.Name) Then
                Set rng = ws.UsedRange

                Set hit = rng.Find( _
                    What:=phrase, _
                    After:=rng.Cells(rng.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False _
                )

                If Not hit Is Nothing Then
                    firstAddress = hit.Address(External:=True)

                    Do
                        CountPhraseInSelectedSheets = _
                            CountPhraseInSelectedSheets + 1

                        Set hit = rng.FindNext(hit)

                        If hit Is Nothing Then Exit Do
                    Loop While _
                        hit.Address(External:=True) <> firstAddress
                End If
            End If
        End If
    Next ws

SafeExit:
End Function

Public Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then
        SafeText = "#CHYBA"
    Else
        SafeText = CStr(value)
    End If
End Function
