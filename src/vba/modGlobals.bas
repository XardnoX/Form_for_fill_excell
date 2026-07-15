Option Explicit

Public gValues As Object
Public gChanges As Collection
Public gTargetWorkbook As Workbook
Public gSelectedSheets As Object

Public Const APP_NAME As String = "FormularProPredvyplneni"
Public Const APP_SECTION As String = "Settings"
Public Const PHRASES_KEY As String = "Phrases"
Public Const LEGACY_TERMS_KEY As String = "Terms"
Public Const VALUE_HISTORY_KEY As String = "ValueHistory"
Public Const TOOL_TITLE As String = "Formulář pro předvyplnění"

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

Public Function LoadPhrases() As Collection
    Const MISSING_VALUE As String = "{FORMULAR_PRO_PREDVYPLNENI_MISSING}"

    Dim result As New Collection
    Dim raw As String
    Dim legacyRaw As String
    Dim parts() As String
    Dim i As Long
    Dim phrase As String
    Dim initializeDefaults As Boolean

    raw = GetSetting( _
        APP_NAME, _
        APP_SECTION, _
        PHRASES_KEY, _
        MISSING_VALUE _
    )

    If StrComp(raw, MISSING_VALUE, vbBinaryCompare) = 0 Then
        legacyRaw = GetSetting( _
            APP_NAME, _
            APP_SECTION, _
            LEGACY_TERMS_KEY, _
            vbNullString _
        )

        If Len(legacyRaw) > 0 Then
            raw = legacyRaw
        Else
            raw = vbNullString
        End If

        initializeDefaults = True
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

    If initializeDefaults Then
        AddUnique result, "Vypracoval:"
        AddUnique result, "Ověřil a schválil:"
        AddUnique result, "Rozdělovník:"
        AddUnique result, "První vydání:"
        AddUnique result, "Číslo/datum revize:"

        SavePhrases result
    End If

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
