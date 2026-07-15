Attribute VB_Name = "modGlobals"
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
Public Const TOOL_TITLE As String = "Formuláø pro pøedvyplnìní"

Public Sub ShowPrefillForm(Optional control As Object)
    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveWorkbook Is ThisWorkbook Then
        MsgBox "Aktivujte sešit, který chcete pøedvyplnit.", vbExclamation, TOOL_TITLE
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
    Dim result As New Collection
    Dim raw As String, parts() As String, i As Long, phrase As String

    raw = GetSetting(APP_NAME, APP_SECTION, PHRASES_KEY, vbNullString)
    If Len(raw) = 0 Then raw = GetSetting(APP_NAME, APP_SECTION, LEGACY_TERMS_KEY, vbNullString)

    If Len(raw) > 0 Then
        parts = Split(raw, vbTab)
        For i = LBound(parts) To UBound(parts)
            phrase = Trim$(parts(i))
            If Len(phrase) > 0 And Not IsLegacyPhrase(phrase) Then AddUnique result, phrase
        Next i
    End If

    AddUnique result, "Vypracoval:"
    AddUnique result, "Ovìøil a schválil:"
    AddUnique result, "Rozdìlovník:"
    AddUnique result, "První vydání:"
    AddUnique result, "Èíslo/datum revize:"

    SavePhrases result
    Set LoadPhrases = result
End Function

Public Sub SavePhrases(ByVal phrases As Collection)
    Dim i As Long, raw As String
    For i = 1 To phrases.Count
        If InStr(CStr(phrases(i)), vbTab) = 0 Then
            If Len(raw) > 0 Then raw = raw & vbTab
            raw = raw & CStr(phrases(i))
        End If
    Next i
    SaveSetting APP_NAME, APP_SECTION, PHRASES_KEY, raw
End Sub

Private Function IsLegacyPhrase(ByVal phrase As String) As Boolean
    IsLegacyPhrase = (StrComp(phrase, "Motor", vbTextCompare) = 0 Or _
                      StrComp(phrase, "Výkon", vbTextCompare) = 0 Or _
                      StrComp(phrase, "Napìtí", vbTextCompare) = 0)
End Function

Public Sub AddUnique(ByVal items As Collection, ByVal value As String)
    Dim item As Variant
    For Each item In items
        If StrComp(CStr(item), value, vbTextCompare) = 0 Then Exit Sub
    Next item
    items.Add value
End Sub

Public Function PhrasesInSelectedSheets(ByVal allPhrases As Collection) As Collection
    Dim visible As New Collection, phrase As Variant
    For Each phrase In allPhrases
        If CountPhraseInSelectedSheets(CStr(phrase)) > 0 Then visible.Add CStr(phrase)
    Next phrase
    Set PhrasesInSelectedSheets = visible
End Function

Public Function CountPhraseInSelectedSheets(ByVal phrase As String) As Long
    Dim ws As Worksheet, rng As Range, hit As Range, firstAddress As String
    If gTargetWorkbook Is Nothing Or gSelectedSheets Is Nothing Then Exit Function

    On Error GoTo SafeExit
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And gSelectedSheets.Exists(ws.Name) Then
            Set rng = ws.UsedRange
            Set hit = rng.Find(What:=phrase, After:=rng.Cells(rng.Cells.Count), _
                LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Not hit Is Nothing Then
                firstAddress = hit.Address(External:=True)
                Do
                    CountPhraseInSelectedSheets = CountPhraseInSelectedSheets + 1
                    Set hit = rng.FindNext(hit)
                    If hit Is Nothing Then Exit Do
                Loop While hit.Address(External:=True) <> firstAddress
            End If
        End If
    Next ws
SafeExit:
End Function

Public Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then SafeText = "#CHYBA" Else SafeText = CStr(value)
End Function
