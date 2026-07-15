Attribute VB_Name = "modValueHistory"
Option Explicit

Private Const SEP_PHRASE_CODE As Long = 28
Private Const SEP_PAIR_CODE As Long = 29
Private Const SEP_VALUE_CODE As Long = 30
Private Const MAX_VALUES_PER_PHRASE As Long = 25

Public Function LoadValueHistory() As Object
    Dim history As Object, raw As String, rows() As String, pair() As String, i As Long
    Set history = NewTextDictionary()
    raw = GetSetting(APP_NAME, APP_SECTION, VALUE_HISTORY_KEY, vbNullString)

    If Len(raw) > 0 Then
        rows = Split(raw, Chr$(SEP_PHRASE_CODE))
        For i = LBound(rows) To UBound(rows)
            pair = Split(rows(i), Chr$(SEP_PAIR_CODE), 2)
            If UBound(pair) = 1 Then history(pair(0)) = pair(1)
        Next i
    End If
    Set LoadValueHistory = history
End Function

Public Sub AddValueToHistory(ByVal history As Object, ByVal phrase As String, ByVal value As String)
    Dim values() As String, item As Variant, result As String, count As Long
    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub

    result = value
    count = 1
    If history.Exists(phrase) Then
        values = Split(CStr(history(phrase)), Chr$(SEP_VALUE_CODE))
        For Each item In values
            If Len(CStr(item)) > 0 And StrComp(CStr(item), value, vbTextCompare) <> 0 Then
                If count >= MAX_VALUES_PER_PHRASE Then Exit For
                result = result & Chr$(SEP_VALUE_CODE) & CStr(item)
                count = count + 1
            End If
        Next item
    End If
    history(phrase) = result
End Sub

Public Sub SaveValueHistory(ByVal history As Object)
    Dim phrase As Variant, raw As String
    For Each phrase In history.Keys
        If Len(raw) > 0 Then raw = raw & Chr$(SEP_PHRASE_CODE)
        raw = raw & CStr(phrase) & Chr$(SEP_PAIR_CODE) & CStr(history(phrase))
    Next phrase
    SaveSetting APP_NAME, APP_SECTION, VALUE_HISTORY_KEY, raw
End Sub

Public Sub FillSuggestions(ByVal combo As Object, ByVal history As Object, ByVal phrase As String)
    Dim values() As String, item As Variant
    combo.Clear
    If history.Exists(phrase) Then
        values = Split(CStr(history(phrase)), Chr$(SEP_VALUE_CODE))
        For Each item In values
            If Len(CStr(item)) > 0 Then combo.AddItem CStr(item)
        Next item
    End If
    combo.Value = vbNullString
End Sub

Public Sub RenameHistoryPhrase(ByVal oldPhrase As String, ByVal newPhrase As String)
    Dim history As Object
    Set history = LoadValueHistory()
    If history.Exists(oldPhrase) Then
        history(newPhrase) = history(oldPhrase)
        history.Remove oldPhrase
        SaveValueHistory history
    End If
End Sub

Public Sub DeleteHistoryPhrase(ByVal phrase As String)
    Dim history As Object
    Set history = LoadValueHistory()
    If history.Exists(phrase) Then
        history.Remove phrase
        SaveValueHistory history
    End If
End Sub
