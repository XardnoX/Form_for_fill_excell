Attribute VB_Name = "modPrefillEngine"
Option Explicit

Public Sub ApplyPrefillChanges()
    Dim ws As Worksheet, phrase As Variant, enteredValue As Variant, rng As Range
    Dim hit As Range, target As Range, hits As Collection
    Dim record As CChangeRecord, touched As Object, key As String
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean

    On Error GoTo Fatal
    Set gChanges = New Collection
    Set touched = NewTextDictionary()

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And gSelectedSheets.Exists(ws.Name) And Not ws.ProtectContents Then
            Set rng = ws.UsedRange
            For Each phrase In gValues.Keys
                enteredValue = gValues(phrase)
                If Len(CStr(enteredValue)) > 0 Then
                    ' Zápis do listu během FindNext může změnit interní stav
                    ' hledání Excelu a způsobit přeskočení dalších výskytů.
                    ' Proto nejdříve sesbíráme všechny nálezy a až potom zapisujeme.
                    Set hits = FindPhraseHits(rng, CStr(phrase))
                    For Each hit In hits
                            Set target = RightCellFor(hit)
                            If Not target Is Nothing Then
                                key = ws.Name & "!" & target.Address(False, False)
                                If Not touched.Exists(key) Then
                                    touched.Add key, CStr(phrase)
                                    Set record = New CChangeRecord
                                    CaptureOriginalState record, target, CStr(phrase)
                                    record.NewValue = CoerceValue(enteredValue, target)
                                    target.Value2 = record.NewValue
                                    target.NumberFormat = record.OldNumberFormat
                                    target.Interior.Pattern = xlSolid
                                    target.Interior.Color = RGB(255, 235, 59)
                                    gChanges.Add record
                                End If
                            End If
                    Next hit
                End If
            Next phrase
        End If
    Next ws

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    If gChanges.Count = 0 Then
        MsgBox "Nebyly provedeny žádné změny. Skryté, nevybrané a zamčené listy byly přeskočeny.", vbInformation, TOOL_TITLE
    Else
        frmReview.Show
    End If
    Exit Sub
Fatal:
    MsgBox "Operaci se nepodařilo dokončit: " & Err.Description, vbCritical, TOOL_TITLE
    Resume CleanExit
End Sub

Private Function FindPhraseHits(ByVal searchRange As Range, ByVal phrase As String) As Collection
    Dim result As New Collection
    Dim hit As Range
    Dim firstAddress As String

    Set hit = searchRange.Find(What:=phrase, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

    If Not hit Is Nothing Then
        firstAddress = hit.Address
        Do
            result.Add hit
            Set hit = searchRange.Find(What:=phrase, After:=hit, _
                LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If hit Is Nothing Then Exit Do
        Loop While hit.Address <> firstAddress
    End If

    Set FindPhraseHits = result
End Function

Private Sub CaptureOriginalState(ByVal record As CChangeRecord, ByVal target As Range, ByVal phrase As String)
    record.SheetName = target.Worksheet.Name
    record.CellAddress = target.Address(False, False)
    record.Phrase = phrase
    record.OldValue = target.Value2
    record.OldHasFormula = target.HasFormula
    If record.OldHasFormula Then record.OldFormula = target.Formula2
    record.OldPattern = target.Interior.Pattern
    record.OldColor = target.Interior.Color
    record.OldNumberFormat = target.NumberFormat
End Sub

Private Function CoerceValue(ByVal inputValue As Variant, ByVal target As Range) As Variant
    Dim s As String
    s = CStr(inputValue)
    If VarType(target.Value2) = vbBoolean Then
        CoerceValue = (LCase$(s) = "ano" Or LCase$(s) = "true" Or s = "1")
    ElseIf IsDate(target.Value) And IsDate(s) Then
        CoerceValue = CDbl(CDate(s))
    ElseIf IsNumeric(target.Value2) And Len(CStr(target.Value2)) > 0 And IsNumeric(s) Then
        CoerceValue = CDbl(s)
    Else
        CoerceValue = s
    End If
End Function

Private Function RightCellFor(ByVal source As Range) As Range
    Dim col As Long, row As Long
    If source.MergeCells Then
        row = source.MergeArea.Row
        col = source.MergeArea.Column + source.MergeArea.Columns.Count
    Else
        row = source.Row
        col = source.Column + 1
    End If
    If col > source.Worksheet.Columns.Count Then Exit Function
    Set RightCellFor = source.Worksheet.Cells(row, col)
    If RightCellFor.MergeCells Then Set RightCellFor = RightCellFor.MergeArea.Cells(1, 1)
End Function

Private Sub RestoreRecord(ByVal record As CChangeRecord)
    Dim c As Range
    Set c = gTargetWorkbook.Worksheets(record.SheetName).Range(record.CellAddress)
    If record.OldHasFormula Then c.Formula2 = record.OldFormula Else c.Value2 = record.OldValue
    c.NumberFormat = record.OldNumberFormat
    c.Interior.Pattern = record.OldPattern
    If record.OldPattern <> xlNone Then c.Interior.Color = record.OldColor
End Sub

Public Sub UndoChange(ByVal index As Long)
    Dim record As CChangeRecord
    If index < 1 Or index > gChanges.Count Then Exit Sub
    Set record = gChanges(index)
    If gTargetWorkbook.Worksheets(record.SheetName).ProtectContents Then
        MsgBox "List je nyní zamčený. Změnu nelze zrušit.", vbExclamation, TOOL_TITLE
        Exit Sub
    End If
    RestoreRecord record
    gChanges.Remove index
End Sub

Public Sub UndoAllChanges()
    Dim i As Long, record As CChangeRecord
    For i = gChanges.Count To 1 Step -1
        Set record = gChanges(i)
        If Not gTargetWorkbook.Worksheets(record.SheetName).ProtectContents Then
            RestoreRecord record
            gChanges.Remove i
        End If
    Next i
End Sub

Public Sub FinishReview()
    Dim record As CChangeRecord, c As Range
    On Error Resume Next
    For Each record In gChanges
        Set c = gTargetWorkbook.Worksheets(record.SheetName).Range(record.CellAddress)
        c.Interior.Pattern = record.OldPattern
        If record.OldPattern <> xlNone Then c.Interior.Color = record.OldColor
    Next record
    On Error GoTo 0
End Sub
