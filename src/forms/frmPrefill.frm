VERSION 5.00
Begin VB.UserForm frmPrefill
   Caption         =   "Formuláø pro pøedvyplnìní"
   ClientHeight    =   7800
   ClientWidth     =   11200
   StartUpPosition =   1
   Begin VB.CommandButton cmdSave
      Caption         =   "Použít zmìny"
      Default         =   -1
      Height          =   420
      Left            =   9360
      Top             =   7140
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Zrušit"
      Height          =   420
      Left            =   8040
      Top             =   7140
      Width           =   1200
   End
   Begin VB.CommandButton cmdManage
      Caption         =   "Spravovat slovní spojení"
      Height          =   420
      Left            =   3240
      Top             =   6420
      Width           =   2400
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Obnovit výskyty"
      Height          =   420
      Left            =   240
      Top             =   6420
      Width           =   2400
   End
   Begin VB.ListBox lstSheets
      Height          =   4920
      Left            =   240
      MultiSelect     =   1
      Top             =   1260
      Width           =   2640
   End
   Begin VB.Frame fraPhrases
      Caption         =   "Nalezená slovní spojení a hodnoty"
      Height          =   4920
      Left            =   3120
      ScrollBars      =   2
      Top             =   1260
      Width           =   7680
   End
   Begin VB.Label lblCount
      Caption         =   ""
      Height          =   300
      Left            =   240
      Top             =   840
      Width           =   10320
   End
   Begin VB.Label lblHeader
      BackColor       =   &H00726A62&
      Caption         =   "FORMULÁØ PRO PØEDVYPLNÌNÍ"
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   11200
   End
End
Attribute VB_Name = "frmPrefill"
Option Explicit

Private mAllPhrases As Collection
Private mVisiblePhrases As Collection
Private mInputs As Collection
Private mValueHistory As Object

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set mAllPhrases = LoadPhrases()
    Set mValueHistory = LoadValueHistory()
    lstSheets.Clear
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            lstSheets.AddItem ws.Name
            lstSheets.Selected(lstSheets.ListCount - 1) = True
        End If
    Next ws
    RefreshContent
End Sub

Private Sub CaptureSelectedSheets()
    Dim i As Long
    Set gSelectedSheets = NewTextDictionary()
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then gSelectedSheets(lstSheets.List(i)) = True
    Next i
End Sub

Private Sub RefreshContent()
    Dim i As Long, y As Single, lbl As Object, combo As Object, occurrenceCount As Long
    CaptureSelectedSheets
    For i = fraPhrases.Controls.Count - 1 To 0 Step -1
        fraPhrases.Controls.Remove fraPhrases.Controls(i).Name
    Next i
    Set mVisiblePhrases = PhrasesInSelectedSheets(mAllPhrases)
    Set mInputs = New Collection
    y = 18
    For i = 1 To mVisiblePhrases.Count
        occurrenceCount = CountPhraseInSelectedSheets(CStr(mVisiblePhrases(i)))
        Set lbl = fraPhrases.Controls.Add("Forms.Label.1", "lblPhrase" & i, True)
        lbl.Caption = CStr(mVisiblePhrases(i)) & "  (" & occurrenceCount & "x)"
        lbl.Left = 18: lbl.Top = y + 4: lbl.Width = 310: lbl.Height = 20

        Set combo = fraPhrases.Controls.Add("Forms.ComboBox.1", "cboValue" & i, True)
        combo.Left = 330: combo.Top = y: combo.Width = 330: combo.Height = 24
        combo.Style = 0
        combo.MatchEntry = 1
        combo.MatchRequired = False
        FillSuggestions combo, mValueHistory, CStr(mVisiblePhrases(i))
        mInputs.Add combo
        y = y + 32
    Next i
    If mVisiblePhrases.Count = 0 Then
        Set lbl = fraPhrases.Controls.Add("Forms.Label.1", "lblEmpty", True)
        lbl.Caption = "Na vybraných listech nebylo nalezeno žádné uložené slovní spojení."
        lbl.Left = 18: lbl.Top = 28: lbl.Width = 600
    End If
    fraPhrases.ScrollHeight = y + 18
    lblCount.Caption = CStr(mVisiblePhrases.Count) & " slovních spojení na " & CStr(gSelectedSheets.Count) & " vybraných listech"
End Sub

Private Sub cmdRefresh_Click()
    RefreshContent
End Sub

Private Sub cmdManage_Click()
    frmPhraseManager.Show
    Set mAllPhrases = LoadPhrases()
    RefreshContent
End Sub

Private Sub cmdSave_Click()
    Dim i As Long, phrase As String, enteredValue As String
    CaptureSelectedSheets
    If gSelectedSheets.Count = 0 Then
        MsgBox "Vyberte alespoò jeden viditelný list.", vbExclamation, TOOL_TITLE
        Exit Sub
    End If

    Set gValues = NewTextDictionary()
    For i = 1 To mVisiblePhrases.Count
        phrase = CStr(mVisiblePhrases(i))
        enteredValue = CStr(mInputs(i).Value)
        gValues(phrase) = enteredValue
        If Len(Trim$(enteredValue)) > 0 Then AddValueToHistory mValueHistory, phrase, enteredValue
    Next i
    SaveValueHistory mValueHistory
    Me.Hide
    ApplyPrefillChanges
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
