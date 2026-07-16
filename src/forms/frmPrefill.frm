VERSION 5.00
Begin VB.UserForm frmPrefill
   Caption         =   "Formulář pro předvyplnění"
   BackColor       =   &H00F4F1ED&
   ClientHeight    =   7500
   ClientWidth     =   9200
   StartUpPosition =   1
   Begin VB.CommandButton cmdSave
      Caption         =   "Použít změny"
      Default         =   -1
      Height          =   420
      Left            =   7440
      Top             =   6840
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Zrušit"
      Height          =   420
      Left            =   6120
      Top             =   6840
      Width           =   1200
   End
   Begin VB.CommandButton cmdManage
      Caption         =   "Spravovat slovní spojení"
      Height          =   420
      Left            =   2520
      Top             =   6120
      Width           =   2400
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Obnovit výskyty"
      Height          =   420
      Left            =   240
      Top             =   6120
      Width           =   2040
   End
   Begin VB.ListBox lstSheets
      Height          =   4620
      Left            =   240
      MultiSelect     =   1
      Top             =   1260
      Width           =   2040
   End
   Begin VB.Frame fraPhrases
      Caption         =   "Nalezená slovní spojení a hodnoty"
      Height          =   4620
      Left            =   2400
      ScrollBars      =   2
      Top             =   1260
      Width           =   6560
   End
   Begin VB.Label lblCount
      Caption         =   ""
      Height          =   300
      Left            =   240
      Top             =   840
      Width           =   8480
   End
   Begin VB.Label lblHeader
      BackColor       =   &H00726A62&
      Caption         =   "FORMULÁŘ PRO PŘEDVYPLNĚNÍ"
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   9200
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
    ApplyVisualStyle
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

Private Sub ApplyVisualStyle()
    Me.BackColor = RGB(244, 241, 237)
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9

    With lblHeader
        .BackColor = RGB(54, 69, 79)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 13
        .Font.Bold = True
    End With

    lblCount.ForeColor = RGB(75, 85, 99)
    lblCount.Font.Bold = True
    fraPhrases.BackColor = RGB(255, 255, 255)
    fraPhrases.ForeColor = RGB(54, 69, 79)
    lstSheets.BackColor = RGB(255, 255, 255)

    StyleButton cmdRefresh, False
    StyleButton cmdManage, False
    StyleButton cmdCancel, False
    StyleButton cmdSave, True
End Sub

Private Sub StyleButton(ByVal button As Object, ByVal primary As Boolean)
    button.Font.Name = "Segoe UI"
    button.Font.Size = 9
    If primary Then
        button.BackColor = RGB(54, 69, 79)
        button.ForeColor = RGB(255, 255, 255)
        button.Font.Bold = True
    Else
        button.BackColor = RGB(226, 232, 240)
        button.ForeColor = RGB(31, 41, 55)
    End If
End Sub

Private Sub CaptureSelectedSheets()
    Dim i As Long
    Set gSelectedSheets = NewTextDictionary()
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then gSelectedSheets(lstSheets.List(i)) = True
    Next i
End Sub

Private Sub RefreshContent()
    Dim i As Long, y As Single, lbl As Object, inputBox As Object, occurrenceCount As Long
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
        lbl.Left = 18: lbl.Top = y: lbl.Width = 370: lbl.Height = 18
        lbl.ForeColor = RGB(55, 65, 81)
        lbl.Font.Bold = True

        Set inputBox = fraPhrases.Controls.Add("Forms.ComboBox.1", "cboValue" & i, True)
        inputBox.Left = 18: inputBox.Top = y + 20: inputBox.Width = 370: inputBox.Height = 24
        inputBox.BackColor = RGB(255, 255, 255)
        inputBox.Style = 0
        inputBox.MatchEntry = 1
        inputBox.MatchRequired = False
        inputBox.ShowDropButtonWhen = 0
        FillSuggestions inputBox, mValueHistory, CStr(mVisiblePhrases(i))
        mInputs.Add inputBox
        y = y + 52
    Next i
    If mVisiblePhrases.Count = 0 Then
        Set lbl = fraPhrases.Controls.Add("Forms.Label.1", "lblEmpty", True)
        lbl.Caption = "Na vybraných listech nebylo nalezeno žádné uložené slovní spojení."
        lbl.Left = 18: lbl.Top = 28: lbl.Width = 370: lbl.Height = 36
        lbl.ForeColor = RGB(107, 114, 128)
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
        MsgBox "Vyberte alespoň jeden viditelný list.", vbExclamation, TOOL_TITLE
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
    gReturnToPrefill = False
    Me.Hide
    ApplyPrefillChanges
    If gReturnToPrefill Then
        gReturnToPrefill = False
        Me.Show
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
