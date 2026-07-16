VERSION 5.00
Begin VB.UserForm frmPrefill
   Caption         =   "MR_Helper"
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
      Height          =   3420
      Left            =   240
      MultiSelect     =   1
      Top             =   2460
      Width           =   2040
   End
   Begin VB.CheckBox chkAllSheets
      Caption         =   ""
      Height          =   300
      Left            =   240
      Top             =   2040
      Value           =   -1
      Width           =   240
   End
   Begin VB.Label lblAllSheets
      Caption         =   "Vybrat vše / zrušit výběr"
      Height          =   240
      Left            =   540
      Top             =   2070
      Width           =   1740
   End
   Begin VB.TextBox txtSheetSearch
      Height          =   360
      Left            =   240
      Top             =   1560
      Width           =   2040
   End
   Begin VB.Label lblSheetSearch
      Caption         =   "Hledat list:"
      Height          =   240
      Left            =   240
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "MR_HELPER"
      ForeColor       =   &H004F4536&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   9200
   End
   Begin VB.Image imgLogo
      BackStyle       =   0
      Height          =   540
      Left            =   120
      PictureSizeMode =   3
      Top             =   30
      Width           =   960
   End
End
Attribute VB_Name = "frmPrefill"
Option Explicit

Private mAllPhrases As Collection
Private mVisiblePhrases As Collection
Private mInputs As Collection
Private mValueHistory As Object
Private mSheetSelection As Object
Private mUpdatingSheetControls As Boolean

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    ApplyVisualStyle
    ApplyBrandLogo imgLogo
    Set mAllPhrases = LoadPhrases()
    Set mValueHistory = LoadValueHistory()
    Set mSheetSelection = NewTextDictionary()
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            mSheetSelection(ws.Name) = True
        End If
    Next ws
    txtSheetSearch.Value = vbNullString
    mUpdatingSheetControls = True
    chkAllSheets.Value = True
    mUpdatingSheetControls = False
    PopulateSheetList
    RefreshContent
End Sub

Private Sub ApplyVisualStyle()
    Me.BackColor = RGB(244, 241, 237)
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9

    With lblHeader
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(31, 41, 55)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 13
        .Font.Bold = True
    End With

    lblCount.ForeColor = RGB(75, 85, 99)
    lblCount.Font.Bold = True
    fraPhrases.BackColor = RGB(255, 255, 255)
    fraPhrases.ForeColor = RGB(54, 69, 79)
    lstSheets.BackColor = RGB(255, 255, 255)
    txtSheetSearch.BackColor = RGB(255, 255, 255)
    lblSheetSearch.ForeColor = RGB(75, 85, 99)
    lblSheetSearch.Font.Bold = True
    lblAllSheets.ForeColor = RGB(55, 65, 81)

    StyleButton cmdRefresh, False
    StyleButton cmdManage, False
    StyleButton cmdCancel, False
    StyleButton cmdSave, True
End Sub

Private Sub PopulateSheetList()
    Dim ws As Worksheet, query As String, index As Long
    query = LCase$(Trim$(CStr(txtSheetSearch.Value)))
    mUpdatingSheetControls = True
    lstSheets.Clear
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Len(query) = 0 Or InStr(1, LCase$(ws.Name), query, vbTextCompare) > 0 Then
                lstSheets.AddItem ws.Name
                index = lstSheets.ListCount - 1
                If mSheetSelection.Exists(ws.Name) Then
                    lstSheets.Selected(index) = CBool(mSheetSelection(ws.Name))
                End If
            End If
        End If
    Next ws
    mUpdatingSheetControls = False
End Sub

Private Sub CaptureDisplayedSheetSelection()
    Dim i As Long
    If mSheetSelection Is Nothing Then Exit Sub
    For i = 0 To lstSheets.ListCount - 1
        mSheetSelection(lstSheets.List(i)) = lstSheets.Selected(i)
    Next i
End Sub

Private Sub txtSheetSearch_Change()
    If mUpdatingSheetControls Then Exit Sub
    CaptureDisplayedSheetSelection
    PopulateSheetList
End Sub

Private Sub lstSheets_Change()
    If mUpdatingSheetControls Then Exit Sub
    CaptureDisplayedSheetSelection
    mUpdatingSheetControls = True
    chkAllSheets.Value = False
    mUpdatingSheetControls = False
End Sub

Private Sub chkAllSheets_Click()
    Dim ws As Worksheet, selectAll As Boolean
    If mUpdatingSheetControls Then Exit Sub
    selectAll = CBool(chkAllSheets.Value)
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then mSheetSelection(ws.Name) = selectAll
    Next ws
    PopulateSheetList
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
    Dim ws As Worksheet
    CaptureDisplayedSheetSelection
    Set gSelectedSheets = NewTextDictionary()
    For Each ws In gTargetWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And mSheetSelection.Exists(ws.Name) Then
            If CBool(mSheetSelection(ws.Name)) Then gSelectedSheets(ws.Name) = True
        End If
    Next ws
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
