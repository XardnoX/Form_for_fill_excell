VERSION 5.00
Begin VB.UserForm frmReview
   Caption         =   "MR_Helper – Kontrola provedených změn"
   BackColor       =   &H00F4F1ED&
   ClientHeight    =   6600
   ClientWidth     =   7200
   StartUpPosition =   0
   Begin VB.CommandButton cmdBack
      Caption         =   "Zpět do formuláře"
      Height          =   420
      Left            =   240
      Top             =   5100
      Width           =   1440
   End
   Begin VB.CommandButton cmdFinish
      Caption         =   "Dokončit"
      Default         =   -1
      Height          =   420
      Left            =   5520
      Top             =   5760
      Width           =   1440
   End
   Begin VB.CommandButton cmdUndoAll
      Caption         =   "Vrátit vše"
      Height          =   420
      Left            =   4800
      Top             =   5100
      Width           =   1020
   End
   Begin VB.CommandButton cmdUndo
      Caption         =   "Zrušit změnu"
      Height          =   420
      Left            =   3480
      Top             =   5100
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetail
      Caption         =   "Detail"
      Height          =   420
      Left            =   2460
      Top             =   5100
      Width           =   900
   End
   Begin VB.CommandButton cmdPrevious
      Caption         =   "▲"
      Height          =   420
      Left            =   1800
      Top             =   5100
      Width           =   540
   End
   Begin VB.CommandButton cmdNext
      Caption         =   "▼"
      Height          =   420
      Left            =   1800
      Top             =   5580
      Width           =   540
   End
   Begin VB.ListBox lstChanges
      Height          =   3300
      Left            =   240
      Top             =   1500
      Width           =   6240
   End
   Begin VB.Label lblCount
      Caption         =   "Provedené změny: 0"
      Height          =   300
      Left            =   240
      Top             =   1020
      Width           =   6240
   End
   Begin VB.Label lblHeader
      BackColor       =   &H00FFFFFF&
      Caption         =   "MR_HELPER – KONTROLA PROVEDENÝCH ZMĚN"
      ForeColor       =   &H004F4536&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   7200
   End
   Begin VB.Image imgLogo
      BackStyle       =   0
      BorderStyle     =   0
      Height          =   540
      Left            =   120
      PictureSizeMode =   3
      SpecialEffect   =   0
      Top             =   30
      Width           =   960
   End
End
Attribute VB_Name = "frmReview"
Option Explicit
Private mUpdatingList As Boolean

Private Sub UserForm_Initialize()
    ApplyVisualStyle
    ApplyBrandLogo imgLogo
    PositionReviewWindow
    RefreshList
End Sub

Private Sub PositionReviewWindow()
    Dim screenWidth As Double, screenHeight As Double
    Dim rightMargin As Double
    On Error Resume Next
    screenWidth = Application.UsableWidth
    screenHeight = Application.UsableHeight
    If screenWidth <= 0 Or screenHeight <= 0 Then Exit Sub
    rightMargin = screenWidth * 0.05
    Me.Left = screenWidth - Me.Width - rightMargin
    Me.Top = screenHeight * 0.1
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
    EnableMouseWheel Me
End Sub

Private Sub UserForm_Terminate()
    DisableMouseWheel
End Sub

Public Function HandleMouseWheel(ByVal screenX As Long, ByVal screenY As Long, _
                                 ByVal delta As Long) As Boolean
    Dim x As Single, y As Single
    On Error GoTo Unsupported
    If Not MousePointInForm(Me, screenX, screenY, x, y) Then Exit Function
    If x >= lstChanges.Left And x <= lstChanges.Left + lstChanges.Width And _
       y >= lstChanges.Top And y <= lstChanges.Top + lstChanges.Height Then
        ScrollListByWheel lstChanges, delta
        HandleMouseWheel = True
    End If
Unsupported:
End Function

Private Sub ApplyVisualStyle()
    Dim accentBar As Object, changesCard As Object, footerLine As Object

    On Error Resume Next
    Me.BackColor = RGB(245, 247, 250)
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9
    With lblHeader
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(15, 53, 77)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 13
        .Font.Bold = True
    End With

    Set accentBar = Me.Controls.Add("Forms.Label.1", "lblAccentBar", True)
    accentBar.Left = 0
    accentBar.Top = 42
    accentBar.Width = Me.InsideWidth
    accentBar.Height = 3
    accentBar.BackColor = RGB(0, 137, 200)

    Set changesCard = AddSurface("pnlChanges", 10, 50, 438, 260)
    Set footerLine = Me.Controls.Add("Forms.Label.1", "lblFooterLine", True)
    footerLine.Left = 10
    footerLine.Top = 312
    footerLine.Width = Me.InsideWidth - 20
    footerLine.Height = 1
    footerLine.BackColor = RGB(218, 226, 234)

    lblCount.Left = 16
    lblCount.Top = 54
    lblCount.Width = 416
    lblCount.Height = 22
    lblCount.BackColor = RGB(232, 244, 249)
    lblCount.ForeColor = RGB(15, 76, 110)
    lblCount.Font.Bold = True
    lstChanges.BackColor = RGB(255, 255, 255)
    lstChanges.BorderStyle = 1
    lstChanges.BorderColor = RGB(203, 213, 225)
    lstChanges.SpecialEffect = 0
    StyleReviewButton cmdBack, False
    StyleReviewButton cmdPrevious, False
    StyleReviewButton cmdNext, False
    cmdPrevious.ControlTipText = "Předchozí změna"
    cmdNext.ControlTipText = "Následující změna"
    StyleReviewButton cmdDetail, False
    StyleReviewButton cmdUndo, False
    StyleReviewButton cmdUndoAll, False
    StyleReviewButton cmdFinish, True
    On Error GoTo 0
End Sub

Private Function AddSurface(ByVal controlName As String, ByVal controlLeft As Single, _
                            ByVal controlTop As Single, ByVal controlWidth As Single, _
                            ByVal controlHeight As Single) As Object
    Dim surface As Object

    Set surface = Me.Controls.Add("Forms.Label.1", controlName, True)
    surface.Caption = vbNullString
    surface.Left = controlLeft
    surface.Top = controlTop
    surface.Width = controlWidth
    surface.Height = controlHeight
    surface.BackColor = RGB(255, 255, 255)
    surface.BorderStyle = 1
    surface.BorderColor = RGB(226, 232, 240)
    surface.ZOrder 1
    Set AddSurface = surface
End Function

Private Sub StyleReviewButton(ByVal button As Object, ByVal primary As Boolean)
    On Error Resume Next
    button.Font.Name = "Segoe UI"
    button.Font.Size = 9
    button.SpecialEffect = 0
    If primary Then
        button.BackColor = RGB(0, 122, 184)
        button.ForeColor = RGB(255, 255, 255)
        button.Font.Bold = True
    Else
        button.BackColor = RGB(255, 255, 255)
        button.ForeColor = RGB(15, 76, 110)
    End If
    On Error GoTo 0
End Sub

Private Sub RefreshList()
    Dim record As CChangeRecord
    mUpdatingList = True
    lstChanges.Clear
    For Each record In gChanges
        lstChanges.AddItem record.Summary
    Next record
    lblCount.Caption = "Provedené změny: " & CStr(gChanges.Count)
    mUpdatingList = False
End Sub

Private Sub cmdBack_Click()
    UndoAllChanges
    If gChanges.Count > 0 Then
        RefreshList
        MsgBox "Některé změny nelze vrátit, proto se zatím nelze vrátit do formuláře. " & _
               "Zkontrolujte, zda nejsou listy zamčené.", vbExclamation, TOOL_TITLE
        Exit Sub
    End If
    gReturnToPrefill = True
    Unload Me
End Sub

Private Function SelectedIndex() As Long
    If lstChanges.ListIndex < 0 Then
        MsgBox "Nejprve vyberte změnu.", vbInformation, TOOL_TITLE
    Else
        SelectedIndex = lstChanges.ListIndex + 1
    End If
End Function

Private Sub lstChanges_Click()
    If mUpdatingList Then Exit Sub
    GoToSelectedChange
End Sub

Private Sub cmdPrevious_Click()
    Dim newIndex As Long
    If lstChanges.ListCount = 0 Then Exit Sub
    If lstChanges.ListIndex <= 0 Then
        newIndex = lstChanges.ListCount - 1
    Else
        newIndex = lstChanges.ListIndex - 1
    End If
    mUpdatingList = True
    lstChanges.ListIndex = newIndex
    mUpdatingList = False
    GoToSelectedChange
End Sub

Private Sub cmdNext_Click()
    Dim newIndex As Long
    If lstChanges.ListCount = 0 Then Exit Sub
    If lstChanges.ListIndex < 0 Or lstChanges.ListIndex >= lstChanges.ListCount - 1 Then
        newIndex = 0
    Else
        newIndex = lstChanges.ListIndex + 1
    End If
    mUpdatingList = True
    lstChanges.ListIndex = newIndex
    mUpdatingList = False
    GoToSelectedChange
End Sub

Private Sub GoToSelectedChange()
    Dim index As Long, record As CChangeRecord, target As Range
    index = lstChanges.ListIndex + 1
    If index < 1 Or index > gChanges.Count Then Exit Sub
    Set record = gChanges(index)
    gTargetWorkbook.Activate
    gTargetWorkbook.Worksheets(record.SheetName).Activate
    Set target = gTargetWorkbook.Worksheets(record.SheetName).Range(record.CellAddress)
    CenterOnCell target
End Sub

Private Sub CenterOnCell(ByVal target As Range)
    Dim visibleRows As Long, visibleColumns As Long
    Dim newTopRow As Long, newLeftColumn As Long
    Dim maxTopRow As Long, maxLeftColumn As Long
    On Error GoTo Fallback

    Application.Goto target, False
    visibleRows = ActiveWindow.VisibleRange.Rows.Count
    visibleColumns = ActiveWindow.VisibleRange.Columns.Count
    newTopRow = target.Row - visibleRows \ 2
    newLeftColumn = target.Column - visibleColumns \ 2
    If newTopRow < 1 Then newTopRow = 1
    If newLeftColumn < 1 Then newLeftColumn = 1
    maxTopRow = target.Worksheet.Rows.Count - visibleRows + 1
    maxLeftColumn = target.Worksheet.Columns.Count - visibleColumns + 1
    If newTopRow > maxTopRow Then newTopRow = maxTopRow
    If newLeftColumn > maxLeftColumn Then newLeftColumn = maxLeftColumn
    ActiveWindow.ScrollRow = newTopRow
    ActiveWindow.ScrollColumn = newLeftColumn
    target.Select
    Exit Sub

Fallback:
    On Error Resume Next
    Application.Goto target, False
    On Error GoTo 0
End Sub

Private Sub cmdDetail_Click()
    Dim index As Long, record As CChangeRecord
    index = SelectedIndex
    If index = 0 Then Exit Sub
    Set record = gChanges(index)
    MsgBox "List: " & record.SheetName & vbCrLf & _
           "Buňka: " & record.CellAddress & vbCrLf & _
           "Slovní spojení: " & record.Phrase & vbCrLf & _
           "Původní hodnota: " & SafeText(record.OldValue) & vbCrLf & _
           "Nová hodnota: " & SafeText(record.NewValue), vbInformation, "Detail změny"
End Sub

Private Sub cmdUndo_Click()
    Dim index As Long
    index = SelectedIndex
    If index > 0 Then
        UndoChange index
        RefreshList
    End If
End Sub

Private Sub cmdUndoAll_Click()
    UndoAllChanges
    RefreshList
    MsgBox "Všechny dostupné změny byly vráceny.", vbInformation, TOOL_TITLE
End Sub

Private Sub cmdFinish_Click()
    FinishReview
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim answer As VbMsgBoxResult

    If CloseMode <> 0 Then Exit Sub
    If gChanges Is Nothing Then Exit Sub
    If gChanges.Count = 0 Then Exit Sub

    answer = MsgBox( _
        "Opravdu chcete zavřít okno bez uložení změn?" & vbCrLf & _
        "Provedené změny budou vráceny.", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        TOOL_TITLE)

    If answer = vbNo Then
        Cancel = True
        Exit Sub
    End If

    UndoAllChanges
    If gChanges.Count > 0 Then
        Cancel = True
        RefreshList
        MsgBox "Některé změny nelze vrátit. Zkontrolujte, zda nejsou listy zamčené.", _
               vbExclamation, TOOL_TITLE
    End If
End Sub
