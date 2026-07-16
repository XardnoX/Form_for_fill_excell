VERSION 5.00
Begin VB.UserForm frmReview
   Caption         =   "MR_Helper – Kontrola provedených změn"
   BackColor       =   &H00F4F1ED&
   ClientHeight    =   6000
   ClientWidth     =   11400
   StartUpPosition =   1
   Begin VB.CommandButton cmdBack
      Caption         =   "Zpět do formuláře"
      Height          =   420
      Left            =   240
      Top             =   5340
      Width           =   1920
   End
   Begin VB.CommandButton cmdFinish
      Caption         =   "Dokončit"
      Default         =   -1
      Height          =   420
      Left            =   9720
      Top             =   5340
      Width           =   1440
   End
   Begin VB.CommandButton cmdUndoAll
      Caption         =   "Vrátit vše"
      Height          =   420
      Left            =   8280
      Top             =   5340
      Width           =   1320
   End
   Begin VB.CommandButton cmdUndo
      Caption         =   "Zrušit změnu"
      Height          =   420
      Left            =   6600
      Top             =   5340
      Width           =   1560
   End
   Begin VB.CommandButton cmdDetail
      Caption         =   "Detail"
      Height          =   420
      Left            =   5160
      Top             =   5340
      Width           =   1320
   End
   Begin VB.CommandButton cmdGoTo
      Caption         =   "Přejít na buňku"
      Height          =   420
      Left            =   3360
      Top             =   5340
      Width           =   1680
   End
   Begin VB.ListBox lstChanges
      Height          =   3540
      Left            =   240
      Top             =   1500
      Width           =   10920
   End
   Begin VB.Label lblCount
      Caption         =   "Provedené změny: 0"
      Height          =   300
      Left            =   240
      Top             =   1020
      Width           =   10920
   End
   Begin VB.Label lblHeader
      BackColor       =   &H00FFFFFF&
      Caption         =   "MR_HELPER – KONTROLA PROVEDENÝCH ZMĚN"
      ForeColor       =   &H004F4536&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   11400
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
Attribute VB_Name = "frmReview"
Option Explicit

Private Sub UserForm_Initialize()
    ApplyVisualStyle
    ApplyBrandLogo imgLogo
    RefreshList
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
    lstChanges.BackColor = RGB(255, 255, 255)
    StyleReviewButton cmdBack, False
    StyleReviewButton cmdGoTo, False
    StyleReviewButton cmdDetail, False
    StyleReviewButton cmdUndo, False
    StyleReviewButton cmdUndoAll, False
    StyleReviewButton cmdFinish, True
End Sub

Private Sub StyleReviewButton(ByVal button As Object, ByVal primary As Boolean)
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

Private Sub RefreshList()
    Dim record As CChangeRecord
    lstChanges.Clear
    For Each record In gChanges
        lstChanges.AddItem record.Summary
    Next record
    lblCount.Caption = "Provedené změny: " & CStr(gChanges.Count)
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

Private Sub cmdGoTo_Click()
    Dim index As Long, record As CChangeRecord
    index = SelectedIndex
    If index = 0 Then Exit Sub
    Set record = gChanges(index)
    gTargetWorkbook.Activate
    gTargetWorkbook.Worksheets(record.SheetName).Activate
    Application.Goto gTargetWorkbook.Worksheets(record.SheetName).Range(record.CellAddress), True
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
    If index > 0 Then UndoChange index: RefreshList
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
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "Kontrolu ukončete tlačítkem Dokončit.", vbInformation, TOOL_TITLE
    End If
End Sub
