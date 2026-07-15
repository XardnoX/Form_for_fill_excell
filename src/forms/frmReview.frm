VERSION 5.00
Begin VB.UserForm frmReview
   Caption         =   "Kontrola provedených změn"
   ClientHeight    =   6000
   ClientWidth     =   11400
   StartUpPosition =   1
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
      Height          =   4680
      Left            =   240
      Top             =   240
      Width           =   10920
   End
End
Attribute VB_Name = "frmReview"
Option Explicit

Private Sub UserForm_Initialize()
    RefreshList
End Sub

Private Sub RefreshList()
    Dim record As CChangeRecord
    lstChanges.Clear
    For Each record In gChanges
        lstChanges.AddItem record.Summary
    Next record
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
