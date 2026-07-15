VERSION 5.00
Begin VB.UserForm frmPhraseManager
   Caption         =   "Správa slovních spojení"
   ClientHeight    =   5400
   ClientWidth     =   7800
   StartUpPosition =   1
   Begin VB.CommandButton cmdClose
      Caption         =   "Zavřít"
      Height          =   420
      Left            =   6360
      Top             =   4740
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete
      Caption         =   "Odstranit"
      Height          =   420
      Left            =   6360
      Top             =   4020
      Width           =   1200
   End
   Begin VB.CommandButton cmdRename
      Caption         =   "Přejmenovat"
      Height          =   420
      Left            =   4920
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Přidat"
      Height          =   420
      Left            =   3720
      Top             =   4020
      Width           =   1080
   End
   Begin VB.TextBox txtPhrase
      Height          =   360
      Left            =   240
      Top             =   4080
      Width           =   3240
   End
   Begin VB.ListBox lstPhrases
      Height          =   3240
      Left            =   240
      Top             =   600
      Width           =   7320
   End
   Begin VB.Label lblInfo
      Caption         =   "Uložená slovní spojení"
      Height          =   300
      Left            =   240
      Top             =   180
      Width           =   3000
   End
End
Attribute VB_Name = "frmPhraseManager"
Option Explicit
Private mPhrases As Collection

Private Sub UserForm_Initialize()
    ReloadPhrases
End Sub

Private Sub ReloadPhrases()
    Dim phrase As Variant
    Set mPhrases = LoadPhrases()
    lstPhrases.Clear
    For Each phrase In mPhrases
        lstPhrases.AddItem CStr(phrase)
    Next phrase
End Sub

Private Sub lstPhrases_Click()
    If lstPhrases.ListIndex >= 0 Then txtPhrase.Text = lstPhrases.List(lstPhrases.ListIndex)
End Sub

Private Sub cmdAdd_Click()
    Dim phrase As String, beforeCount As Long
    phrase = Trim$(txtPhrase.Text)
    If Len(phrase) = 0 Then Exit Sub
    beforeCount = mPhrases.Count
    AddUnique mPhrases, phrase
    If mPhrases.Count = beforeCount Then
        MsgBox "Toto slovní spojení již existuje.", vbInformation, TOOL_TITLE
        Exit Sub
    End If
    SavePhrases mPhrases
    txtPhrase.Text = vbNullString
    ReloadPhrases
End Sub

Private Sub cmdRename_Click()
    Dim oldPhrase As String, newPhrase As String, result As New Collection, item As Variant
    If lstPhrases.ListIndex < 0 Then Exit Sub
    oldPhrase = lstPhrases.List(lstPhrases.ListIndex)
    newPhrase = Trim$(txtPhrase.Text)
    If Len(newPhrase) = 0 Then Exit Sub
    For Each item In mPhrases
        If StrComp(CStr(item), oldPhrase, vbTextCompare) = 0 Then AddUnique result, newPhrase Else AddUnique result, CStr(item)
    Next item
    SavePhrases result
    RenameHistoryPhrase oldPhrase, newPhrase
    ReloadPhrases
End Sub

Private Sub cmdDelete_Click()
    Dim phrase As String, result As New Collection, item As Variant
    If lstPhrases.ListIndex < 0 Then Exit Sub
    phrase = lstPhrases.List(lstPhrases.ListIndex)
    For Each item In mPhrases
        If StrComp(CStr(item), phrase, vbTextCompare) <> 0 Then result.Add CStr(item)
    Next item
    SavePhrases result
    DeleteHistoryPhrase phrase
    txtPhrase.Text = vbNullString
    ReloadPhrases
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
