VERSION 5.00
Begin VB.UserForm frmPhraseManager
   Caption         =   "MR_Helper – Správa slovních spojení"
   BackColor       =   &H00F4F1ED&
   ClientHeight    =   5400
   ClientWidth     =   7800
   StartUpPosition =   1
   Begin VB.CommandButton cmdHelp
      Caption         =   "?"
      Height          =   360
      Left            =   7200
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton cmdImport
      Caption         =   "Import TXT"
      Height          =   420
      Left            =   240
      Top             =   4740
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport
      Caption         =   "Export TXT"
      Height          =   420
      Left            =   1800
      Top             =   4740
      Width           =   1440
   End
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
      Height          =   3000
      Left            =   240
      Top             =   840
      Width           =   7320
   End
   Begin VB.Label lblInfo
      BackColor       =   &H00FFFFFF&
      Caption         =   "MR_HELPER – ULOŽENÁ SLOVNÍ SPOJENÍ"
      ForeColor       =   &H004F4536&
      Height          =   600
      Left            =   0
      TextAlign       =   2
      Top             =   0
      Width           =   7800
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
Attribute VB_Name = "frmPhraseManager"
Option Explicit
Private mPhrases As Collection

Private Sub UserForm_Initialize()
    ApplyVisualStyle
    ApplyBrandLogo imgLogo
    ReloadPhrases
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
    If x >= lstPhrases.left And x <= lstPhrases.left + lstPhrases.Width And _
       y >= lstPhrases.top And y <= lstPhrases.top + lstPhrases.Height Then
        ScrollListByWheel lstPhrases, delta
        HandleMouseWheel = True
    End If
Unsupported:
End Function

Private Sub ApplyVisualStyle()
    Me.BackColor = RGB(244, 241, 237)
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9
    With lblInfo
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(31, 41, 55)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 11
        .Font.Bold = True
    End With
    lstPhrases.BackColor = RGB(255, 255, 255)
    txtPhrase.BackColor = RGB(255, 255, 255)
    StyleManagerButton cmdAdd, True
    StyleManagerButton cmdDelete, False
    StyleManagerButton cmdImport, False
    StyleManagerButton cmdExport, False
    StyleManagerButton cmdClose, False
    With cmdHelp
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(14, 116, 180)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10
        .Font.Bold = True
        .ControlTipText = "Nápověda"
        .ZOrder 0
    End With
End Sub

Private Sub StyleManagerButton(ByVal button As Object, ByVal primary As Boolean)
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

Private Sub cmdExport_Click()
    Dim item As Variant, content As String, exportPath As Variant
    Dim stream As Object, exportedCount As Long, exportError As String

    For Each item In mPhrases
        If Not IsRequiredPhrase(CStr(item)) Then
            If Len(content) > 0 Then content = content & vbCrLf
            content = content & CStr(item)
            exportedCount = exportedCount + 1
        End If
    Next item

    If exportedCount = 0 Then
        MsgBox "Nejsou k dispozici žádná uživatelsky přidaná slovní spojení.", _
               vbInformation, TOOL_TITLE
        Exit Sub
    End If

    exportPath = Application.GetSaveAsFilename( _
        InitialFileName:="MR_Helper_slovni_spojeni.txt", _
        FileFilter:="Textové soubory (*.txt), *.txt", _
        Title:="Export slovních spojení")
    If VarType(exportPath) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(exportPath), 4)) <> ".txt" Then exportPath = CStr(exportPath) & ".txt"

    On Error GoTo ExportFailed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile CStr(exportPath), 2
    stream.Close
    MsgBox "Exportováno slovních spojení: " & CStr(exportedCount), vbInformation, TOOL_TITLE
    Exit Sub

ExportFailed:
    exportError = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    MsgBox "Export se nepodařil: " & exportError, vbExclamation, TOOL_TITLE
End Sub

Private Sub cmdImport_Click()
    Dim importPath As Variant, stream As Object, content As String
    Dim lines() As String, line As Variant, phrase As String
    Dim beforeCount As Long, importedCount As Long, importError As String

    importPath = Application.GetOpenFilename( _
        FileFilter:="Textové soubory (*.txt), *.txt", _
        Title:="Import slovních spojení")
    If VarType(importPath) = vbBoolean Then Exit Sub

    On Error GoTo ImportFailed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile CStr(importPath)
    content = stream.ReadText
    stream.Close

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)
    For Each line In lines
        phrase = Trim$(CStr(line))
        If Len(phrase) > 0 Then
            beforeCount = mPhrases.Count
            AddUnique mPhrases, phrase
            If mPhrases.Count > beforeCount Then importedCount = importedCount + 1
        End If
    Next line

    If importedCount > 0 Then SavePhrases mPhrases
    ReloadPhrases
    MsgBox "Importováno nových slovních spojení: " & CStr(importedCount), _
           vbInformation, TOOL_TITLE
    Exit Sub

ImportFailed:
    importError = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    MsgBox "Import se nepodařil: " & importError, vbExclamation, TOOL_TITLE
End Sub

Private Sub cmdHelp_Click()
    MsgBox "Přidání: Napište vše co je obsáhlé v konkrétní buňce (vyhněte se zbytečných mezer a znaků) a klikněte na Přidat." & vbCrLf & vbCrLf & _
           "Import: Načte jednu frázi z každého řádku; duplicity se přeskočí. Nejlépe načítejte pouze soubory, které byly exportovány." & vbCrLf & vbCrLf & _
           "Export: Uloží uživatelsky přidaná slovní spojení do .txt souboru.", _
           vbOKOnly + vbInformation, "Nápověda – slovní spojení"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
