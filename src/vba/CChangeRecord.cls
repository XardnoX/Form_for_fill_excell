VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CChangeRecord"
Option Explicit

Public SheetName As String
Public CellAddress As String
Public Phrase As String
Public OldValue As Variant
Public OldFormula As Variant
Public OldHasFormula As Boolean
Public NewValue As Variant
Public OldPattern As Long
Public OldColor As Long
Public OldNumberFormat As String

Public Function Summary() As String
    Summary = SheetName & " | " & CellAddress & " | " & Phrase & " | " & _
              SafeText(OldValue) & " -> " & SafeText(NewValue)
End Function
