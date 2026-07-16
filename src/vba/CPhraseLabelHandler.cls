VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CPhraseLabelHandler"
Option Explicit

Private WithEvents mLabel As MSForms.Label
Private mOwner As Object
Private mPhrase As String

Public Sub Initialize(ByVal labelControl As MSForms.Label, _
                      ByVal owner As Object, ByVal phrase As String)
    Set mLabel = labelControl
    Set mOwner = owner
    mPhrase = phrase
End Sub

Private Sub mLabel_Click()
    mOwner.TogglePhraseOccurrences mPhrase
End Sub
