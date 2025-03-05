Attribute VB_Name = "TestTrackChanges"
Option Explicit

Sub TestTrackChanges()
    Dim strOutput As String
    Dim userName As String
    userName = InputBox("Enter your name for tracked changes:", "User Name", Application.UserName)

    strOutput = "This is a another test." & vbNewLine & "Another line here."
      
    ' Replace selected text with corrected text
    Dim rngSelection As Range
    Set rngSelection = Selection.Range
    rngSelection.Text = strOutput

    ' Apply tracked changes
    With rngSelection.Revisions
      .Add userName, Date, wdReplaceAll
      .AcceptAll
    End With    

End Sub