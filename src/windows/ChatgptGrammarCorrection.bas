Attribute VB_Name = "ChatgptGrammarCorrection"
Option Explicit

Sub GrammarCorrection()
  ' Grammar Correction Macro - Version 0.0.1

  ' Check if a selection is made
  If Selection.Type = wdSelectionIP Then
    Exit Sub
  End If

  ' Check if selected text is empty
  If Trim(Selection.Text) = "" Then
    Exit Sub
  End If

  ' Prompt user for name to be displayed in tracked changes
  Dim userName As String
  userName = InputBox("Enter your name for tracked changes:", "User Name", Application.userName)

  Dim strAPIKey As String
  Dim strURL As String
  Dim strModel As String
  Dim intMaxTokens As Integer
  Dim strResponse As String
  Dim objCurlHttp As Object
  Dim strJSONdata As String

  ' API configuration
  strAPIKey = Environ("OPENAI_API_KEY")
  strURL = "https://api.openai.com/v1/completions"
  strModel = "text-davinci-003"
  strURL = "https://api.openai.com/v1/chat/completions"
  strModel = "gpt-3.5-turbo"
  intMaxTokens = 2048

  ' Prepare prompt and JSON data
  Dim strPrompt As String

  strPrompt = Replace(Selection.Text, ChrW$(13), "")
  strPrompt = "Revise the text by correct only grammar mistakes. Do not change anything else. --- " & strPrompt
  strJSONdata = "{""model"":""" & strModel & """,""prompt"":""" & strPrompt & """,""max_tokens"":" & intMaxTokens & "}"

  Set objCurlHttp = CreateObject("MSXML2.serverXMLHTTP")

  With objCurlHttp
    .Open "POST", strURL, False
    .SetRequestHeader "Content-type", "application/json"
    .SetRequestHeader "Authorization", "Bearer " & strAPIKey
    .Send (strJSONdata)

    ' Check response status
    Dim strStatus As Integer
    strStatus = .Status

    If strStatus <> 200 Then
      MsgBox Prompt:="The OpenAI servers have experienced an error while processing your request! Please try again shortly, or for continued downtime please check the Chat status at: https://status.openai.com/"
      Exit Sub
    End If

    strResponse = .ResponseText

    ' Check if error occurred in the response
    If Mid(strResponse, 8, 5) = "error" Then
      MsgBox Prompt:="The server had an error while processing your request. Sorry about that! Please try again shortly."
      Exit Sub
    End If

    ' Extract the corrected text from the response
    Dim intStartPos As Integer
    intStartPos = InStr(1, strResponse, Chr(34) & "text" & Chr(34)) + 12

    If intStartPos = 12 Then
      MsgBox Prompt:="ChatGPT is at capacity right now. Please wait a minute and try again."
      Exit Sub
    End If

    Dim intEndPos As Integer
    intEndPos = InStr(1, strResponse, Chr(34) & "index" & Chr(34)) - 2

    Dim intLength As Integer
    intLength = intEndPos - intStartPos

    Dim strOutput As String
    strOutput = Mid(strResponse, intStartPos, intLength)

    Dim strOutputFormatted As String, strOutputFormatted1 As String, strOutputFormatted2 As String
    strOutputFormatted1 = Replace(strOutput, "\n\n", vbCrLf)
    strOutputFormatted2 = Replace(strOutputFormatted1, "\n", vbCrLf)
    strOutputFormatted = strOutputFormatted2

    ' Replace selected text with corrected text
    Dim rngSelection As Range
    Set rngSelection = Selection.Range
    rngSelection.Text = strOutputFormatted


  End With

  Set objCurlHttp = Nothing
End Sub

