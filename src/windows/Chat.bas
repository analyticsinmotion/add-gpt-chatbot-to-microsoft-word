Attribute VB_Name = "Chat"
'
'
' GPT ChatBot - Version 0.1.1
'
'
Option Explicit

Sub Chat()

  If Selection.Type = wdSelectionIP Then
    Exit Sub
  End If

  If Selection.Text = ChrW$(13) Then
    Exit Sub
  End If
 


  Dim strDocumentVariables As String
  Dim objDocVar As Object
  For Each objDocVar In ActiveDocument.Variables
    strDocumentVariables = strDocumentVariables + objDocVar.Name + " "
  Next objDocVar






  Dim ChatHistoryExists As Integer
  ChatHistoryExists = InStr(1, strDocumentVariables, "ChatHistory")
  If ChatHistoryExists = 0 Then
    ActiveDocument.Variables.Add Name:="ChatHistory", Value:="{""role"": ""system"", ""content"": ""You are a helpful assistant.""}"
  End If



  Dim strChatHistory As String
  strChatHistory = ActiveDocument.Variables("ChatHistory").Value




  Dim strAPIKey As String
  strAPIKey = Environ("OPENAI_API_KEY")
  
  Dim strURL As String
  strURL = "https://api.openai.com/v1/chat/completions"
  
  Dim strModel As String
  strModel = "gpt-3.5-turbo"
  
  Dim intMaxTokens As Integer
  intMaxTokens = 3584
  
  Dim strPrompt As String
  strPrompt = Replace(Selection, ChrW$(13), "")
  Dim strFormattedPrompt As String
  strFormattedPrompt = "{""role"": ""user"", ""content"": """ & strPrompt & """}"

  
  Dim strMessage As String
  strMessage = strChatHistory & "," & strFormattedPrompt

  
  Dim strFormattedMessage As String
  strFormattedMessage = "[" & strMessage & "]"

  
  Dim strJSONdata As String
  strJSONdata = "{""model"": """ & strModel & """, ""messages"": " & strFormattedMessage & "}"

  
  

  Dim objCurlHttp As Object
  Set objCurlHttp = CreateObject("MSXML2.serverXMLHTTP")

  With objCurlHttp
    .Open "POST", strURL, False
    .SetRequestHeader "Content-type", "application/json"
    .SetRequestHeader "Authorization", "Bearer " + strAPIKey
    .Send (strJSONdata)
    
    Dim strStatus As Integer
    strStatus = .Status
    Dim strStatusText As String
    strStatusText = .StatusText
    
    If strStatus <> 200 Then
      MsgBox Prompt:="The OpenAI servers have experienced an error while processing your request! Please try again shortly, or for continued downtime please check the Chat status at: https://status.openai.com/"
      Exit Sub
    End If
    
    Dim strResponse As String
    strResponse = .ResponseText
      

    If Mid(strResponse, 8, 5) = "error" Then
      MsgBox Prompt:="The ChatGPT model is currently overloaded with other requests. You can retry your request, or contact us through our help center at help.openai.com if the error persists."
      Exit Sub
    End If
    
    
    Dim intStartPos As Integer
    intStartPos = InStr(1, strResponse, Chr(34) & "content" & Chr(34)) + 11

    

    If intStartPos = 11 Then
      MsgBox Prompt:="ChatGPT is at capacity right now. Please wait a minute and try again."
      Exit Sub
    End If
    
    Dim intEndPos As Integer
    intEndPos = InStr(1, strResponse, Chr(34) & "finish_reason" & Chr(34)) - 3

    
    Dim intLength As Integer
    intLength = intEndPos - intStartPos

    
    Dim strOutput As String
    strOutput = Mid(strResponse, intStartPos, intLength)

    
    Dim strOutputFormatted As String, strOutputFormatted1 As String, strOutputFormatted2 As String
    strOutputFormatted1 = Replace(strOutput, "\n\n", vbCrLf)
    strOutputFormatted2 = Replace(strOutputFormatted1, "\n", vbCrLf)
    strOutputFormatted = strOutputFormatted2
    
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertAfter vbCr & strOutputFormatted
    Selection.Font.Name = "Courier New"
    Selection.Font.Size = 9
    Selection.Font.ColorIndex = wdViolet
    Selection.Paragraphs.Alignment = wdAlignParagraphJustify
    Selection.InsertAfter vbCr
    Selection.Collapse Direction:=wdCollapseEnd


  End With
  
  Set objCurlHttp = Nothing
 
 
  Dim strAssistantResponse As String
  strAssistantResponse = "{""role"": ""assistant"", ""content"": """ & strOutput & """}"
  
  Dim strMessageExtended As String
  strMessageExtended = strMessage & "," & strAssistantResponse
  

  ActiveDocument.Variables("ChatHistory").Delete
  ActiveDocument.Variables.Add Name:="ChatHistory", Value:=strMessageExtended
  
  Dim strChatHistory2 As String
  strChatHistory2 = ActiveDocument.Variables("ChatHistory").Value


End Sub
