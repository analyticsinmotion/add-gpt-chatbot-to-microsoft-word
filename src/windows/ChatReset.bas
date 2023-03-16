Attribute VB_Name = "ChatReset"
'
'
' GPT ChatBot Reset - Version 0.1.0
'
'
Option Explicit

Sub ChatReset()

  Dim strDocumentVariables As String
  Dim objDocVar As Object
  For Each objDocVar In ActiveDocument.Variables
    strDocumentVariables = strDocumentVariables + objDocVar.Name + " "
  Next objDocVar
  

  Dim ChatHistoryExists As Integer
  ChatHistoryExists = InStr(1, strDocumentVariables, "ChatHistory")
  If ChatHistoryExists = 0 Then
    MsgBox Prompt:="There is no message history to reset!"
    Exit Sub
  End If
  
  ActiveDocument.Variables("ChatHistory").Delete
  
  Dim strChatResetMessage As String
    strChatResetMessage = "  Your previous conversation history has been removed from the chatbot's memory.  "
    
    With Selection
      .InsertAfter vbCr & strChatResetMessage
      .Font.Name = "Courier New"
      .Font.Size = 9
      .Font.ColorIndex = wdWhite
      .Range.HighlightColorIndex = wdViolet
      .Paragraphs.Alignment = wdAlignParagraphJustify
      .InsertAfter vbCr
      .Collapse Direction:=wdCollapseEnd
    End With
    
  
End Sub
