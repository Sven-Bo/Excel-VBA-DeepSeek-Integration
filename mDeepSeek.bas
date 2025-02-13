Attribute VB_Name = "mDeepSeek"
Option Explicit

' ========================================================
' DeepSeek API Integration for Excel
' Author: Sven Bosau
' Website: https://pythonandvba.com
' YouTube: https://youtube.com/@CodingIsFun
' Version: 1.0
' Description: This module provides a custom Excel function
'              to interact with the DeepSeek API for chat completions.
' Requirements: JSON Converter Module (Required for parsing JSON responses)
'               Import the JsonConverter.bas module into your VBA project.
'               Download from: https://github.com/VBA-tools/VBA-JSON
' ========================================================

' API Configuration
Private Const API_KEY As String = "YOUR_API_KEY" ' Replace with your DeepSeek API key (https://platform.deepseek.com)
Private Const API_URL As String = "https://api.deepseek.com/v1/chat/completions" ' DeepSeek API endpoint
Private Const DEFAULT_MODEL As String = "deepseek-chat" ' Default model to use
Private Const DEFAULT_TEMPERATURE As Double = 0.7 ' Default temperature for creativity (0 to 2)
Private Const DEFAULT_MAX_TOKENS As Long = 300 ' Maximum number of tokens to generate (between 1 and 8192)
Private Const DEFAULT_TOP_P As Double = 1 ' Default top_p for nucleus sampling (0 to 1)
Private Const DEFAULT_FREQUENCY_PENALTY As Double = 0 ' Default frequency penalty (-2 to 2)
Private Const DEFAULT_PRESENCE_PENALTY As Double = 0 ' Default presence penalty (-2 to 2)

' HTTP Object Configuration
Private Const USE_SERVER_XMLHTTP As Boolean = False ' Set to True to use MSXML2.ServerXMLHTTP (for server environments)

' Available Models:
' - deepseek-chat: General-purpose chat model
' - deepseek-reasoner: Model optimized for reasoning tasks

' Get Your API Key:
' Visit https://platform.deepseek.com to sign up and get your API key.

' ========================================================
' Custom Excel Function: DS_Chat
' Description: Sends a prompt to the DeepSeek API and returns the response.
' Parameters:
'   - prompt: The input text/prompt for the API.
'   - model (Optional): The model to use (default is "deepseek-chat").
'   - temperature (Optional): Controls randomness (0 = deterministic, 2 = creative).
'   - max_tokens (Optional): Maximum number of tokens to generate (default is 4096).
'   - top_p (Optional): Nucleus sampling parameter (0 to 1, default is 1).
'   - frequency_penalty (Optional): Penalizes repeated tokens (-2 to 2, default is 0).
'   - presence_penalty (Optional): Encourages new topics (-2 to 2, default is 0).
' Returns: The API response as a string.
' ========================================================

' ========================================================
' Usage Example:
' Below are examples of how to use the DS_Chat function in Excel:
'
' 1. Basic Usage:
'    =DS_Chat("What is the capital of France?")
'    Output: "The capital of France is Paris."
'
' 2. Advanced Usage with Custom Parameters:
'    =DS_Chat("Explain quantum computing in simple terms.", "deepseek-chat", 0.8, 512, 0.9, 0.5, 0.5)
'    Output: "Quantum computing is a type of computing that uses quantum bits (qubits) to perform calculations..."
'
' 3. Using a Named Range for API Key:
'    - Create a named range "DS_API_KEY" in your workbook and enter your API key in the corresponding cell.
'    - If the named range is missing or empty, the function will fall back to the constant API_KEY.
'
' Notes:
' - Ensure the "Microsoft XML, v6.0" & "Microsoft Scripting Runtime" library is enabled in your VBA project references.
' - Import the JsonConverter.bas module for JSON parsing.
' ========================================================

Function DS_Chat(prompt As String, Optional model As String = DEFAULT_MODEL, _
                 Optional temperature As Double = DEFAULT_TEMPERATURE, Optional max_tokens As Long = DEFAULT_MAX_TOKENS, _
                 Optional top_p As Double = DEFAULT_TOP_P, Optional frequency_penalty As Double = DEFAULT_FREQUENCY_PENALTY, _
                 Optional presence_penalty As Double = DEFAULT_PRESENCE_PENALTY) As String
    On Error GoTo ErrorHandler
    
    ' Validate input
    If prompt = "" Then
        DS_Chat = "Error: Prompt cannot be empty."
        Exit Function
    End If
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Calling DeepSeek API..."
    
    ' Create HTTP object based on configuration
    Dim http As Object
    If USE_SERVER_XMLHTTP Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP") ' Use this for server environments
    Else
        Set http = CreateObject("MSXML2.XMLHTTP") ' Use this for desktop environments
    End If
    
    ' Get API key from named range or use the constant
    Dim apiKey As String
    On Error Resume Next
    apiKey = ActiveWorkbook.Names("DS_API_KEY").RefersToRange.Value
    On Error GoTo ErrorHandler
    If apiKey = "" Then
        apiKey = API_KEY ' Fallback to the constant if named range is missing or empty
    End If
    
    ' Clean prompt
    prompt = JsonEscape(prompt)
    
    ' Prepare the request body
    Dim requestBody As String
    requestBody = "{""model"": """ & model & """, ""messages"": [{""role"": ""user"", ""content"": """ & prompt & """}], " & _
                  """temperature"": " & temperature & ", ""max_tokens"": " & max_tokens & ", ""top_p"": " & top_p & ", " & _
                  """frequency_penalty"": " & frequency_penalty & ", ""presence_penalty"": " & presence_penalty & "}"
    
    ' Send the request
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send requestBody
    
    ' Check the response status
    If http.Status = 200 Then
        ' Parse the response
        Dim response As String
        response = http.responseText
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)
        
        ' Extract the completion text
        DS_Chat = json("choices")(1)("message")("content")
    Else
        ' Handle API errors
        DS_Chat = "API Error: " & http.Status & " - " & http.statusText
    End If
    
    ' Clear status bar and restore settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Function
    
ErrorHandler:
    ' Handle VBA errors
    DS_Chat = "VBA Error: " & Err.Description
    ' Clear status bar and restore settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Function

'Source: https://stackoverflow.com/a/50854612
Private Function JsonEscape(sText As String) As String

    Const STR_CODES As String = "\u0000|\u0001|\u0002|\u0003|\u0004|\u0005|\u0006|\u0007|\b|\t|\n|\u000B|\f|\r|\u000E|\u000F|\u0010|\u0011|\u0012|\u0013|\u0014|\u0015|\u0016|\u0017|\u0018|\u0019|\u001A|\u001B|\u001C|\u001D|\u001E|\u001F"
    Static vTranscode As Variant
    Dim lIdx As Long
    Dim lAsc As Long

    If IsEmpty(vTranscode) Then
        vTranscode = Split(STR_CODES, "|")
    End If
    
    For lIdx = 1 To Len(sText)
        lAsc = AscW(Mid$(sText, lIdx, 1))
        If lAsc = 92 Or lAsc = 34 Then '--- \ and "
            JsonEscape = JsonEscape & "\" & ChrW$(lAsc)
        ElseIf lAsc >= 32 And lAsc < 256 Then
            JsonEscape = JsonEscape & ChrW$(lAsc)
        ElseIf lAsc >= 0 And lAsc < 32 Then
            JsonEscape = JsonEscape & vTranscode(lAsc)
        ElseIf Asc(Mid$(sText, lIdx, 1)) <> 63 Or Mid$(sText, lIdx, 1) = "?" Then '--- ?
            JsonEscape = JsonEscape & ChrW$(AscW(Mid$(sText, lIdx, 1)))
        Else
            JsonEscape = JsonEscape & "\u" & Right$("0000" & Hex$(lAsc), 4)
        End If
    Next
    
    ' Replace actual newline characters with \n
    JsonEscape = Replace(JsonEscape, vbCrLf, "\n")
    JsonEscape = Replace(JsonEscape, vbLf, "\n")

End Function
