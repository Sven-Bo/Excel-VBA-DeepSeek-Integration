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
    Public API_KEY As String                                   ' Replace with your DeepSeek API key (https://platform.deepseek.com)
    Private Const API_URL As String = "https://api.deepseek.com/v1/chat/completions"    ' DeepSeek API endpoint
    Public DEFAULT_MODEL As String
    Private Const DEFAULT_TEMPERATURE As Double = 0.7          ' Default temperature for creativity (0 to 2)
    Private Const DEFAULT_MAX_TOKENS As Long = 4000            ' Maximum number of tokens to generate (between 1 and 8192)
    Private Const DEFAULT_TOP_P As Double = 1                  ' Default top_p for nucleus sampling (0 to 1)
    Private Const DEFAULT_FREQUENCY_PENALTY As Double = 0      ' Default frequency penalty (-2 to 2)
    Private Const DEFAULT_PRESENCE_PENALTY As Double = 0       ' Default presence penalty (-2 to 2)
    
    ' HTTP Object Configuration
    Private Const USE_SERVER_XMLHTTP As Boolean = False        ' Set to True to use MSXML2.ServerXMLHTTP (for server environments)
    
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
    
Function DS_Chat(prompt As String, _
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
    
    Dim model As String
    model = DEFAULT_MODEL
    
    ' Create HTTP object based on configuration
    Dim http As Object
    '    If USE_SERVER_XMLHTTP Then
    Set http = CreateObject("MSXML2.ServerXMLHTTP")            ' Use this for server environments
    '    Else
    '        Set http = CreateObject("MSXML2.XMLHTTP") ' Use this for desktop environments
    '    End If
    
    ' Get API key from named range or use the constant
    Dim apiKey As String
    On Error Resume Next
    apiKey = ThisWorkbook.Names("DS_API_KEY").RefersToRange.Value
    On Error GoTo ErrorHandler
    If apiKey = "" Then
        apiKey = API_KEY                                       ' Fallback to the constant if named range is missing or empty
    End If
    
    ' Prepare the request body
    Dim requestBody As String
    requestBody = "{""model"": """ & model & """, ""messages"": [{""role"": ""user"", ""content"": """ & prompt & """}], " & _
            """temperature"": " & temperature & ", ""max_tokens"": " & max_tokens & ", ""top_p"": " & top_p & ", " & _
            """frequency_penalty"": " & frequency_penalty & ", ""presence_penalty"": " & presence_penalty & "}"
    
    ' Send the request
    http.Open "POST", API_URL, True
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send requestBody
    http.settimeouts 30000, 30000, 30000, 30000                '域名解析，连接服务器，发送数据，接收响应
    
    Do While http.readyState <> 4
        DoEvents
    Loop
    
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
