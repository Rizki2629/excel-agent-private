Attribute VB_Name = "modConfig"
'==============================================================================
' Excel Agent Add-in - Configuration Manager Module
' Handles configuration storage and API key encryption
' Uses VBA built-in SaveSetting/GetSetting (HKCU\Software\VB and VBA Program Settings)
'==============================================================================

Option Explicit

'==============================================================================
' Constants
'==============================================================================
Private Const APP_NAME As String = "ExcelAgent"
Private Const SECTION_NAME As String = "Config"
Private Const ENCRYPTION_KEY As String = "ExcelAgentSecretKey2026"

'==============================================================================
' SaveConfig - Save a configuration value to registry
'==============================================================================
Public Sub SaveConfig(key As String, value As String)
    On Error GoTo ErrorHandler
    SaveSetting APP_NAME, SECTION_NAME, key, value
    Exit Sub
ErrorHandler:
    ' Silently fail
End Sub

'==============================================================================
' GetConfig - Get a configuration value from registry
'==============================================================================
Public Function GetConfig(key As String, Optional defaultValue As String = "") As String
    On Error GoTo ErrorHandler
    Dim value As String
    value = GetSetting(APP_NAME, SECTION_NAME, key, defaultValue)
    GetConfig = value
    Exit Function
ErrorHandler:
    GetConfig = defaultValue
End Function

'==============================================================================
' DeleteConfig - Delete a configuration value from registry
'==============================================================================
Public Sub DeleteConfig(key As String)
    On Error Resume Next
    DeleteSetting APP_NAME, SECTION_NAME, key
    On Error GoTo 0
End Sub

'==============================================================================
' SaveAPIKey - Save API key with encryption
'==============================================================================
Public Sub SaveAPIKey(apiKey As String)
    Dim encrypted As String
    encrypted = EncryptString(apiKey, ENCRYPTION_KEY)
    SaveConfig "APIKey", encrypted
End Sub

'==============================================================================
' GetAPIKey - Get decrypted API key
'==============================================================================
Public Function GetAPIKey() As String
    Dim encrypted As String
    encrypted = GetConfig("APIKey", "")

    If encrypted <> "" Then
        GetAPIKey = DecryptString(encrypted, ENCRYPTION_KEY)
    Else
        GetAPIKey = ""
    End If
End Function

'==============================================================================
' SaveEndpoint - Save API endpoint URL
'==============================================================================
Public Sub SaveEndpoint(endpointURL As String)
    SaveConfig "APIEndpoint", endpointURL
End Sub

'==============================================================================
' GetEndpoint - Get API endpoint URL
'==============================================================================
Public Function GetEndpoint() As String
    GetEndpoint = GetConfig("APIEndpoint", "https://api.example.com/v1/chat/completions")
End Function

'==============================================================================
' SaveModel - Save model name
'==============================================================================
Public Sub SaveModel(modelName As String)
    SaveConfig "Model", modelName
End Sub

'==============================================================================
' GetModel - Get model name
'==============================================================================
Public Function GetModel() As String
    GetModel = GetConfig("Model", "gpt-4")
End Function

'==============================================================================
' IsAutoExecute - Check if auto-execute is enabled
'==============================================================================
Public Function IsAutoExecute() As Boolean
    IsAutoExecute = (GetConfig("AutoExecute", "true") = "true")
End Function

'==============================================================================
' SetAutoExecute - Set auto-execute mode
'==============================================================================
Public Sub SetAutoExecute(enabled As Boolean)
    If enabled Then
        SaveConfig "AutoExecute", "true"
    Else
        SaveConfig "AutoExecute", "false"
    End If
End Sub

'==============================================================================
' IsLogCommands - Check if command logging is enabled
'==============================================================================
Public Function IsLogCommands() As Boolean
    IsLogCommands = (GetConfig("LogCommands", "true") = "true")
End Function

'==============================================================================
' SetLogCommands - Set command logging mode
'==============================================================================
Public Sub SetLogCommands(enabled As Boolean)
    If enabled Then
        SaveConfig "LogCommands", "true"
    Else
        SaveConfig "LogCommands", "false"
    End If
End Sub

'==============================================================================
' GetLanguage - Get interface language
'==============================================================================
Public Function GetLanguage() As String
    GetLanguage = GetConfig("Language", "id")
End Function

'==============================================================================
' SetLanguage - Set interface language
'==============================================================================
Public Sub SetLanguage(lang As String)
    SaveConfig "Language", lang
End Sub

'==============================================================================
' TestConnection - Test API connection
'==============================================================================
Public Function TestConnection() As Boolean
    On Error GoTo ErrorHandler

    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    Dim endpointURL As String
    endpointURL = GetEndpoint()

    ' Simple GET request to test connection
    objHTTP.Open "GET", endpointURL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.Send

    TestConnection = (objHTTP.Status = 200 Or objHTTP.Status = 405)

    Set objHTTP = Nothing

    Exit Function

ErrorHandler:
    TestConnection = False
End Function

'==============================================================================
' ExportConfig - Export configuration to JSON file
'==============================================================================
Public Function ExportConfig(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim jsonConfig As String
    jsonConfig = "{" & vbCrLf
    jsonConfig = jsonConfig & "  ""endpoint"": """ & GetEndpoint() & """," & vbCrLf
    jsonConfig = jsonConfig & "  ""model"": """ & GetModel() & """," & vbCrLf
    jsonConfig = jsonConfig & "  ""autoExecute"": " & CStr(IsAutoExecute()) & "," & vbCrLf
    jsonConfig = jsonConfig & "  ""logCommands"": " & CStr(IsLogCommands()) & "," & vbCrLf
    jsonConfig = jsonConfig & "  ""language"": """ & GetLanguage() & """"
    jsonConfig = jsonConfig & "}"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim file As Object
    Set file = fso.CreateTextFile(filePath, True)
    file.Write jsonConfig
    file.Close

    ExportConfig = True

    Exit Function

ErrorHandler:
    ExportConfig = False
End Function

'==============================================================================
' ImportConfig - Import configuration from JSON file
'==============================================================================
Public Function ImportConfig(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        ImportConfig = False
        Exit Function
    End If

    Dim file As Object
    Set file = fso.OpenTextFile(filePath, 1)
    Dim jsonConfig As String
    jsonConfig = file.ReadAll
    file.Close

    ' Parse JSON (simple parser)
    Dim endpoint As String
    endpoint = ExtractJSONValue(jsonConfig, "endpoint")
    If endpoint <> "" Then SaveEndpoint endpoint

    Dim model As String
    model = ExtractJSONValue(jsonConfig, "model")
    If model <> "" Then SaveModel model

    Dim autoExec As String
    autoExec = ExtractJSONValue(jsonConfig, "autoExecute")
    If autoExec <> "" Then SetAutoExecute (autoExec = "true")

    Dim logCmds As String
    logCmds = ExtractJSONValue(jsonConfig, "logCommands")
    If logCmds <> "" Then SetLogCommands (logCmds = "true")

    Dim lang As String
    lang = ExtractJSONValue(jsonConfig, "language")
    If lang <> "" Then SetLanguage lang

    ImportConfig = True

    Exit Function

ErrorHandler:
    ImportConfig = False
End Function

'==============================================================================
' ExtractJSONValue - Simple JSON value extractor
'==============================================================================
Private Function ExtractJSONValue(json As String, key As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchStr As String

    searchStr = """" & key & """:"""
    startPos = InStr(1, json, searchStr)

    If startPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    startPos = startPos + Len(searchStr)
    endPos = InStr(startPos, json, """")

    If endPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    ExtractJSONValue = Mid(json, startPos, endPos - startPos)
End Function

'==============================================================================
' EncryptString - XOR encryption
'==============================================================================
Private Function EncryptString(text As String, key As String) As String
    Dim i As Long
    Dim result As String
    Dim textChar As Long
    Dim keyChar As Long

    For i = 1 To Len(text)
        textChar = Asc(Mid(text, i, 1))
        keyChar = Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))
        result = result & Chr(textChar Xor keyChar)
    Next i

    ' Convert to hex for safe storage
    Dim hexResult As String
    For i = 1 To Len(result)
        hexResult = hexResult & Right("0" & Hex(Asc(Mid(result, i, 1))), 2)
    Next i

    EncryptString = hexResult
End Function

'==============================================================================
' DecryptString - XOR decryption
'==============================================================================
Private Function DecryptString(hexText As String, key As String) As String
    ' Convert hex back to string
    Dim result As String
    Dim i As Long

    For i = 1 To Len(hexText) Step 2
        result = result & Chr(CLng("&H" & Mid(hexText, i, 2)))
    Next i

    ' XOR decrypt
    Dim decrypted As String
    Dim textChar As Long
    Dim keyChar As Long

    For i = 1 To Len(result)
        textChar = Asc(Mid(result, i, 1))
        keyChar = Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))
        decrypted = decrypted & Chr(textChar Xor keyChar)
    Next i

    DecryptString = decrypted
End Function
