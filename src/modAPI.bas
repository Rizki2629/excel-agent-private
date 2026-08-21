Attribute VB_Name = "modAPI"
'==============================================================================
' Excel Agent Add-in - API Communication Module
' Handles HTTP communication with API gateway
'==============================================================================

Option Explicit

'==============================================================================
' Constants
'==============================================================================
Private Const SYSTEM_PROMPT As String = "You are an expert Excel VBA developer. " & _
    "When user gives you a task, generate VBA code and return it in this JSON format:" & vbCrLf & _
    "{" & vbCrLf & _
    "  ""code"": ""VBA code here""," & vbCrLf & _
    "  ""explanation"": ""What the code does in Indonesian""," & vbCrLf & _
    "  ""action"": ""execute_vba|create_module|modify_code|run_macro""," & vbCrLf & _
    "  ""target"": ""active_workbook|specific_sheet_name""" & vbCrLf & _
    "}" & vbCrLf & vbCrLf & _
    "Rules:" & vbCrLf & _
    "- Always return valid JSON" & vbCrLf & _
    "- Code must be complete and ready to execute" & vbCrLf & _
    "- Explanation must be in Indonesian" & vbCrLf & _
    "- Choose the most appropriate action for the task" & vbCrLf & _
    "- If creating a module, use a descriptive name" & vbCrLf & _
    "- Handle errors gracefully in the VBA code"

'==============================================================================
' APIResponse - Structure to hold API response
'==============================================================================
Public Type APIResponse
    code As String
    explanation As String
    action As String
    target As String
    success As Boolean
    errorMessage As String
End Type

'==============================================================================
' SendToAPI - Send message to API and get response
'==============================================================================
Public Function SendToAPI(userMessage As String) As APIResponse
    On Error GoTo ErrorHandler

    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    Dim endpointURL As String
    endpointURL = GetEndpoint()

    Dim apiKey As String
    apiKey = GetAPIKey()

    Dim modelName As String
    modelName = GetModel()

    ' Build JSON payload
    Dim jsonPayload As String
    jsonPayload = BuildRequestPayload(userMessage, modelName)

    ' Send HTTP request
    With objHTTP
        .Open "POST", endpointURL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send jsonPayload

        If .Status = 200 Then
            ' Parse response
            SendToAPI = ParseAPIResponse(.responseText)
        Else
            SendToAPI.success = False
            SendToAPI.errorMessage = "API Error: " & .Status & " - " & .StatusText
        End If
    End With

    Set objHTTP = Nothing

    Exit Function

ErrorHandler:
    SendToAPI.success = False
    SendToAPI.errorMessage = "Connection Error: " & Err.Description
End Function

'==============================================================================
' BuildRequestPayload - Build JSON request payload
'==============================================================================
Private Function BuildRequestPayload(userMessage As String, modelName As String) As String
    Dim json As String

    json = "{" & vbCrLf
    json = json & "  ""model"": """ & EscapeJSON(modelName) & """," & vbCrLf
    json = json & "  ""messages"": [" & vbCrLf
    json = json & "    {" & vbCrLf
    json = json & "      ""role"": ""system""," & vbCrLf
    json = json & "      ""content"": """ & EscapeJSON(SYSTEM_PROMPT) & """" & vbCrLf
    json = json & "    }," & vbCrLf
    json = json & "    {" & vbCrLf
    json = json & "      ""role"": ""user""," & vbCrLf
    json = json & "      ""content"": """ & EscapeJSON(userMessage) & """" & vbCrLf
    json = json & "    }" & vbCrLf
    json = json & "  ]," & vbCrLf
    json = json & "  ""response_format"": {""type"": ""json_object""}" & vbCrLf
    json = json & "}"

    BuildRequestPayload = json
End Function

'==============================================================================
' ParseAPIResponse - Parse JSON response from API
'==============================================================================
Private Function ParseAPIResponse(jsonResponse As String) As APIResponse
    Dim result As APIResponse

    On Error GoTo ErrorHandler

    ' Extract fields from JSON
    result.code = ExtractJSONString(jsonResponse, "code")
    result.explanation = ExtractJSONString(jsonResponse, "explanation")
    result.action = ExtractJSONString(jsonResponse, "action")
    result.target = ExtractJSONString(jsonResponse, "target")

    ' Validate required fields
    If result.code = "" Then
        result.success = False
        result.errorMessage = "No VBA code in response"
    ElseIf result.action = "" Then
        result.success = False
        result.errorMessage = "No action specified in response"
    Else
        result.success = True
    End If

    ParseAPIResponse = result

    Exit Function

ErrorHandler:
    result.success = False
    result.errorMessage = "Parse Error: " & Err.Description
    ParseAPIResponse = result
End Function

'==============================================================================
' ExtractJSONString - Extract string value from JSON
'==============================================================================
Private Function ExtractJSONString(json As String, key As String) As String
    Dim searchStr As String
    Dim startPos As Long
    Dim endPos As Long

    searchStr = """" & key & """:"""
    startPos = InStr(1, json, searchStr)

    If startPos = 0 Then
        ExtractJSONString = ""
        Exit Function
    End If

    startPos = startPos + Len(searchStr)

    ' Handle escaped quotes
    endPos = startPos
    Do
        endPos = InStr(endPos, json, """")
        If endPos = 0 Then Exit Do

        ' Check if quote is escaped
        If endPos > 1 And Mid(json, endPos - 1, 1) = "\" Then
            endPos = endPos + 1
        Else
            Exit Do
        End If
    Loop

    If endPos = 0 Then
        ExtractJSONString = ""
        Exit Function
    End If

    Dim value As String
    value = Mid(json, startPos, endPos - startPos)

    ' Unescape JSON string
    value = UnescapeJSON(value)

    ExtractJSONString = value
End Function

'==============================================================================
' EscapeJSON - Escape special characters for JSON
'==============================================================================
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text

    ' Replace special characters
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    EscapeJSON = result
End Function

'==============================================================================
' UnescapeJSON - Unescape JSON string
'==============================================================================
Private Function UnescapeJSON(text As String) As String
    Dim result As String
    result = text

    ' Replace escaped characters
    result = Replace(result, "\n", vbCrLf)
    result = Replace(result, "\t", vbTab)
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")

    UnescapeJSON = result
End Function

'==============================================================================
' GetContextInfo - Get current Excel context for API
'==============================================================================
Public Function GetContextInfo() As String
    Dim context As String

    context = "Current Excel Context:" & vbCrLf
    context = context & "- Active Sheet: " & ActiveSheet.Name & vbCrLf
    context = context & "- Active Cell: " & ActiveCell.Address & vbCrLf

    If TypeName(Selection) = "Range" Then
        context = context & "- Selection: " & Selection.Address & vbCrLf
        context = context & "- Selection Value: " & CStr(Selection.Value) & vbCrLf
    End If

    context = context & "- Workbook: " & ActiveWorkbook.Name & vbCrLf

    GetContextInfo = context
End Function

'==============================================================================
' FormatAPIError - Format error message for display
'==============================================================================
Public Function FormatAPIError(errorMsg As String) As String
    FormatAPIError = "Error: " & errorMsg & vbCrLf & vbCrLf & _
                     "Pastikan:" & vbCrLf & _
                     "- Endpoint API sudah benar" & vbCrLf & _
                     "- API key sudah valid" & vbCrLf & _
                     "- Koneksi internet aktif"
End Function
