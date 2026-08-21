Attribute VB_Name = "modJSON"
'==============================================================================
' Excel Agent Add-in - JSON Parser Module
' Simple JSON parser for API communication
'==============================================================================

Option Explicit

'==============================================================================
' JSONValue - Structure to hold parsed JSON values
'==============================================================================
Public Type JSONValue
    valueType As String  ' "string", "number", "boolean", "null", "object", "array"
    stringValue As String
    numberValue As Double
    booleanValue As Boolean
    children() As JSONValue
    keys() As String
    childCount As Long
End Type

'==============================================================================
' ParseJSON - Parse JSON string into JSONValue structure
'==============================================================================
Public Function ParseJSON(jsonStr As String) As JSONValue
    Dim pos As Long
    pos = 1

    ParseJSON = ParseValue(jsonStr, pos)
End Function

'==============================================================================
' ParseValue - Parse a JSON value
'==============================================================================
Private Function ParseValue(jsonStr As String, pos As Long) As JSONValue
    SkipWhitespace jsonStr, pos

    If pos > Len(jsonStr) Then
        ParseValue.valueType = "null"
        Exit Function
    End If

    Dim ch As String
    ch = Mid(jsonStr, pos, 1)

    Select Case ch
        Case """"
            ParseValue = ParseString(jsonStr, pos)
        Case "{"
            ParseValue = ParseObject(jsonStr, pos)
        Case "["
            ParseValue = ParseArray(jsonStr, pos)
        Case "t", "f"
            ParseValue = ParseBoolean(jsonStr, pos)
        Case "n"
            ParseValue = ParseNull(jsonStr, pos)
        Case Else
            If IsNumeric(ch) Or ch = "-" Then
                ParseValue = ParseNumber(jsonStr, pos)
            Else
                ParseValue.valueType = "null"
                pos = pos + 1
            End If
    End Select
End Function

'==============================================================================
' ParseString - Parse a JSON string
'==============================================================================
Private Function ParseString(jsonStr As String, pos As Long) As JSONValue
    pos = pos + 1  ' Skip opening quote

    Dim result As String
    result = ""

    Do While pos <= Len(jsonStr)
        Dim ch As String
        ch = Mid(jsonStr, pos, 1)

        If ch = """" Then
            pos = pos + 1  ' Skip closing quote
            Exit Do
        ElseIf ch = "\" Then
            pos = pos + 1
            ch = Mid(jsonStr, pos, 1)
            Select Case ch
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "/": result = result & "/"
                Case "b": result = result & vbBack
                Case "f": result = result & vbFormFeed
                Case "n": result = result & vbCrLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    Dim hex As String
                    hex = Mid(jsonStr, pos + 1, 4)
                    result = result & Chr(CLng("&H" & hex))
                    pos = pos + 4
                Case Else
                    result = result & ch
            End Select
        Else
            result = result & ch
        End If

        pos = pos + 1
    Loop

    ParseString.valueType = "string"
    ParseString.stringValue = result
End Function

'==============================================================================
' ParseNumber - Parse a JSON number
'==============================================================================
Private Function ParseNumber(jsonStr As String, pos As Long) As JSONValue
    Dim startPos As Long
    startPos = pos

    ' Handle negative sign
    If Mid(jsonStr, pos, 1) = "-" Then
        pos = pos + 1
    End If

    ' Integer part
    Do While pos <= Len(jsonStr) And IsNumeric(Mid(jsonStr, pos, 1))
        pos = pos + 1
    Loop

    ' Decimal part
    If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "." Then
        pos = pos + 1
        Do While pos <= Len(jsonStr) And IsNumeric(Mid(jsonStr, pos, 1))
            pos = pos + 1
        Loop
    End If

    ' Exponent part
    If pos <= Len(jsonStr) And (Mid(jsonStr, pos, 1) = "e" Or Mid(jsonStr, pos, 1) = "E") Then
        pos = pos + 1
        If pos <= Len(jsonStr) And (Mid(jsonStr, pos, 1) = "+" Or Mid(jsonStr, pos, 1) = "-") Then
            pos = pos + 1
        End If
        Do While pos <= Len(jsonStr) And IsNumeric(Mid(jsonStr, pos, 1))
            pos = pos + 1
        Loop
    End If

    Dim numStr As String
    numStr = Mid(jsonStr, startPos, pos - startPos)

    ParseNumber.valueType = "number"
    ParseNumber.numberValue = CDbl(numStr)
End Function

'==============================================================================
' ParseBoolean - Parse a JSON boolean
'==============================================================================
Private Function ParseBoolean(jsonStr As String, pos As Long) As JSONValue
    Dim remaining As String
    remaining = Mid(jsonStr, pos)

    If Left(remaining, 4) = "true" Then
        ParseBoolean.valueType = "boolean"
        ParseBoolean.booleanValue = True
        pos = pos + 4
    ElseIf Left(remaining, 5) = "false" Then
        ParseBoolean.valueType = "boolean"
        ParseBoolean.booleanValue = False
        pos = pos + 5
    Else
        ParseBoolean.valueType = "null"
    End If
End Function

'==============================================================================
' ParseNull - Parse a JSON null
'==============================================================================
Private Function ParseNull(jsonStr As String, pos As Long) As JSONValue
    Dim remaining As String
    remaining = Mid(jsonStr, pos)

    If Left(remaining, 4) = "null" Then
        ParseNull.valueType = "null"
        pos = pos + 4
    Else
        ParseNull.valueType = "null"
        pos = pos + 1
    End If
End Function

'==============================================================================
' ParseObject - Parse a JSON object
'==============================================================================
Private Function ParseObject(jsonStr As String, pos As Long) As JSONValue
    pos = pos + 1  ' Skip opening brace

    Dim obj As JSONValue
    obj.valueType = "object"
    obj.childCount = 0

    SkipWhitespace jsonStr, pos

    If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "}" Then
        pos = pos + 1
        ParseObject = obj
        Exit Function
    End If

    Do While pos <= Len(jsonStr)
        SkipWhitespace jsonStr, pos

        ' Parse key
        Dim key As JSONValue
        key = ParseValue(jsonStr, pos)

        SkipWhitespace jsonStr, pos

        ' Skip colon
        If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = ":" Then
            pos = pos + 1
        End If

        ' Parse value
        Dim value As JSONValue
        value = ParseValue(jsonStr, pos)

        ' Add to object
        obj.childCount = obj.childCount + 1
        ReDim Preserve obj.children(1 To obj.childCount)
        ReDim Preserve obj.keys(1 To obj.childCount)
        obj.children(obj.childCount) = value
        obj.keys(obj.childCount) = key.stringValue

        SkipWhitespace jsonStr, pos

        ' Check for comma or end
        If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "," Then
            pos = pos + 1
        ElseIf pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        End If
    Loop

    ParseObject = obj
End Function

'==============================================================================
' ParseArray - Parse a JSON array
'==============================================================================
Private Function ParseArray(jsonStr As String, pos As Long) As JSONValue
    pos = pos + 1  ' Skip opening bracket

    Dim arr As JSONValue
    arr.valueType = "array"
    arr.childCount = 0

    SkipWhitespace jsonStr, pos

    If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "]" Then
        pos = pos + 1
        ParseArray = arr
        Exit Function
    End If

    Do While pos <= Len(jsonStr)
        SkipWhitespace jsonStr, pos

        ' Parse value
        Dim value As JSONValue
        value = ParseValue(jsonStr, pos)

        ' Add to array
        arr.childCount = arr.childCount + 1
        ReDim Preserve arr.children(1 To arr.childCount)
        arr.children(arr.childCount) = value

        SkipWhitespace jsonStr, pos

        ' Check for comma or end
        If pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "," Then
            pos = pos + 1
        ElseIf pos <= Len(jsonStr) And Mid(jsonStr, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        End If
    Loop

    ParseArray = arr
End Function

'==============================================================================
' SkipWhitespace - Skip whitespace characters
'==============================================================================
Private Sub SkipWhitespace(jsonStr As String, pos As Long)
    Do While pos <= Len(jsonStr)
        Dim ch As String
        ch = Mid(jsonStr, pos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

'==============================================================================
' GetJSONValue - Get value by key from parsed JSON
'==============================================================================
Public Function GetJSONValue(json As JSONValue, key As String) As JSONValue
    Dim i As Long

    If json.valueType = "object" Then
        For i = 1 To json.childCount
            If json.keys(i) = key Then
                GetJSONValue = json.children(i)
                Exit Function
            End If
        Next i
    End If

    GetJSONValue.valueType = "null"
End Function

'==============================================================================
' GetJSONString - Get string value by key
'==============================================================================
Public Function GetJSONString(json As JSONValue, key As String) As String
    Dim val As JSONValue
    val = GetJSONValue(json, key)

    If val.valueType = "string" Then
        GetJSONString = val.stringValue
    Else
        GetJSONString = ""
    End If
End Function

'==============================================================================
' GetJSONNumber - Get number value by key
'==============================================================================
Public Function GetJSONNumber(json As JSONValue, key As String) As Double
    Dim val As JSONValue
    val = GetJSONValue(json, key)

    If val.valueType = "number" Then
        GetJSONNumber = val.numberValue
    Else
        GetJSONNumber = 0
    End If
End Function

'==============================================================================
' GetJSONBoolean - Get boolean value by key
'==============================================================================
Public Function GetJSONBoolean(json As JSONValue, key As String) As Boolean
    Dim val As JSONValue
    val = GetJSONValue(json, key)

    If val.valueType = "boolean" Then
        GetJSONBoolean = val.booleanValue
    Else
        GetJSONBoolean = False
    End If
End Function

'==============================================================================
' ToJSONString - Convert JSONValue back to JSON string
'==============================================================================
Public Function ToJSONString(json As JSONValue) As String
    Dim result As String
    Dim i As Long

    Select Case json.valueType
        Case "string"
            result = """" & EscapeJSONStr(json.stringValue) & """"

        Case "number"
            result = CStr(json.numberValue)

        Case "boolean"
            If json.booleanValue Then
                result = "true"
            Else
                result = "false"
            End If

        Case "null"
            result = "null"

        Case "object"
            result = "{"
            For i = 1 To json.childCount
                If i > 1 Then result = result & ","
                result = result & """" & json.keys(i) & """:" & ToJSONString(json.children(i))
            Next i
            result = result & "}"

        Case "array"
            result = "["
            For i = 1 To json.childCount
                If i > 1 Then result = result & ","
                result = result & ToJSONString(json.children(i))
            Next i
            result = result & "]"
    End Select

    ToJSONString = result
End Function

'==============================================================================
' EscapeJSONStr - Escape string for JSON
'==============================================================================
Private Function EscapeJSONStr(text As String) As String
    Dim result As String
    result = text

    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    EscapeJSONStr = result
End Function
