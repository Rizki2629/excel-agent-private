Attribute VB_Name = "modChatHistory"
'==============================================================================
' Excel Agent Add-in - Chat History Module
' Handles chat history logging and retrieval
'==============================================================================

Option Explicit

'==============================================================================
' LogChat - Log a chat entry to history sheet
'==============================================================================
Public Sub LogChat(userMessage As String, agentResponse As String, _
                   Optional code As String = "", Optional status As String = "SUCCESS")
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create chat history sheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "ChatHistory"
        ws.Visible = xlSheetVeryHidden

        ' Add headers
        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Sender"
        ws.Range("C1").Value = "Message"
        ws.Range("D1").Value = "Status"
        ws.Range("E1").Value = "Code"

        ' Format header
        With ws.Range("A1:E1")
            .Font.Bold = True
            .Interior.Color = RGB(52, 152, 219)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Log user message
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "User"
    ws.Cells(nextRow, 3).Value = userMessage
    ws.Cells(nextRow, 4).Value = status

    ' Log agent response
    nextRow = nextRow + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "Agent"
    ws.Cells(nextRow, 3).Value = agentResponse
    ws.Cells(nextRow, 4).Value = status
    ws.Cells(nextRow, 5).Value = code

    ' Keep only last 500 entries (1000 rows for user+agent)
    If nextRow > 1000 Then
        ws.Rows("2:" & (nextRow - 500)).Delete
    End If
End Sub

'==============================================================================
' GetChatHistory - Get chat history as string
'==============================================================================
Public Function GetChatHistory(Optional lastN As Long = 20) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        GetChatHistory = "No chat history"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        GetChatHistory = "No chat history"
        Exit Function
    End If

    Dim startRow As Long
    startRow = Application.Max(2, lastRow - (lastN * 2) + 1)

    Dim result As String
    Dim i As Long

    For i = startRow To lastRow
        Dim timestamp As String
        Dim sender As String
        Dim message As String

        timestamp = Format(ws.Cells(i, 1).Value, "hh:mm:ss")
        sender = ws.Cells(i, 2).Value
        message = ws.Cells(i, 3).Value

        If sender <> "[SYSTEM]" Then
            result = result & "[" & timestamp & "] " & sender & ": " & message & vbCrLf
        End If
    Next i

    GetChatHistory = result
End Function

'==============================================================================
' GetLastUserMessage - Get last user message
'==============================================================================
Public Function GetLastUserMessage() As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        GetLastUserMessage = ""
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Search backwards for last user message
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 2).Value = "User" Then
            GetLastUserMessage = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i

    GetLastUserMessage = ""
End Function

'==============================================================================
' GetLastAgentResponse - Get last agent response
'==============================================================================
Public Function GetLastAgentResponse() As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        GetLastAgentResponse = ""
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Search backwards for last agent response
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 2).Value = "Agent" Then
            GetLastAgentResponse = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i

    GetLastAgentResponse = ""
End Function

'==============================================================================
' ClearChatHistory - Clear all chat history
'==============================================================================
Public Sub ClearChatHistory()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        If lastRow >= 2 Then
            ws.Range("A2:E" & lastRow).ClearContents
        End If
    End If
End Sub

'==============================================================================
' ExportChatHistory - Export chat history to text file
'==============================================================================
Public Function ExportChatHistory(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim history As String
    history = GetChatHistory(100) ' Last 100 entries

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim file As Object
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine "Excel Agent Chat History"
    file.WriteLine "Exported: " & Now
    file.WriteLine String(50, "=")
    file.WriteLine
    file.Write history
    file.Close

    ExportChatHistory = True

    Exit Function

ErrorHandler:
    ExportChatHistory = False
End Function

'==============================================================================
' GetChatStats - Get chat statistics
'==============================================================================
Public Function GetChatStats() As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        GetChatStats = "No chat history"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        GetChatStats = "No chat history"
        Exit Function
    End If

    Dim userCount As Long
    Dim agentCount As Long
    Dim successCount As Long
    Dim errorCount As Long

    Dim i As Long
    For i = 2 To lastRow
        Select Case ws.Cells(i, 2).Value
            Case "User"
                userCount = userCount + 1
            Case "Agent"
                agentCount = agentCount + 1
        End Select

        Select Case ws.Cells(i, 4).Value
            Case "SUCCESS"
                successCount = successCount + 1
            Case "ERROR"
                errorCount = errorCount + 1
        End Select
    Next i

    Dim stats As String
    stats = "Chat Statistics:" & vbCrLf
    stats = stats & "- Total entries: " & (lastRow - 1) & vbCrLf
    stats = stats & "- User messages: " & userCount & vbCrLf
    stats = stats & "- Agent responses: " & agentCount & vbCrLf
    stats = stats & "- Successful: " & successCount & vbCrLf
    stats = stats & "- Errors: " & errorCount

    GetChatStats = stats
End Function
