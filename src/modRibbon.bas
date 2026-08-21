Attribute VB_Name = "modRibbon"
'==============================================================================
' Excel Agent Add-in - UI Callbacks Module
' Handles toolbar button callbacks
'==============================================================================

Option Explicit

'==============================================================================
' OnChatClick - Open Chat form
'==============================================================================
Public Sub OnChatClick()
    On Error GoTo ErrorHandler
    frmChat.Show vbModeless
    Exit Sub
ErrorHandler:
    MsgBox "Error opening chat: " & Err.Description, vbExclamation, "Excel Agent"
End Sub

'==============================================================================
' OnRunLastClick - Re-run last user message
'==============================================================================
Public Sub OnRunLastClick()
    On Error GoTo ErrorHandler

    Dim lastMessage As String
    lastMessage = GetLastUserMessageFromHistory()

    If lastMessage = "" Then
        MsgBox "No previous command to run.", vbInformation, "Excel Agent"
        Exit Sub
    End If

    Dim apiResponse As Object
    Set apiResponse = SendToAPI(lastMessage)

    If apiResponse.success Then
        If apiResponse.code <> "" Then
            Dim execResult As Object
            Set execResult = ExecuteAction(apiResponse.code, apiResponse.action, apiResponse.target)

            If execResult.success Then
                MsgBox "Command executed successfully!", vbInformation, "Excel Agent"
            Else
                MsgBox "Execution error: " & execResult.errorMessage, vbExclamation, "Excel Agent"
            End If
        End If
    Else
        MsgBox "API error: " & apiResponse.errorMessage, vbExclamation, "Excel Agent"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error running last command: " & Err.Description, vbExclamation, "Excel Agent"
End Sub

'==============================================================================
' OnClearLogClick - Clear chat history
'==============================================================================
Public Sub OnClearLogClick()
    On Error GoTo ErrorHandler

    Dim result As VbMsgBoxResult
    result = MsgBox("Clear all chat history?", vbYesNo + vbQuestion, "Excel Agent")

    If result = vbYes Then
        ClearAllChatHistoryData
        MsgBox "Chat history cleared.", vbInformation, "Excel Agent"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error clearing log: " & Err.Description, vbExclamation, "Excel Agent"
End Sub

'==============================================================================
' OnSettingsClick - Open Settings form
'==============================================================================
Public Sub OnSettingsClick()
    On Error GoTo ErrorHandler
    frmSettings.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening settings: " & Err.Description, vbExclamation, "Excel Agent"
End Sub

'==============================================================================
' OnHelpClick - Show help
'==============================================================================
Public Sub OnHelpClick()
    On Error GoTo ErrorHandler

    Dim helpText As String
    helpText = "Excel Agent - AI Assistant" & vbCrLf & vbCrLf
    helpText = helpText & "Cara Penggunaan:" & vbCrLf
    helpText = helpText & "1. Klik 'Chat' untuk membuka panel chat" & vbCrLf
    helpText = helpText & "2. Ketik perintah dalam bahasa Indonesia" & vbCrLf
    helpText = helpText & "3. Tekan Enter atau klik Send" & vbCrLf
    helpText = helpText & "4. Agent akan membuat dan menjalankan VBA" & vbCrLf & vbCrLf
    helpText = helpText & "Contoh Perintah:" & vbCrLf
    helpText = helpText & "- 'Buat chart penjualan bulanan'" & vbCrLf
    helpText = helpText & "- 'Format sel A1:F10 dengan warna biru'" & vbCrLf
    helpText = helpText & "- 'Ekspor data ke PDF'" & vbCrLf
    helpText = helpText & "- 'Buat sheet baru bernama Laporan'" & vbCrLf

    MsgBox helpText, vbInformation, "Excel Agent Help"

    Exit Sub

ErrorHandler:
    MsgBox "Error showing help: " & Err.Description, vbExclamation, "Excel Agent"
End Sub

'==============================================================================
' GetLastUserMessageFromHistory - Get last user message
'==============================================================================
Private Function GetLastUserMessageFromHistory() As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ChatHistory")
    On Error GoTo 0

    If ws Is Nothing Then
        GetLastUserMessageFromHistory = ""
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 2).Value = "User" Then
            GetLastUserMessageFromHistory = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i

    GetLastUserMessageFromHistory = ""
End Function

'==============================================================================
' ClearAllChatHistoryData - Clear all chat history
'==============================================================================
Private Sub ClearAllChatHistoryData()
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
