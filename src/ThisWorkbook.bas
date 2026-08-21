Attribute VB_Name = "ThisWorkbook"
'==============================================================================
' Excel Agent Add-in - ThisWorkbook Module
' Creates toolbar on load, removes on unload
'==============================================================================

Option Explicit

Private Const TOOLBAR_NAME As String = "AI Agent"

'==============================================================================
' Workbook_Open - Initialize add-in when Excel opens
'==============================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    ' Load default configuration
    Call LoadDefaultConfig

    ' Create toolbar
    Call CreateToolbar

    Exit Sub

ErrorHandler:
    ' Silent fail - don't block Excel startup
    Debug.Print "Excel Agent init error: " & Err.Description
End Sub

'==============================================================================
' Workbook_BeforeClose - Clean up toolbar
'==============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub

'==============================================================================
' CreateToolbar - Create the AI Agent toolbar
'==============================================================================
Private Sub CreateToolbar()
    On Error Resume Next
    ' Delete existing toolbar if present
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0

    Dim cb As CommandBar
    Set cb = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=msoBarTop, Temporary:=True)
    cb.Visible = True

    Dim btn As CommandBarButton

    ' Chat button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Chat"
        .Style = msoButtonIconAndCaption
        .FaceId = 589
        .OnAction = "OnChatClick"
        .TooltipText = "Buka AI Agent Chat"
    End With

    ' Settings button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Settings"
        .Style = msoButtonIconAndCaption
        .FaceId = 2950
        .OnAction = "OnSettingsClick"
        .TooltipText = "Pengaturan AI Agent"
    End With

    ' Run Last button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Run Last"
        .Style = msoButtonIconAndCaption
        .FaceId = 186
        .OnAction = "OnRunLastClick"
        .TooltipText = "Jalankan perintah terakhir"
    End With

    ' Clear Log button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Clear Log"
        .Style = msoButtonIconAndCaption
        .FaceId = 1786
        .OnAction = "OnClearLogClick"
        .TooltipText = "Hapus riwayat chat"
    End With

    ' Help button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Help"
        .Style = msoButtonIconAndCaption
        .FaceId = 984
        .OnAction = "OnHelpClick"
        .TooltipText = "Bantuan AI Agent"
    End With
End Sub

'==============================================================================
' LoadDefaultConfig - Load default configuration values
'==============================================================================
Private Sub LoadDefaultConfig()
    If GetConfig("APIEndpoint", "") = "" Then
        SaveConfig "APIEndpoint", "https://api.example.com/v1/chat/completions"
    End If

    If GetConfig("Model", "") = "" Then
        SaveConfig "Model", "gpt-4"
    End If

    If GetConfig("Language", "") = "" Then
        SaveConfig "Language", "id"
    End If

    If GetConfig("AutoExecute", "") = "" Then
        SaveConfig "AutoExecute", "true"
    End If

    If GetConfig("LogCommands", "") = "" Then
        SaveConfig "LogCommands", "true"
    End If
End Sub
