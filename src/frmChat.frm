VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00513573} frmChat 
   Caption         =   "AI Agent - Excel Assistant"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label StatusLabel 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8160
      Caption         =   "Status: Ready"
      ForeColor       =   0
   End
   Begin MSForms.ListBox ChatHistory 
      Height          =   3600
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8160
      ColumnCount     =   2
      ColumnWidths    =   "100;7060"
   End
   Begin MSForms.TextBox CodePreview 
      Height          =   1200
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      Locked          =   -1  'True
      FontName        =   "Consolas"
      FontSize        =   9
   End
   Begin MSForms.TextBox InputBox 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   5880
      FontName        =   "Segoe UI"
      FontSize        =   10
   End
   Begin MSForms.CommandButton SendBtn 
      Caption         =   "Send"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   5520
      Width           =   1080
      FontName        =   "Segoe UI"
      FontSize        =   10
      BackColor       =   3407872
      ForeColor       =   16777215
   End
   Begin MSForms.CommandButton ClearBtn 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   5520
      Width           =   960
      FontName        =   "Segoe UI"
      FontSize        =   10
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    UpdateStatus "Ready", RGB(0, 128, 0)
End Sub

Private Sub SendBtn_Click()
    SendMessage
End Sub

Private Sub ClearBtn_Click()
    ChatHistory.Clear
    CodePreview.Text = ""
End Sub

Private Sub SendMessage()
    Dim message As String
    message = Trim(InputBox.Text)
    If message = "" Then Exit Sub
    
    ChatHistory.AddItem "You"
    If ChatHistory.ListCount > 0 Then
        ChatHistory.List(ChatHistory.ListCount - 1, 1) = message
    End If
    
    UpdateStatus "Processing...", RGB(255, 165, 0)
    DoEvents
    
    Dim apiResponse As Object
    Set apiResponse = CallByName(Me, "SendToAPIMessage", VbMethod, message)
    
    If Not apiResponse Is Nothing Then
        If apiResponse.success Then
            ChatHistory.AddItem "Agent"
            If ChatHistory.ListCount > 0 Then
                ChatHistory.List(ChatHistory.ListCount - 1, 1) = apiResponse.explanation
            End If
            
            If apiResponse.code <> "" Then
                CodePreview.Text = apiResponse.code
            End If
            
            UpdateStatus "Ready", RGB(0, 128, 0)
        Else
            ChatHistory.AddItem "Error"
            If ChatHistory.ListCount > 0 Then
                ChatHistory.List(ChatHistory.ListCount - 1, 1) = apiResponse.errorMessage
            End If
            UpdateStatus "Error", RGB(200, 0, 0)
        End If
    End If
    
    InputBox.Text = ""
End Sub

Private Sub UpdateStatus(status As String, color As Long)
    StatusLabel.Caption = "Status: " & status
    StatusLabel.ForeColor = color
End Sub
