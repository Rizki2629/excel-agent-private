VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00513573} frmSettings 
   Caption         =   "AI Agent Settings"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label Label1 
      Caption         =   "API Endpoint:"
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   1200
   End
   Begin MSForms.TextBox txtEndpoint 
      Height          =   375
      Left            =   1440
      Top             =   120
      Width           =   4440
   End
   Begin MSForms.Label Label2 
      Caption         =   "API Key:"
      Height          =   255
      Left            =   120
      Top             =   600
      Width           =   1200
   End
   Begin MSForms.TextBox txtAPIKey 
      Height          =   375
      Left            =   1440
      Top             =   600
      Width           =   4440
   End
   Begin MSForms.Label Label3 
      Caption         =   "Model:"
      Height          =   255
      Left            =   120
      Top             =   1080
      Width           =   1200
   End
   Begin MSForms.TextBox txtModel 
      Height          =   375
      Left            =   1440
      Top             =   1080
      Width           =   4440
   End
   Begin MSForms.CheckBox chkAutoExecute 
      Caption         =   "Auto-execute VBA (no confirmation)"
      Height          =   375
      Left            =   120
      Top             =   1560
      Width           =   5760
   End
   Begin MSForms.CheckBox chkLogCommands 
      Caption         =   "Log all commands"
      Height          =   375
      Left            =   120
      Top             =   2040
      Width           =   5760
   End
   Begin MSForms.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   120
      Top             =   2640
      Width           =   1200
   End
   Begin MSForms.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1440
      Top             =   2640
      Width           =   1200
   End
   Begin MSForms.CommandButton btnTest 
      Caption         =   "Test Connection"
      Height          =   495
      Left            =   2760
      Top             =   2640
      Width           =   1800
   End
   Begin MSForms.Label lblStatus 
      Caption         =   ""
      Height          =   255
      Left            =   120
      Top             =   3360
      Width           =   5760
      ForeColor       =   0
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    txtEndpoint.Text = GetConfig("APIEndpoint", "https://api.example.com/v1/chat/completions")
    txtAPIKey.Text = GetConfig("APIKey", "")
    txtModel.Text = GetConfig("Model", "gpt-4")
    chkAutoExecute.Value = (GetConfig("AutoExecute", "true") = "true")
    chkLogCommands.Value = (GetConfig("LogCommands", "true") = "true")
    On Error GoTo 0
End Sub

Private Sub btnSave_Click()
    On Error Resume Next
    SaveConfig "APIEndpoint", txtEndpoint.Text
    SaveConfig "APIKey", txtAPIKey.Text
    SaveConfig "Model", txtModel.Text
    If chkAutoExecute.Value Then
        SaveConfig "AutoExecute", "true"
    Else
        SaveConfig "AutoExecute", "false"
    End If
    If chkLogCommands.Value Then
        SaveConfig "LogCommands", "true"
    Else
        SaveConfig "LogCommands", "false"
    End If
    On Error GoTo 0
    
    MsgBox "Settings saved!", vbInformation, "Excel Agent"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnTest_Click()
    lblStatus.Caption = "Testing connection..."
    lblStatus.ForeColor = RGB(255, 165, 0)
    
    SaveConfig "APIEndpoint", txtEndpoint.Text
    SaveConfig "APIKey", txtAPIKey.Text
    
    On Error Resume Next
    Dim result As Boolean
    result = TestConnection()
    On Error GoTo 0
    
    If result Then
        lblStatus.Caption = "Connection successful!"
        lblStatus.ForeColor = RGB(0, 128, 0)
    Else
        lblStatus.Caption = "Connection failed. Check endpoint and API key."
        lblStatus.ForeColor = RGB(200, 0, 0)
    End If
End Sub
