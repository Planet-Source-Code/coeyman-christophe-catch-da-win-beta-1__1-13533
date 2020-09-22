VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CheckBox Check8 
      Caption         =   "protect by pasword if reactivate"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "add invisible unnamed win on startup"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "add invisible named win on startup"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "hide program manager on startup"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "hide supposed explorer on startup"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "hide visible named form on startup"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "add supposed explorer on startup"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "add visible named form on startup"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If (Check1 = 0) Then
    Check3.Value = 0
    Check5.Value = 0
End If
End Sub

Private Sub Check2_Click()
If (Check2 = 0) Then
   Check4.Value = 0
End If
End Sub

Private Sub Check3_Click()
If (Check3 = 1) Then Check1.Value = 1
End Sub

Private Sub Check4_Click()
If (Check4 = 1) Then Check2.Value = 1
End Sub

Private Sub Check5_Click()
If (Check5 = 1) Then Check1.Value = 1
End Sub

Private Sub Check8_Click()
If (Check8 = 1) Then
    Text1.Enabled = True
Else
    Text1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
SaveSetting lbAppName, "visible", "add", Check1.Value
SaveSetting lbAppName, "explorer", "add", Check2.Value
SaveSetting lbAppName, "visible", "hide", Check3.Value
SaveSetting lbAppName, "explorer", "hide", Check4.Value
SaveSetting lbAppName, "progman", "hide", Check5.Value
SaveSetting lbAppName, "invinamed", "add", Check6.Value
SaveSetting lbAppName, "inviunamed", "add", Check7.Value
SaveSetting lbAppName, "passwordproc", "actif", Check8.Value
SaveSetting lbAppName, "passwordproc", "the", Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Check1.Value = GetSetting(lbAppName, "visible", "add")
Check2.Value = GetSetting(lbAppName, "explorer", "add")
Check3.Value = GetSetting(lbAppName, "visible", "hide")
Check4.Value = GetSetting(lbAppName, "explorer", "hide")
Check5.Value = GetSetting(lbAppName, "progman", "hide")
Check6.Value = GetSetting(lbAppName, "invinamed", "add")
Check7.Value = GetSetting(lbAppName, "inviunamed", "add")
Check8.Value = GetSetting(lbAppName, "passwordproc", "actif")
If (Check8.Value = 0) Then
        Text1.Text = ""
        Text1.Enabled = False
Else
        Text1.Text = GetSetting(lbAppName, "passwordproc", "the")
        Text1.Enabled = True
End If


End Sub


