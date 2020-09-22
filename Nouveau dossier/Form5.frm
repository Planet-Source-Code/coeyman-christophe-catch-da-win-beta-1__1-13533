VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "code?"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2820
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
passok = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (passok = False) Then
    
    Cancel = True
    
Else
    Call hidetaskman(1)
    Exit Sub
End If

End Sub

Private Sub Text1_Change()
If (GetSetting(lbAppName, "passwordproc", "the") = Text1.Text) Then
    passok = True
    Unload Me
End If

End Sub

Private Sub Timer1_Timer()
If (GetSetting(lbAppName, "passwordproc", "actif") = "1") Then Call hidetaskman(0)
DoEvents
End Sub
