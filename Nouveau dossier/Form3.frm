VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " crazy hacker 2000"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1710
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   1080
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "catch da win by crazy hacker 2000"
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
passok = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (passok = True) Then
Form1.Timer2 = True
Form1.Top = Me.Top
Form1.Left = Me.Left + Me.Width - Form1.Width
Form1.Show
Unload Me

Else
Cancel = True
Me.Show
End If
End Sub

Private Sub Image1_Click()

If (GetSetting(lbAppName, "passwordproc", "actif") = "1") Then
     passok = False
     Form5.Show vbModal
     If (passok = True) Then
        Form1.Show
        Form1.Timer2 = True
        Form1.Top = Me.Top
        Form1.Left = Me.Left + Me.Width - Form1.Width
        Call hidetaskman(1)
        Unload Me
        Exit Sub
     End If
Else
passok = True
Form1.Show
Form1.Timer2 = True
Form1.Top = Me.Top
Form1.Left = Me.Left + Me.Width - Form1.Width
Call hidetaskman(1)
Unload Me
Exit Sub
End If
Me.SetFocus

End Sub

Private Sub Timer1_Timer()
If (GetSetting(lbAppName, "passwordproc", "actif") = "1") Then Call hidetaskman(0)
DoEvents
End Sub
