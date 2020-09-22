VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form2"
   ScaleHeight     =   1665
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Nom?"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ID As Long
Private Sub Command1_Click()

If (text1.Text = "") Then
    If (Me.Caption = "") Then
    text1.Text = "frm num=" + CStr(ptwin) + Me.Caption
    Form1.List1.AddItem "frm num=" + CStr(ptwin) + Me.Caption
    Else
    
    text1.Text = Me.Caption
    Form1.List1.AddItem Me.Caption
    
    End If
Else
     Form1.List1.AddItem text1
End If
Unload Me


End Sub



Private Sub Form_Load()
Dim Taille As Long
Dim Texte As String
Dim chemin_exe As String
Dim itmx As ListItem
Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim Parent As Long

    
Dim Version As Long
Dim Version_Min As Long
Dim Version_Max As Long
Dim WS_Min As Long
Dim WS_Max As Long
    
Dim Priorite As Long
Dim HThread As Long
Dim Status As Long
Dim Modules(1 To 200) As Long
Dim cbNeeded2 As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim I, J As Long
Dim EXE As String
Dim Compteur As Long
    
    
    
    'Text1.SetFocus
        Taille = GetWindowTextLength(activ) + 1

        Texte = String(Taille, " ")

        Status = GetWindowText(activ, Texte, Taille)
        Call GetWindowThreadProcessId(activ, ID)
        HThread = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ID)
        lRet = EnumProcessModules(HThread, Modules(1), 200, _
                  cbNeeded2)
         ModuleName = Space(260)
                    nSize = 500
        lRet = GetModuleFileNameExA(HThread, Modules(1), _
                    ModuleName, nSize)
                    CloseHandle (HThread)
                    
        
        If (Mid(Texte, 1, 1) > " ") Then
           Me.Caption = Texte
           text1.Text = Me.Caption
           'Form1.List1.AddItem Me.Caption
           
           Exit Sub
        Else
            Me.Caption = ModuleName
        End If

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Call Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
    If (Me.Caption <> "") Then
    text1.Text = Me.Caption
    Form1.List1.AddItem Me.Caption
    winnam(ptwin) = Me.Caption
    winId(ptwin) = ID
    Unload Me
    Else
    winnam(ptwin) = ""
    End If
    
End Sub
