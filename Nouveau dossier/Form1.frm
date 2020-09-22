VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "catch da win crazy hacker 2000"
   ClientHeight    =   1890
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView LV_1 
      Height          =   5655
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9975
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6240
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1860
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":031E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":04F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufileinf 
      Caption         =   "fileinfo"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type strinf
    tabr(1 To 10000) As Byte
End Type

Dim ptf As strinf
Dim pt2 As String
Dim correc As Boolean
Dim hr_poste_deb As Date
Dim old_hr_poste As Date
Dim dep_deb_poste As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim explorerwin As Long
Dim IDeplo As Long
Dim clmX As ColumnHeader

Private Sub readparam()
On Error Resume Next

    opt(1) = "0"
    opt(2) = "0"
    opt(3) = "0"
    opt(4) = "0"
    opt(5) = "0"
    opt(6) = "0"
    opt(7) = "0"
    opt(8) = "0"
    
    
opt(1) = GetSetting(lbAppName, "visible", "add")
opt(2) = GetSetting(lbAppName, "explorer", "add")
opt(3) = GetSetting(lbAppName, "visible", "hide")
opt(4) = GetSetting(lbAppName, "explorer", "hide")
opt(5) = GetSetting(lbAppName, "progman", "hide")
opt(6) = GetSetting(lbAppName, "invinamed", "add")
opt(7) = GetSetting(lbAppName, "inviunamed", "add")
opt(8) = GetSetting(lbAppName, "passwordproc", "actif")
End Sub




Private Sub Command1_Click()
Dim tmpi As Integer
Dim Trouve As Boolean
Dim Taille As Long
Dim Texte As String
Dim metxt As String
Timer2 = False
    'Text1.SetFocus
Taille = 13
Texte = String(13, " ")
metxt = String(13, " ")
activ = Me.hwnd
Call GetWindowText(activ, metxt, Taille)
Trouve = False
Do While (Me.hwnd = activ Or activ = 0)
    activ = GetForegroundWindow
    DoEvents
Loop
Call GetWindowText(activ, Texte, Taille)
If (Texte <> metxt) Then

    For tmpi = 1 To ptwin
        If (activ = wintab(tmpi)) Then
            Trouve = True
            Exit For
        End If
    Next tmpi
    If (Trouve = False) Then
        Form1.SetFocus
        ptwin = ptwin + 1
        Form2.Show vbModal
        
        Call ShowWindow(activ, 0)
        wintab(ptwin) = activ
        
    Else
        
        List1.Selected(tmpi - 1) = False
        Call ShowWindow(wintab(tmpi), 0)
    End If
End If
Timer2 = True
End Sub


Private Sub Command2_Click()

Form3.Top = Me.Top
Form3.Left = Me.Left + Me.Width - Form3.Width
Timer2 = False
Form3.Show
Me.Hide

End Sub

Private Sub Command3_Click()
Dim tmpi As Integer
For tmpi = 1 To ptwin
    List1.Selected(tmpi - 1) = False
    Call ShowWindow(wintab(tmpi), 0)
Next tmpi
Command2_Click
End Sub



Private Sub Command4_Click()
Form4.Show vbModal
Me.SetFocus
Call readparam

End Sub

Private Sub Command5_Click()
If (Me.Height = 6375) Then
Me.Height = 2205
Else
Me.Height = 6375
End If
End Sub

Private Sub Form_Load()
Set clmX = LV_1.ColumnHeaders.Add(, , "Parametres", 2000)
Set clmX = LV_1.ColumnHeaders. _
       Add(, , "Valeur", 3350, lvwColumnLeft)
        

    
LV_1.BorderStyle = ccFixedSingle
    
    LV_1.View = lvwReport
    
    LV_1.SmallIcons = ImageList1

    LV_1.View = lvwReport
lbAppName = App.Title
Call readparam
Timer2 = False
Option1 = True
Me.Show
DoEvents
ptwin = 0
correc = True
Call Run_catch
correc = False
Timer2 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tmpi As Integer
For tmpi = 1 To ptwin

If (Mid(winnam(tmpi), 1, 1) = "&") Then
Call ShowWindow(wintab(tmpi), 0)
Else

Call ShowWindow(wintab(tmpi), 1)

End If
Next tmpi
End
End Sub



Private Sub List1_Click()
On Error GoTo finlist

    Dim itmx As ListItem
    Dim CurrWnd As Long
    Dim Length As Long
    Dim TaskName As String
    Dim Parent As Long
    Dim ID As Long
    Dim Time_Creation As FILETIME
    Dim Time_CPU As FILETIME
    Dim Time_User As FILETIME
    Dim Time_Exit As FILETIME
    Dim Version As Long
    Dim Version_Min As Long
    Dim Version_Max As Long
    Dim filevers As String
    Dim WS_Min As Long
    Dim WS_Max As Long
    Dim sClassName As String * 100
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
        
    Dim Date_Trv As SYSTEMTIME
    Dim Time_Trv As SYSTEMTIME
   
    


Timer2.Enabled = False
If (correc = False) Then

    Compteur = 0
    
    LV_1.ListItems.Clear
    
    CurrWnd = wintab(List1.ListIndex + 1) 'GetWindow(Frm_taches.hwnd, GW_HWNDFIRST)
 
        Parent = GetParent(CurrWnd)
        Length = GetWindowTextLength(CurrWnd)
        TaskName = Space$(Length + 1)
        Length = GetWindowText(CurrWnd, TaskName, Length + 1)
        TaskName = Left$(TaskName, Len(TaskName) - 1)
                

            Compteur = Compteur + 1
            
            If TaskName <> Me.Caption Then
                
                Call GetWindowThreadProcessId(CurrWnd, ID)
                HThread = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ID)
                If HThread <> 0 Then
                  lRet = EnumProcessModules(HThread, Modules(1), 200, _
                  cbNeeded2)
                  If lRet <> 0 Then
                    ModuleName = Space(260)
                    nSize = 500
                    lRet = GetModuleFileNameExA(HThread, Modules(1), _
                    ModuleName, nSize)
                    EXE = ""
                                      
                    For I = nSize To 1 Step -1
                        If (Mid(ModuleName, I, 1) > " ") Then
                            EXE = Mid(ModuleName, 1, I)
                            Exit For
                        End If
                    Next I
              
                  End If
                End If
                
                'ID du thread
                Call GetProcessTimes(HThread, Time_Creation, Time_Exit, Time_CPU, Time_User)
                Version = GetProcessVersion(ID)
                Version_Max = Version / 65535
                Version_Min = Version - (Version_Max * 65535)
'                Version_Max = CLng(Hex(Version) And &HFFFF0000)
'                Version_Min = CLng(Hex(Version) And &HFFFF)
                Call GetProcessWorkingSetSize(HThread, WS_Min, WS_Max)
                
                Priorite = GetThreadPriority(HThread)
                Infos.Cle = "Cle " & Compteur
                Infos.EXE = EXE
                Infos.Thread_Handle = HThread
                Infos.Window_Handle = CurrWnd
                ret = GetFileVersionInfoSize(EXE, yo)
                
                ret2 = GetFileVersionInfo(EXE, Infos.Window_Handle, ret, ptf)
                
                For tmpconv = 1 To ret
                    vartmpcnv = Chr$(ptf.tabr(tmpconv * 2))
                    strcnv = strcnv + vartmpcnv
                Next tmpconv
                LSet strcnv = ptf.tabr
                 Infos.filevers = ""
                If (InStr(1, strcnv, "FileVersion")) Then
                    filevers = Trim$(Mid$(strcnv, InStr(1, strcnv, "FileVersion") + 13, 30))
                    filevers = Mid(filevers, 1, InStr(1, filevers, Chr$(0)) - 1)
                    Infos.filevers = filevers
                End If
                
                Infos.organisation = ""
                If (InStr(1, strcnv, "CompanyName")) Then
                    filevers = Trim$(Mid$(strcnv, InStr(1, strcnv, "CompanyName") + 13, 30))
                    filevers = Mid(filevers, 1, InStr(1, filevers, Chr$(0)) - 1)
                    Infos.organisation = filevers
                
                End If
                
                Infos.Description = ""
                If (InStr(1, strcnv, "FileDescription")) Then
                    filevers = Trim$(Mid$(strcnv, InStr(1, strcnv, "FileDescription") + 17, 30))
                    filevers = Mid(filevers, 1, InStr(1, filevers, Chr$(0)) - 1)
                    Infos.Description = filevers
                End If
 
                
                Infos.ID = ID
                Infos.Nom = TaskName
                Infos.Parent_Handle = Parent
                Infos.Lu = 1
                Infos.Time_Creation = Time_Creation
                Infos.Time_Exit = Time_Exit
                Infos.Time_User = Time_User
                Infos.Time_CPU = Time_CPU
                Infos.Priorite = Priorite
                Infos.WS_Max = WS_Max
                Infos.WS_Min = WS_Min
                 Infos.Version_Max = Version_Max
                Infos.Version_Min = Version_Min
                
                CloseHandle (HThread)
            
            End If
 
        DoEvents
   
    LV_1.ListItems.Clear
            Set itmx = LV_1.ListItems.Add(, , "Process")
            itmx.SubItems(1) = Infos.Nom
            Set itmx = LV_1.ListItems.Add(, , "ID du process")
            itmx.SubItems(1) = Infos.ID
            Set itmx = LV_1.ListItems.Add(, , "Handle de la fenêtre")
            itmx.SubItems(1) = Infos.Window_Handle
            Set itmx = LV_1.ListItems.Add(, , "Handle du thread")
            itmx.SubItems(1) = Infos.Thread_Handle
            Set itmx = LV_1.ListItems.Add(, , "Executable")
            itmx.SubItems(1) = Infos.EXE
            Set itmx = LV_1.ListItems.Add(, , "File version")
            itmx.SubItems(1) = Infos.filevers
            Set itmx = LV_1.ListItems.Add(, , "Organisation")
            itmx.SubItems(1) = Infos.organisation
            Set itmx = LV_1.ListItems.Add(, , "Description")
            itmx.SubItems(1) = Infos.Description

            Call FileTimeToSystemTime(Infos.Time_CPU, Time_Trv)

            Set itmx = LV_1.ListItems.Add(, , "CPU Time")
            itmx.SubItems(1) = Format(Time_Trv.wHour, "00") & ":" & _
                               Format(Time_Trv.wMinute, "00") & ":" & _
                               Format(Time_Trv.wSecond, "00") & "." & _
                               Format(Time_Trv.wMilliseconds, "00")
            
            Call FileTimeToSystemTime(Infos.Time_User, Time_Trv)
            
            Set itmx = LV_1.ListItems.Add(, , "User Time")
            itmx.SubItems(1) = Format(Time_Trv.wHour, "00") & ":" & _
                               Format(Time_Trv.wMinute, "00") & ":" & _
                               Format(Time_Trv.wSecond, "00") & "." & _
                               Format(Time_Trv.wMilliseconds, "00")
            
            
            Call FileTimeToSystemTime(Infos.Time_Creation, Time_Trv)
            Set itmx = LV_1.ListItems.Add(, , "Creation Time")
            itmx.SubItems(1) = Format(Time_Trv.wDay, "00") & "/" & _
                               Format(Time_Trv.wMonth, "00") & "/" & _
                               Format(Time_Trv.wYear, "0000") & " " & _
                               Format(Time_Trv.wHour, "00") & ":" & _
                               Format(Time_Trv.wMinute, "00") & ":" & _
                               Format(Time_Trv.wSecond, "00") & "." & _
                               Format(Time_Trv.wMilliseconds, "00")
            
            
            Call FileTimeToSystemTime(Infos.Time_Exit, Time_Trv)
            Set itmx = LV_1.ListItems.Add(, , "Exit Time")
            itmx.SubItems(1) = Format(Time_Trv.wHour, "00") & ":" & _
                               Format(Time_Trv.wMinute, "00") & ":" & _
                               Format(Time_Trv.wSecond, "00") & "." & _
                               Format(Time_Trv.wMilliseconds, "00")

            Set itmx = LV_1.ListItems.Add(, , "Version")
            itmx.SubItems(1) = Infos.Version_Max & "." & Infos.Version_Min
            
            Set itmx = LV_1.ListItems.Add(, , "WorkingSetSize")
            itmx.SubItems(1) = Infos.WS_Min & " (min)/" & Infos.WS_Max & " (Max)"
             
             Call GetClassName(Infos.Window_Handle, sClassName, 100)
             Set itmx = LV_1.ListItems.Add(, , "ClassName")
            itmx.SubItems(1) = sClassName
             
             
                   
             Set itmx = LV_1.ListItems.Add(, , "PArent")
            itmx.SubItems(1) = Infos.Parent_Handle
        
            
If (List1.Selected(List1.ListIndex) = True) Then
    Call ShowWindow(wintab(List1.ListIndex + 1), 1)
    Call ShowWindow(wintab(List1.ListIndex + 1), 0)
    Call ShowWindow(wintab(List1.ListIndex + 1), 1)
    Form1.SetFocus
Else
    Call ShowWindow(wintab(List1.ListIndex + 1), 0)
End If
End If



finlist:
Timer2.Enabled = True
End Sub






Private Sub text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If (KeyAscii = 13) Then
    Call Shell(Text1.Text, vbNormalFocus)
    Text1.Text = ""
End If
End Sub



Private Sub mnufileinf_Click()
If (Me.Height > 2505) Then
        Me.Height = 2505
Else
 Me.Height = 8205
End If
End Sub

Private Sub Timer2_Timer()
Dim Taille As Long
Dim Texte As String
Dim tmpi
Dim tmpi2 As Integer
Dim cpt_change As Integer
cpt_change = 0
    'Text1.SetFocus
If (IsWindowVisible(Me.hwnd) = 0) Then
    Call ShowWindow(Me.hwnd, 1)
End If

correc = True
For tmpi = 1 To ptwin
        Taille = GetWindowTextLength(wintab(tmpi)) + 1
        Texte = String(Taille, " ")
        Call GetWindowText(wintab(tmpi), Texte, Taille)
        'Debug.Print "iswindows " + CStr(IsWindow(wintab(tmpi)))
        'Debug.Print "IsWindowEnabled " + CStr(IsWindowEnabled(wintab(tmpi)))
        If (IsWindowVisible(wintab(tmpi)) = 0 And List1.Selected(tmpi - 1) = True) Then
                List1.Selected(tmpi - 1) = False
                cpt_change = cpt_change + 1
        Else
                If (IsWindowVisible(wintab(tmpi)) = 1 And List1.Selected(tmpi - 1) = False) Then
                      List1.Selected(tmpi - 1) = True
                      cpt_change = cpt_change + 1
                End If
        End If
        If (Mid(Texte, 1, 1) > " " And winnam(tmpi) <> "") Then
            If (InStr(1, winnam(tmpi), Texte) = 0) Then
                    If (Mid(winnam(tmpi), 1, 1) = "&") Then
                    List1.List(tmpi - 1) = Texte
                    cpt_change = cpt_change + 1
                    winnam(tmpi) = "&" + Texte
                    Else
                    List1.List(tmpi - 1) = Texte
                    cpt_change = cpt_change + 1
                    winnam(tmpi) = Texte
                    End If
            End If
           
        Else
            If (IsWindow(wintab(tmpi)) = 0) Then
               For tmpi2 = tmpi + 1 To ptwin
               winnam(tmpi2 - 1) = winnam(tmpi2)
               wintab(tmpi2 - 1) = wintab(tmpi2)
               winId(tmpi2 - 1) = winId(tmpi2)
               Next tmpi2
               List1.RemoveItem (tmpi - 1)
               cpt_change = cpt_change + 1
               ptwin = ptwin - 1
               Exit For
            End If
        End If
DoEvents
Next tmpi
correc = False
If (cpt_change <> 0) Then List1.Refresh
End Sub
Public Function Run_catch()
    Dim chemin_exe As String
    Dim itmx As ListItem
    Dim CurrWnd As Long
    Dim Length As Long
    Dim TaskName As String
    Dim Parent As Long
    Dim ID As Long
    
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
    
    
    'Dim Compteur As Integer
    Dim Handle As Long
    Dim Handle_Rdt_Glo As Long
    Dim Handle2 As Long
    Dim Taille As Long
    Dim Texte As String
    Dim Retour As Long
    Dim Taille_Texte As Long
    Dim Trouve As Boolean
    Dim sClassName          As String * 100
    'Dim Status As Long
    'Dim I As Long
    Dim tmpi As Integer
    Trouve = False
    'Dim Modules(1 To 200) As Long
  
    'Handle = GetWindow(frm_synopt.hWnd, GWL_HWNDPARENT)
    Handle = GetDesktopWindow()
    
    Handle2 = GetWindow(Handle, GW_CHILD)
    'dbg.Show
 
    While (Handle2 <> 0)

        Taille = GetWindowTextLength(Handle2) + 1

        Texte = String(Taille, " ")

        Status = GetWindowText(Handle2, Texte, Taille)
        Call GetWindowThreadProcessId(Handle2, ID)
        HThread = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ID)
        lRet = EnumProcessModules(HThread, Modules(1), 200, _
                  cbNeeded2)
         ModuleName = Space(260)
                    nSize = 500
        lRet = GetModuleFileNameExA(HThread, Modules(1), _
                    ModuleName, nSize)
                    CloseHandle (HThread)
        Parent = GetParent(Handle2)
        
        Trouve = False
        'Debug.Print ModuleName
        
        If (IsWindowVisible(Handle2) = 1) Then
                    
                    'dbg.Text1.Text = dbg.Text1.Text + Chr$(13) + Chr$(10) + "Fenêtre visible: " + CStr(Handle2) + " PID = " + CStr(ID) + " text: " + Mid$(Texte, 1, InStr(1, Texte, Chr$(0)) - 1) + " from :" + Mid$(ModuleName, 1, InStr(1, ModuleName, Chr$(0)) - 1)
                    
                    For tmpi = 1 To ptwin
                        If (Handle2 = wintab(tmpi)) Then
                            Trouve = True
                            Exit For
                        End If
                        If (ID = winId(tmpi) And InStr(1, ModuleName, "explorer") = 0 And IDeplo <> ID) Then
                            If (GetParent(wintab(tmpi)) <> 0 And GetParent(wintab(tmpi)) = Handle2) Then
                            
                             'on laisse courire
                            
                            Else
                            'Trouve = True
                            Exit For
                            End If
                            
                        Else
                            If (InStr(1, ModuleName, "explorer") <> 0) Then
                                IDeplo = ID
                                
                            End If
                        End If
                        
                    Next tmpi


                    If (Trouve = False) Then
                        
                        If (InStr(1, Texte, "catch da win") = 0) Then
                            If (opt(1) = "1" And Mid(Texte, 1, 1) > " ") Then
                                    ptwin = ptwin + 1
                                    
                                        If (Mid(Texte, 1, 1) < " ") Then
                                            Form1.List1.AddItem ModuleName
                                        Else
                                            Form1.List1.AddItem Texte
                                        End If
                                    List1.Selected(ptwin - 1) = True
                                    wintab(ptwin) = Handle2
                                    winnam(ptwin) = Texte
                                    winId(ptwin) = ID
                                    
                                    If (opt(3) = "1" And InStr(1, UCase(Texte), UCase("Program Manager")) = 0) Then
                                        Call ShowWindow(Handle2, 0)
                                         List1.Selected(ptwin - 1) = False
                                    End If
                                    If (opt(5) = "1" And InStr(1, UCase(Texte), UCase("Program Manager")) <> 0) Then
                                        Call ShowWindow(Handle2, 0)
                                         List1.Selected(ptwin - 1) = False
                                    End If
                                    
                            Else
                                If (opt(2) = "1" And Mid(Texte, 1, 1) < " ") Then
                                        ptwin = ptwin + 1
                                        
                                            If (Mid(Texte, 1, 1) < " ") Then
                                                Form1.List1.AddItem ModuleName
                                            Else
                                                Form1.List1.AddItem Texte
                                            End If
                                        List1.Selected(ptwin - 1) = True
                                        wintab(ptwin) = Handle2
                                        winnam(ptwin) = ModuleName
                                        winId(ptwin) = ID
                                        If (opt(4) = "1") Then
                                            Call ShowWindow(Handle2, 0)
                                             List1.Selected(ptwin - 1) = False
                                        End If
                                End If
                            End If
                        End If
                    End If
        Else
            If (opt(6) = "1" And Mid(Texte, 1, 1) > " ") Then
                                    ptwin = ptwin + 1
                                    
                                        If (Mid(Texte, 1, 1) < " ") Then
                                            Form1.List1.AddItem "explorer?"
                                        Else
                                            Form1.List1.AddItem "&" + Texte
                                        End If
                                    List1.Selected(ptwin - 1) = False
                                    wintab(ptwin) = Handle2
                                    'winId(ptwin) = ID
                                    winnam(ptwin) = "&" + Texte
            Else
            
            If (opt(7) = "1" And Mid(Texte, 1, 1) < " ") Then
                                    ptwin = ptwin + 1
                                    
                                        If (Mid(Texte, 1, 1) < " ") Then
                                            Form1.List1.AddItem "&" + CStr(Handle2) + " from " + ModuleName
                                        Else
                                            Form1.List1.AddItem Texte
                                        End If
                                    List1.Selected(ptwin - 1) = False
                                    wintab(ptwin) = Handle2
                                    'winId(ptwin) = ID
                                    winnam(ptwin) = "&" + CStr(Handle2)
                End If
            End If
        End If
                    
        
        If (Mid(Texte, 1, 1) > " ") Then
            
            Debug.Print Mid(Texte, 1, Taille)
           
        End If
        For I = 1 To 25
            If (Mid(Texte, I, 1) < " ") Then
                Taille_Texte = I - 1
                Exit For
            End If
        Next I
        If (Taille_Texte >= 13) Then

            If (InStr(1, Texte, "catch da win N")) Then
                DoEvents
                cpt = cpt + 1
            End If
        End If
        Handle2 = GetNextWindow(Handle2, GW_HWNDNEXT)
    Wend

Suite:
    cpt = cpt + 1
    Me.Caption = "catch da win N°" + CStr(cpt)

End Function
