Attribute VB_Name = "Module1"

Option Explicit
Public passok As Boolean
Public lbAppName As String
Public activ As Long
Public wintab(1 To 1000) As Long
Public winnam(1 To 1000) As String
Public winId(1 To 1000) As Long
Public opt(1 To 10) As Variant

Public ptwin As Integer
Public dirname As String
Public ptdir As Integer
Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
Dim lShowCursor As Long
Declare Function GetActiveWindow Lib "User32" () As Long
Declare Function GetForegroundWindow Lib "User32" () As Long
Declare Function GetParent Lib "User32" _
(ByVal hwnd As Long) As Long
Declare Function EnumProcessModules Lib "psapi.dll" _
(ByVal hProcess As Long, ByRef lphModule As Long, _
ByVal cb As Long, ByRef cbNeeded As Long) As Long
Declare Function GetModuleFileNameExA Lib "psapi.dll" _
(ByVal hProcess As Long, ByVal hModule As Long, _
ByVal ModuleName As String, ByVal nSize As Long) As Long
 Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
 Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
 Declare Function GetProcessVersion Lib "kernel32" (ByVal hProcess As Long) As Long
 Declare Function GetThreadPriority Lib "kernel32" (ByVal hProcess As Long) As Long
 Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
 Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
 Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function GetWindowTextLength Lib _
"User32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
 Declare Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function GetClassName Lib "User32" Alias _
   "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As _
   String, ByVal nMaxCount As Long) As Long
   
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_VM_READ = &H10
' Définit les constantes (à partir de WIN32API.TXT).
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Public Const GWL_HWNDPARENT = (-8)

Declare Function IsWindowEnabled Lib "User32" (ByVal hwnd As Long) As Long
Declare Function IsWindow Lib "User32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "User32" (ByVal hwnd As Long) As Long

Declare Function SetWindowLong Lib _
    "User32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewLong As Long) As Long


Public Taille_ComputerName As Long     ' taille du nom de la machine local
Public ComputerName As String * 16     ' nom de la machine local
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2


 Public Type tb_defaut
    libelle    As String       ' texte du defaut
    cd_clr     As Integer      ' couleur du defaut
 End Type
Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long           ' Ce processus
  th32DefaultHeapID As Long
  th32ModuleID As Long            ' exe associé
  cntThreads As Long
  th32ParentProcessID As Long     ' Processus parent du processus
  pcPriClassBase As Long          ' Priorité de base des threads du
                                  ' processus
  dwFlags As Long
  szExeFile As String * 260       ' MAX_PATH
End Type

Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long           '1 = Windows 95.
                                 '2 = Windows NT
  szCSDVersion As String * 128
End Type


 Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


 Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Type PROCESS_INFORMATION
        hProcess As Long
        HThread As Long
        dwProcessId As Long
        dwThreadID As Long
End Type
Private Type Str_Infos
    Lu As Long
    Cle As String
    Nom As String
    Priorite As Long
    Window_Handle As Long
    Thread_Handle As Long
    ID As Long
    Parent_Handle As Long
    EXE As String
    Time_Creation As FILETIME
    Time_CPU As FILETIME
    Time_Exit As FILETIME
    Time_User As FILETIME
    Version_Min As Long
    Version_Max As Long
    filevers As String
    organisation As String
    Description As String
    WS_Min As Long
    WS_Max As Long
End Type

Public Infos As Str_Infos

Public proc As PROCESS_INFORMATION
Public ExitCode As Long
Public Fl_Vue_Maintenance As Boolean

Public StartInf As STARTUPINFO
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetNextWindow Lib "User32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
     lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
     lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
     ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
     ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
     lpStartupInfo As STARTUPINFO, lpProcessInformation As _
     PROCESS_INFORMATION) As Long
     


Declare Function GetDesktopWindow Lib "User32" () As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_CHILD = 5

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Const SW_ERASE = &H4
Public Const SW_HIDE = 0
Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_PARENTOPENING = 3
Public Const SW_RESTORE = 9
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Hwnd_Syn As Long

Public Modules(1 To 1000) As Long

Public Function hidetaskman(ETAT)
  Dim chemin_exe As String

    Dim Handle As Long
    Dim Handle_Rdt_Glo As Long
    Dim Handle2 As Long
    Dim Taille As Long
    Dim Texte As String
    Dim Retour As Long
    Dim Taille_Texte As Long
    Dim Trouve As Boolean
    Dim Status As Long
    Dim I As Long

    Trouve = False

  
    'Handle = GetWindow(frm_synopt.hWnd, GWL_HWNDPARENT)
    Handle = GetDesktopWindow()
    
    Handle2 = GetWindow(Handle, GW_CHILD)


    While (Handle2 <> 0)

        Taille = 50

        Texte = String(50, " ")

        Status = GetWindowText(Handle2, Texte, Taille)
        
        If (Mid(Texte, 1, 1) > " ") Then
            'Debug.Print Mid(Texte, 1, Taille)
        End If
        For I = 1 To 50
            If (Mid(Texte, I, 1) < " ") Then
                Taille_Texte = I - 1
                Exit For
            End If
        Next I
        If (Taille_Texte >= 6) Then

            If (InStr(1, Mid(Texte, 1, 20), Mid("Gestionnaire des tâches de Windows NT", 1, 20)) <> 0) Then

                DoEvents

                Trouve = True
                Handle_Rdt_Glo = Handle2
                Handle2 = 0
                GoTo Suite
            Else

                Handle = Handle2
                Handle2 = GetNextWindow(Handle, GW_HWNDNEXT)

            End If
        Else

            Handle = Handle2
            Handle2 = GetNextWindow(Handle, GW_HWNDNEXT)

        End If
    Wend

Suite:

    If (Trouve = True) Then
        Call ShowWindow(Handle_Rdt_Glo, ETAT)
    Else
       
   End If
End Function

