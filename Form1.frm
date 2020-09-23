VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Shell 2"
      Height          =   375
      Left            =   2550
      TabIndex        =   1
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shell 1"
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   2340
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PId As String
Private FileName$

Private Sub Form_Load()
   Dim OSInfo As OSVERSIONINFO, retval As Long
   
   OSInfo.dwOSVersionInfoSize = Len(OSInfo)

   retval = GetVersionEx(OSInfo)
   If retval <> 0 Then
      Select Case OSInfo.dwPlatformId
          Case 0
              PId = "Win32"
          Case 1
              PId = "Win9x"
          Case 2
              PId = "WinNT"
      End Select
   End If
   FileName = "C:\windows\notepad.exe"
End Sub

'*********************************************************
'This one from hardcore vb with a little modification
'*********************************************************
Private Sub Command1_Click()
   Dim AppID As Single
   Command1.Enabled = False
   AppID = Shell(FileName, vbNormalFocus)
   
   Do While Inst(AppID)
      DoEvents
   Loop
   MsgBox FileName & " Completed.", vbInformation, App.EXEName
   Command1.Enabled = True
End Sub

Function Inst(ByVal idProc As Long) As Boolean
    Dim f As Long, hModule As Long, c As Long
    Inst = False
    If PId = "Win9x" Then
        Dim process As PROCESSENTRY32, module As MODULEENTRY32
        Dim hSnap As Long, idModule As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        ' Loop through to find matching process
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        Do While f
            If process.th32ProcessID = idProc Then
                ' Save module ID
                Inst = True
                Exit Do
            End If
            f = Process32Next(hSnap, process)
        Loop
        CloseHandle hSnap
        
    ElseIf PId = "WinNT" Then
        ' First module is the main executable
        f = EnumProcessModules(ProcFromProcID(idProc), hModule, 4, c)
        If f = 0 Then Exit Function
        Dim modinfo As MODULEINFO
        f = GetModuleInformation(ProcFromProcID(idProc), hModule, modinfo, c)
        If f Then Inst = True
    End If
End Function

Function ProcFromProcID(idProc As Long) As Long
    ProcFromProcID = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                 PROCESS_VM_READ, 0, idProc)
End Function

'*********************************************************
'This one from Desaware top 7 API
'*********************************************************
Private Sub Command2_Click()
    Dim retval As Long
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    sinfo.cb = Len(sinfo)
    sinfo.lpReserved = vbNullString
    sinfo.lpDesktop = vbNullString
    sinfo.lpTitle = vbNullString
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    sinfo.wShowWindow = SW_NORMAL
        
    retval = CreateProcessBynum(vbNullString, FileName, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, sinfo, pinfo)
    If retval Then
        WaitForTerm2 pinfo 'launch and wait
    Else
        MsgBox FileName & " could not be launched", App.EXEName
        Exit Sub
    End If
    MsgBox FileName & " has terminated", vbInformation, App.EXEName

End Sub

Private Sub WaitForTerm2(pinfo As PROCESS_INFORMATION)
    Dim res&
    ' Let the process initialize
    Call WaitForInputIdle(pinfo.hProcess, INFINITE)
    ' We don't need the thread handle
    Call CloseHandle(pinfo.hThread)
    ' Disable the button to prevent reentrancy
    Command2.Enabled = False
    Do
        res = WaitForSingleObject(pinfo.hProcess, 0)
        If res <> WAIT_TIMEOUT Then
            ' No timeout, app is terminated
            Exit Do
        End If
        DoEvents
    Loop While True
    
    Command2.Enabled = True
    ' Kill the last handle of the process
    Call CloseHandle(pinfo.hProcess)
End Sub

