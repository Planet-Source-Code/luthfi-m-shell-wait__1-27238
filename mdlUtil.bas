Attribute VB_Name = "mdlUtil"
Option Explicit

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
  ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Declare Function Process32First Lib "kernel32" ( _
  ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Declare Function Process32Next Lib "kernel32" ( _
  ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long
Declare Function EnumProcessModules Lib "PSAPI" ( _
    ByVal hProcess As Long, lphModule As Long, _
    ByVal cb As Long, lpcbNeeded As Long) As Long
Declare Function GetModuleInformation Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Declare Function CreateProcessBynum Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Const hNull = &O0
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const TH32CS_SNAPMODULE = &H8&
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_VM_READ = &H10
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_NORMAL = 1

Public Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const WAIT_TIMEOUT = &H102&
Public Const CREATE_NO_WINDOW = &H8000000


Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process's threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type

Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long        ' This module
    th32ProcessID As Long       ' owning process
    GlblcntUsage As Long        ' Global usage count on the module
    ProccntUsage As Long        ' Module usage count in th32ProcessID's context
    modBaseAddr As Long         ' Base address of module in th32ProcessID's context
    modBaseSize As Long         ' Size in bytes of module starting at modBaseAddr
    hModule As Long             ' The hModule of this module in th32ProcessID's context
    szModule As String * 256    ' MAX_MODULE_NAME32 + 1
    szExePath As String * 260   ' MAX_PATH
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
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type
