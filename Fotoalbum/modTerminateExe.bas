Attribute VB_Name = "modTerminateExe"
Option Explicit
 
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type
 
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
 
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
 
Private Const PROCESS_TERMINATE = &H1
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const TH32CS_SNAPPROCESS = &H2
 
Private Function CheckVersion() As Long
  Dim tOS As OSVERSIONINFO
  tOS.dwOSVersionInfoSize = Len(tOS)
  Call GetVersionEx(tOS)
  CheckVersion = tOS.dwPlatformId
End Function
 
Public Function GetEXEProcessID(ByVal sEXE As String) As Long
  Dim aPID() As Long
  Dim lProcesses As Long
  Dim lProcess As Long
  Dim lModule As Long
  Dim sName As String
  Dim iIndex As Integer
  Dim bCopied As Long
  Dim lSnapShot As Long
  Dim tPE As PROCESSENTRY32
  Dim bDone As Boolean
  
  If CheckVersion() = VER_PLATFORM_WIN32_WINDOWS Then
    'Windows 9x
    'Create a SnapShot of the Currently Running Processes
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lSnapShot < 0 Then Exit Function
    tPE.dwSize = Len(tPE)
    'Buffer the First Processes Info..
    bCopied = Process32First(lSnapShot, tPE)
    Do While bCopied
      'While there are Processes List them..
      sName = Left$(tPE.szExeFile, InStr(tPE.szExeFile, Chr(0)) - 1)
      sName = Mid(sName, InStrRev(sName, "\") + 1)
      If InStr(sName, Chr(0)) Then
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
      End If
      bCopied = Process32Next(lSnapShot, tPE)
      If StrComp(sEXE, sName, vbTextCompare) = 0 Then
        GetEXEProcessID = tPE.th32ProcessID
        Exit Do
      End If
    Loop
    
  Else
    'Windows NT
    'The EnumProcesses Function doesn't indicate how many Process there are,
    'so you need to pass a large array and trim off the empty elements
    'as cbNeeded will return the no. of Processes copied.
    ReDim aPID(255)
    Call EnumProcesses(aPID(0), 1024, lProcesses)
    lProcesses = lProcesses / 4
    ReDim Preserve aPID(lProcesses)
    
    For iIndex = 0 To lProcesses - 1
      'Get the Process Handle, by Opening the Process
      lProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, aPID(iIndex))
      If lProcess Then
        'Just get the First Module, all we need is the Handle to get
        'the Filename..
        If EnumProcessModules(lProcess, lModule, 4, 0&) Then
          sName = Space(260)
          Call GetModuleFileNameExA(lProcess, lModule, sName, Len(sName))
          If InStr(sName, "\") > 0 Then
            sName = Mid(sName, InStrRev(sName, "\") + 1)
          End If
          If InStr(sName, Chr(0)) Then
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
          End If
          If StrComp(sEXE, sName, vbTextCompare) = 0 Then
            GetEXEProcessID = aPID(iIndex)
            bDone = True
          End If
        End If
        'Close the Process Handle
        CloseHandle lProcess
        If bDone Then Exit For
      End If
    Next
  End If
End Function
 
Public Function TerminateEXE(ByVal sEXE As String) As Boolean
  Dim lPID As Long
  Dim lProcess As Long
  
  Do
    lPID = GetEXEProcessID(sEXE)
    If lPID <> 0 Then
        lProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
        Call TerminateProcess(lProcess, 0&)
        Call CloseHandle(lProcess)
    End If
  Loop Until lPID = 0
  TerminateEXE = True
End Function

