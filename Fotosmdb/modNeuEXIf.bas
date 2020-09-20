Attribute VB_Name = "modNeuEXIf"
Option Explicit

'*************************************************************************************************
'    Module     : m_CommandLine (new)
'*************************************************************************************************

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type
      
Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Long
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type
      
Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessID             As Long
    dwThreadID              As Long
End Type
         
Private Const NORMAL_PRIORITY_CLASS         As Long = &H20&

Private Const STARTF_USESTDHANDLES          As Long = &H100&
Private Const STARTF_USESHOWWINDOW          As Long = &H1

Private Const LOGON_WITH_PROFILE            As Long = &H1
Private Const LOGON_NETCREDENTIALS_ONLY     As Long = &H2

Private Const LOGON32_LOGON_INTERACTIVE     As Long = 2
Private Const LOGON32_PROVIDER_DEFAULT      As Long = 0

Private Const CREATE_DEFAULT_ERROR_MODE     As Long = &H4000000
Private Const CREATE_NEW_CONSOLE            As Long = &H10&
Private Const CREATE_NEW_PROCESS_GROUP      As Long = &H200&
Private Const CREATE_SEPARATE_WOW_VDM       As Long = &H800&
Private Const CREATE_SUSPENDED              As Long = &H4&
Private Const CREATE_UNICODE_ENVIRONMENT    As Long = &H400&



Public Function ExecuteCommandLine(Optional ByVal UserName As String, Optional ByVal Password As String, Optional ByVal Domain As String, Optional ByVal strDirectory As String = vbNullString, Optional CommandLine As String) As String
    Dim typProcess      As PROCESS_INFORMATION
    Dim typStartup      As STARTUPINFO
    Dim typSecurity     As SECURITY_ATTRIBUTES
    Dim lngReadPipe     As Long
    Dim lngWritePipe    As Long
    Dim lngBytesRead    As Long
    Dim lngResult       As Long
    Dim lngSuccess      As Long
    Dim strBuffer       As String
    Dim strReturn       As String
    Dim lngToken        As Long
    Dim blnResult           As Boolean
    
    typSecurity.nLength = Len(typSecurity)
    typSecurity.bInheritHandle = 1&
    typSecurity.lpSecurityDescriptor = 0&
    
    lngResult = CreatePipe(lngReadPipe, lngWritePipe, typSecurity, 0)
   
    If lngResult = 0 Then
        MsgBox "CreatePipe failed Error!"
        Exit Function
    End If
   
    typStartup.cb = Len(typStartup)
    typStartup.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    typStartup.hStdInput = lngWritePipe
    typStartup.hStdOutput = lngWritePipe
    typStartup.hStdError = lngWritePipe
    
    If Len(UserName) <> 0 Or Len(Password) <> 0 Then

        lngResult = CreateProcessWithLogon(StrPtr(UserName), StrPtr(Domain), StrPtr(Password), LOGON_WITH_PROFILE, StrPtr(vbNullString), StrPtr(CommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP Or CREATE_UNICODE_ENVIRONMENT, ByVal 0&, StrPtr(strDirectory), typStartup, typProcess)
 
    Else
        lngResult = CreateProcess(StrPtr(vbNullString), StrPtr(CommandLine), typSecurity, typSecurity, ByVal 1&, NORMAL_PRIORITY_CLASS Or CREATE_UNICODE_ENVIRONMENT, ByVal 0&, StrPtr(strDirectory), typStartup, typProcess)
    End If
    
    
    If lngResult <> 0 Then
       
        Dim lngPeekData As Long
        
        Do
            Call PeekNamedPipe(lngReadPipe, ByVal 0&, 0&, ByVal 0&, lngPeekData, ByVal 0&)
            
            If lngPeekData > 0 Then
                strBuffer = Space$(lngPeekData)
                lngSuccess = ReadFile(lngReadPipe, strBuffer, Len(strBuffer), lngBytesRead, 0&)
                
                If lngSuccess = 1 Then
                    strReturn = strReturn & Left$(strBuffer, lngBytesRead)
                Else
                    MsgBox "ReadFile failed!"
                End If
            Else
                lngSuccess = WaitForSingleObject(typProcess.hProcess, 0&)
                        
                If lngSuccess = 0 Then
                    Exit Do
                End If
            End If
            
            DoEvents
        Loop
    Else
        'MsgBox GetSystemErrorMessageText(Err.LastDllError)
        MsgBox "Fehler"
    End If
    
    Call CloseHandle(typProcess.hProcess)
    Call CloseHandle(typProcess.hThread)
    Call CloseHandle(lngReadPipe)
    Call CloseHandle(lngWritePipe)

    ExecuteCommandLine = strReturn
End Function

