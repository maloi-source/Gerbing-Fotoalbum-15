Attribute VB_Name = "modVbZentrum"
'http://www.vb-zentrum.de/uniallgemein.html
Option Explicit

    Private isUnicode As Boolean
    
    ' Deklaration:
    Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
    
    ' Typendefinition:
    Private Type OSVERSIONINFO
      dwOSVersionInfoSize As Long
      dwMajorVersion As Long
      dwMinorVersion As Long
      dwBuildNumber As Long
      dwPlatformId As Long
      szCSDVersion As String * 128
    End Type
    
    Private Const VER_PLATFORM_WIN32_NT = 2
    
    ' API declarations.
    Private Const OFS_MAXPATHNAME As Long = 128
    'Private Const OF_WRITE       As Long = &H1
    Private Const OF_READ         As Long = &H0
    Private Const OF_CREATE       As Long = &H1000
    Private Const WM_SETTEXT = &HC&

    Private Type OFSTRUCT
       cBytes               As Byte
       fFixedDisk           As Byte
       nErrCode             As Integer
       Reserved1            As Integer
       Reserved2            As Integer
       szPathName           As String * OFS_MAXPATHNAME
    End Type
    
    Private Type OVERLAPPED
       Internal             As Long
       InternalHigh         As Long
       Offset               As Long
       OffsetHigh           As Long
       hEvent               As Long
    End Type
    
    
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' library command line parameters
    Private Declare Function GetCommandLineW Lib "kernel32" () As Long
    Private Declare Function PathGetArgsW Lib "shlwapi" (ByVal pszPath As Long) As Long
    Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Long, _
                                                              ByVal pszStrPtr As Long) As Long

    
    Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
    Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
    Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
    Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal FileAttributes As Long) As Long
    
    Public Const INVALID_FILE_ATTRIBUTES As Long = -1



    ' Funktion:
    ' Prüft auf NT-Betriebssysteme: True ab Windows NT/2000 und neuer, sonst False
Public Function isSystemNT() As Boolean
    Dim info As OSVERSIONINFO
    
    info.dwOSVersionInfoSize = Len(info)
    GetVersionExA info
    isSystemNT = (info.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

'Public Sub formCaption(ByRef hwnd As Long, ByVal UniCaption As String)
Public Sub formCaption(ByVal hwnd As Long, ByVal UniCaption As String)
    Dim rc As Long
    
    'rc = DefWindowProcW(hwnd, WM_SETTEXT, ByVal &H0&, ByVal StrPtr(UniCaption))
    rc = DefWindowProcW(hwnd, WM_SETTEXT, 0, ByVal StrPtr(UniCaption))
End Sub
'
'Public Function isIDE() As Boolean
'    isIDE = (App.LogMode = 0)
'End Function

'Purpose:Returns True if string has a Unicode char.
Public Function IsStringUnicode(S As String) As Boolean
   Dim i As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(S) Then
        Map = S
        bLen = UBound(Map)
        For i = 1 To bLen Step 2
            If (Map(i) > 0) Then
                IsStringUnicode = True
                Exit Function
            End If
        Next
   End If
End Function

