Attribute VB_Name = "UnivbzGlobal"
' path_create - Sichere Funktion zum erstellen eines Verzeichnisses auch über mehrere Verzeichnis-Ebenen
' Kill -' VBA replacement for "Kill(PathName)" with UNICODE support, no "On Error..." needed, if useVBAErr = False!
'       Overwrites VBA Kill function, to use VBA explicit write: "VBA.FileSystem.Kill FileName"
' NameAs - is a VBA replacement of: Name 'source' As 'destination'
'       it includes Unicode support, ShellNotification (VBA does not),
'       advanced options (see Enum MOVEFILE_FLAGS) and returns a BOOL for success
' file_rename - ' Umbenennen einer einzelnen Datei/Ordner von 'src' nach 'dst'
'       im Gegensatz zu VBs "Name SourceFile As DestFile" wird hierbei
'       auch ein ShellNotification-Ereignis ausgelöst, sodass andere
'       Anwendungen darauf reagieren können!
' file_move - ' Verschiebt alle an 'src' übergebenen Dateien nach 'dst'
'       Bei Übergabe mehrerer Dateien in 'src' müssen diese zuvor
'       durch Chr$(0) getrennt werden, 'dst' ist dann nur eine Pfadangabe!
' file_path_exist - check, if a file or folder exists (no wildcard allowed)
' anyfile_exist - check, if a file exists, supports unicode and wildcards (? and *)
' file_exist - check, if a single file exists, supports unicode but NO wildcards!!!
' file_delete - Löscht alle an 'fname' übergebenen Dateien, bei Übergabe mehrerer
'       Dateien im 'fname' müssen diese zuvor durch Chr$(0) getrennt werden!
'       Ist 'undo' = True, werden die Dateien in den Papierkorb verschoben!
' file_copy - Kopiert alle an 'src' übergebenen Dateien nach 'dst'
'       Bei Übergabe mehrerer Dateien in 'src' müssen diese zuvor
'       durch Chr$(0) getrennt werden, 'dst' ist dann nur eine Pfadangabe!
' FileLen - replacement for VBA.FileLen: get the file size in bytes (VBA-Overwrite)
' ext_sep - Gibt die Erweiterung eines Dateinamens zurück, also wenn
'       'fname' = 'C:\Dateien\Dokument.doc" ist 'ext_sep' = 'doc'
' ext_change - Ändert die Erweiterung eines Dateinamens gemäß der
'       übergebenen Variable 'ext' - ist keine Erweiterung vorhanden,
'       so wird eine Extension gemäß 'ext' erzeugt:
' ext_appLink - Gibt die Anwendung (incl. path) zurück, mit deren Erweiterung die
'       übergebene Datei 'fname' verknüpft ist. Ist keine Anwendung registriert,
'       so ist das Ergebnis ein Leerstring:

' get volume name of a drive
' gets window position (and dimensions) from registry
' saves window position (and dimensions) to registry
' enable / disable redrawing of a window:
' last function call of a program, makes sure any api function will close!
' This function overwrites VBAs Command$ to get unicode support
' check, if we a running in IDE
' get volume name of a drive
' returns the name of a net drive
' driveToNetpath
' Gibt den Resource-String einer DLL anhand der übergebenen ID zuück
' get the temp directory from environment:
' getTempFile
' getDefaultPrinter
' setDefaultPrinter
' get actual drive incl. path
' set current directory - also on network!
' run a programm or file ShellExecuteW
' get all drive letters and return as an string array
' returns only a folder name, e.g.:' getDirname "C:\Windows\System32\Drivers\" returns "Drivers"
' reconvert a shorten path (8.3 formatted) back to "normal"
' UAC friendly account type function
' has user full admin rights?
' tests, if string contains pure ANSI code
' tests, if a string contains unicode chars
' returns True, if given string (drive, directory,' or complete filename) belongs to a CD drive:
' returns True, if given string (drive, directory,' or complete filename) belongs to a network drive:
' check for NT compatible operating systems:
' send a text with CrLf to a printer object, or an object' which has a DC handle, to actual position
' drawUniText
' MessageBox API variant without window handle

'-------------------------------------------------------------------------------------------
' Modulname : vbzGlobal.bas
' Version   : 4.32
' Date      : 2013-01-24
' Source    : http://www.vb-zentrum.de
' Remarks   : BEFORE you use any of this functions: call "init_global" first!!!
'           : 2012-12-25 replaced function "file_kill" with "Kill", VBA-Overwrite
'           : 2012-01-06 added: "FileLen" (for unicode support and big files, VBA-Overwrite)
'           : 2012-01-21 removed "file_rename" funktion (use "NameAs" or "file_move" instead)
'
' FOR MORE DOCUMENTATION PLEASE VISITE www.vb-zentrum.de/tip_unicode.html

Option Explicit

' IMPORTANT NOTE: ANSISupport
' by default this module creates pure unicode code, which means
' your application supports only Win2K and newer.
' This makes your application smaller and faster!
' If your Application should also support Win95/98/ME do it like this:
' - open the project properties from menu 'project'
' - change to tab 'Make'
' - add to 'Conditional Compilation Arguments': ANSISupport = 1

' GERMAN: ANSISupport
' Standardmäßig erzeugt dieses Modul (fast) reinen Unicode Code und ist
' somit für Programme ab Windows 2000 und neuer geeignet:
' der compilierte Code ist so kompakter und schneller!
' Soll Ihre Anwendung zusätzlich Windows95/98 und ME unterstützen,
' gehen Sie folgendermaßen vor:
' - öffnen Sie die Projekteigenschaften aus dem Menü 'Projekt'
' - wechsel Sie auf die Registerkarte 'Erstellen'
' - tragen Sie unter 'Argumente für bedingte Kompilierung' folgendes ein: ANSISupport = 1

#If False Then                          ' just to keep "ANSISupport" case sensitive
  Private Const ANSISupport = 1
#End If

' API error handling, global tags for MessageBox or LogFiles:                                                                                  # '
Public Const vbCustomError = &H80040200 ' vbObjectError + 512 is free for own use (see MSDN)
Public apiErrorDescription As String    ' error description of API /DLL (api version of VBA.Err.Description)
Public apiErrorNumber As Long           ' error number of API /DLL (api version of VBA.Err.Number)

' country section:
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

' global memory constants
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000&
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_LOWER = GMEM_NOT_BANKED
' global memory functions:
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' command line parameters (for Command$ replacement)
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function PathGetArgsW Lib "shlwapi" (ByVal pszPath As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Long, _
                                                          ByVal pszStrPtr As Long) As Long

' library functions
Public Declare Function LoadLibraryA Lib "kernel32" (ByVal LibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

' 64-bit support
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' program termination:
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

' shell32.dll resource string IDs
Private Declare Function LoadStringA Lib "user32" (ByVal hInstance As Long, ByVal uID As Long, _
                         ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadStringW Lib "user32" (ByVal hInstance As Long, ByVal uID As Long, _
                         ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long
Public Enum IDS_SHELL32_Enum
  IDS_SHELL32_DESKTOP = 4162
  IDS_SHELL32_EXPLORE = 8502
  IDS_SHELL32_NAME = 8976
  IDS_SHELL32_SIZE = 8978
  IDS_SHELL32_TYPE = 8979
  IDS_SHELL32_MODIFIED = 8980
  IDS_SHELL32_CREATED = 8996
  IDS_SHELL32_LASTACCESS = 8997
  IDS_SHELL32_ATTIBUTES = 8987
  IDS_SHELL32_MYCOMPUTER = 9216
  IDS_SHELL32_AFFICHAGES = 33585
  IDS_SHELL32_LARGEICONS = 33577
  IDS_SHELL32_SMALLICONS = 33578
  IDS_SHELL32_LIST = 33579
  IDS_SHELL32_REPORT = 33580
End Enum
#If False Then  ' just to keep Enum in upper case
  Const IDS_SHELL32_DESKTOP = 4162
  Const IDS_SHELL32_EXPLORE = 8502
  Const IDS_SHELL32_NAME = 8976
  Const IDS_SHELL32_SIZE = 8978
  Const IDS_SHELL32_TYPE = 8979
  Const IDS_SHELL32_MODIFIED = 8980
  Const IDS_SHELL32_CREATED = 8996
  Const IDS_SHELL32_LASTACCESS = 8997
  Const IDS_SHELL32_ATTIBUTES = 8987
  Const IDS_SHELL32_MYCOMPUTER = 9216
  Const IDS_SHELL32_AFFICHAGES = 33585
  Const IDS_SHELL32_LARGEICONS = 33577
  Const IDS_SHELL32_SMALLICONS = 33578
  Const IDS_SHELL32_LIST = 33579
  Const IDS_SHELL32_REPORT = 33580
#End If

' user account type (will be set by 'init_global')
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Any, _
        UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal lpBuffer As Long) As Long
' UserInfo structure
Private Type USER_INFO_1
  usri1_name         As Long
  usri1_password     As Long
  usri1_password_age As Long
  usri1_priv         As Long
  usri1_home_dir     As Long
  usri1_comment      As Long
  usri1_flags        As Long
  usri1_script_path  As Long
End Type
Public Enum vbzUserAccount
  USER_PRIV_GUEST = 0                 ' guest account
  USER_PRIV_USER = 1                  ' user account
  USER_PRIV_ADMIN = 2                 ' admin account
  USER_PRIV_ADMIN_RESTRICT = 1002     ' resticted admin (UAC activated!)
End Enum
#If False Then  ' just to keep Enum in upper case
  Const USER_PRIV_GUEST = 0
  Const USER_PRIV_USER = 1
  Const USER_PRIV_ADMIN = 2
  Const USER_PRIV_ADMIN_RESTRICT = 1002
#End If
Public userAccount As vbzUserAccount

' system functions:
#If ANSISupport Then
  Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, _
          ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function PostMessageA Lib "user32" (ByVal hWnd As Long, _
          ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function MessageBoxA Lib "user32" (ByVal hWnd As Long, _
          ByVal lpText As String, ByVal lpCaption As String, ByVal msgType As VbMsgBoxStyle) As VbMsgBoxResult
  Private Declare Function FormatMessageA Lib "kernel32" (ByVal dwFlags As Long, lpSource As Any, _
          ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
          Arguments As Long) As Long
#End If
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, _
        ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, _
        ByVal lpText As Long, ByVal lpCaption As Long, ByVal msgType As VbMsgBoxStyle) As VbMsgBoxResult
Private Declare Function FormatMessageW Lib "kernel32" (ByVal dwFlags As Long, lpSource As Any, _
        ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, _
        Arguments As Long) As Long
Private Const LANG_NEUTRAL = &H0
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
       ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
       
Public Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeout As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
Public Const NIIF_GUID = &H4
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3

       
#If ANSISupport Then
  Public Declare Function ShellAboutA Lib "shell32" (ByVal hWnd As Long, _
         ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
  Private Declare Function ShellExecuteA Lib "shell32" (ByVal hWnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
          ByVal lpDirectory As String, ByVal nShowCmd As VbAppWinStyle) As Long
#End If
Public Declare Function ShellAboutW Lib "shell32" (ByVal hWnd As Long, _
       ByVal szApp As Long, ByVal szOtherStuff As Long, ByVal hIcon As Long) As Long
Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, _
        ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, _
        ByVal lpDirectory As Long, ByVal nShowCmd As VbAppWinStyle) As Long
' use ANSI version of "GetVersionEx" to be compatible to all systems:
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
' theme support
Private Declare Function GetCurrentThemeName Lib "uxtheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
        
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long                                                   'Gerbing 16.11.2015 Public
       
' network functions:
#If ANSISupport Then
  Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare Function GetUserNameA Lib "advapi32" (ByVal lpBuffer As String, nSize As Long) As Long
#End If
Private Declare Function GetComputerNameW Lib "kernel32" (ByVal lpBuffer As Long, nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32" (ByVal lpBuffer As Long, nSize As Long) As Long

' file system functions:
Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const INVALID_FILE_ATTRIBUTES As Long = -1

' *** Begin file and folder operations
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Public Enum FOF_CONSTANTS
  FOF_MULTIDESTFILES = &H1
  FOF_CONFIRMMOUSE = &H2
  FOF_SILENT = &H4
  FOF_RENAMEONCOLLISION = &H8
  FOF_NOCONFIRMATION = &H10
  FOF_WANTMAPPINGHANDLE = &H20
  FOF_ALLOWUNDO = &H40
  FOF_FILESONLY = &H80
  FOF_SIMPLEPROGRESS = &H100
  FOF_NOCONFIRMMKDIR = &H200
  FOF_NOERRORUI = &H400
  FOF_NOCOPYSECURITYATTRIBS = &H800
  FOF_NORECURSION = &H1000
  FOF_NO_CONNECTED_ELEMENTS = &H2000
  FOF_WANTNUKEWARNING = &H4000
End Enum
#If False Then  ' just to keep Enum in upper case
  Private Const FOF_MULTIDESTFILES = &H1
  Private Const FOF_CONFIRMMOUSE = &H2
  Private Const FOF_SILENT = &H4
  Private Const FOF_RENAMEONCOLLISION = &H8
  Private Const FOF_NOCONFIRMATION = &H10
  Private Const FOF_WANTMAPPINGHANDLE = &H20
  Private Const FOF_ALLOWUNDO = &H40
  Private Const FOF_FILESONLY = &H80
  Private Const FOF_SIMPLEPROGRESS = &H100
  Private Const FOF_NOCONFIRMMKDIR = &H200
  Private Const FOF_NOERRORUI = &H400
  Private Const FOF_NOCOPYSECURITYATTRIBS = &H800
  Private Const FOF_NORECURSION = &H1000
  Private Const FOF_NO_CONNECTED_ELEMENTS = &H2000
  Private Const FOF_WANTNUKEWARNING = &H4000
#End If

#If ANSISupport Then
  Private Type SHFILEOPSTRUCTA
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    sProgress As String
  End Type
  Private Declare Function SHFileOperationA Lib "shell32" (lpFileOp As SHFILEOPSTRUCTA) As Long
  Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
          ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
  Private Declare Function MoveFileExA Lib "kernel32" (ByVal lpExistingFileName As String, _
          ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
#End If
Public Type SHFILEOPSTRUCTW
  hWnd As Long
  wFunc As Long
  pFrom As Long         ' Stringpointer
  pTo As Long           ' Stringpointer
  fFlags As Integer
  fAborted As Boolean
  hNameMaps As Long
  sProgress As Long     ' Stringpointer
End Type
Public Declare Function SHFileOperationW Lib "shell32" (lpFileOp As SHFILEOPSTRUCTW) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, _
                                                     ByVal dwFlags As Long) As Long
' CreateFile const
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1&
Private Const FILE_SHARE_WRITE = &H2&
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const OPEN_EXISTING = &H3
Private Const OPEN_ALWAYS = &H4&
' file positions
Private Const FILE_BEGIN = &H0&
Private Const FILE_CURRENT = &H1&
Private Const FILE_END = &H2&
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
        ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
        ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
        ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Enum MOVEFILE_FLAGS
  MOVEFILE_REPLACE_EXISTING = 1   ' Overwrites an existing destination file
  MOVEFILE_COPY_ALLOWED = 2       ' Allow use of CopyFile and DeleteFile instead of directly MoveFile
  MOVEFILE_DELAY_UNTIL_REBOOT = 4 ' The system does not move the file until the operating system is restarted
                                  ' the user must have full admin rights for this flag because of registry settings!
  MOVEFILE_WRITE_THROUGH = 8      ' Function does not return until the file is actually moved on the disk (sync move)
End Enum
#If False Then  ' just to keep Enum in upper case
  Const MOVEFILE_REPLACE_EXISTING = 1
  Const MOVEFILE_COPY_ALLOWED = 2
  Const MOVEFILE_DELAY_UNTIL_REBOOT = 4
  Const MOVEFILE_WRITE_THROUGH = 8
#End If
' *** End file and folder operations

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'Public Type WIN32_FIND_DATA       ' structure for ANSI and UNICODE
'  FileAttributes As vbzFileAttrib
'  CreationTime As FILETIME
'  LastAccessTime As FILETIME
'  LastWriteTime As FILETIME
'  nFileSizeBig As Currency
'  Reserved0 As Long
'  Reserved1 As Long
'  FileName As String * MAX_PATH
'  AlternateFileName As String * 14
'End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


#If ANSISupport Then
  Public Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, lpFFData As WIN32_FIND_DATA) As Long
  Public Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, lpFFData As WIN32_FIND_DATA) As Long
#End If
Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFFData As Long) As Long    'Gerbing 24.06.2014
Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFFData As Long) As Long      'Gerbing 24.06.2014
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function PathMakeUniqueName Lib "shell32" (ByVal pszUniqueName As Long, _
        ByVal cchMax As Long, ByVal pszTemplate As Long, ByVal pszLongPlate As Long, _
        ByVal pszDir As Long) As Boolean

#If ANSISupport Then
  Private Declare Function OemToCharA Lib "user32" (ByVal lpszSrc As String, _
          ByVal lpszDst As String) As Long    ' ASCII to ANSI
  Private Declare Function CharToOemA Lib "user32" (ByVal lpszSrc As String, _
          ByVal lpszDst As String) As Long    ' ANSI to ASCII
  Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
  Private Declare Function GetCurrentDirectoryA Lib "kernel32" _
          (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
  Private Declare Function PathFileExistsA Lib "shlwapi" (ByVal pszPath As String) As Long
  Private Declare Function PathCompactPathExA Lib "shlwapi" (ByVal pszOut As String, _
          ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
  Private Declare Function FindExecutableA Lib "shell32" (ByVal lpFile As String, _
          ByVal lpDirectory As String, ByVal lpResult As String) As Long
  Private Declare Function GetFullPathNameA Lib "kernel32" (ByVal lpFileName As String, _
          ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
#End If
Private Declare Function OemToCharW Lib "user32" (ByVal lpszSrc As Long, _
        ByVal lpszDst As Long) As Long    ' ASCII to ANSI
Private Declare Function CharToOemW Lib "user32" (ByVal lpszSrc As Long, _
        ByVal lpszDst As Long) As Long    ' ANSI to ASCII
Private Declare Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Private Declare Function GetCurrentDirectoryW Lib "kernel32" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function PathFileExistsW Lib "shlwapi" (ByVal pszPath As Long) As Long
Private Declare Function PathCompactPathExW Lib "shlwapi" (ByVal pszOut As Long, _
        ByVal pszSrc As Long, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function FindExecutableW Lib "shell32" (ByVal lpFile As Long, _
        ByVal lpDirectory As Long, ByVal lpResult As Long) As Long
Private Declare Function GetFullPathNameW Lib "kernel32" (ByVal lpFileName As Long, _
        ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal lpPath As String) As Long ' pure ANSI !
Private Declare Function GetTempPathA Lib "kernel32" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, _
        ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
      
' Begin Extended file attributes:
#If ANSISupport Then
  Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
  Private Declare Function SetFileAttributesA Lib "kernel32" (ByVal lpFileName As String, ByVal FileAttributes As Long) As Long
#End If
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal FileAttributes As Long) As Long

Public Enum vbzFileAttrib
  FILE_ATTRIBUTE_READONLY = &H1
  FILE_ATTRIBUTE_HIDDEN = &H2
  FILE_ATTRIBUTE_SYSTEM = &H4
  FILE_ATTRIBUTE_VOLUME = &H8           ' Readonly Attribut! do not use with "SetAttr"!
  FILE_ATTRIBUTE_DIRECTORY = &H10
  FILE_ATTRIBUTE_ARCHIVE = &H20
  FILE_ATTRIBUTE_ALIAS = &H40
  FILE_ATTRIBUTE_NORMAL = &H80
  FILE_ATTRIBUTE_TEMPORARY = &H100
  FILE_ATTRIBUTE_REPARSE_POINT = &H400
  FILE_ATTRIBUTE_COMPRESSED = &H800
  FILE_ATTRIBUTE_OFFLINE = &H1000
  FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
  FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum
#If False Then
  Const FILE_ATTRIBUTE_READONLY = &H1
  Const FILE_ATTRIBUTE_HIDDEN = &H2
  Const FILE_ATTRIBUTE_SYSTEM = &H4
  Const FILE_ATTRIBUTE_VOLUME = &H8
  Const FILE_ATTRIBUTE_DIRECTORY = &H10
  Const FILE_ATTRIBUTE_ARCHIVE = &H20
  Const FILE_ATTRIBUTE_ALIAS = &H40
  Const FILE_ATTRIBUTE_NORMAL = &H80
  Const FILE_ATTRIBUTE_TEMPORARY = &H100
  Const FILE_ATTRIBUTE_REPARSE_POINT = &H400
  Const FILE_ATTRIBUTE_COMPRESSED = &H800
  Const FILE_ATTRIBUTE_OFFLINE = &H1000
  Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
  Const FILE_ATTRIBUTE_ENCRYPTED = &H4000
#End If
' End Extended file attributes
        
' Begin Special Folder:
Private Type ITEMID
  cb As Long
  abID As Byte
End Type
Private Type ITEMIDLIST
  mkid As ITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
        (ByVal hWndOwner As Long, ByVal nFolder As Long, PIDL As ITEMIDLIST) As Long
#If ANSISupport Then
  Private Declare Function SHGetPathFromIDListA Lib "shell32" _
          (ByVal pidList As Long, ByVal lpBuffer As String) As Long
#End If
Private Declare Function SHGetPathFromIDListW Lib "shell32" _
        (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Public Enum CSIDLConstants
  CSIDL_FIRST = &H0                 ' first object in PIDL tree
  CSIDL_LAST = &H3D
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_MYDOCUMENTS = &HC
  CSIDL_MYMUSIC = &HD
  CSIDL_MYVIDEO = &HE
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_LOCAL_APPDATA = &H1C
  CSIDL_ALTSTARTUP = &H1D
  CSIDL_COMMON_ALTSTARTUP = &H1E
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
  CSIDL_COMMON_APPDATA = &H23
  CSIDL_WINDOWS = &H24
  CSIDL_SYSTEM = &H25
  CSIDL_PROGRAM_FILES = &H26
  CSIDL_MYPICTURES = &H27
  CSIDL_PROFILE = &H28
  CSIDL_SYSTEMX86 = &H29
  CSIDL_PROGRAM_FILESX86 = &H2A
  CSIDL_PROGRAM_FILES_COMMON = &H2B
  'CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
  CSIDL_COMMON_TEMPLATES = &H2D
  CSIDL_COMMON_DOCUMENTS = &H2E
  CSIDL_COMMON_ADMINTOOLS = &H2F
  CSIDL_ADMINTOOLS = &H30
  CSIDL_CONNECTIONS = &H31
  CSIDL_COMMON_MUSIC = &H35
  CSIDL_COMMON_PICTURES = &H36
  CSIDL_COMMON_VIDEO = &H37
  CSIDL_RESOURCES = &H38
  CSIDL_RESOURCES_LOCALIZED = &H39
  CSIDL_COMMON_OEM_LINKS = &H3A
  CSIDL_CDBURN_AREA = &H3B
  CSIDL_COMPUTERSNEARME = &H3D
  CSIDL_FLAG_CREATE = &H8000
  CSIDL_FLAG_DONT_VERIFY = &H4000
  CSIDL_FLAG_MASK = &HFF
  CSIDL_FLAG_NO_ALIAS = &H1000
  CSIDL_FLAG_PER_USER_INIT = &H800
End Enum
' End Special Folder

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
        (ByVal nDrive As String) As Long    ' scan drive type
Public Const DRV_REMOVABLE = &H2
Public Const DRV_FIXED = &H3
Public Const DRV_REMOTE = &H4
Public Const DRV_CDROM = &H5
Public Const DRV_RAMDISK = &H6

Private Declare Function GetLogicalDriveStringsA Lib "kernel32" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformationA Lib "kernel32" _
        (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
        ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
        lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
  
Private Declare Function WNetGetConnectionA Lib "Mpr.dll" _
        (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function WNetGetConnectionW Lib "Mpr.dll" _
        (ByVal lpszLocalName As Long, ByVal lpszRemoteName As Long, cbRemoteName As Long) As Long
 
' file properties:
#If ANSISupport Then
  Private Declare Function ShellExecuteExA Lib "shell32" (ShellExExInfo As SHELLEXECUTEINFOA) As Long
  Private Type SHELLEXECUTEINFOA
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
  End Type
#End If
Private Declare Function ShellExecuteExW Lib "shell32" (ShellExExInfo As SHELLEXECUTEINFOW) As Long
Private Type SHELLEXECUTEINFOW
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As Long
  lpFile As Long
  lpParameters As Long
  lpDirectory As Long
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

' Declaration window and object handling:
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, _
        ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long

#If ANSISupport Then
  Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long

Public Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long  ' window is minimized
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long  ' window is maximized
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As VbAppWinStyle) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
       ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function apiSetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

' multimedia:
#If ANSISupport Then
  Private Declare Function sndPlaySoundA Lib "winmm" (ByVal sndName As String, ByVal flags As Long) As Long
#End If
Private Declare Function sndPlaySoundW Lib "winmm" (ByVal sndName As Long, ByVal flags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
        
' input devices
Public Declare Function GetInputState Lib "user32" () As Long ' check mouse or key event
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long   ' text cursor off
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long   ' text cursor on
' private section input device
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursorA Lib "user32" (ByVal hInstance As Long, lpCursorName As Any) As Long
Private Const IDC_WAIT = 32514&    ' Hourglass
Private Const IDC_ARROW = 32512&   ' Arrow

' global WindowMessage mouse constants
Public Const WM_SETREDRAW = &HB&
Public Const WM_SETTEXT = &HC&
Public Const WM_GETTEXT = &HD&
Public Const WM_GETTEXTLENGTH = &HE&
Public Const WM_PAINT = &HF&
Public Const WM_GETMINMAXINFO = &H24&
Public Const WM_NCLBUTTONDOWN = &HA1&
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSEHOVER = &H2A1   ' WINVER >= 2K or COMMCTL >=4.71
Public Const WM_MOUSELEAVE = &H2A3   ' WINVER >= 2K or COMMCTL >=4.71
Public Const WM_DROPFILES = &H233&
Public Const HTCAPTION = &H2&

' Edit messages (as in general editing, not the control)
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

' printer section:
Private Declare Function SetDefaultPrinterA Lib "winspool.drv" (ByVal sPrinterName As String) As Long
Private Declare Function GetDefaultPrinterA Lib "winspool.drv" (ByVal sPrinterName As String, lPrinterNameBufferSize As Long) As Long
#If ANSISupport Then
  Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, _
          ByVal nCount As Long, lpRect As RECT, ByVal wFormat As DT_Format) As Long
#End If
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, _
        ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Enum DT_Format
  DT_LEFT = &H0
  DT_CENTER = &H1
  DT_RIGHT = &H2
  DT_TOP = &H0
  DT_VCENTER = &H4
  DT_BOTTOM = &H8
  DT_WORDBREAK = &H10
  DT_SINGLELINE = &H20
  DT_EXPANDTABS = &H40
  DT_TABSTOP = &H80
  DT_NOCLIP = &H100
  DT_EXTERNALLEADING = &H200
  DT_CALCRECT = &H400
  DT_NOPREFIX = &H800
  DT_INTERNAL = &H1000
  DT_EDITCONTROL = &H2000
  DT_PATH_ELLIPSIS = &H4000
  DT_END_ELLIPSIS = &H8000
  DT_MODIFYSTRING = &H10000
  DT_RTLREADING = &H20000   ' Right to left
End Enum

' system data:
Public Enum vbzSysInfo
  ARW_BOTTOMLEFT = 0
  ARW_BOTTOMRIGHT = 1
  ARW_DOWN = 4
  ARW_HIDE = 8
  ARW_LEFT = 0
  ARW_RIGHT = 4
  ARW_STARTRIGHT = 1
  ARW_STARTTOP = 2
  ARW_TOPLEFT = 2
  ARW_TOPRIGHT = 3
  ARW_UP = 0
  SM_CXSCREEN = 0             ' width of primary screen in pixels
  SM_CYSCREEN = 1             ' height of primary screen in pixels
  SM_CXVSCROLL = 2
  SM_CYHSCROLL = 3
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXDLGFRAME = 7
  SM_CYDLGFRAME = 8
  SM_CYVTHUMB = 9
  SM_CXHTHUMB = 10
  SM_CXICON = 11
  SM_CYICON = 12
  SM_CXCURSOR = 13
  SM_CYCURSOR = 14
  SM_CYMENU = 15
  SM_CXFULLSCREEN = 16
  SM_CYFULLSCREEN = 17
  SM_CYKANJIWINDOW = 18
  SM_MOUSEPRESENT = 19
  SM_CYVSCROLL = 20
  SM_CXHSCROLL = 21
  SM_DEBUG = 22
  SM_SWAPBUTTON = 23
  SM_CXMIN = 28
  SM_CYMIN = 29
  SM_CXSIZE = 30
  SM_CYSIZE = 31
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CXSIZEFRAME = 32
  SM_CYSIZEFRAME = 33
  SM_CXMINTRACK = 34
  SM_CYMINTRACK = 35
  SM_CXDOUBLECLK = 36
  SM_CYDOUBLECLK = 37
  SM_CXICONSPACING = 38
  SM_CYICONSPACING = 39
  SM_MENUDROPALIGNMENT = 40
  SM_PENWINDOWS = 41
  SM_DBCSENABLED = 42
  SM_CMOUSEBUTTONS = 43
  SM_SECURE = 44              ' always 0
  SM_CMETRICS = 44            ' always 0
  SM_CXEDGE = 45
  SM_CYEDGE = 46
  SM_CXMINSPACING = 47
  SM_CYMINSPACING = 48
  SM_CXSMICON = 49
  SM_CYSMICON = 50
  SM_CYSMCAPTION = 51
  SM_CXSMSIZE = 52
  SM_CYSMSIZE = 53
  SM_CXMENUSIZE = 54
  SM_CYMENUSIZE = 55
  SM_ARRANGE = 56
  SM_CXMINIMIZED = 57
  SM_CYMINIMIZED = 58
  SM_CXMAXTRACK = 59
  SM_CYMAXTRACK = 60
  SM_CXMAXIMIZED = 61
  SM_CYMAXIMIZED = 62
  SM_NETWORK = 63
  SM_CLEANBOOT = 67           ' value that specifies how the system is started (boot):
                              ' 0 Normal, 1 Fail-safe, 2 Fail-safe with network
  SM_CXDRAG = 68
  SM_CYDRAG = 69
  SM_SHOWSOUNDS = 70
  SM_CXMENUCHECK = 71         ' Breite der Checkmark Bitmap von Menüs in Pixel
  SM_CYMENUCHECK = 72         ' Höhe der Checkmark Bitmap von Menüs in Pixel
  SM_SLOWMACHINE = 73         ' <> 0 wenn Windows den Rechner für "langsam" hält
  SM_MIDEASTENABLED = 74      ' <> 0 auf Hebräischen und Arabischen Systemen
  SM_MOUSEWHEELPRESENT = 75   ' <> 0 if a mouse with a vertical scroll wheel is installed
  SM_XVIRTUALSCREEN = 76      ' linke Koordinate des Desktops (normalerweise 0)
  SM_YVIRTUALSCREEN = 77      ' obere Koordinate des Desktops (normalerweise 0)
  SM_CXVIRTUALSCREEN = 78     ' gibt die Gesamtbreite des Desktops zurück
  SM_CYVIRTUALSCREEN = 79     ' gibt die Gesamthöhe des Desktops zurück
  SM_CMONITORS = 80           ' number of physical monitors
  SM_SAMEDISPLAYFORMAT = 81   ' <> 0 if all the display monitors have the same color format
  SM_CXFOCUSBORDER = 83       ' width of the left and right edges of the focus rectangle in pixels
  SM_MEDIACENTER = 87         ' <> 0 wenn Media Center Edition installiert ist (ab XP)
End Enum
Public Declare Function GetSystemMetrics Lib "user32" (ByVal Value As vbzSysInfo) As Long

' Typendefinitionen Fensterhandling:
Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As VbAppWinStyle
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Public Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type

Private Type SIZEPAR
  xMin As Long
  yMin As Long
  xMax As Long
  yMax As Long
End Type
' end window handling

' Typ definition of operating system:
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Public Enum SystemTyp
  osUnknown = 0
  osWin95 = 1
  osWin98 = 2
  osWinNT = 3
  osWin2K = 4
  osWinXP = 5
  osVista = 6
  osSeven = 7
  osEight = 8
End Enum
#If False Then  ' just to keep Enum in upper case
  Const osUnknown = 0
  Const osWin95 = 1
  Const osWin98 = 2
  Const osWinNT = 3
  Const osWin2K = 4
  Const osWinXP = 5
  Const osVista = 6
  Const osSeven = 7
  Const osEight = 8
#End If
' end operating system

' global variables:
Public IC As NOTIFYICONDATA           ' ShellNotification-Icon
Public glbDialogRet As VbMsgBoxResult ' Rückgabewert einer Dialogform
Public glbResultString As String      ' globaler Rückgabewert
Public glbPrinterCopies As Integer    ' Anzahl Ausdrucke (Printer)
Public glbCancel As Boolean           ' globaler Abbruch

' global variables (will be initialized by init_global)
Public glbGerman As Boolean           ' True = German / False = English
Public glbComputer As String          ' computer name
Public glbUserName As String          ' user name
Public winVersion As SystemTyp        ' type of operation system
Public isUnicode As Boolean           ' True, if Windows is Win2k or newer
Public isThemedWin As Boolean         ' True, if Windows theming is active
Public is64BitWin As Boolean          ' True, if Windows version is 64 bit
Public lngVersion As Long             ' version number of application as long value
Public strVersion As String           ' version number as formatted string
Public glbAppTitle As String          ' replacement for App.Title (faster and smaller code)
Public glbInstance As Long            ' replacement for App.hInstance (faster and smaller code)
Public glbAppPath As String           ' application path incl. '\'
Public glbTmpPath As String           ' user "Temp" path incl. '\'
Public screenWidth As Long            ' monitor width in pixel
Public screenHeight As Long           ' monitor height in pixel
Public glbDeskHwnd As Long            ' desktop window handle
Public glbHelpFile As String          ' path and name of help file (for vbzDialog.bas)
Public glbMyDocuments As String       ' sys folder 'My documents'
Public glbMyPictures As String        ' sys folder 'My pictures'
Public glbMyMusic As String           ' sys folder 'My music'
Public glbMyVideo As String           ' sys folder 'My videos'
Public glbAppData As String           ' program data folder (common)
Public glbAppLocal As String          ' user local data folder (user depending)

' switch on/off: form always on top
Public Sub always_on_top(ByVal hWnd As Long, ByVal prop As Boolean)
  Dim mode As Long
  
  mode = IIf(prop, -1, -2) ' Effekt ein- / ausschalten
  SetWindowPos hWnd, mode, 0, 0, 0, 0, &H13
End Sub

Public Function isGermanKeyboard() As Boolean
  ' if primary language identifier is "&H7", we have a German layout!
  isGermanKeyboard = ((GetKeyboardLayout(0) And &HFF) = &H7)
End Function

' Gibt die Mauskoordinaten auf den Bildschirm bezogen zurück
' mit Button läßt sich bestimmen wann das geschehen soll:
' 0 = immer, oder vbKeyLButton = bei Linksklick, oder vbKeyRButton = bei Rechtsklick
Public Function getMousePos(ByVal Button As Long, ByRef xPos As Long, ByRef yPos As Long) As Boolean
  Dim cur As POINTAPI
  
  If (GetAsyncKeyState(Button) <> 0) Or Button = 0 Then
    Call GetCursorPos(cur)
    yPos = cur.Y
    xPos = cur.X
    getMousePos = True
  Else
    getMousePos = False
  End If
End Function

Public Function getShiftState() As ShiftConstants
  getShiftState = (-vbShiftMask * pKeyPressed(vbKeyShift))
  getShiftState = getShiftState Or (-vbAltMask * pKeyPressed(vbKeyMenu))
  getShiftState = getShiftState Or (-vbCtrlMask * pKeyPressed(vbKeyControl))
End Function
' helper for "getShiftState"
Private Function pKeyPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean
  pKeyPressed = (GetAsyncKeyState(nVirtKeyCode) And &H8000& = &H8000&)
End Function

' converts an ANSI-String to ASCII-String:
Public Function ANSItoASCII(ByVal Text As String) As String
  #If ANSISupport Then
    If isUnicode Then
      CharToOemW StrPtr(Text), StrPtr(Text)
    Else
      CharToOemA Text, Text
    End If
  #Else
    CharToOemW StrPtr(Text), StrPtr(Text)
  #End If
  ANSItoASCII = Text
End Function

' converts an ASCII-String to ANSI-String:
Public Function ASCIItoANSI(ByVal Text As String) As String
  #If ANSISupport Then
    If isUnicode Then
      OemToCharW StrPtr(Text), StrPtr(Text)
    Else
      OemToCharA Text, Text
    End If
  #Else
    OemToCharW StrPtr(Text), StrPtr(Text)
  #End If
  ASCIItoANSI = Text
End Function

' always reutrn a save string, prevent from NULL an EMPTY, as known
' from e.g. databases and API return values
Public Function asString(ByVal S As Variant) As String
  On Local Error Resume Next
  If IsNull(S) Or IsEmpty(S) Then
    asString = vbNullString
  Else
    asString = RTrim$(CStr(S))
  End If
End Function

' Betrachtet die ersten 8 Zeichen eines Strings als Binärzeichen
' und wandelt diese in die Wertigkeit eines Bytes um, Beispiel: "01000001" -> 65 bzw. 'A'
Public Function BinaryStringToByte(ByVal bString As String) As Byte
  Dim i As Integer, Rv As Byte

  For i = 1 To 8
    If Mid$(bString, i, 1) = "1" Then Rv = Rv + 2 ^ (8 - i)
  Next
  BinaryStringToByte = Rv
End Function
  
' add backslash, if neccessary
Public Function checkPath(ByVal Path As String) As String
  If Right$(Path, 1) <> "\" Then
    checkPath = Path & "\"
  Else
    checkPath = Path
  End If
End Function

' Gibt die Anwendung (incl. path) zurück, mit deren Erweiterung die
' übergebene Datei 'fname' verknüpft ist. Ist keine Anwendung registriert,
' so ist das Ergebnis ein Leerstring:
Public Function ext_appLink(ByVal fName As String) As String
  Dim ret As Long
  Dim Path As String
  
  Path = String(MAX_PATH, 0)
  #If ANSISupport Then
    If isUnicode Then
      ret = FindExecutableW(StrPtr(fName), 0&, StrPtr(Path))
    Else
      ret = FindExecutableA(fName, vbNullString, Path)
    End If
  #Else
    ret = FindExecutableW(StrPtr(fName), 0&, StrPtr(Path))
  #End If
  If ret > 32 Then ext_appLink = vbzNullTrim(Path)    ' ret = 0 to 32 means error (see MSDN)
End Function

' Änder die Erweiterung eines Dateinamens gemäß der
' übergebenen Variable 'ext' - ist keine Erweiterung vorhanden,
' so wird eine Extension gemäß 'ext' erzeugt:
Public Function ext_change(ByVal fName As String, ByVal ext As String) As String
  Dim pos As Long
  
  pos = InStrRev(fName, ".")
  If pos Then
    ext_change = Left$(fName, pos) & ext
  Else
    ext_change = fName & "." & ext
  End If
End Function

' Gibt die Erweiterung eines Dateinamens zurück, also wenn
' 'fname' = 'C:\Dateien\Dokument.doc" ist 'ext_sep' = 'doc'
Public Function ext_sep(ByVal fName As String, Optional ByVal pipe As Boolean = False) As String
  Dim pos As Long
  
  fName = Mid$(fName, InStrRev(fName, "\") + 1) ' get the last part, if fName is a full path
  pos = InStrRev(fName, ".")                    ' search for extension
  If pos Then
    ext_sep = LCase$(Mid$(fName, pos + 1))
  Else
    ext_sep = ""
  End If
  If pipe Then ext_sep = "|" & ext_sep & "|"
End Function

' replacement for VBA.FileLen: get the file size in bytes (VBA-Overwrite)
Public Function FileLen(ByVal fName As String) As Currency
  Dim fHandle As Long
  
  #If ANSISupport Then
    If isUnicode Then
      fHandle = CreateFileW(StrPtr(fName), GENERIC_READ, FILE_SHARE_READ, _
                            ByVal 0&, OPEN_EXISTING, 0&, 0&)
    Else
      fHandle = CreateFileA(fName, GENERIC_READ, FILE_SHARE_READ, _
                            ByVal 0&, OPEN_EXISTING, 0&, 0&)
    End If
  #Else
    fHandle = CreateFileW(StrPtr(fName), GENERIC_READ, FILE_SHARE_READ, _
                          ByVal 0&, OPEN_EXISTING, 0&, 0&)
  #End If
  If fHandle > 0 Then
    If winVersion > osWinNT Then  ' GetFileSizeEx available since Win2K
      Dim fileSize As Currency
      If GetFileSizeEx(fHandle, fileSize) Then
        FileLen = fileSize * 10000
      End If
    Else
      FileLen = GetFileSize(fHandle, 0)
    End If
    Call CloseHandle(fHandle)
  End If
End Function

' Kopiert alle an 'src' übergebenen Dateien nach 'dst'
' Bei Übergabe mehrerer Dateien in 'src' müssen diese zuvor
' durch Chr$(0) getrennt werden, 'dst' ist dann nur eine Pfadangabe!
Public Function file_copy(ByVal src As String, ByVal dst As String, _
                          Optional blSilent As Boolean = False) As Boolean
  Dim ShellInfoW As SHFILEOPSTRUCTW
  
  DoEvents
  src = src & Chr$(0) & Chr$(0)
  dst = dst & Chr$(0) & Chr$(0)
  #If ANSISupport Then
    If isUnicode Then
      With ShellInfoW
        .hWnd = GetForegroundWindow
        .wFunc = FO_COPY
        .pFrom = StrPtr(src)
        .pTo = StrPtr(dst)
        .fFlags = FOF_ALLOWUNDO Or FOF_MULTIDESTFILES Or FOF_NOCONFIRMMKDIR
        If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
        If path_sep(src) = path_sep(dst) Then           ' if path is the same
          .fFlags = .fFlags Or FOF_RENAMEONCOLLISION    ' create copy
        End If
      End With
      file_copy = (SHFileOperationW(ShellInfoW) = &H0&)
    Else
      Dim ShellInfoA As SHFILEOPSTRUCTA
      With ShellInfoA
        .hWnd = GetForegroundWindow
        .wFunc = FO_COPY
        .pFrom = src
        .pTo = dst
        .fFlags = FOF_ALLOWUNDO Or FOF_MULTIDESTFILES Or FOF_NOCONFIRMMKDIR
        If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
        If path_sep(src) = path_sep(dst) Then         ' if path is the same
          .fFlags = .fFlags Or FOF_RENAMEONCOLLISION  ' create copy
        End If
      End With
      file_copy = (SHFileOperationA(ShellInfoA) = &H0&)
    End If
  #Else
    With ShellInfoW
      .hWnd = GetForegroundWindow
      .wFunc = FO_COPY
      .pFrom = StrPtr(src)
      .pTo = StrPtr(dst)
      .fFlags = FOF_ALLOWUNDO Or FOF_MULTIDESTFILES Or FOF_NOCONFIRMMKDIR
      If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
      If path_sep(src) = path_sep(dst) Then           ' if path is the same
        .fFlags = .fFlags Or FOF_RENAMEONCOLLISION    ' create copy
      End If
    End With
    file_copy = (SHFileOperationW(ShellInfoW) = &H0&)
  #End If
End Function

' Löscht alle an 'fname' übergebenen Dateien, bei Übergabe mehrerer
' Dateien im 'fname' müssen diese zuvor durch Chr$(0) getrennt werden!
' Ist 'undo' = True, werden die Dateien in den Papierkorb verschoben!
' Bei blsilent = False kommt ein Bestätigungsdialog
Public Function file_delete(ByVal fName As String, _
                            Optional blUndo As Boolean = False, _
                            Optional blSilent As Boolean = False) As Boolean
  Dim ShellInfoW As SHFILEOPSTRUCTW
  
  DoEvents
  fName = fName & Chr$(0) & Chr$(0)
  If Trim$(fName) = "" Then Exit Function
  #If ANSISupport Then
    If isUnicode Then
      With ShellInfoW
        .hWnd = GetForegroundWindow
        .wFunc = FO_DELETE
        .pFrom = StrPtr(fName)
        .pTo = 0
        .fFlags = FOF_MULTIDESTFILES
        If blUndo Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
      End With
      file_delete = (SHFileOperationW(ShellInfoW) = &H0&)
    Else
      Dim ShInfoA As SHFILEOPSTRUCTA
      With ShInfoA
        .hWnd = GetForegroundWindow
        .wFunc = FO_DELETE
        .pFrom = fName
        .pTo = vbNullChar
        .fFlags = FOF_MULTIDESTFILES
        If blUndo Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
      End With
      file_delete = (SHFileOperationA(ShInfoA) = &H0&)
    End If
  #Else
    With ShellInfoW
      .hWnd = GetForegroundWindow
      .wFunc = FO_DELETE
      .pFrom = StrPtr(fName)
      .pTo = 0
      .fFlags = FOF_MULTIDESTFILES
      If blUndo Then .fFlags = .fFlags Or FOF_ALLOWUNDO
      If blSilent Then .fFlags = .fFlags Or FOF_SILENT Or FOF_NOCONFIRMATION
    End With
    file_delete = (SHFileOperationW(ShellInfoW) = &H0&)
  #End If
End Function

' check, if a single file exists, supports unicode but NO wildcards!!!
Public Function file_exist(ByVal file As String) As Boolean
  Dim dwAttributes As Long
  
  If (file = vbNullString) Then Exit Function
  #If ANSISupport Then
    If isUnicode Then
      dwAttributes = GetFileAttributesW(StrPtr(file))
    Else
      dwAttributes = GetFileAttributesA(file)
    End If
  #Else
    dwAttributes = GetFileAttributesW(StrPtr(file))
  #End If
  If dwAttributes <> INVALID_FILE_ATTRIBUTES Then
    file_exist = (dwAttributes And vbDirectory) = 0
  End If
End Function

'' check, if a file exists, supports unicode and wildcards (? and *)
'Public Function anyfile_exist(ByVal file As String) As Boolean
'  Dim hFile As Long
'  Dim FD As WIN32_FIND_DATA
'
'  #If ANSISupport Then
'    If isUnicode Then
'      hFile = FindFirstFileW(StrPtr(file), VarPtr(FD))
'    Else
'      hFile = FindFirstFileA(file, FD)
'    End If
'  #Else
'    hFile = FindFirstFileW(StrPtr(file), VarPtr(FD))
'  #End If
'  If hFile <> INVALID_HANDLE_VALUE Then
'    anyfile_exist = True
'    FindClose hFile
'  End If
'End Function

' check, if a file or folder exists (no wildcard allowed)
Public Function file_path_exist(ByVal Path As String) As Boolean
  #If ANSISupport Then
    If isUnicode Then
      file_path_exist = CBool(PathFileExistsW(StrPtr(Path)))
    Else
      file_path_exist = CBool(PathFileExistsA(Path))
    End If
  #Else
    file_path_exist = CBool(PathFileExistsW(StrPtr(Path)))
  #End If
End Function

' Verschiebt alle an 'src' übergebenen Dateien nach 'dst'
' Bei Übergabe mehrerer Dateien in 'src' müssen diese zuvor
' durch Chr$(0) getrennt werden, 'dst' ist dann nur eine Pfadangabe!
Public Function file_move(ByVal src As String, ByVal dst, Optional ByVal flags As FOF_CONSTANTS) As Boolean
  Dim ShellInfoW As SHFILEOPSTRUCTW
  
  DoEvents
  src = src & Chr$(0) & Chr$(0)
  dst = dst & Chr$(0) & Chr$(0)
  #If ANSISupport Then
    If isUnicode Then
      With ShellInfoW
        .hWnd = GetForegroundWindow
        .wFunc = FO_MOVE
        .pFrom = StrPtr(src)
        .pTo = StrPtr(dst)
        .fFlags = FOF_MULTIDESTFILES Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or flags
      End With
      file_move = (SHFileOperationW(ShellInfoW) = &H0&)
    Else
      Dim ShellInfoA As SHFILEOPSTRUCTA
      With ShellInfoA
        .hWnd = GetForegroundWindow
        .wFunc = FO_MOVE
        .pFrom = src
        .pTo = dst
        .fFlags = FOF_MULTIDESTFILES Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or flags
      End With
      file_move = (SHFileOperationA(ShellInfoA) = &H0&)
    End If
  #Else
    With ShellInfoW
      .hWnd = GetForegroundWindow
      .wFunc = FO_MOVE
      .pFrom = StrPtr(src)
      .pTo = StrPtr(dst)
      .fFlags = FOF_MULTIDESTFILES Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or flags
    End With
    file_move = (SHFileOperationW(ShellInfoW) = &H0&)
  #End If
End Function


' Umbenennen einer einzelnen Datei/Ordner von 'src' nach 'dst'
' im Gegensatz zu VBs "Name SourceFile As DestFile" wird hierbei
' auch ein ShellNotification-Ereignis ausgelöst, sodass andere
' Anwendungen darauf reagieren können!
Public Function file_rename(ByVal src As String, ByVal dst As String) As Boolean
  Dim ShellInfoW As SHFILEOPSTRUCTW
  Dim flags As Long
  
  src = src & Chr$(0) & Chr$(0)
  dst = dst & Chr$(0) & Chr$(0)
  flags = FOF_ALLOWUNDO Or FOF_RENAMEONCOLLISION Or FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
  #If ANSISupport Then
    If isUnicode Then
      With ShellInfoW
        .hWnd = GetForegroundWindow
        .wFunc = FO_RENAME
        .pFrom = StrPtr(src)
        .pTo = StrPtr(dst)
        .fFlags = flags
      End With
      file_rename = (SHFileOperationW(ShellInfoW) = &H0&)
    Else
      Dim ShellInfoA As SHFILEOPSTRUCTA
      With ShellInfoA
        .hWnd = GetForegroundWindow
        .wFunc = FO_RENAME
        .pFrom = src
        .pTo = dst
        .fFlags = flags
      End With
      file_rename = (SHFileOperationA(ShellInfoA) = &H0&)
    End If
  #Else
    With ShellInfoW
      .hWnd = GetForegroundWindow
      .wFunc = FO_RENAME
      .pFrom = StrPtr(src)
      .pTo = StrPtr(dst)
      .fFlags = flags
    End With
    file_rename = (SHFileOperationW(ShellInfoW) = &H0&)
  #End If
End Function

' NameAs is a VBA replacement of: Name 'source' As 'destination'
' it includes Unicode support, ShellNotification (VBA does not),
' advanced options (see Enum MOVEFILE_FLAGS) and returns a BOOL for success
Public Function NameAs(ByVal src As String, ByVal dst, Optional ByVal useVBAErr As Boolean = True, _
                       Optional ByVal flags As MOVEFILE_FLAGS = MOVEFILE_WRITE_THROUGH) As Boolean
  #If ANSISupport Then
    If isUnicode Then
      NameAs = MoveFileExW(StrPtr(src), StrPtr(dst), flags)
    Else
      NameAs = MoveFileExA(src, dst, flags)
    End If
  #Else
    NameAs = MoveFileExW(StrPtr(src), StrPtr(dst), flags)
  #End If
  getAPIError "NameAs", useVBAErr
End Function

Public Sub file_properties(ByVal fName As String)
  Dim ShExInfo As SHELLEXECUTEINFOW           ' Unicode structure
  
  #If ANSISupport Then                        ' support of ANSI and Unicode
    If isUnicode Then                         ' check for Unicode
      With ShExInfo                           ' fill structure with values...
        .cbSize = Len(ShExInfo)
        .fMask = &H54C
        .hWnd = glbDeskHwnd
        .lpVerb = StrPtr("properties")
        .lpFile = StrPtr(fName & Chr$(0))
      End With
      ShellExecuteExW ShExInfo                ' Unicode API call
    Else                                      ' ANSI Version
      Dim ShExInfoA As SHELLEXECUTEINFOA      ' ANSI structure
      With ShExInfoA                          ' fill structure with values...
        .cbSize = Len(ShExInfoA)
        .fMask = &H54C
        .hWnd = glbDeskHwnd
        .lpVerb = "properties"
        .lpFile = fName & Chr$(0)
      End With
      ShellExecuteExA ShExInfoA               ' ANSI API call
    End If
  #Else                                       ' pure Unicode (small and fast!)
    With ShExInfo                             ' fill structure with values...
      .cbSize = Len(ShExInfo)
      .fMask = &H54C
      .hWnd = glbDeskHwnd
      .lpVerb = StrPtr("properties")
      .lpFile = StrPtr(fName & Chr$(0))
    End With
    ShellExecuteExW ShExInfo                  ' Unicode API call
  #End If
End Sub

' separate file name from full file path
Public Function file_sep(ByVal Path As String, Optional extension As Boolean = True) As String
  Dim pos As Long
  
  file_sep = Mid$(Path, InStrRev(Path, "\") + 1)
  If Not extension Then ' remove extension from file name
    pos = InStrRev(file_sep, ".") - 1
    If pos > 0 Then file_sep = Left$(file_sep, pos)
  End If
End Function

' splits a complete file name into directory, file name and extension:
Public Sub file_split(ByVal src As String, ByRef Path As String, ByRef file As String, ByRef ext As String)
  Dim i As Long
  
  i = InStrRev(src, "\")                ' get last backslash
  Path = Left$(src, i)                  ' Pfad steht links davon
  file = Mid$(src, i + 1)               ' Dateiname rechts davon
  i = InStrRev(file, ".")
  If i = 0 Then
    ext = ""                            ' no extension
  Else
    ext = Mid$(file, i + 1)             ' extension
    file = Left$(file, i - 1)           ' pure filename
  End If
End Sub

' VBA replacement for "Kill(PathName)" with UNICODE support, no "On Error..." needed, if useVBAErr = False!
' Overwrites VBA Kill function, to use VBA explicit write: "VBA.FileSystem.Kill FileName"
Public Sub Kill(ByVal fName As String, Optional ByVal useVBAErr As Boolean = False)
  Dim ret As Long
  Dim hFile As Long
  Dim pName As String
  Dim wfd As WIN32_FIND_DATA
  
  #If ANSISupport Then
    If isUnicode Then
      hFile = FindFirstFileW(StrPtr(fName), VarPtr(FD))
    Else
      hFile = FindFirstFileA(fName, FD)
    End If
  #Else
    hFile = FindFirstFileW(StrPtr(fName), VarPtr(wfd))
    'hFile = FindFirstFileW(StrPtr(fName), wfd)              'Gerbing 04.03.2013
  #End If

  On Error GoTo KillError                                   ' affects only if useVBAErr = True !
  If hFile <> INVALID_HANDLE_VALUE Then
    pName = path_sep(fName)                                 ' save path name
    Do
      'fName = pName & vbzNullTrim(FD.Filename)             ' get full path/file name
      fName = pName & vbzNullTrim(wfd.cFileName)            ' get full path/file name
      SetAttr fName, FILE_ATTRIBUTE_NORMAL                  ' try to reset file attribute
      #If ANSISupport Then
        If isUnicode Then
          If DeleteFileW(StrPtr(fName)) = 0 Then
            getAPIError "Kill", useVBAErr
          End If
          ret = FindNextFileW(hFile, VarPtr(FD))
        Else
          If DeleteFileA(fName) = 0 Then
            getAPIError "Kill", useVBAErr
          End If
          ret = FindNextFileA(hFile, FD)
        End If
      #Else
        If DeleteFileW(StrPtr(fName)) = 0 Then
          getAPIError "Kill", useVBAErr
        End If
        ret = FindNextFileW(hFile, VarPtr(wfd))
        'ret = FindNextFileW(hFile, wfd)
      #End If
    Loop While ret
    FindClose hFile
  End If
  Exit Sub
KillError:
  MessageBox "Error: " & apiErrorNumber & vbCr & vbCr & _
              apiErrorDescription, "Delete " & file_sep(fName), vbExclamation
  Resume Next
End Sub

' shortend a path to MaxChars chars, example:
' path_compact("C:\Programme\Microsoft Visual Studio\VB98", 16)
' returns "C:\Progr...\VB98"
Public Function path_compact(ByVal Path As String, ByVal maxChars As Long) As String
  Dim sBuf As String
  
  sBuf = String(maxChars + 1, 0)
  #If ANSISupport Then
    If isUnicode Then
      PathCompactPathExW StrPtr(sBuf), StrPtr(Path), maxChars, 0&
    Else
      PathCompactPathExA sBuf, Path, maxChars, 0&
    End If
  #Else
    PathCompactPathExW StrPtr(sBuf), StrPtr(Path), maxChars, 0&
  #End If
  path_compact = vbzNullTrim(sBuf)
End Function

' Sicher Funktion zum erstellen eines Verzeichnisses
' auch über mehrere Verzeichnis-Ebenen
Public Function path_create(ByVal Path As String) As Long
  If Not path_exist(Path) Then
    path_create = MakeSureDirectoryPathExists(Path & "\")
  End If
End Function

' same as file_delte (this function is here only to be backward compatible to older versions!)
Public Function path_delete(ByVal hWnd As Long, ByVal PathName As String, _
                       Optional ByVal undo As Boolean = False, _
                       Optional ByVal silent As Boolean = False) As Boolean
  path_delete = file_delete(PathName, undo, silent)
End Function

' check wether a path exists or not
Public Function path_exist(ByVal Path As String) As Boolean
  Dim dwAttributes As Long
  
  If Right$(Path, 1) = "\" Then Path = Left$(Path, Len(Path) - 1)
  #If ANSISupport Then
    If isUnicode Then
      dwAttributes = GetFileAttributesW(StrPtr(Path))
    Else
      dwAttributes = GetFileAttributesA(Path)
    End If
  #Else
    dwAttributes = GetFileAttributesW(StrPtr(Path))
  #End If
  If dwAttributes = INVALID_FILE_ATTRIBUTES Then
    path_exist = False
  Else
    path_exist = (dwAttributes And vbDirectory) <> 0
  End If
End Function

Public Function file_MakeUniqueName(ByVal templateName As String, ByVal folder As String) As String
  Dim uniqueName As String
  
  uniqueName = String(MAX_PATH, 0)
  If PathMakeUniqueName(StrPtr(uniqueName), MAX_PATH, 0, StrPtr(templateName), StrPtr(folder)) Then
    uniqueName = Split(uniqueName, vbNullChar, 2)(0)   ' remove NullChars
    file_MakeUniqueName = file_sep(uniqueName)
  End If
End Function

' returns the path part of a full file name, including backslash
Public Function path_sep(ByVal fName As String) As String
  path_sep = Left$(fName, InStrRev(fName, "\"))
End Function

'' returns a window handle of a given class name or window title,
'' ONE parameter may also be a vbNullString
'Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'  #If ANSISupport Then
'    If isUnicode Then
'      FindWindow = FindWindowW(StrPtr(lpClassName), StrPtr(lpWindowName))
'    Else
'      FindWindow = FindWindowA(lpClassName, lpWindowName)
'    End If
'  #Else
'    FindWindow = FindWindowW(StrPtr(lpClassName), StrPtr(lpWindowName))
'  #End If
'End Function

' centers a form on the screen, or on an optional parent form
Public Sub form_center(ByVal frm As Form, Optional parent As Object = Nothing)
  If parent Is Nothing Then
    frm.Left = (Screen.Width - frm.Width) \ 2
    frm.Top = (Screen.Height - frm.Height) \ 2
  Else
    frm.Left = parent.Left + (parent.Width - frm.Width) \ 2
    frm.Top = parent.Top + (parent.Height - frm.Height) \ 2
  End If
End Sub

' move a form by mouse (inside the form)
' call this on Mouse_Down event of a form, parameter is form handle itself!
Public Sub form_move(ByVal hWnd As Long)
  ReleaseCapture
  SendMessageLong hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

' returns no. of selected elements inside a listbox
Public Function lstGetSelCount(ByVal hWnd As Long) As Long
  lstGetSelCount = SendMessageLong(hWnd, &H191, 0&, 0&)
End Function

' gets window position (and dimensions) from registry
Public Sub load_winpos(ByVal frm As Form)
  Dim fTop As Long, fLeft As Long
  Dim fWidth As Long, fHeight As Long
  Dim dskWidth As Long, dskHeight As Long
    
  dskWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN) * Screen.TwipsPerPixelX
  dskHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN) * Screen.TwipsPerPixelY
  With frm
    fLeft = GetSetting(glbAppTitle, "Settings", .Name & "Left", 1000)
    fTop = GetSetting(glbAppTitle, "Settings", .Name & "Top", 1000)
    If .BorderStyle And (vbSizable Or vbSizableToolWindow) Then
      fWidth = GetSetting(glbAppTitle, "Settings", .Name & "Width", .Width)
      fHeight = GetSetting(glbAppTitle, "Settings", .Name & "Height", .Height)
    Else
      fWidth = .Width
      fHeight = .Height
    End If
    If fLeft > dskWidth - 1000 Then fLeft = dskWidth - fWidth
    If fTop > dskHeight - 1000 Then fTop = dskHeight - fHeight
    .Move fLeft, fTop, fWidth, fHeight           ' move and raise a resize event!
  End With
End Sub

' saves window position (and dimensions) to registry
' these values can be restored by 'load_winpos'
Public Sub save_winpos(ByVal frm As Form)
  With frm
    If (.WindowState <> vbMinimized) Then
      SaveSetting glbAppTitle, "Settings", .Name & "Left", .Left
      SaveSetting glbAppTitle, "Settings", .Name & "Top", .Top
      If .BorderStyle And (vbSizable Or vbSizableToolWindow) Then
        SaveSetting glbAppTitle, "Settings", .Name & "Width", .Width
        SaveSetting glbAppTitle, "Settings", .Name & "Height", .Height
      End If
    End If
  End With
End Sub

' enable / disable redrawing of a window:
Public Sub lockWindow(ByVal hWnd As Long, ByVal wLock As Boolean)
  Dim clientRect As RECT
  
  If wLock Then
    SendMessageLong hWnd, WM_SETREDRAW, False, 0&
  Else
    SendMessageLong hWnd, WM_SETREDRAW, True, 0&
    apiSetFocus hWnd
    GetClientRect hWnd, clientRect
    RedrawWindow hWnd, clientRect, 0&, &H185&
  End If
End Sub

Public Sub hourglass(ByVal hWnd As Long, ByVal flag As Boolean)
  If flag Then  ' Hourglass on!
    Call SetCapture(hWnd)
    Call SetCursor(LoadCursorA(glbInstance, ByVal IDC_WAIT))
  Else          ' Hourglass off!
    Call ReleaseCapture
    Call SetCursor(LoadCursorA(glbInstance, ByVal IDC_ARROW))
  End If
End Sub

' Progressbar in StatusBar-Panel setzen
Public Sub setProgressToStatusbar(ByVal hWnd_PBar As Long, ByVal hWnd_SBar As Long, ByVal nPanel As Long)
  Dim R As RECT
  ' Ausmaße des Panel ermitteln
  SendMessage hWnd_SBar, &H40A, nPanel - 1, R
  ' ProgressBar ein neues Zuhause geben...
  SetParent hWnd_PBar, hWnd_SBar
  ' ... und korrekt positionieren
  MoveWindow hWnd_PBar, R.Left + 1, R.Top + 1, R.Right - R.Left - 1, R.Bottom - R.Top - 2, True
End Sub

' Initialisiert diverse globale Variabeln
' Sollte im Load-Ereignis der Startform oder in 'Sub Main'
' aufgerufen werden. wird multiInstance nicht auf 'TRUE' gesetzt,
' so wird die Applikation zur Highlander-Anwendung, will sagen
' es kann nur eine geben!
Public Function init_global(Optional multiInstance As Boolean = False) As Boolean
  Dim ret As Long
  Dim Buffer As String
  Dim verInfo As OSVERSIONINFO
  
  ' get region by LanguageCodeID:
  Buffer = "|" & CStr(GetUserDefaultLCID) & "|"
  glbGerman = InStrB(1, "|1031|2055|3079|4103|5127|", Buffer) > 0
  
  ' get operating system, Unicode flag and Theming:
  verInfo.dwOSVersionInfoSize = Len(verInfo)
  GetVersionExA verInfo
  isUnicode = (verInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
  If isUnicode Then
    Select Case verInfo.dwMajorVersion
      Case 4
        winVersion = osWinNT
      Case 5
        If verInfo.dwMinorVersion Then
          winVersion = osWinXP
        Else
          winVersion = osWin2K
        End If
      Case 6
        If verInfo.dwMinorVersion = 0 Then
          winVersion = osVista
        ElseIf verInfo.dwMinorVersion = 1 Then
          winVersion = osSeven
        ElseIf verInfo.dwMinorVersion = 2 Then
          winVersion = osEight
        End If
      Case Else
        winVersion = osUnknown
    End Select
    isThemedWin = isThemedWindows
    ' check for 64bit OS
    If GetProcAddress(GetModuleHandleA("kernel32.dll"), "IsWow64Process") Then
      IsWow64Process GetCurrentProcess, ret
      is64BitWin = ret <> 0
    End If
  Else
    winVersion = osWin98
  End If
  
  With App
    ' get applications version:
'    lngVersion = .Major * 10000 + .Minor * 1000 + .Revision  '  12345
'    strVersion = Format$(lngVersion, "##\.#\.###")           '1.2.345
    glbAppTitle = .Title
    glbInstance = .hInstance
    If Not multiInstance Then
      ' only one instance:
      If .PrevInstance Then
        ' bring application to top:
        Buffer = glbAppTitle
        glbAppTitle = " "
        AppActivate Buffer
        init_global = False
        Exit Function
      End If
    End If
    glbTmpPath = checkPath(getTempFolder)
    glbAppPath = checkPath(.Path)
    setCurrentDir glbAppPath
  End With
  
  ' computer- and user name
  glbComputer = String(128, 0)
  glbUserName = String(128, 0)
  #If ANSISupport Then
    If isUnicode Then
      GetComputerNameW StrPtr(glbComputer), Len(glbComputer)
      glbComputer = vbzNullTrim(glbComputer)
      GetUserNameW StrPtr(glbUserName), Len(glbUserName)
      glbUserName = vbzNullTrim(glbUserName)
    Else
      GetComputerNameA glbComputer, Len(glbComputer)
      glbComputer = vbzNullTrim(glbComputer)
      GetUserNameA glbUserName, Len(glbUserName)
      glbUserName = vbzNullTrim(glbUserName)
    End If
  #Else
    GetComputerNameW StrPtr(glbComputer), Len(glbComputer)
    glbComputer = vbzNullTrim(glbComputer)
    GetUserNameW StrPtr(glbUserName), Len(glbUserName)
    glbUserName = vbzNullTrim(glbUserName)
  #End If
  ' User Konto Typ
  userAccount = getUserAccountType
  ' Pfaddefinitionen:
  ' Benutzer-Verzeichnisse
  glbMyDocuments = getSpecialFolder(CSIDL_PERSONAL)
  If glbMyDocuments = "" Then
    glbMyDocuments = "C:\" & IIf(glbGerman, "Eigene Dateien", "MyDocuments")
    path_create glbMyDocuments
  End If
  ' Eigene Bilder
  glbMyPictures = getSpecialFolder(CSIDL_MYPICTURES)
  If glbMyPictures = "" Then
    glbMyPictures = glbMyDocuments & IIf(glbGerman, "\Eigene Bilder", "\MyPictures")
    path_create glbMyPictures
  End If
  ' Eigene Musik
  glbMyMusic = getSpecialFolder(CSIDL_MYMUSIC)
  If glbMyMusic = "" Then
    glbMyMusic = glbMyDocuments & IIf(glbGerman, "\Eigene Musik", "\MyMusic")
    path_create glbMyMusic
  End If
  ' Eigene Videos
  glbMyVideo = getSpecialFolder(CSIDL_MYVIDEO)
  If glbMyVideo = "" Then
    glbMyVideo = glbMyDocuments & IIf(glbGerman, "\Eigene Videos", "\MyVideo")
    path_create glbMyVideo
  End If
  
  ' program  data directory
  glbAppData = getSpecialFolder(CSIDL_COMMON_APPDATA)
  If glbAppData = vbNullString Then
    glbAppData = glbAppPath
  Else
    ' ab Windows 2000 unter ProgramData ablegen:
    glbAppData = glbAppData & "\vbZentrum\" & App.EXEName & "\"
    path_create glbAppData
  End If
  ' Lokales Programm-Datenverzeichnis
  glbAppLocal = getSpecialFolder(CSIDL_LOCAL_APPDATA)
  If glbAppLocal = vbNullString Then
    glbAppLocal = glbAppPath
  Else
    ' ab Windows 2000 unter lokale ProgramData ablegen:
    glbAppLocal = glbAppLocal & "\vbZentrum\" & App.EXEName & "\"
    path_create glbAppLocal
  End If
  
  ' Desktop-Infos
  glbDeskHwnd = GetDesktopWindow()
  screenWidth = GetSystemMetrics(SM_CXSCREEN)
  screenHeight = GetSystemMetrics(SM_CYSCREEN)
  init_global = True
End Function

' last function call of a program, makes sure any api function will close!
Public Sub term_global()
  If App.LogMode Then     ' do this only in compiled code
    ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
  End If
End Sub

' This function overwrites VBAs Command$ to get unicode support
' no code change on your project neccessary! ;-)
Public Function command() As String
  If isSystemNT() And Not isIDE Then
    SysReAllocString VarPtr(command), PathGetArgsW(GetCommandLineW)
    isUnicode = True
  Else
    command = VBA.command$
  End If
End Function

' check, if we a running in IDE
Public Function isIDE() As Boolean
  isIDE = (App.LogMode = 0)
End Function

' get volume name of a drive
Public Function getVolumeName(ByVal drive As String) As String
  Dim DrvVolumeName As String, FSName As String * 32
  Dim Unused1 As Long, Unused2 As Long, Unused3 As Long
  
  DrvVolumeName = Space$(41)
  If GetVolumeInformationA(drive, DrvVolumeName, Len(DrvVolumeName), _
                          Unused1, Unused2, Unused3, FSName, Len(FSName)) Then
    getVolumeName = vbzNullTrim(DrvVolumeName)
  End If
End Function

' delay function, parameter on milliseconds
Public Function Wait(ByVal mSek As Long)
  WaitForSingleObject -1, mSek
End Function

' returns the name of a net drive
Public Function WNetGetVolumeName(ByVal drive As String) As String
  Dim ret As Long
  Dim nSize As Long
  Dim volName As String
  Dim Text As String
  
  On Local Error Resume Next
  drive = Left$(drive, 2)
  volName = Space$(255)
  nSize = Len(volName)
  #If ANSISupport Then
    If isUnicode Then
      ret = WNetGetConnectionW(StrPtr(drive), StrPtr(volName), nSize)
    Else
      ret = WNetGetConnectionA(drive, volName, nSize)
    End If
  #Else
    ret = WNetGetConnectionW(StrPtr(drive), StrPtr(volName), nSize)
  #End If
  If ret = 0 Then
    volName = vbzNullTrim(volName)
    If Left$(volName, 2) = "\\" Then
      volName = Mid$(volName, 3)
      ret = InStr(volName, "\")
      If ret > 0 Then
        Text = Left$(volName, ret - 1)
        volName = Mid$(volName, ret + 1) & " on """ & _
                  StrConv(Text, vbProperCase) & """"
      End If
    End If
    WNetGetVolumeName = volName
  End If
End Function

Public Function driveToNetpath(ByVal drive As String) As String
  Dim ret As Long
  Dim nSize As Long
  Dim volName As String
  
  drive = Left$(drive, 2)
  volName = Space$(255)
  nSize = Len(volName)
  #If ANSISupport Then
    If isUnicode Then
      ret = WNetGetConnectionW(StrPtr(drive), StrPtr(volName), nSize)
    Else
      ret = WNetGetConnectionA(drive, volName, nSize)
    End If
  #Else
    ret = WNetGetConnectionW(StrPtr(drive), StrPtr(volName), nSize)
  #End If
  If ret = 0 Then driveToNetpath = vbzNullTrim(volName)
End Function

' retrieve the windows directory
Public Function getWinFolder() As String
  getWinFolder = getSpecialFolder(CSIDL_WINDOWS)
End Function

' Gibt den Resource-String einer DLL anhand der übergebenen ID zuück
Public Function getResourceString(ByVal sModule As String, ByVal ids As IDS_SHELL32_Enum) As String
  Dim hModule As Long
  Dim sBuf As String * MAX_PATH
  Dim ln As Long
  
  hModule = LoadLibraryA(sModule)
  If hModule Then
    ln = LoadStringA(hModule, ids, sBuf, MAX_PATH)
    If ln Then getResourceString = Left$(sBuf, ln)
    Call FreeLibrary(hModule)
  End If
End Function

' request for Special Folders like: 'Documents', 'My Music' etc.
' by sending the parameter 'num' from CSIDLConstants Enum
Public Function getSpecialFolder(ByVal num As CSIDLConstants) As String
  Dim temp As String
  Dim idl As ITEMIDLIST

  temp = String(MAX_PATH, 0)
  If (SHGetSpecialFolderLocation(0&, num, idl) = 0) Then
    #If ANSISupport Then
      If isUnicode Then
        If SHGetPathFromIDListW(ByVal idl.mkid.cb, StrPtr(temp)) Then
          getSpecialFolder = vbzNullTrim(temp)
        End If
      ElseIf SHGetPathFromIDListA(ByVal idl.mkid.cb, temp) Then
        getSpecialFolder = vbzNullTrim(temp)
      End If
    #Else
      If SHGetPathFromIDListW(ByVal idl.mkid.cb, StrPtr(temp)) Then
        getSpecialFolder = vbzNullTrim(temp)
      End If
    #End If
  End If
End Function

'  get the systems directory:
Public Function getSystemFolder() As String
  getSystemFolder = getSpecialFolder(CSIDL_SYSTEM)
End Function

' get the temp directory from environment:
Public Function getTempFolder() As String
  Dim vLen As Long
  Dim temp As String * MAX_PATH

  vLen = GetTempPathA(Len(temp), temp)
  getTempFolder = Left$(temp, vLen)
End Function

Public Function getTempFile(Optional Prefix As String, Optional extension As String = "tmp") As String
  Dim tempFile As String
  
  If extension <> "tmp" Then
    Do
      tempFile = glbTmpPath & Prefix & CStr(GetTickCount) & "." & extension
    Loop While file_exist(tempFile)
  Else
    tempFile = Space$(MAX_PATH)
    GetTempFileNameA glbTmpPath, Prefix, 0, tempFile
    tempFile = vbzNullTrim(tempFile)
  End If
  getTempFile = tempFile
End Function

Public Function getDefaultPrinter() As String
  Dim Buffer As String * 128
  Dim lbuf As Long
  
  lbuf = Len(Buffer)
  If GetDefaultPrinterA(Buffer, lbuf) Then
    getDefaultPrinter = Left$(Buffer, lbuf - 1)
  Else
    getDefaultPrinter = vbNullString
  End If
End Function

Public Function setDefaultPrinter(ByVal PrinterName As String) As Boolean
  setDefaultPrinter = CBool(SetDefaultPrinterA(PrinterName))
End Function

' VBA replacement for GetAttr, supports unicode and network
Public Function GetAttr(ByVal fName As String) As vbzFileAttrib
  #If ANSISupport Then
    If isUnicode Then
      If Left$(fName, 2) = "\\" Then fName = "UNC\" & Mid$(fName, 3)
      GetAttr = GetFileAttributesW(StrPtr("\\?\" & fName))
    Else
      GetAttr = GetFileAttributesA(fName)
    End If
  #Else
    If Left$(fName, 2) = "\\" Then fName = "UNC\" & Mid$(fName, 3)
    GetAttr = GetFileAttributesW(StrPtr("\\?\" & fName))
  #End If
End Function

' VBA replacement for SetAttr, supports unicode and network
Public Function SetAttr(ByVal fName As String, ByVal Attributes As vbzFileAttrib) As Boolean
  #If ANSISupport Then
    If isUnicode Then
      If Left$(fName, 2) = "\\" Then fName = "UNC\" & Mid$(fName, 3)
      SetAttr = CBool(SetFileAttributesW(StrPtr("\\?\" & fName), Attributes))
    Else
      SetAttr = CBool(SetFileAttributesA(fName, Attributes))
    End If
  #Else
    If Left$(fName, 2) = "\\" Then fName = "UNC\" & Mid$(fName, 3)
    SetAttr = CBool(SetFileAttributesW(StrPtr("\\?\" & fName), Attributes))
  #End If
End Function

' get actual drive incl. path
Public Function getCurrentDir() As String
  Dim ln As Long
  Dim sBuf As String
  
  sBuf = String$(MAX_PATH, 0)
  #If ANSISupport Then
    If isUnicode Then
      ln = GetCurrentDirectoryW(Len(sBuf), StrPtr(sBuf))
    Else
      ln = GetCurrentDirectoryA(Len(sBuf), sBuf)
    End If
  #Else
    ln = GetCurrentDirectoryW(Len(sBuf), StrPtr(sBuf))
  #End If
  If ln Then getCurrentDir = Left$(sBuf, ln)
End Function

' set current directory - also on network!
Public Function setCurrentDir(ByVal Path As String) As Boolean
  #If ANSISupport Then
    If isUnicode Then
      setCurrentDir = CBool(SetCurrentDirectoryW(StrPtr(Path)))
    Else
      setCurrentDir = CBool(SetCurrentDirectoryA(Path))
    End If
  #Else
    setCurrentDir = CBool(SetCurrentDirectoryW(StrPtr(Path)))
  #End If
End Function

' run a programm or file
Public Function RunShellExecute(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As VbAppWinStyle) As Long
  #If ANSISupport Then
    If isUnicode Then
      RunShellExecute = ShellExecuteW(hWnd, StrPtr(lpOperation), StrPtr(lpFile), StrPtr(lpParameters), _
                                   StrPtr(lpDirectory), nShowCmd)
    Else
      RunShellExecute = ShellExecuteA(hWnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd)
    End If
  #Else
    RunShellExecute = ShellExecuteW(hWnd, StrPtr(lpOperation), StrPtr(lpFile), StrPtr(lpParameters), _
                                 StrPtr(lpDirectory), nShowCmd)
  #End If
End Function

' get all drive letters and return as an string array
Public Function getDriveLetters() As String()
  Dim ret As Long
  Dim sBuf As String * 256
 
  ret = GetLogicalDriveStringsA(Len(sBuf), sBuf)
  If ret Then
    sBuf = Left$(sBuf, ret - 1)
    getDriveLetters = Split(sBuf, vbNullChar)
  End If
End Function

' returns only a folder name, e.g.:
' getDirname "C:\Windows\System32\Drivers\" returns "Drivers"
Public Function getDirName(ByVal Path As String) As String
  Dim parts() As String
  
  If Not Right$(Path, 1) = "\" Then Path = Path & "\"
  parts = Split(Path, "\")
  getDirName = parts(UBound(parts) - 1)
End Function

' reconvert a shorten path (8.3 formatted) back to "normal"
Public Function getLongPathName(ByVal Path As String) As String
  Dim ret As Long
  Dim sBuf As String
  
  sBuf = Space$(MAX_PATH)
  #If ANSISupport Then
    If isUnicode Then
      ret = GetFullPathNameW(StrPtr(Path), MAX_PATH, StrPtr(sBuf), 0)
    Else
      ret = GetFullPathNameA(Path, MAX_PATH, sBuf, vbNullString)
    End If
  #Else
    ret = GetFullPathNameW(StrPtr(Path), MAX_PATH, StrPtr(sBuf), 0)
  #End If
  If ret Then getLongPathName = Left$(sBuf, ret)
End Function

' UAC friendly account type function
Public Function getUserAccountType() As vbzUserAccount
  Dim nBuffer As Long
  Dim bUser() As Byte
  Dim bServer() As Byte
  Dim uInfo As USER_INFO_1
   
  If winVersion > osWin98 Then
    ' convert computer and user name to byte array
    bServer = "" & vbNullChar
    bUser = glbUserName & vbNullChar
    ' get user unfo
    If NetUserGetInfo(bServer(0), bUser(0), &H1, nBuffer) = 0 Then
      MoveMemory uInfo, ByVal nBuffer, Len(uInfo)
      NetApiBufferFree nBuffer
      ' evaluate user rights
      getUserAccountType = uInfo.usri1_priv
      If winVersion > osWin2K And getUserAccountType = USER_PRIV_ADMIN Then
        If IsUserAnAdmin() = 0 Then getUserAccountType = USER_PRIV_ADMIN_RESTRICT
      End If
    Else
      getUserAccountType = USER_PRIV_GUEST
    End If  ' NetUserGetInfo
  Else
    getUserAccountType = USER_PRIV_ADMIN  ' before Win2k: always admin rights
  End If  ' winVersion > osWin98
End Function

' has user full admin rights?
Public Function isAdmin() As Boolean
  isAdmin = CBool(userAccount = USER_PRIV_ADMIN)
End Function

' tests, if string contains pure ANSI code
Public Function isAnsiString(ByVal uString As String) As Boolean
  Dim b() As Byte
  
  b = StrConv(uString, vbFromUnicode)
  isAnsiString = (uString = StrConv(b(), vbUnicode))
End Function

' tests, if a string contains unicode chars
Public Function isUnicodeString(ByVal uString As String) As Boolean
  Dim i As Long
  Dim b() As Byte
  
  If LenB(uString) Then
    b = uString
    For i = 1 To UBound(b) Step 2 ' check every 2nd byte
      If b(i) Then                ' if not 0...
        isUnicodeString = True    ' ... must be unicode
        Exit Function             ' do not test anymore
      End If
    Next
  End If
End Function

' returns True, if given string (drive, directory,
' or complete filename) belongs to a CD drive:
Public Function isCDDrive(ByVal fName As String) As Boolean
  Dim drv As String
  drv = Left$(fName, 1) & ":\"
  isCDDrive = (GetDriveType(drv) = DRV_CDROM)
End Function

' returns True, if given string (drive, directory,
' or complete filename) belongs to a network drive:
Public Function isNetFile(ByVal fName As String) As Boolean
  Dim drv As String
  
  drv = Left$(fName, 1) & ":\"
  isNetFile = (GetDriveType(drv) = DRV_REMOTE)
End Function

' check for NT compatible operating systems:
Public Function isSystemNT() As Boolean
  Dim info As OSVERSIONINFO

  info.dwOSVersionInfoSize = Len(info)
  GetVersionExA info
  isSystemNT = (info.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function isThemedWindows() As Boolean
  Dim hModule  As Long
  Dim strTheme As String
  Dim strName  As String
  
  If winVersion > osWin2K Then
    hModule = LoadLibraryA("UXTheme.dll")
    If hModule Then
      strTheme = String(255, 0)
      GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
      strTheme = Split(strTheme & vbNullChar, vbNullChar, 2)(0)
      If Len(strTheme) Then
        strName = String(255, 0)
        GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
        isThemedWindows = (Split(strName & vbNullChar, vbNullChar, 2)(0) <> "") ' special kind of NullTrim
      End If
      FreeLibrary hModule
    End If
  End If
End Function

' remove Null chars from and return a safe string
Public Function vbzNullTrim(ByVal S As String) As String
    If isUnicodeString(S) = True Then
        S = StrConv(S, vbFromUnicode)                               'Gerbing 11.06.2013
    End If
    vbzNullTrim = Split(S & vbNullChar, vbNullChar, 2)(0)
End Function

' remove Null and space chars
Public Function vbzTrim(ByVal S As String) As String
  vbzTrim = Trim$(Split(S & vbNullChar, vbNullChar, 2)(0))
End Function

' mark the whole text of an object, that must contain a text property
Public Sub obj_select(ByVal obj As Object)
  obj.SelStart = 0
  obj.SelLength = Len(obj.Text)
End Sub

' send a text with CrLf to a printer object, or an object
' which has a DC handle, to actual position
Public Sub print_out(ByVal Text As String, Optional ByVal obj As Object, _
                     Optional fmt As DT_Format = DT_WORDBREAK)
  Dim mode As Integer
  Dim R As RECT
   
  If obj Is Nothing Then Set obj = Printer
  With obj
    mode = .ScaleMode
    .ScaleMode = vbPixels
    
    If obj Is Printer Then
      .FontName = "Arial"
      .Font.Size = 10
    Else
      If fmt And DT_VCENTER Then
        .CurrentY = .ScaleHeight \ 2
      End If
    End If
    SetRect R, .CurrentX, .CurrentY, .ScaleWidth - (.CurrentX \ 2), .ScaleHeight
    #If ANSISupport Then
      If isUnicode Then
        DrawTextW .hdc, StrPtr(Text), Len(Text), R, fmt
      Else
        DrawTextA .hdc, Text, Len(Text), R, fmt
      End If
    #Else
      DrawTextW .hdc, StrPtr(Text), Len(Text), R, fmt
    #End If
    .ScaleMode = mode
  End With
End Sub

Public Sub drawUniText(ByVal Container As Object, ByVal strText As String, _
                      Optional ByVal X As Long, Optional ByVal Y As Long, _
                      Optional fmt As DT_Format = DT_WORDBREAK)
  Dim R As RECT
  
  With Container
    Set .Picture = Nothing
    SetRect R, X, Y, .ScaleWidth - X, .ScaleHeight - Y
    #If ANSISupport Then
      If isUnicode Then
        DrawTextW .hdc, StrPtr(strText), Len(strText), R, fmt
      Else
        DrawTextA .hdc, strText, Len(strText), R, fmt
      End If
    #Else
      DrawTextW .hdc, StrPtr(strText), Len(strText), R, fmt
    #End If
    Set .Picture = .image
  End With
End Sub

' play a WAV file asynchon!
Public Sub playSound(ByVal sound As String)
  On Error Resume Next
  
  #If ANSISupport Then
    If isUnicode Then
      sndPlaySoundW StrPtr(sound), SND_ASYNC
    Else
      sndPlaySoundA sound, SND_ASYNC
    End If
  #Else
    sndPlaySoundW StrPtr(sound), SND_ASYNC
  #End If
End Sub

' function collection of runDLL (as known from control panel)
Public Function runDLL(ByVal Index As Long, Optional fName As String) As Double
  Dim strParameter As String

  Select Case Index
    Case 0
      strParameter = "shell32,Control_RunDLL access.cpl,,5"
    Case 1
      strParameter = "SHELL32,OpenAs_RunDLL " & fName
    Case 2
      strParameter = "shell32,Control_RunDLL sysdm.cpl,,1"
    Case 3  ' install screen saver, properties if fName = ""
      strParameter = "DESK.CPL,InstallScreenSaver " & fName
    Case 4  ' open fonts folder (to install more fonts)
      strParameter = "shell32,SHHelpShortcuts_RunDLL FontsFolder"
    Case 5
      strParameter = "shell32,Control_RunDLL sysdm.cpl @1"
    Case 6
      strParameter = "shell32,SHHelpShortcuts_RunDLL AddPrinter"
    Case 7
      strParameter = "shell32,Control_RunDLL appwiz.cpl"
    Case 8
      strParameter = "mshtml,PrintHTML " & fName
    Case 9  ' create briefcase (for WIN version < Vista!)
      strParameter = "syncui, Briefcase_Create"
    Case 10 ' disk copy
      strParameter = "diskcopy,DiskCopyRunDll"
    Case 11
      strParameter = "shell32,Control_RunDLL joy.cpl"
    Case 12
      strParameter = "shell32,Control_RunDLL TIMEDATE.CPL,@0,0"
    Case 13
      strParameter = "shell32,Control_RunDLL main.cpl,@1,1"
    Case 14
      strParameter = "shell32,Control_RunDLL DESK.CPL,@0,1"
    Case 15
      strParameter = "shell32,Control_RunDLL NCPA.CPL,@0,2"
    Case 16
      strParameter = "shell32,Control_RunDLL mmsys.cpl"
    Case 17
      strParameter = "shell32,Control_RunDLL inetcpl.cpl users"
    Case 18
      strParameter = "shell32,Control_RunDLL ncpa.cpl,,1"
    Case 19
      strParameter = "shell32,Control_RunDLL intl.cpl,,0"
    Case 20
      strParameter = "shell32,Control_RunDLL modem.cpl"
    Case 21
      strParameter = "appwiz.cpl, NewLinkHere " & fName
    Case 22
     strParameter = "shell32, Control_FillCache_RunDLL"
    Case 23
      strParameter = "url, FileProtocolHandler " & fName
    Case 24
      strParameter = "url, MailToProtocolHandler " & fName
    Case 25
      strParameter = "url, NewsProtocolHandler " & fName
    Case Else
      Exit Function
  End Select
  runDLL = Shell("rundll32.exe " & strParameter)
End Function

' MessageBox API variant without window handle
Public Function MessageBox(ByVal lpText As String, _
       ByVal lpCaption As String, ByVal mbStyle As VbMsgBoxStyle) As VbMsgBoxResult
  Dim hWnd As Long
    
  If (mbStyle And vbSystemModal) Then
    hWnd = GetForegroundWindow  ' get the active window handle of the system
  Else
    hWnd = GetActiveWindow      ' get the active window handle of your application
  End If
  
  #If ANSISupport Then          ' allow both variants, is ANSI should be supported
    If isUnicode Then
      MessageBox = MessageBoxW(hWnd, StrPtr(lpText), StrPtr(lpCaption), mbStyle)
    Else
      MessageBox = MessageBoxA(hWnd, lpText, lpCaption, mbStyle)
    End If
  #Else                         ' fast: only UNICODE version
    MessageBox = MessageBoxW(hWnd, StrPtr(lpText), StrPtr(lpCaption), mbStyle)
  #End If
End Function

' PostMessage API with unicode support
Public Function PostMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  #If ANSISupport Then          ' allow both variants, is ANSI should be supported
    If isUnicode Then
      PostMessage = PostMessageW(hWnd, Msg, wParam, lParam)
    Else
      PostMessage = PostMessageA(hWnd, Msg, wParam, lParam)
    End If
  #Else                         ' fast: only UNICODE version
    PostMessage = PostMessageW(hWnd, Msg, wParam, lParam)
  #End If
End Function

' old function, only to keep backwards compatible
Public Function RunningIDE(ByVal hWnd As Long) As Boolean
  RunningIDE = isIDE
End Function

' Start WORD operations:
Public Function HIWORD(ByVal DWord As Long) As Long
  HIWORD = DWord \ &H10000 And &HFFFF&                  ' This is the HIWORD of the DWord
End Function
Public Function LOWORD(ByVal DWord As Long) As Long
  LOWORD = DWord And &HFFFF&                            ' This is the LOWORD of the DWord
End Function
Public Function MakeLong(ByVal LOWORD As Long, ByVal HIWORD As Long) As Long
  MakeLong = (LOWORD And &HFFFF&) Or (HIWORD * &H10000) ' Replacement for the C++ Function MAKELONG
End Function
' End WORD operations

Public Function HIBYTE(ByVal wParam As Integer) As Integer
  HIBYTE = (Abs(wParam) \ &H100) And &HFF&
End Function
Public Function LOBYTE(ByVal wParam As Integer) As Integer
  LOBYTE = Abs(wParam) And &HFF&
End Function
Public Function MAKEWORD(ByVal wLow As Integer, ByVal wHigh As Integer) As Integer
  If wHigh And &H80 Then
    MAKEWORD = (((wHigh And &H7F) * 256) + wLow) Or &H8000
  Else
    MAKEWORD = (wHigh * 256) + wLow
  End If
End Function

' ######################################################################################################## '
' # API Error Handling                                                                                   # '
' ######################################################################################################## '

Public Function getAPIError(ByVal Source As String, Optional ByVal flRaiseVBerr As Boolean) As Boolean
  Dim flags As Long, ret As Long
  Dim Msg As String
 
  apiErrorNumber = ERR.LastDllError         ' store the last DLL error, called by an API function
  If apiErrorNumber Then                    ' if we have one
    Msg = String(256, 0)                    ' try to get the description text
    flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
    #If ANSISupport Then
      If isUnicode Then
        ret = FormatMessageW(flags, 0&, apiErrorNumber, LANG_NEUTRAL, StrPtr(Msg), Len(Msg), 0&)
      Else
        ret = FormatMessageA(flags, 0&, apiErrorNumber, LANG_NEUTRAL, Msg, Len(Msg), 0&)
      End If
    #Else
      ret = FormatMessageW(flags, 0&, apiErrorNumber, LANG_NEUTRAL, StrPtr(Msg), Len(Msg), 0&)
    #End If
    If ret Then                             ' if we have a message
      apiErrorDescription = Left$(Msg, ret) ' store it in a public variable
    Else
      apiErrorDescription = "Unknown API error no. " & apiErrorNumber & " on " & Source
    End If
    getAPIError = True
    If flRaiseVBerr Then                    ' do we want to raise an error?
      ' vbCustomError indicates an API ErrorMessage, to show the "REAL"
      ' error number in a MsgBox use apiErrorNumber instead of Err.Number
      ERR.Raise vbCustomError, Source, apiErrorDescription
    End If
  Else
    apiErrorDescription = vbNullString      ' clear up the description
  End If
End Function

