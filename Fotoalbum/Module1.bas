Attribute VB_Name = "Module1"
Option Explicit
    
    Public gstrEXIFOrientation As String                                                    'Gerbing 10.05.2019
    Public gblnComeFromBeenden As Boolean                                                   'Gerbing 30.11.2018
    Public glngStrgG As Long                                                                'Gerbing 16.10.2018
    Public gstrNetzwerkDir As String                                                        'Gerbing 27.08.2017
    Public gblnDieseNachrichtNichtMehrZeigen As Boolean                                     'Gerbing 30.11.2016
    Public gstrGEOStartPunkt As String                                                      'Gerbing 05.09.2016
    Public gstrGEOEndPunkt As String                                                        'Gerbing 05.09.2016
    Public gstrGEOPosition As String                                                        'Gerbing 27.08.2015
    Public gstrLat As String                                                                'Gerbing 29.09.2018
    Public gstrLong As String                                                               'Gerbing 29.09.2018
    Public gstrLatXMP As String                                                             'Gerbing 08.04.2019
    Public gstrLongXMP As String                                                            'Gerbing 08.04.2019
    Public gblnWasOptThumbClick As Boolean                                                  'Gerbing 04.05.2015
    Public glngGridline As Long                                                             'Gerbing 19.04.2015
    Public gblnComeFromButtonF8 As Boolean                                                  'Gerbing 29.03.2015
    Public gblnWasHeadClick As Boolean                                                      'Gerbing 29.03.2015
    Public gblnComeFromThumbs As Boolean                                                    'Gerbing 29.03.2015
    Public Gefundenexifdatetimeoriginal As Boolean                                          'Gerbing 16.10.2014
    Public gblnF5Alt As Boolean                                                             'Gerbing 22.04.2014
    Public gblnComeFromF2F3 As Boolean                                                      'Gerbing 27.03.2014
    Public gstrGerbingSoftLogOrdner As String
    Public glngKommentarTop As Long                                                         'Gerbing 30.05.2013
    Public glngKommentarLeft As Long
    Public glngKommentarWidth As Long
    Public glngKommentarHeight As Long
    Public Fso As Scripting.FileSystemObject
    Public IniFso As Scripting.FileSystemObject
    Public oStream As Scripting.TextStream
    Public gblnFotosMitFET As Boolean                                                       'Gerbing 15.02.2013
    Public glngSaveForm1Width As Long                                                       'Gerbing 15.01.2012
    Public glngSaveForm1Height As Long                                                      'Gerbing 15.01.2012
    Public gblnComeFromBildanzeigen As Boolean                                              'Gerbing 31.12.2012
    Public gblnShowXYPos As Boolean                                                         'Gerbing 04.12.2012
    Public gblnComeFromAboutMenue As Boolean                                                'Gerbing 19.09.2012
    Public gblnComeFromF8 As Boolean                                                        'Gerbing 20.06.2012
    Public gblnComefromVideo As Boolean                                                     'Gerbing 16.06.2012
    Public gblnErsterStart As Boolean                                                       'Gerbing 29.10.2013
    Public gvarVideoAppid As Variant                                                        'Gerbing 21.05.2012
    Public glngForm1Hwnd As Long                                                            'Gerbing 09.05.2012
    Public glngQueryHwnd As Long                                                            'Gerbing 20.06.2012
    Public gblnIPTCAusgewählt As Boolean                                                    'Gerbing 29.03.2012
    Public gblnCheckSpeichernBildPosition As Boolean                                        'Gerbing 29.03.2012
    Public gblnWertxFormOptZufall As Boolean                                                'Gerbing 29.03.2012
    Public glngZoomProzent As Long                                                          'Gerbing 29.03.2012
    Public glngDiffX As Long            'Maß der Verschiebung mit der Maus
    Public glngDiffY As Long            'Maß der Verschiebung mit der Maus
    Public glngSaveX As Long            'rettet X zum Ausrechnen von DifferenzX für die Rechtecklupe
    Public glngSaveY As Long            'rettet Y zum Ausrechnen von DifferenzX für die Rechtecklupe
    Public DifferenzX As Long           'muss für die Rechtecklupe mit dem gdblZoomFaktor multipliziert werden
    Public DifferenzY As Long           'muss für die Rechtecklupe mit dem gdblZoomFaktor multipliziert werden
    Public gdblZoomFaktor As Double                                                         'Gerbing 29.03.2012
    Public gblnRechteckLupeScharf As Boolean                                                'Gerbing 29.03.2012
    Public gblnComeFromRechtecklupe As Boolean                                              'Gerbing 14.10.2014
    Public gblnMouseSichtbar As Boolean                                                     'Gerbing 29.03.2012
    
    Public gblnBildBeschreibung As Boolean                                                  'Gerbing 03.03.2012
    Public gstrLoggedInName As String                                                       'Gerbing 29.12.2011
    Public gblnSQLServerConnected As Boolean                                                'Gerbing 29.12.2011
    Public gblnSQLServerVersion As Boolean                                                  'Gerbing 29.12.2011
    Public gintNumberOfUsers As Integer                                                     'Gerbing 29.12.2011
    Public gstrAllowedlicenses As String                                                    'Gerbing 29.12.2011
    Public gdtDatumFotosMdb As Date                                                         'Gerbing 29.12.2011
    Public gblstrSystemDirectory As String              'Gerbing 04.12.2011
    Public gblnInstallSprache As Boolean                'Gerbing 20.11.2007
    Public gblnMsg As Boolean                                                               'Gerbing 02.09.2008
    Public AppPath As String
    Public gstrFotosMdbLocation As String
    Global NL As String                                 'Gerbing 07.11.2011
    

    Public gstrCommandLine As String
    Public PublicExportTargetFolder As String
    Public PublicLanguage As String
    Public PublicAutomaticInterval As Long                                                  'Gerbing 02.09.2008
    Public PublicBackgroundColor As String                                                  'Gerbing 22.08.2007
    Public PublicZoomToFullscreen As String                                                 'Gerbing 22.08.2007
    Public PublicPlayVideosWith As String                                                   'Gerbing 01.09.2008
    Public PublicCheckForDPI As String                                                      'Gerbing 23.06.2011
    Public PublicSQLServer As String                                                        'Gerbing 29.12.2011
    Public PublicSQLDatabase As String                                                      'Gerbing 29.12.2011
    Public PublicSQLServerUserName As String                                                'Gerbing 29.12.2011
    Public PublicWindowsAuthentication As String                                            'Gerbing 29.12.2011
    Public PublicSQLServerPassword As String                                                'Gerbing 29.12.2011
    Public PublicLocationFotos As String                                                    'Gerbing 29.12.2011
    Public PublicDatagridCaption As String                                                  'Gerbing 29.12.2011

    
    Public gstrFRODN As String
    Declare Function IsUserAnAdmin Lib "shell32" () As Long             'Gerbing 04.12.2011 08.11.2015
    
    Declare Function GetVersionEx1 Lib "kernel32.dll" Alias _
            "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO _
            ) As Long                                                                                   'Gerbing 21.05.2012

    Declare Function GetSystemDirectoryA Lib "kernel32" _
       (ByVal lpBuffer As String, ByVal nSize As Long) As Long

    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long

    Declare Function GetDesktopWindow Lib "user32" () _
            As Long                                                                                     'Gerbing 11.03.2009
    
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long                             'Gerbing 21.05.2012

    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long                      'Gerbing 21.05.2012
    
    Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long                          'Gerbing 21.05.2012

    Declare Function GetWindow Lib "user32" (ByVal hWnd _
            As Long, ByVal wCmd As Long) As Long
            
    Declare Function GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As _
            Long) As Long
            
    Declare Function GetWindowTextLength Lib "user32" _
            Alias "GetWindowTextLengthA" (ByVal hWnd As Long) _
            As Long
            
    Declare Function GetWindowText Lib "user32" Alias _
            "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString _
            As String, ByVal cch As Long) As Long
                   
    Declare Function GetParent Lib "user32" (ByVal hWnd _
            As Long) As Long
            
    Declare Function GetWindowThreadProcessId Lib "user32" _
            (ByVal hWnd As Long, lpdwProcessId As Long) As Long
            
    Declare Function timeGetTime Lib "winmm.dll" () As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
    
    ' Einige Konstantenwerte (aus WIN32API.TXT) setzen.
    Public Const conHwndTopmost = -1
    Public Const conHwndNoTopmost = -2
    Public Const conSwpNoActivate = &H10
    Public Const conSwpShowWindow = &H40
    
    Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
            (ByVal hWndOwner As Long, _
            ByVal lpszPath As String, _
            ByVal nFolder As Long, _
            ByVal fCreate As Long) As Long                      'Gerbing 19.10.2007
            
    Public Const MAX_PATH As Long = 260                         'Gerbing 19.10.2007
    Public gstrWmplayerFolder As String                         'Gerbing 19.10.2007
    Public glngVideoDuration As Long                            'Gerbing 19.10.2007

    Public Declare Function GetPrivateProfileStringW Lib "kernel32.dll" _
            (ByVal lpApplicationName As Any, _
            ByVal lpKeyName As Any, _
            ByVal lpDefault As Any, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
    
    Public Declare Function WritePrivateProfileStringW Lib "kernel32.dll" _
            (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpString As Long, _
            ByVal lpFileName As Long) As Long

    Dim FotosIniFile As String
    Dim ABSCHNITT As String * 300            'was nicht reinpaßt wird abgeschnitten
    Dim absch As String
    Dim FolderNames As String           'Abschnitt [FolderNames]
    Dim ExportTargetFolder As String
    Dim Language As String
    Dim Adjustments As Long                         'Gerbing 02.09.2008
    Dim AutomaticInterval As String                 'Gerbing 22.08.2007
    Dim BackgroundColor As String                   'Gerbing 22.08.2007
    Dim ZoomToFullscreen As String                  'Gerbing 22.08.2007
    Dim PlayVideosWith As String                    'Gerbing 01.09.2008
    Dim StartSoundFiles As String                   'Gerbing 09.12.2009
    Dim CheckForDPI As String                       'Gerbing 23.06.2011
    Dim Server As String                            'Gerbing 29.12.2011
    Dim Database As String                          'Gerbing 29.12.2011
    Dim WindowsAuthentication As String             'Gerbing 29.12.2011
    Dim UserName As String                          'Gerbing 29.12.2011
    Dim LocationFotos As String                     'Gerbing 29.12.2011
    Dim Password As String
    Dim zeichen As Integer
    Dim Zeile As String
    Dim msg As String
    Dim StartPos As Integer
    Dim temp As String
    Dim DateiNummer As Long
    Public SQLWurdeBearbeitet As Boolean
    
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Public Const HORZRES = 8
    Public Const VERTRES = 10
'    Public screenWidth As Long
'    Public screenHeight As Long
    Public gstrDateChoose As Date
    Public gblnSchreibgeschützt As Boolean
    Public gblnWeiterMitLeererDatenbank As Boolean
    Public gblnWeiterMitAnderemFotosStandort As Boolean                                         'Gerbing 29.12.2011
    Public glngStartMillisek As Long
    Public glngEndMillisek As Long
    Public gblnVollversion As Boolean
    Public gblnProversion As Boolean
    Public gintDiffTage As Integer
    Public gstrFreischalteschlüssel As String
    Public gblnfrmGridAndThumbDblClick As Boolean        'Gerbing 06.06.2005
    Public Sprache As Long                          'Gerbing 08.11.2005
    Public gstrRowColChangeName As String           'Gerbing 04.01.2006
    
    Public Type Bildposition                        'Gerbing 29.07.2006
        Top As Long
        Left As Long
        ZoomPercent As Long                         'Gerbing 29.03.2012
        Dateiname As String * 1000
    End Type
    Public BildPosList() As Bildposition
    
    Public Type TrefferProJahr                      'Gerbing 09.02.2013
        Jahr As String * 4
        Anzahl As Long
    End Type
    Public TrefferProJahrList() As TrefferProJahr

    Declare Function GetVersionEx Lib "kernel32" Alias _
        "GetVersionExA" (lpVersionInformation As _
         OSVERSIONINFO) As Long

    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNummer As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
    End Type
    
    Public OSInfo As OSVERSIONINFO
    'Public HyperlinkFieldColumns As New Collection
    Public Const CSIDL_APPDATA As Long = &H1A&  '' <username>\Application Data              'Gerbing 20.12.2010
    Public gstrFotosIniAnwendungsOrdner As String                                           'Gerbing 20.12.2010
    
    '-----------Gerbing 23.06.2011------------------------------------------
    Private Declare Function GetDC Lib "user32" ( _
            ByVal hWnd As Long) As Long
     
    Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hDC As Long) As Long
      
    
    Public walto As Long                                                                    'Gerbing 26.10.2011
    Public wancho As Long                                                                   'Gerbing 26.10.2011
    
    Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    
    Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type
    
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'KPDTeam@Allapi.net
    Public m_Date As Date
    Public lngHandle As Long
    Public udtFileTime As FILETIME
    Public udtLocalTime As FILETIME
    Public udtSystemTime As SYSTEMTIME
    
    Public Const GENERIC_WRITE = &H40000000
    Public Const OPEN_EXISTING = 3
    Public Const FILE_SHARE_READ = &H1
    Public Const FILE_SHARE_WRITE = &H2
    Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    
    Public oCat As ADOX.Catalog                                                         'Gerbing 16.10.2014
'    Public DBsql As ADODB.Connection
'    Public DollarDBsql As ADODB.Connection
    Public DBado As ADODB.Connection                                                    'Gerbing 23.11.2017
    Public DollarDBado As ADODB.Connection                                              'Gerbing 23.11.2017
    Public rstsql As ADODB.Recordset
    
    '-----------Gerbing 29.03.2012------------------------------------------
 
    Public Declare Function SetTimer Lib "user32.dll" ( _
      ByVal hWnd As Long, _
      ByVal nIDEvent As Long, _
      ByVal uElapse As Long, _
      ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32.dll" ( _
      ByVal hWnd As Long, _
      ByVal nIDEvent As Long) As Long
    Public Const WM_TIMER = &H113 ' Timer-Ereignis trifft ein
    Public hEvent As Long
    
    Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2B&  'Gerbing 19.10.2007
    
    Private Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
     
    Private Declare Function ShowWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nCmdShow As Long) As Long
     
'    Private Declare Function GetDesktopWindow Lib "user32" () As Long
     
    Private Declare Function LockWindowUpdate Lib "user32" ( _
        ByVal hwndLock As Long) As Long
     
    Private Const SW_HIDE = 0
    Private Const SW_SHOW = 5
    Private Const GWL_EXSTYLE = (-20)
     
    Private Const WS_EX_APPWINDOW = &H40000
    Public udtData As GdiplusStartupInput                           'Gerbing 27.08.2012
    '---------------------------------------------------------------'Gerbing 04.09.2012
    Private Const GWL_STYLE = (-16)
    Private Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
    Private Const WS_MAXIMIZEBOX = &H10000
    Private Const WS_MINIMIZEBOX = &H20000
    Private Const WS_SYSMENU = &H80000
    
    
    Private Enum ESetWindowPosStyles
        SWP_SHOWWINDOW = &H40
        SWP_HIDEWINDOW = &H80
        SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
        SWP_NOACTIVATE = &H10
        SWP_NOCOPYBITS = &H100
        SWP_NOMOVE = &H2
        SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
        SWP_NOREDRAW = &H8
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSIZE = &H1
        SWP_NOZORDER = &H4
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        HWND_NOTOPMOST = -2
    End Enum
    
    Private Declare Function GetWindowRect Lib "user32" ( _
          ByVal hWnd As Long, lpRect As RECT) As Long
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    '-----------Gerbing 04.03.2013------------------------------------------
    Public gblnIsUni As Boolean
    Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

    'Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
    'Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
    
    Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
    Public Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Public Const MB_ICONINFORMATION As Long = &H40&
    Public Const MB_TASKMODAL As Long = &H2000&
    
    Public Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long 'Gerbing 21.01.2018

Sub ReadFotosIniFile()
    Dim j As Long
    
    j = GetSystemDirectoryA("", 0)
    gblstrSystemDirectory = Space(j - 1)
    Call GetSystemDirectoryA(gblstrSystemDirectory, j)

    gstrGerbingSoftLogOrdner = AppPath                         'Gerbing 10.08.2017
    'ab Version 14.0.0 c:\Users\Public\Documents\GERBING Fotoalbum 15\gerbingsoft.log
    'ab Version 15.0.1 AppPath
    gstrFotosIniAnwendungsOrdner = AppPath
    'im XP          x:\Dokumente und Einstellungen\user\Anwendungsdaten
    'im Windows7    C:\Users\gottfried\AppData\Roaming
    'ab Version 14.0.0 AppPath
    'FotosIniFile = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15\fotos.ini"      'Pfad der fotos.ini                 Gerbing 20.12.2010
    FotosIniFile = gstrFotosIniAnwendungsOrdner & "\fotos.ini"      'Pfad der fotos.ini                 Gerbing 20.12.2010
    On Error Resume Next
    If file_path_exist(FotosIniFile) = False Then   'prüfe ob fotos.ini existiert
    'Standardwerte eintragen, wenn es keine Datei fotos.ini gibt    'Gerbing 27.09.2010
        PublicLanguage = "9"                      'Gerbing 27.09.2010   'Gerbing 04.12.2011
        PublicAutomaticInterval = "3"            'Gerbing 27.09.2010    'Gerbing 04.12.2011
        PublicBackgroundColor = "black"         'Gerbing 27.09.2010
        PublicZoomToFullscreen = "0"              'Gerbing 27.09.2010
        PublicPlayVideosWith = "10"             'Gerbing 26.10.2011
        PublicCheckForDPI = "1"                 'Gerbing 23.06.2011
        Exit Sub
    End If
    '------------------------------------------------------------------------------
    Call EZV                                'Prüfe [FolderNames] ExportTargetFolder
    Call GlL                                'Prüfe [Global] Language
    Call AAI                                'Prüfe [Adjustments] AutomaticInterval          'Gerbing 22.08.2007
    Call ABC                                'Prüfe [Adjustments] BackgroundColor            'Gerbing 22.08.2007
    Call AZF                                'Prüfe [Adjustments] ZoomToFullscreen           'Gerbing 22.08.2007
    Call MPVW                               'Prüfe [Mediaplayer] PlayVideoWith              'Gerbing 01.09.2008
    Call DPI                                'Prüfe [Adjustments] CheckForDPI                'Gerbing 23.06.2011
    Call SSRV                               'Prüfe [SQL] Server
    Call SDB                                'Prüfe [SQL] Database
    Call SWA                                'Prüfe [SQL] WindowsAuthentication
    Call SUN                                'Prüfe [SQL] username
    Call SPW                                'Prüfe [SQL] password
    Call SLF                                'Prüfe [SQL] LocationFotos
End Sub

Public Function INIReadString(ByVal Section As String, ByVal key As String, ByVal Default As String, ByVal FileName As String) As String
    Dim cSize As Long
    Dim strReturn As String
    Dim RetVal As Long
    
    cSize = 300
    strReturn = String(cSize, 0)
    RetVal = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(strReturn), cSize, StrPtr(FileName))
    If RetVal > 0 Then
        INIReadString = Left(strReturn, RetVal)
    End If
End Function

Sub EZV()
    'Prüfe [FolderNames] ExportTargetFolder
    absch = "FolderNames"
    ExportTargetFolder = "ExportTargetFolder"
    Zeile = INIReadString(absch, ExportTargetFolder, "", FotosIniFile)
    PublicExportTargetFolder = Zeile
End Sub

Sub GlL()
    'Prüfe [Global] Language
    absch = "Global"
    Language = "Language"
    Zeile = INIReadString(absch, Language, "", FotosIniFile)
    If Zeile = "" Then Sprache = 0            'Deutsch
    If Zeile = "0" Then Sprache = 0            'Deutsch
    If Zeile = "1" Then Sprache = 3000         'English
    'If Dir(AppPath & "\fotos.mdb") <> "" Then
    If file_path_exist(AppPath & "\fotos.mdb") = True Then
        'Nur bei der Access-Shareware-Version ist es nötig, daß beim ersten Start von fotos.exe Language = "9" ist
        'nur dann wird msdmo.log erzeugt
        'mit Hilfe des Alters von msdmo.log nerve ich die Shareware-Nutzer mit Einblendung des Shareware-Hinweises
        'Das Datum 30.12.2011 ist das Datum der Fotos.mdb im Auslieferungszustand
        If DatePart("d", gdtDatumFotosMdb) = 30 And DatePart("m", gdtDatumFotosMdb) = 12 And DatePart("yyyy", gdtDatumFotosMdb) = 2011 Then
            Zeile = "9"
        End If
    End If
    PublicLanguage = Zeile
End Sub

Sub AAI()
    'Prüfe [Adjustments] AutomaticInterval
    On Error Resume Next
    absch = "Adjustments"
    AutomaticInterval = "AutomaticInterval"
    PublicAutomaticInterval = 3
    Zeile = INIReadString(absch, AutomaticInterval, "", FotosIniFile)
    If Zeile = "" Or Zeile = "0" Then                                         'Gerbing 02.09.2008
        PublicAutomaticInterval = 3
    End If
    If Not IsNumeric(Zeile) Then
        PublicAutomaticInterval = 3
    End If
    If Zeile < 1 Then PublicAutomaticInterval = 3             '3=Standard
    If Zeile > 60 Then PublicAutomaticInterval = 3
    If Zeile <> "3" Then
        PublicAutomaticInterval = CLng(Zeile)
    End If
End Sub

Sub ABC()
    Dim strTemp As String
    
    'Prüfe [Adjustments] BackgroundColor
    absch = "Adjustments"
    BackgroundColor = "BackgroundColor"
    Zeile = INIReadString(absch, BackgroundColor, "", FotosIniFile)
    strTemp = "black"
    If StrComp(Zeile, "gray", vbTextCompare) = 0 Then
        strTemp = "gray"
    End If
    If StrComp(Zeile, "black", vbTextCompare) = 0 Then
        strTemp = "black"
    End If
    PublicBackgroundColor = strTemp
End Sub

Sub MPVW()
    Dim strTemp As String
    Dim Pos As Long                                                     'Gerbing 28.12.2011
    Dim FileName As String

    'Prüfe [Mediaplayer] PlayVideosWith
    absch = "Mediaplayer"
    PlayVideosWith = "PlayVideosWith"
    Zeile = INIReadString(absch, PlayVideosWith, "", FotosIniFile)
    PublicPlayVideosWith = Zeile                               'Gerbing 28.12.2011
    Zeile = LCase(Zeile)
    Pos = InStr(1, Zeile, "exe", vbTextCompare)
    Select Case Zeile
        Case ""                                                         'Gerbing 18.01.2013
            WriteMPVW ("10")
            PublicPlayVideosWith = "10"
        Case "10"
            '
        Case "external"
            'ermitteln wo der Windows Mediaplayer steht
            gstrWmplayerFolder = getSpecialFolder(CSIDL_PROGRAM_FILES_COMMONX86)
            'liefert zB C:\Programme\Gemeinsame Dateien
            Pos = InStr(4, gstrWmplayerFolder, "\")
            If Pos <> 0 Then
                gstrWmplayerFolder = Left(gstrWmplayerFolder, Pos) & "Windows Media Player" & "\wmplayer.exe"
                'strTemp = Dir(gstrWmplayerFolder)
                'If strTemp = "" Then Pos = 0
                If file_path_exist(gstrWmplayerFolder) = False Then Pos = 0
            End If
            If Pos = 0 Then
                FileName = ShowOpenUnicodeExternalVideoPlayer(Form1)
                FileName = ConvertFileName(FileName)
                Pos = InStr(1, FileName, "wmplayer.exe")
                If Pos = 0 Then
                    strTemp = "10"
                End If
                gstrWmplayerFolder = FileName
            End If
        Case Else
'            strTemp = Dir(zeile)                     'Gerbing 18.01.2013
'            If strTemp = "" Then
            If file_path_exist(PublicPlayVideosWith) = False Then
                msg = "error - external video player not found" & vbNewLine
                msg = msg & "settings - video are replaced to internal (10)"
                MsgBox msg
                WriteMPVW ("10")
                PublicPlayVideosWith = "10"
                Exit Sub
            End If
            gstrWmplayerFolder = PublicPlayVideosWith
    End Select
End Sub

Sub AZF()
    'Prüfe [Adjustments] ZoomToFullscreen
    absch = "Adjustments"
    ZoomToFullscreen = "ZoomToFullscreen"
    Zeile = INIReadString(absch, ZoomToFullscreen, "", FotosIniFile)
    Adjustments = 0
    If Not IsNumeric(Zeile) Then
        Adjustments = 0
    Else
        If Zeile = "0" Then
            Adjustments = 0
        End If
        If Zeile = "1" Then                                                             'Gerbing 29.04.2013
            Adjustments = 1
        End If
        If Zeile = "640" Then
            Adjustments = 640
        End If
        If Zeile = "800" Then
            Adjustments = 800
        End If
        If Zeile = "1024" Then
            Adjustments = 1024
        End If
        If Zeile = "1024768" Then                                                      'Gerbing 16.03.2009
            Adjustments = 1024768
        End If
    End If
    PublicZoomToFullscreen = Adjustments
End Sub

Sub DPI()                                                                                       'Gerbing 09.12.2009
    'Prüfe [Adjustments] CheckForDPI
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    Zeile = INIReadString(absch, CheckForDPI, "", FotosIniFile)
    If Zeile = "" Then PublicCheckForDPI = "1"            '1=berücksichtigen
    If Zeile = "0" Then PublicCheckForDPI = "0"             '0=ignorieren
    If Zeile = "1" Then PublicCheckForDPI = "1"             '1=berücksichtigen
    If Zeile = "2" Then PublicCheckForDPI = "2"             '2=berücksichtigen                  'Gerbing 11.03.2017
    If Zeile = "3" Then PublicCheckForDPI = "3"             '3=berücksichtigen                  'Gerbing 11.03.2017
End Sub

Sub SSRV()
    'Prüfe [SQL] Server
    absch = "SQL"
    Server = "Server"
    Zeile = INIReadString(absch, Server, "", FotosIniFile)
    PublicSQLServer = Zeile
End Sub

Sub SDB()
    'Prüfe [SQL] Database
    absch = "SQL"
    Database = "Database"
    Zeile = INIReadString(absch, Database, "", FotosIniFile)
    PublicSQLDatabase = Zeile
End Sub

Sub SWA()
    'Prüfe [SQL] WindowsAuthentication
    absch = "SQL"
    WindowsAuthentication = "WindowsAuthentication"
    Zeile = INIReadString(absch, WindowsAuthentication, "", FotosIniFile)
    PublicWindowsAuthentication = Zeile
End Sub

Sub SUN()
    'Prüfe [SQL] username
    absch = "SQL"
    UserName = "username"
    Zeile = INIReadString(absch, UserName, "", FotosIniFile)
    PublicSQLServerUserName = Zeile
End Sub

Sub SPW()
    'Prüfe [SQL] password
    absch = "SQL"
    Password = "password"
    Zeile = INIReadString(absch, Password, "", FotosIniFile)
    PublicSQLServerPassword = Zeile
End Sub

Sub SLF()
    'Prüfe [SQL] LocationFotos
    absch = "SQL"
    LocationFotos = "LocationFotos"
    Zeile = INIReadString(absch, LocationFotos, "", FotosIniFile)
    PublicLocationFotos = Zeile
End Sub

Sub WriteEZV(NeuerInhalt As String)
    'Schreibe [FolderNames] ExportTargetFolder
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "FolderNames"
    ExportTargetFolder = "ExportTargetFolder"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(ExportTargetFolder), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteGlL(NeuerInhalt As String)
    'Schreibe [Global] Language
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Global"
    Language = "Language"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(Language), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteMPVW(NeuerInhalt As String)
    'Schreibe [Mediaplayer] PlayVideosWith
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Mediaplayer"
    PlayVideosWith = "PlayVideosWith"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(PlayVideosWith), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteAAI(NeuerInhalt As String)                                                         'Gerbing 22.08.2007
    'Schreibe [Adjustments] AutomaticInterval
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    AutomaticInterval = "AutomaticInterval"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(AutomaticInterval), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteABC(NeuerInhalt As String)                                                        'Gerbing 22.08.2007
    'Schreibe [Adjustments] BackgroundColor
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    BackgroundColor = "BackgroundColor"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(BackgroundColor), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteAZF(NeuerInhalt As String)                                                        'Gerbing 22.08.2007
    'Schreibe [Adjustments] ZoomToFullscreen
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    ZoomToFullscreen = "ZoomToFullscreen"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(ZoomToFullscreen), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteCDS(NeuerInhalt As String)                                                        'Gerbing 08.03.2020
    'Schreibe [CheckCompactDatabase] CompactDatabaseStarted
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "CheckCompactDatabase"
    ZoomToFullscreen = "CompactDatabaseStarted"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(ZoomToFullscreen), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteCDE(NeuerInhalt As String)                                                        'Gerbing 08.03.2020
    'Schreibe [CheckCompactDatabase] CompactDatabaseEnded
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "CheckCompactDatabase"
    ZoomToFullscreen = "CompactDatabaseEnded"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(ZoomToFullscreen), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSSF(NeuerInhalt As String)                                                        'Gerbing 09.12.2009
    'Schreibe [Adjustments] StartSoundFiles
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    StartSoundFiles = "StartSoundFiles"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(StartSoundFiles), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteDPI(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [Adjustments] CheckForDPI
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(CheckForDPI), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSSRV(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] Server
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    Server = "Server"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(Server), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSDB(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] Database
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    Database = "Database"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(Database), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSWA(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] WindowsAuthentication
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    WindowsAuthentication = "WindowsAuthentication"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(WindowsAuthentication), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSUN(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] username
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    UserName = "username"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(UserName), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSPW(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] password
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    Password = "password"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(Password), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteSLF(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] LocationFotos
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "SQL"
    LocationFotos = "LocationFotos"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(LocationFotos), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub FRODateinameRstsql()
    gstrFRODN = Replace(rstsql(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\") '1028=Dateiname  'Gerbing 07.11.2011
    'gstrFRODN = Replace(frmGridAndThumb.rsDataGrid.Fields("dateiname"), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
    'gstrFRODN = Replace(frmGridAndThumb.rsDataGrid.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
End Sub

Sub FRODateiname()
    'gstrFRODN = Replace(rstsql(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\") '1028=Dateiname  'Gerbing 07.11.2011
    'gstrFRODN = Replace(frmGridAndThumb.rsDataGrid.Fields("dateiname"), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
    gstrFRODN = Replace(frmGridAndThumb.rsDataGrid.Fields(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\")  '1028=Dateiname  'Gerbing 07.11.2011
End Sub

Function Crypt(Inp$, key$, mode As Boolean) As String
    Dim z As String
    Dim i As Integer, Position As Integer
    Dim cptZahl As Long, orgZahl As Long
    Dim keyZahl As Long, cptString As String
    
    On Error GoTo CryptError
    For i = 1 To Len(Inp)
        Position = Position + 1
        If Position > Len(key) Then Position = 1
        keyZahl = Asc(Mid(key, Position, 1))
        
        If mode Then
            
            'Verschlüsseln
            orgZahl = Asc(Mid(Inp, i, 1))
            cptZahl = orgZahl Xor keyZahl
            cptString = Hex(cptZahl)
            If Len(cptString) < 2 Then cptString = "0" & cptString
            z = z & cptString
            
        Else
            
            'Entschlüsseln
            If i > Len(Inp) \ 2 Then Exit For
            cptZahl = CByte("&H" & Mid$(Inp, i * 2 - 1, 2))
            orgZahl = cptZahl Xor keyZahl
            z = z & Chr$(orgZahl)
            
        End If
    Next i
    On Error GoTo 0
    Crypt = z
    Exit Function
CryptError:
    Crypt = "error"
End Function

Public Function GetShellError(lErrorCode As Long) As String
    Const SE_ERR_FNF = 2&, SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&, SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&, SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&, SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&, SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&, ERROR_BAD_FORMAT = 11&

    Select Case lErrorCode
        Case SE_ERR_FNF
            GetShellError = "File not found"
        Case SE_ERR_PNF
            GetShellError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            GetShellError = "Access denied"
        Case SE_ERR_OOM
            GetShellError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            GetShellError = "DLL not found"
        Case SE_ERR_SHARE
            GetShellError = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            GetShellError = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            GetShellError = "DDE Time out"
        Case SE_ERR_DDEFAIL
            GetShellError = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            GetShellError = "DDE busy"
        Case SE_ERR_NOASSOC
            GetShellError = "No association for file extension"
        Case ERROR_BAD_FORMAT
            GetShellError = "Invalid EXE file or error in EXE image"
        Case Else
            GetShellError = "Unknown error"
    End Select
End Function

Public Sub AudioDateiMitkopieren(Quellname, Zielname)
    Dim WavQuellname As String
    Dim Mp3Quellname As String
    Dim WavZielname As String
    Dim Mp3Zielname As String
    Dim rc As Boolean

    'FileCopy Quellname, Zielname
    'wenn es eine gleichnamige .wav oder .mp3 Datei gibt wie Quellname                      'Gerbing 26.09.2007
    'dann wird diese mitkopiert
    WavQuellname = Left(Quellname, Len(Quellname) - 4) & ".wav"
    Mp3Quellname = Left(Quellname, Len(Quellname) - 4) & ".mp3"
    WavZielname = Left(Zielname, Len(Zielname) - 4) & ".wav"
    Mp3Zielname = Left(Zielname, Len(Zielname) - 4) & ".mp3"
    On Error Resume Next
    'If Dir(WavQuellname) <> "" Then
    If file_path_exist(WavQuellname) = True Then
        'FileCopy WavQuellname, WavZielname
        rc = file_copy(WavQuellname, WavZielname)
    End If
    'If Dir(Mp3Quellname) <> "" Then
    If file_path_exist(Mp3Quellname) = True Then
        'FileCopy Mp3Quellname, Mp3Zielname
        rc = file_copy(Mp3Quellname, Mp3Zielname)
    End If
    On Error GoTo 0
End Sub

Public Sub AudioDateiMitUmnennen(Quellname, Zielname)                                       'Gerbing 26.03.2018
    Dim WavQuellname As String
    Dim Mp3Quellname As String
    Dim WavZielname As String
    Dim Mp3Zielname As String
    Dim rc As Boolean

    'FileCopy Quellname, Zielname
    'wenn es eine gleichnamige .wav oder .mp3 Datei gibt wie Quellname
    'dann wird diese mit umgenannt
    WavQuellname = Left(Quellname, Len(Quellname) - 4) & ".wav"
    Mp3Quellname = Left(Quellname, Len(Quellname) - 4) & ".mp3"
    WavZielname = Left(Zielname, Len(Zielname) - 4) & ".wav"
    Mp3Zielname = Left(Zielname, Len(Zielname) - 4) & ".mp3"
    On Error Resume Next
    'If Dir(WavQuellname) <> "" Then
    If file_path_exist(WavQuellname) = True Then
        'FileCopy WavQuellname, WavZielname
        rc = NameAs(WavQuellname, WavZielname)
    End If
    'If Dir(Mp3Quellname) <> "" Then
    If file_path_exist(Mp3Quellname) = True Then
        'FileCopy Mp3Quellname, Mp3Zielname
        rc = NameAs(Mp3Quellname, Mp3Zielname)
    End If
    On Error GoTo 0
End Sub

'Public Function getSpecialFolder(ByVal FolderID As Long) As String
'    Dim sBuffer As String
'
'    'SHGetSpecialFolderPath is not necessarily supported under Win95. It canbe called
'    'under Win95 if IE4 AND the Desktop Update are installed.
'    'Cheap-shot error handling
'    'If running under Win95 or NT4, merely ignore the error
'    'This will result in a 0-length string being returned.
'    On Error Resume Next
'    sBuffer = String$(MAX_PATH, vbNullChar)
'    'For Win98 and higher
'    Call SHGetSpecialFolderPath(0&, sBuffer, FolderID, 0&)
'    getSpecialFolder = StripNulls(sBuffer)
'End Function

Public Function StripNulls(ByVal sText As String) As String
    'Returns all characters up to a null character.
    'If the string does not contain a null character,
    'the string is returned unmodified.
    Dim lNullPos As Long

    lNullPos = InStr(sText, vbNullChar)
    If lNullPos Then
        StripNulls = Left$(sText, lNullPos - 1)
    Else
        StripNulls = sText
    End If
End Function

Public Sub VierUrsachenFürSchreibsperre()                                                'Gerbing 02.09.2008
    'If gblnMsg = False Then Exit Sub
    'vier mögliche Ursachen
    'Msg = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15\fotos.ini" & vbNewLine
    msg = gstrFotosIniAnwendungsOrdner & "\fotos.ini" & vbNewLine
    'msg = msg & "Die Datei ist schreibgeschützt. Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können." & vbnewline
    msg = msg & LoadResString(2275 + Sprache) & vbNewLine
    'msg = msg & "Es gibt vier mögliche Ursachen für den Lesemodus:" & vbnewline
    msg = msg & LoadResString(2133 + Sprache) & vbNewLine
    'msg = msg & "1. Das Dateiattribut 'Schreibgeschützt' ist gesetzt" & vbnewline
    msg = msg & LoadResString(2134 + Sprache) & vbNewLine
    'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte für Ihren PC" & vbnewline
    msg = msg & LoadResString(2135 + Sprache) & vbNewLine
    'msg = msg & "3. Sie arbeiten mit einer CD oder DVD" & vbnewline
    msg = msg & LoadResString(2136 + Sprache) & vbNewLine
    'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & vbnewline & vbnewline
    msg = msg & LoadResString(2137 + Sprache) & vbNewLine & vbNewLine
    'MsgBox Msg, , LoadResString(1119 + Sprache)
    MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
End Sub

Public Function AnpassenNutzerWunsch(Form)                                              'Gerbing 11.03.2017
    Dim i As Long
    
    'PublicCheckForDPI = 1 = klein
    'PublicCheckForDPI = 2 = mittel
    'PublicCheckForDPI = 3 = groß
    On Error Resume Next
    If PublicCheckForDPI = 1 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = Form1.txtFont.FontName
            Form.Controls(i).Font.Size = 8
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 2 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = Form1.txtFont.FontName
            Form.Controls(i).Font.Size = 10
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 3 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = Form1.txtFont.FontName
            Form.Controls(i).Font.Size = 12
        Next i
    End If
    On Error GoTo 0
End Function

Public Function AnpassenHeadFont(DbGrid)                                                'Gerbing 11.03.2017
    On Error Resume Next
    DbGrid.HeadFont.Bold = False
    DbGrid.HeadFont.Name = Form1.txtFont.FontName
    If PublicCheckForDPI = 1 Then
            DbGrid.HeadFont.Size = 8
            DbGrid.RowHeight = 220 \ Screen.TwipsPerPixelY
    End If
    If PublicCheckForDPI = 2 Then
            DbGrid.HeadFont.Size = 10
            DbGrid.RowHeight = 260 \ Screen.TwipsPerPixelY
    End If
    If PublicCheckForDPI = 3 Then
            DbGrid.HeadFont.Size = 12
            DbGrid.RowHeight = 300 \ Screen.TwipsPerPixelY
    End If
    On Error GoTo 0
End Function

' Startet den Timer
Public Function EnableTimer(ByVal msInterval As Long)
    If hEvent <> 0 Then Exit Function
    hEvent = SetTimer(0&, 0&, msInterval, AddressOf Timer1Ersatz)
End Function

' Beendet den Timer
Public Function DisableTimer()
    On Error Resume Next                                                                    'Gerbing 29.03.2015
    If hEvent = 0 Then Exit Function
    KillTimer 0&, hEvent
    hEvent = 0
End Function

' Timer-Prozedur, welche im Abstand der festgelegten
' Millisekunden ein Ereignis sendet
Public Sub Timer1Ersatz(ByVal hWnd As Long, ByVal uMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long)
     Dim st As SYSTEMTIME
     If uMsg = WM_TIMER Then
        On Error Resume Next                                                                'Gerbing 11.08.2012
        frmGridAndThumb.Hide
        Hilfebx.Hide
        'KommentarForm.Hide                                                                 'Gerbing 05.08.2013
        On Error GoTo 0
        Call Form1.GeheEinBildVorwärts
        Call Form1.BildAnzeigen
     End If
End Sub

Public Function ShowTitleBar(ByVal bState As Boolean, blnIsFoto As Boolean)                   'Gerbing 04.09.2012
    Dim lStyle As Long
    Dim tR As RECT

   'blnIsFoto = True bei Foto=Form1
   'blnIsFoto = False bei Video=frmVideo
   ' Get the window's position:
    If blnIsFoto = True Then                                 'es ist Foto
       ' Get the window's position:
       GetWindowRect Form1.hWnd, tR
       ' Modify whether title bar will be visible:
       lStyle = GetWindowLong(Form1.hWnd, GWL_STYLE)
       If (bState) Then
          'Form1.Caption = Form1.FotoAlbumTitle
          'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
          formCaption Form1.hWnd, Form1.FotoAlbumTitle
          If Form1.ControlBox Then
             lStyle = lStyle Or WS_SYSMENU
          End If
          If Form1.MaxButton Then
             lStyle = lStyle Or WS_MAXIMIZEBOX
          End If
          If Form1.MinButton Then
             lStyle = lStyle Or WS_MINIMIZEBOX
          End If
          If Form1.Caption <> "" Then
             lStyle = lStyle Or WS_CAPTION
          End If
       Else
          'Form1.Caption = ""
          formCaption Form1.hWnd, ""
          lStyle = lStyle And Not WS_SYSMENU
          lStyle = lStyle And Not WS_MAXIMIZEBOX
          lStyle = lStyle And Not WS_MINIMIZEBOX
          lStyle = lStyle And Not WS_CAPTION
       End If
       SetWindowLong Form1.hWnd, GWL_STYLE, lStyle
       ' Ensure the style takes and make the window the
       ' same size, regardless that the title bar etc
       ' is now a different size:
       SetWindowPos Form1.hWnd, _
           0, tR.Left, tR.Top, _
           tR.Right - tR.Left, tR.Bottom - tR.Top, _
           SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
       Form1.Refresh
       ' Ensure that your resize code is fired, as the client area
       ' has changed:
       Form1.Form_Resize
    Else                                                            'es ist Video
       ' Get the window's position:
       GetWindowRect frmVideo.hWnd, tR
       ' Modify whether title bar will be visible:
       lStyle = GetWindowLong(frmVideo.hWnd, GWL_STYLE)
       If (bState) Then
            formCaption frmVideo.hWnd, Form1.FotoAlbumTitle                         'Gerbing 02.04.2018
            If frmVideo.ControlBox Then
               lStyle = lStyle Or WS_SYSMENU
            End If
            If frmVideo.MaxButton Then
               lStyle = lStyle Or WS_MAXIMIZEBOX
            End If
            If frmVideo.MinButton Then
               lStyle = lStyle Or WS_MINIMIZEBOX
            End If
            If frmVideo.Caption <> "" Then
               lStyle = lStyle Or WS_CAPTION
            End If
       Else
            frmVideo.Caption = ""
            lStyle = lStyle And Not WS_SYSMENU
            lStyle = lStyle And Not WS_MAXIMIZEBOX
            lStyle = lStyle And Not WS_MINIMIZEBOX
            lStyle = lStyle And Not WS_CAPTION
       End If
       SetWindowLong frmVideo.hWnd, GWL_STYLE, lStyle
       ' Ensure the style takes and make the window the
       ' same size, regardless that the title bar etc
       ' is now a different size:
       
       SetWindowPos frmVideo.hWnd, _
           0, tR.Left, tR.Top, _
           tR.Right - tR.Left, tR.Bottom - tR.Top, _
           SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
       
'auskommentiert Gerbing 27.11.2016       SetWindowPos frmVideo.hwnd, _
'           0, tR.Left, tR.Top, _
'           tR.Right - tR.Left, tR.Bottom - tR.Top, _
'           SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
       
       frmVideo.Refresh
       ' Ensure that your resize code is fired, as the client area
       ' has changed:
       frmVideo.Form_Resize
    End If
End Function

Public Sub HelpFileErrorMsg(RetVal As Long, CHMFile As String)
    Dim DateinamenErweiterung As String
    Dim ErrorText As String
    Dim msg As String

    DateinamenErweiterung = "CHM"
    ErrorText = GetShellError(RetVal)           'Gerbing 20.08.2008
    msg = "Errortext=" & ErrorText & vbNewLine
    msg = msg & "Errornr=" & RetVal & vbNewLine & vbNewLine
    
    msg = msg & CHMFile & vbNewLine
    'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
    msg = msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
    
    'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
    msg = msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
    
    'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
    msg = msg & LoadResString(1378 + Sprache) & vbNewLine
    'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
    msg = msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
    'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
    msg = msg & LoadResString(2012 + Sprache) & vbNewLine
    'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
    msg = msg & LoadResString(2013 + Sprache)
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotosmdb"), vbInformation
End Sub

'From MSDN . . .
'-----------------------------------------------
'Jet OLEDB:Engine Type  Jet x.x Format MDB Files
'---------------------  ------------------------
'       1                     JET10
'       2                     JET11
'       3                     JET2X
'       4                     JET3X
'       5                     JET4X
'-----------------------------------------------
 
Public Function CompactDB(ByVal sSource As String, ByVal sDest As String) As Boolean
    'Ersatz für DBEngine.CompactDatabase                                            'Gerbing 23.11.2017
    'Requires references to:
    ' Microsoft Jet and Replication Objects 2.1 Library (or higher)
    ' Microsoft ActiveX Data Objects 2.5 Library (or higher)
    Dim iEngineType As Integer
    Dim JRO As JRO.JetEngine
    Dim cn As ADODB.Connection
 
   On Error GoTo CompactDB_Error
    'sSource = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sSource
    sSource = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & sSource       'Gerbing 23.11.2017
    ' Find the engine type to use when compacting database
    Set cn = New ADODB.Connection
    With cn
        .Open sSource
       iEngineType = .Properties("Jet OLEDB:Engine Type")
       .Close
    End With
    Set cn = Nothing
    'sDest = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Jet OLEDB:Engine Type=" & iEngineType & ";Data Source=" & sDest
    sDest = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Jet OLEDB:Engine Type=" & iEngineType & ";Data Source=" & sDest 'Gerbing 23.11.2017
    Set JRO = New JRO.JetEngine
    JRO.CompactDatabase sSource, sDest
    CompactDB = True
    Set JRO = Nothing
 
    On Error GoTo 0
    Exit Function
CompactDB_Error:
    CompactDB = False
'    Msg = "Error " & Err.Number & " (" & Err.Description & ")" & " in procedure JRO.CompactDatabase" & vbNewLine '16.12.2017
'    Msg = Msg & "Source=" & sSource & vbNewLine
'    Msg = Msg & "Destination=" & sDest
'    MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CompactDB of Module Module1"
    ' Clean up any "junk" left behind
    On Error Resume Next
    Set cn = Nothing
    Set JRO = Nothing
    On Error GoTo 0
End Function

'Purpose: Unicode aware MkDir 'Gerbing 21.01.2018
    Public Function MkDir(ByVal lpPathName As String, Optional ByVal lpSecurityAttributes As Long = 0) As Boolean
       MkDir = CreateDirectoryW(StrPtr("\\?\" & lpPathName), ByVal lpSecurityAttributes) <> 0
    End Function
