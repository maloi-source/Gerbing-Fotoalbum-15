Attribute VB_Name = "Module1"
Option Explicit
    Public gstrLatXMP As String                                                             'Gerbing 08.04.2019
    Public gstrLongXMP As String                                                            'Gerbing 08.04.2019
    Public gstrLat As String                                                                'Gerbing 08.04.2019
    Public gstrLong As String                                                               'Gerbing 08.04.2019

    Public gblstrSystemDirectory As String                                                  'Gerbing 23.11.2013
    Public gblnSubdirectories As Boolean                                                    'Gerbing 04.03.2013
    Public IniFso As Scripting.FileSystemObject
    Public PruefFso As Scripting.FileSystemObject
    Public oStream As Scripting.TextStream
    Public StartMillisek As Long
    Public EndMillisek As Long

    Public gblnSQLServerConnected As Boolean                                                'Gerbing 29.12.2011
    Public gblnSQLServerVersion As Boolean                                                  'Gerbing 29.12.2011
    Public gblnCommandLineEmpty As Boolean                                                  'Gerbing 29.12.2011
    Public gdtDatumFotosMdb As Date                                                         'Gerbing 29.12.2011
    Public PublicLanguage As String
    Public Sprache As Long                          'Gerbing 08.11.2005
    Public AppPath As String
    Public LogNichtBenutzbar As Boolean                                                     'Gerbing 02.09.2008
    Public ExiftoolNichtBenutzbar As Boolean                                                'Gerbing 16.11.2015

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
    Dim Language As String
    Dim zeichen As Integer
    Dim Zeile As String
    Dim StartPos As Integer
    Dim temp As String
    Dim DateiNummer As Long
    Dim Msg As String
    Dim Server As String                            'Gerbing 29.12.2011
    Dim Database As String                          'Gerbing 29.12.2011
    Dim WindowsAuthentication As String             'Gerbing 29.12.2011
    Dim UserName As String                          'Gerbing 29.12.2011
    Dim LocationFotos As String                     'Gerbing 29.12.2011
    Dim Password As String                          'Gerbing 29.12.2011
    Public HyperlinkFieldColumns As New Collection
    
    Public gblnPrüfen4FormAbbrechen As Boolean                  'Gerbing 30.09.2004
    Public gstrMerkeDateiname As String                         'Gerbing 30.12.2004
    Public gblnProversion As Boolean                            'Gerbing 10.06.2005
    Public gblnAbbrechen As Boolean                             'Gerbing 20.07.2005
    Public gblnSchreibgeschützt As Boolean                      'Gerbing 23.01.2007
    
    Public Const CSIDL_APPDATA As Long = &H1A&  '' <username>\Application Data              'Gerbing 17.02.2011
    Public gstrFotosIniAnwendungsOrdner As String                                           'Gerbing 17.02.2011
    
    Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
            (ByVal hWndOwner As Long, _
            ByVal lpszPath As String, _
            ByVal nFolder As Long, _
            ByVal fCreate As Long) As Long                      'Gerbing 17.02.2011
            
    Public Const MAX_PATH As Long = 260                         'Gerbing 17.02.2011
    Public PruefLogFile As String                               'Gerbing 17.02.2011
    '-----------Gerbing 23.06.2011------------------------------------------
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 
    Private Declare Function GetDC Lib "user32" ( _
            ByVal hWnd As Long) As Long
     
    Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hdc As Long) As Long
      
    Public PublicCheckForDPI As String                                                      'Gerbing 23.06.2011
    Dim CheckForDPI As String                       'Gerbing 23.06.2011
    Public PublicSQLServer As String                                                        'Gerbing 29.12.2011
    Public PublicSQLDatabase As String                                                      'Gerbing 29.12.2011
    Public PublicSQLServerUserName As String                                                'Gerbing 29.12.2011
    Public PublicWindowsAuthentication As String                                            'Gerbing 29.12.2011
    Public PublicSQLServerPassword As String                                                'Gerbing 29.12.2011
    Public PublicLocationFotos As String                                                    'Gerbing 29.12.2011
    Public PublicDatagridCaption As String                                                  'Gerbing 29.12.2011
    
    
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

    'The WideCharToMultiByte function maps a wide-character string to a new character string.
    'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.
    
    'CodePage
    Public Const CP_ACP = 0 'ANSI
    Public Const CP_MACCP = 2 'Mac
    Public Const CP_OEMCP = 1 'OEM
    Public Const CP_UTF7 = 65000
    Public Const CP_UTF8 = 65001
    
    'dwFlags
    Public Const WC_NO_BEST_FIT_CHARS = &H400
    Public Const WC_COMPOSITECHECK = &H200
    Public Const WC_DISCARDNS = &H10
    Public Const WC_SEPCHARS = &H20 'Default
    Public Const WC_DEFAULTCHAR = &H40
    
    Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                        ByVal dwFlags As Long, _
                                                        ByVal lpWideCharStr As Long, _
                                                        ByVal cchWideChar As Long, _
                                                        ByVal lpMultiByteStr As Long, _
                                                        ByVal cbMultiByte As Long, _
                                                        ByVal lpDefaultChar As Long, _
                                                        ByVal lpUsedDefaultChar As Long) As Long
    
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
    

Sub ReadFotosIniFile()
    'gstrFotosIniAnwendungsOrdner = getSpecialFolder(CSIDL_APPDATA)                         'Gerbing 08.06.2013
    gstrFotosIniAnwendungsOrdner = AppPath
    'im XP          x:\Dokumente und Einstellungen\user\Anwendungsdaten
    'im Windows7    C:\Users\gottfried\AppData\Roaming
    'ab Version 14.0.0 AppPath
    'FotosIniFile = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14\fotos.ini"      'Pfad der fotos.ini                 Gerbing 20.12.2010
    FotosIniFile = gstrFotosIniAnwendungsOrdner & "\fotos.ini"      'Pfad der fotos.ini                 Gerbing 20.12.2010
    On Error Resume Next
    If file_path_exist(FotosIniFile) = False Then
    'If Kontrolle = "" Then
        'Standardwerte eintragen, wenn es keine Datei fotos.ini gibt    'Gerbing 27.09.2010
        PublicCheckForDPI = "1"                 'Gerbing 23.06.2011
        Exit Sub
    End If
    Call GlL                                'Prüfe [Global] Language
    Call DPI                                'Prüfe [Adjustments] CheckForDPI                'Gerbing 23.06.2011
    Call SSRV                               'Prüfe [SQL] Server
    Call SDB                                'Prüfe [SQL] Database
    Call SWA                                'Prüfe [SQL] WindowsAuthentication
    Call SUN                                'Prüfe [SQL] username
    Call SPW                                'Prüfe [SQL] password
    Call SLF                                'Prüfe [SQL] LocationFotos
End Sub

Public Function INIReadString(ByVal Section As String, ByVal key As String, ByVal Default As String, ByVal Filename As String) As String
    Dim cSize As Long
    Dim strReturn As String
    Dim retval As Long
    
    cSize = 300
    strReturn = String(cSize, 0)
    retval = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(strReturn), cSize, StrPtr(Filename))
    If retval > 0 Then
        INIReadString = left(strReturn, retval)
    End If
End Function

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
            MsgBox "Starten Sie zuerst fotos.exe. At first you must start fotos.exe."
            End
        End If
    End If
    PublicLanguage = Zeile
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

Sub WriteGlL(NeuerInhalt As String)
    'Schreibe [Global] Language
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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

Sub WriteSSRV(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [SQL] Server
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14"
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
'
'Public Function LoadPicBox(ByVal Dateiname As String, _
'    ByRef PicBox As VB.PictureBox)
'
'    'GdipGetImageDimension ist eine schnelle Funktion zum Ermitteln von Bildbreite und Bildhöhe
'    'möglich mit BMP, DIB, JPG, GIF, PNG, TIF
'    'nicht möglich mit ICO, CUR, PSD
'    'Falsches Ergebnis bei EMF, WMF
'
'    Dim retcode As Long              ' Funktions-Rückgaben
'    Dim Bitmap As Long
'    Dim GDIP_Connection As Long      ' Verbindung zu GDIPlus
'    Dim GDIP_Startup As GDIPlusStartupInput
'
'    On Error GoTo exitfunction
'    ERR.Clear
'
'    gsngPicWidth = 0                'wenn diese Werte = 0 bleiben, lag ein Fehler vor
'    gsngPicHeight = 0
'    GDIP_Startup.Version = 1
'    retcode = GdiplusStartup(GDIP_Connection, GDIP_Startup, ByVal 0&)
'    If retcode <> 0 Then
'       Exit Function
'    End If
'
'    ' Trägt das Bild aus der Datei in die Bitmap ein
'    retcode = GdipLoadImageFromFile(StrPtr(Dateiname), Bitmap)
'    If retcode <> 0 Then
'       GoTo exitfunction
'    End If
'
'    ' Abfrage der Abmessungen der Bitmap
'    retcode = GdipGetImageDimension(Bitmap, gsngPicWidth, gsngPicHeight)
'    If retcode <> 0 Then
'      GoTo exitfunction
'    End If
'exitfunction:
'    ' Ressourcen und GDIPLus freigeben
'    If Bitmap <> 0 Then
'      ' Bitmap löschen
'      GdipDisposeImage Bitmap
'    End If
'
'    If GDIP_Connection <> 0 Then
'      ' GDIPlus-DLL freigeben
'      GdiplusShutdown GDIP_Connection
'    End If
'End Function

Public Function getSpecialFolder(ByVal FolderID As Long) As String
    Dim sBuffer As String

    'SHGetSpecialFolderPath is not necessarily supported under Win95. It canbe called
    'under Win95 if IE4 AND the Desktop Update are installed.
    'Cheap-shot error handling
    'If running under Win95 or NT4, merely ignore the error
    'This will result in a 0-length string being returned.
    On Error Resume Next
    sBuffer = String$(MAX_PATH, vbNullChar)
    'For Win98 and higher
    Call SHGetSpecialFolderPath(0&, sBuffer, FolderID, 0&)
    getSpecialFolder = StripNulls(sBuffer)
End Function

Public Function StripNulls(ByVal sText As String) As String
    'Returns all characters up to a null character.
    'If the string does not contain a null character,
    'the string is returned unmodified.
    Dim lNullPos As Long

    lNullPos = InStr(sText, vbNullChar)
    If lNullPos Then
        StripNulls = left$(sText, lNullPos - 1)
    Else
        StripNulls = sText
    End If
End Function
'
'Public Function AnpassenHeadFont(DbGrid)
'    CheckForSmallFonts                                          'Gerbing 23.06.2011
'    On Error Resume Next
'    If PublicCheckForDPI = "0" Then
'        DbGrid.HeadFont.Size = 8
'        DbGrid.HeadFont.Bold = False
'        DbGrid.RowHeight = 220
'        DbGrid.HeadFont.Name = Form1.txtFont.FontName  'Gerbing 14.02.2012
'        On Error GoTo 0
'        Exit Function
'    End If
'    Select Case gbllogPix                                       'Gerbing 23.06.2011
'      Case Is <= 96
'              DbGrid.HeadFont.Size = 8
'              DbGrid.HeadFont.Bold = False
'              DbGrid.RowHeight = 220
'              DbGrid.HeadFont.Name = Form1.txtFont.FontName  'Gerbing 14.02.2012
'      Case Is <= 120
'              DbGrid.HeadFont.Size = 10
'              DbGrid.HeadFont.Bold = False
'              DbGrid.RowHeight = 260
'              DbGrid.HeadFont.Name = Form1.txtFont.FontName  'Gerbing 14.02.2012
'      Case Else
'              DbGrid.HeadFont.Size = 12
'              DbGrid.HeadFont.Bold = False
'              DbGrid.RowHeight = 300
'              DbGrid.HeadFont.Name = Form1.txtFont.FontName  'Gerbing 14.02.2012
'    End Select
'    On Error GoTo 0
'End Function

Public Function AnpassenHeadFont(DbGrid)                                                'Gerbing 11.03.2017
    On Error Resume Next
    DbGrid.HeadFont.Bold = False
    DbGrid.HeadFont.Name = Form1.txtFont.FontName
    If PublicCheckForDPI = 1 Then
            DbGrid.HeadFont.Size = 8
            DbGrid.RowHeight = 220
    End If
    If PublicCheckForDPI = 2 Then
            DbGrid.HeadFont.Size = 10
            DbGrid.RowHeight = 260
    End If
    If PublicCheckForDPI = 3 Then
            DbGrid.HeadFont.Size = 12
            DbGrid.RowHeight = 300
    End If
    On Error GoTo 0
End Function

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


Function Crypt(Inp$, key$, mode As Boolean) As String
    Dim z As String
    Dim i As Integer, Position As Integer
    Dim cptZahl As Long, orgZahl As Long
    Dim keyZahl As Long, cptString As String
    
    On Error Resume Next
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
End Function

Public Sub VierUrsachenFürSchreibsperre()                                                'Gerbing 02.09.2008
    'vier mögliche Ursachen
    'Msg = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14\fotos.ini" & vbNewLine
    Msg = gstrFotosIniAnwendungsOrdner & "\fotos.ini" & vbNewLine
    'msg = msg & "Die Datei ist schreibgeschützt. Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können." & vbnewline
    Msg = Msg & LoadResString(2275 + Sprache) & vbNewLine
    'msg = msg & "Es gibt vier mögliche Ursachen für den Lesemodus:" & vbnewline
    Msg = Msg & LoadResString(2133 + Sprache) & vbNewLine
    'msg = msg & "1. Das Dateiattribut 'Schreibgeschützt' ist gesetzt" & vbnewline
    Msg = Msg & LoadResString(2134 + Sprache) & vbNewLine
    'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte für Ihren PC" & vbnewline
    Msg = Msg & LoadResString(2135 + Sprache) & vbNewLine
    'msg = msg & "3. Sie arbeiten mit einer CD oder DVD" & vbnewline
    Msg = Msg & LoadResString(2136 + Sprache) & vbNewLine
    'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & vbnewline & vbnewline
    Msg = Msg & LoadResString(2137 + Sprache) & vbNewLine & vbNewLine
    MsgBox Msg, , LoadResString(1119 + Sprache)
End Sub

Public Function StringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        
        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    StringToByteArray = bytBuffer
End Function

Public Function ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    If iUnicode > 256 Then
        gblnIsUni = True
    Else
        gblnIsUni = False
    End If
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
End Function

Public Sub HelpFileErrorMsg(retval As Long, CHMFile As String)
    Dim DateinamenErweiterung As String
    Dim ErrorText As String
    Dim Msg As String

    DateinamenErweiterung = "CHM"
    ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
    Msg = "Errortext=" & ErrorText & vbNewLine
    Msg = Msg & "Errornr=" & retval & vbNewLine & vbNewLine
    
    Msg = Msg & CHMFile & vbNewLine
    'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
    Msg = Msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
    
    'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
    Msg = Msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
    
    'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
    Msg = Msg & LoadResString(1378 + Sprache) & vbNewLine
    'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
    Msg = Msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
    'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
    Msg = Msg & LoadResString(2012 + Sprache) & vbNewLine
    'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
    Msg = Msg & LoadResString(2013 + Sprache)
    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
End Sub

