Attribute VB_Name = "Module1"
Option Explicit
    Public gblnSubdirectories As Boolean                                                    'Gerbing 10.12.2017
    Public fso As New FileSystemObject
    Public oStream As TextStream
    Public gblnProversion As Boolean                                                        'Gerbing 04.03.2012
    Public gblnSQLServerConnected As Boolean                                                'Gerbing 29.12.2011
    Public gblnSQLServerVersion As Boolean                                                  'Gerbing 29.12.2011
    Public gblnCommandLineEmpty As Boolean                                                  'Gerbing 29.12.2011
    Public gdtDatumFotosMdb As Date                                                         'Gerbing 29.12.2011
    
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
    Dim msg As String
    Dim Server As String                            'Gerbing 29.12.2011
    Dim Database As String                          'Gerbing 29.12.2011
    Dim WindowsAuthentication As String             'Gerbing 29.12.2011
    Dim UserName As String                          'Gerbing 29.12.2011
    Dim LocationFotos As String                     'Gerbing 29.12.2011
    Dim Password As String                          'Gerbing 29.12.2011

    Public Sprache As Long                          'Gerbing 08.11.2005
    Public AppPath As String

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
            
    Public Const CSIDL_APPDATA As Long = &H1A&  '' <username>\Application Data              'Gerbing 17.02.2011
    Public gstrFotosIniAnwendungsOrdner As String                                           'Gerbing 17.02.2011
    
    Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
            (ByVal hWndOwner As Long, _
            ByVal lpszPath As String, _
            ByVal nFolder As Long, _
            ByVal fCreate As Long) As Long                      'Gerbing 17.02.2011
            
    Public Const MAX_PATH As Long = 260                         'Gerbing 17.02.2011
    '-----------Gerbing 23.06.2011------------------------------------------
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 
    Private Declare Function GetDC Lib "user32" ( _
            ByVal hWnd As Long) As Long
     
    Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hdc As Long) As Long
      
    Public PublicCheckForDPI As String                                                      'Gerbing 23.06.2011
    Public CheckForDPI As String                       'Gerbing 23.06.2011
    
    Public PublicSQLServer As String                                                        'Gerbing 29.12.2011
    Public PublicSQLDatabase As String                                                      'Gerbing 29.12.2011
    Public PublicSQLServerUserName As String                                                'Gerbing 29.12.2011
    Public PublicWindowsAuthentication As String                                            'Gerbing 29.12.2011
    Public PublicSQLServerPassword As String                                                'Gerbing 29.12.2011
    Public PublicLocationFotos As String                                                    'Gerbing 29.12.2011
    Public PublicDatagridCaption As String                                                  'Gerbing 29.12.2011
    
    Public PublicLanguage As String
    '-----------Gerbing 04.03.2013------------------------------------------
    Public gblnIsUni As Boolean
    Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'    Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'    Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
    
    Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFFData As Long) As Long    'Gerbing 24.06.2014
    Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFFData As Long) As Long      'Gerbing 24.06.2014

    Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
    Public Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Public Const MB_ICONINFORMATION As Long = &H40&
    Public Const MB_TASKMODAL As Long = &H2000&


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
        StripNulls = Left$(sText, lNullPos - 1)
    Else
        StripNulls = sText
    End If
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
            Form.Controls(i).FontName = Renam.txtFont.FontName
            Form.Controls(i).Font.Size = 8
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 2 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = Renam.txtFont.FontName
            Form.Controls(i).Font.Size = 10
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 3 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = Renam.txtFont.FontName
            Form.Controls(i).Font.Size = 12
        Next i
    End If
    On Error GoTo 0
End Function

Public Function AnpassenHeadFont(DbGrid)                                                'Gerbing 11.03.2017
    On Error Resume Next
    DbGrid.HeadFont.Bold = False
    DbGrid.HeadFont.Name = Renam.txtFont.FontName
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
'        msg = "Datei " & FotosIniFile & NL & "nicht gefunden." & NL
'        msg = msg & "Diese Datei ist hilfreich, wenn Sie häufig Export/Import-Funktionen benutzen." & NL
'        msg = msg & "Sie speichert den jeweils zuletzt von Ihnen benutzen Export/Import-Dateinamen."
'        MsgBox msg
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
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
        Set oStream = fso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
        Set oStream = Nothing
    End If
    absch = "SQL"
    LocationFotos = "LocationFotos"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(LocationFotos), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Public Sub VierUrsachenFürSchreibsperre()                                                'Gerbing 02.09.2008
    'vier mögliche Ursachen
    'Msg = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14\fotos.ini" & vbNewLine
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

