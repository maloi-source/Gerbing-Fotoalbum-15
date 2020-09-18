Attribute VB_Name = "Module1"
Option Explicit
    Public gstrEXIFOrientation As String                                                    'Gerbing 10.05.2019
    Public gstrLatXMP As String                                                             'Gerbing 08.04.2019
    Public gstrLongXMP As String                                                            'Gerbing 08.04.2019
    Public glngStrgG As Long                                                                'Gerbing 16.10.2018
    Public gblnHäkchenGefunden As Boolean           'Gerbing 19.10.2017
    Public gstrGEOPosition As String                'Gerbing 27.08.2015
    Public gstrLat As String                        'Gerbing 29.09.2018
    Public gstrLong As String                       'Gerbing 29.09.2018
    Public gblnListBoxNeuDblClick As Boolean        'Gerbing 25.08.2013
    Public IniFso As Scripting.FileSystemObject     'Gerbing 11.03.2017
    Public fso As FileSystemObject
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public StartMillisek As Long
    Public EndMillisek As Long

    Public gblnSubdirectories As Boolean            'Gerbing 04.03.2013
    Public gblnBildBeschreibung As Boolean          'Gerbing 03.03.2012
    Public PublicZoomToFullscreen As String
    Public PublicPathToEverythingExe As String      'Gerbing 30.01.2018
    Public Sprache As Long                          'Gerbing 08.11.2005
    Public gstrSprache As String                    'Gerbing 20.11.2007
    Public AppPath As String
    
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
    
    Public FotosIniFile As String
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
    Public gblStrAktuellGezeigtesBild As String     'Gerbing 29.07.2007
    
    Private Type Bildposition
        Top As Long
        Left As Long
        ZoomPercent As Long                         'Gerbing 29.03.2012
    End Type

    Public BildPosList() As Bildposition

    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long


    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        ()
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const HORZRES = 8
    Public Const VERTRES = 10
    
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
    'Const aus win32api.txt
    Public Const LB_SETHORIZONTALEXTENT = &H194
    
    
    Public Enum SHSpecialFolderIDs
        CSIDL_PROGRAMS = &H2&                            ' Me\StartMenu\Programs
        CSIDL_COMMON_PROGRAMS = &H17&                    ' All Users\StartMenu\Programs
        CSIDL_SYSTEM = &H25&                             ' GetSystemDirectory()
    End Enum
    Public Const MAX_PATH                  As Long = 260
    'Public gblnIsWin98 As Boolean
    'Public JüngsterOrdner As String                                                         'Gerbing 14.11.2007
    
    '-----------Gerbing 23.01.2010------------------------------------------
    Public Declare Function SHFileOperation Lib "shell32.dll" _
            Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) _
            As Long
    
    Public Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Boolean
        hNameMappings As Long
        lpszProgressTitle As String
    End Type
    
    Public Const FO_DELETE = &H3
    
'    Public Const FOF_SILENT = &H4
'    Public Const FOF_NOCONFIRMATION = &H10
'    Public Const FOF_WANTMAPPINGHANDLE = &H20
'    Public Const FOF_ALLOWUNDO = &H40
'    Public Const FOF_FILESONLY = &H80
'    Public Const FOF_SIMPLEPROGRESS = &H100
'    Public Const FOF_NOCONFIRMMKDIR = &H200
'    Public Const FOF_NOERRORUI = &H400
'-----------Gerbing 23.01.2010------------------------------------------
    
    Public Const CSIDL_APPDATA As Long = &H1A&  '' <username>\Application Data              'Gerbing 17.02.2011
    Public gstrFotosIniAnwendungsOrdner As String                                           'Gerbing 17.02.2011
    
    Public Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
            (ByVal hWndOwner As Long, _
            ByVal lpszPath As String, _
            ByVal nFolder As Long, _
            ByVal fCreate As Long) As Long                      'Gerbing 17.02.2011
'-----------Gerbing 23.06.2011------------------------------------------
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
     
    Private Declare Function GetDC Lib "user32" ( _
      ByVal hwnd As Long) As Long
     
    Private Declare Function ReleaseDC Lib "user32" ( _
      ByVal hwnd As Long, _
      ByVal hdc As Long) As Long
      
    Public PublicCheckForDPI As String                                                      'Gerbing 23.06.2011
    Public PathToEverythingExe As String                                                    'Gerbing 30.01.2018
    Dim ZoomToFullscreen As String
    Dim Adjustments As Long
    Dim CheckForDPI As String                       'Gerbing 23.06.2011
'-----------Gerbing 04.03.2013------------------------------------------
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
                                                        
    Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
                            ByVal CodePage As Long, _
                            ByVal dwFlags As Long, _
                            ByVal lpMultiByteStr As Long, _
                            ByVal cbMultiByte As Long, _
                            ByVal lpWideCharStr As Long, _
                            ByVal cchWideChar As Long) As Long
                                         
    Public gblnIsUni As Boolean
    Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
    Public Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Public Const MB_ICONINFORMATION As Long = &H40&
    Public Const MB_TASKMODAL As Long = &H2000&
    
    Public Type DLLVERSIONINFO
        cbSize As Long
        dwMajor As Long
        dwMinor As Long
        dwBuildNumber As Long
        dwPlatformId As Long
    End Type
    
    Public Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
    Public hImgLst As Long
    Public hStateImgLst As Long
    Public bComctl32Version600OrNewer As Boolean
    Public Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
    Public Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
    Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
    Public Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
    Public Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
    Public Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
    Public Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
    Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
    Public Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)              'Gerbing 29.01.2018

    'Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
    'Public Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
    
    'Public Const MAX_PATH = 260
    Public Const MAXDWORD = &HFFFF
    'Public Const INVALID_HANDLE_VALUE = -1
    Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
    'Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
    Public Const FILE_ATTRIBUTE_HIDDEN = &H2
    Public Const FILE_ATTRIBUTE_NORMAL = &H80
    Public Const FILE_ATTRIBUTE_READONLY = &H1
    Public Const FILE_ATTRIBUTE_SYSTEM = &H4
    Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
    
    Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    
    Public gblnMessageAusgeben As Boolean                                                                       'Gerbing 24.01.2009
    Public Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long 'Gerbing 30.01.2018

Sub ReadFotosIniFile()
    gstrFotosIniAnwendungsOrdner = AppPath
    FotosIniFile = gstrFotosIniAnwendungsOrdner & "\fotos.ini"
    On Error Resume Next
    If file_path_exist(FotosIniFile) = False Then 'prüfe ob fotos.ini existiert
'        msg = "Datei " & FotosIniFile & vbNewLine & "nicht gefunden." & vbNewLine          'Gerbing 24.12.2017
'        MsgBox msg
        'End                                                                                 'Gerbing 11.03.2017
        'Standardwerte eintragen, wenn es keine Datei fotos.ini gibt    'Gerbing 27.09.2010
        PublicCheckForDPI = "1"                 'Gerbing 23.06.2011
        Exit Sub
    End If
    Call GlL                                'Prüfe [Global] Language
    Call DPI                                'Prüfe [Adjustments] CheckForDPI                'Gerbing 23.06.2011
    Call AZF                                'Prüfe [Adjustments] ZoomToFullscreen           'Gerbing 29.03.2012
    Call PTE                                'Prüfe [Adjustments] PathToEverythingExe        'Gerbing 30.01.2018
End Sub

Public Function INIReadString(ByVal Section As String, ByVal key As String, ByVal Default As String, ByVal FileName As String) As String
    Dim cSize As Long
    Dim strReturn As String
    Dim retval As Long
    
    cSize = 300
    strReturn = String(cSize, 0)
    retval = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(strReturn), cSize, StrPtr(FileName))
    If retval > 0 Then
        INIReadString = Left(strReturn, retval)
    End If
End Function

Private Sub PTE()                                                                       'Gerbing 30.01.2018
    'Prüfe [Adjustments] PathToEverythingExe
    absch = "Adjustments"
    CheckForDPI = "PathToEverythingExe"
    Zeile = INIReadString(absch, CheckForDPI, "", FotosIniFile)
    If Zeile = "" Then PublicPathToEverythingExe = ""              'leer berücksichtigen
    PublicPathToEverythingExe = Zeile
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

Sub GlL()
    'Prüfe [Global] Language
    absch = "Global"
    Language = "Language"
    Zeile = INIReadString(absch, Language, "", FotosIniFile)
    If Zeile = "" Then Sprache = 0
    If Zeile = "0" Then Sprache = 0            'Deutsch
    If Zeile = "1" Then Sprache = 3000         'English
End Sub

Sub DPI()                                                                                       'Gerbing 11.03.2017
    'Prüfe [Adjustments] CheckForDPI
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    Zeile = INIReadString(absch, CheckForDPI, "", FotosIniFile)
    If Zeile = "" Then PublicCheckForDPI = "1"              'leer berücksichtigen
    If Zeile = "1" Then PublicCheckForDPI = "1"             'klein
    If Zeile = "2" Then PublicCheckForDPI = "2"             'mittel
    If Zeile = "3" Then PublicCheckForDPI = "3"             'gross
End Sub

Sub WritePTE(NeuerInhalt As String)                                                         'Gerbing 30.01.2018
    Dim oStream As TextStream
    
    'Schreibe [Adjustments] PathToEverythingExe
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
    PathToEverythingExe = "PathToEverythingExe"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(PathToEverythingExe), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
End Sub

Sub WriteAZF(NeuerInhalt As String)                                                        'Gerbing 14.03.2018
    Dim oStream As TextStream
    
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

Sub WriteDPI(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    Dim oStream As TextStream
    Dim antwort As Long
    
    'Schreibe [Adjustments] CheckForDPI
    If file_path_exist(FotosIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'If file_path_exist(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15") = False Then
        If file_path_exist(gstrFotosIniAnwendungsOrdner) = False Then
            'MkDir gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 15"
            MkDir gstrFotosIniAnwendungsOrdner
        End If
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        On Error Resume Next                                                            'Gerbing 24.12.2017
        Err.Number = 0
        Set oStream = IniFso.CreateTextFile(FotosIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(CheckForDPI), StrPtr(NeuerInhalt), StrPtr(FotosIniFile))
    If zeichen = 0 Then                                                                 'Gerbing 18.03.2017
            msg = FotosIniFile & vbNewLine
            msg = msg & LoadResString(2276 + Sprache) & vbNewLine & vbNewLine   'Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können.
            msg = msg & LoadResString(1542 + Sprache)                           'Wollen Sie trotzdem weiterarbeiten ?
            antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo, "GERBING Diashow")
            If antwort = vbNo Then
                End                                                                     'Gerbing 24.12.2017
            End If
    On Error GoTo 0
    End If
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

Public Function getSpecialFolder(ByVal FolderID As SHSpecialFolderIDs) As String
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

' Alle Datei-Angaben für SHFileOperation müssen mit vbNullChar+vbNullChar           'Gerbing 23.01.2010
' abgeschlossen werden. Hier wird's noch mal geprüft
Public Function Check_NullChars(S As String) As String
  If Right(S, 2) <> vbNullChar + vbNullChar Then
    If Right(S, 1) <> vbNullChar Then
      S = S + vbNullChar + vbNullChar
    Else
      S = S + vbNullChar
    End If
  End If
  Check_NullChars = S
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
            Form.Controls(i).FontName = MDIForm1.txtFont.FontName
            Form.Controls(i).Font.Size = 8
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 2 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = MDIForm1.txtFont.FontName
            Form.Controls(i).Font.Size = 10
            'Debug.Print Form.Controls(i).Name
        Next i
    End If
    If PublicCheckForDPI = 3 Then
        For i = 0 To Form.Controls.Count - 1
            Form.Controls(i).Font.Bold = False
            Form.Controls(i).FontName = MDIForm1.txtFont.FontName
            Form.Controls(i).Font.Size = 12
        Next i
    End If
    On Error GoTo 0
End Function

Public Function RemoveNulls(OriginalString As String) As String
    Dim pos As Long
    pos = InStr(OriginalString, Chr$(0))
    If pos > 1 Then
        RemoveNulls = Mid$(OriginalString, 1, pos - 1)
    Else
        RemoveNulls = OriginalString
    End If
End Function

Public Function EXIFLesenOrientation()
    If StrComp(Right(gblStrAktuellGezeigtesBild, 3), "JPG", vbTextCompare) <> 0 Then
        ListBoxForm.chkExifAnzeigen.Value = 0
        Exit Function                                           'Gerbing 27.05.2015
    End If
    DiashowForm.EXF.ImageFile = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) 'set the image file property
    '
    'EXF.ListInfo ist ein String mit vbCrLf
    '
    ListBoxForm.txtEXIFInfo.Text = DiashowForm.EXF.ListInfo 'list all tags into the text box
End Function

Public Function EXIFLesen()
    If StrComp(Right(gblStrAktuellGezeigtesBild, 3), "JPG", vbTextCompare) <> 0 Then
        ListBoxForm.chkExifAnzeigen.Value = 0
        Exit Function                                           'Gerbing 27.05.2015
    End If
    
    'chkIptcAnzeigen.Value = 0                                  'Gerbing 18.08.2008
    If ListBoxForm.chkExifAnzeigen.Value = 0 Then
        ListBoxForm.txtIPTCInfo.Visible = False
        ListBoxForm.txtEXIFInfo.Visible = False                             'Gerbing 07.05.2007
    Else
        ListBoxForm.chkIptcAnzeigen.Value = 0                                  'Gerbing 18.08.2008
        ListBoxForm.txtIPTCInfo.Visible = False
        ListBoxForm.txtEXIFInfo.Visible = True
        If ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) = "" Then
            ListBoxForm.txtEXIFInfo.Text = LoadResString(1483 + Sprache)  'keine Datei markiert
        Else
            DiashowForm.EXF.ImageFile = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) 'set the image file property
            '
            'EXF.ListInfo ist ein String mit vbCrLf
            '
            ListBoxForm.txtEXIFInfo.Text = DiashowForm.EXF.ListInfo 'list all tags into the text box
        End If
    End If
End Function

'Purpose: Unicode aware MkDir 'Gerbing 30.01.2018
    Public Function MkDir(ByVal lpPathName As String, Optional ByVal lpSecurityAttributes As Long = 0) As Boolean
       MkDir = CreateDirectoryW(StrPtr("\\?\" & lpPathName), ByVal lpSecurityAttributes) <> 0
    End Function

