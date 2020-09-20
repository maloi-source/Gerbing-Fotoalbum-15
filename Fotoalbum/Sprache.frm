VERSION 5.00
Begin VB.Form frmSprache 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GERBING Fotoalbum 13"
   ClientHeight    =   2256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7260
   Icon            =   "Sprache.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   7260
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sprache - Language"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6852
      Begin VB.OptionButton OptEnglish 
         Caption         =   "English"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3732
      End
      Begin VB.OptionButton OptDeutsch 
         Caption         =   "Deutsch"
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   4692
      End
   End
End
Attribute VB_Name = "frmSprache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim SQLR As String
    Dim blnMerkeDeutsch As Boolean
    
    
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2
    
    ' Auflistung
    Public Enum OfficeVersion
        Office2003 = 11
        Office2007 = 12
        Office2010 = 14
    End Enum
    
    'Hier wird die Registry ausgewertet, ob eine bestimmte Office-Version installiert ist
    ' benötigte API-Deklarationen
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
      Alias "RegOpenKeyExA" ( _
      ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      phkResult As Long) As Long
     
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
      Alias "RegQueryValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      ByVal lpData As String, _
      lpcbData As Long) As Long
     
    Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
      ByVal hKey As Long) As Long
     
    ' Konstanten
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const ERROR_SUCCESS = 0&
    Private Const REG_SZ = 1
    Private Const KEY_QUERY_VALUE = &H1
    

Private Sub btnOK_Click()
    Unload Me                                      'Gerbing 24.12.2007
End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim strVersion As String
    Dim Pos As Long
    Dim Datei As String
    Dim DateiNummer As Long
    
    Call AnpassenNutzerWunsch(Me)                           'Gerbing 11.03.2017
    If PublicLanguage = "9" Then                            'es gibt keine fotos.ini   'Gerbing 20.11.2007 'Gerbing 04.12.2011
        '---------------------------------------------------------------------------------------------------
FrameCaption:
        Frame1.Caption = LoadResString(1041 + Sprache)      'Gerbing 24.12.2007 'Sprache - Language
        Me.Caption = LoadResString(1119 + Sprache)          'Gerbing 24.12.2007 'GERBING Fotoalbum 15
'auskommentiert Gerbing 21.11.2019
'        If gblnVollversion = False Then
'        'es gibt keine fotos.ini Bei Vollversion = True muss der User keine Administratorrechte haben
'            If IsUserAnAdmin = False Then                        'Gerbing 04.12.2011
'                'Datei = Dir(gblstrSystemDirectory & "\msdmo.log")
'                'If Datei = "" Then
'                If file_path_exist(gblstrSystemDirectory & Chr(92) & Chr(109) & Chr(115) & Chr(100) & Chr(109) & Chr(111) & Chr(46) & Chr(108) & Chr(111) & Chr(103)) = False Then
'                    Msg = "Sie brauchen Administratorrechte und Schreibzugriff auf 'fotos.mdb' um die Sprache festzulegen" & vbNewLine
'                    Msg = Msg & "You need administrator rights and write access on 'fotos.mdb' for selecting the language"
'                    MsgBox Msg, , "GERBING Fotoalbum 15"
'                    'jetzt muss fotos.mdb wieder das Datum 30.12.2011 bekommen
'                    m_Date = "30.12.2011 15:01:02"
'                    udtSystemTime.wYear = Year(m_Date)
'                    udtSystemTime.wMonth = Month(m_Date)
'                    udtSystemTime.wDay = Day(m_Date)
'                    udtSystemTime.wDayOfWeek = Weekday(m_Date) - 1
'                    udtSystemTime.wHour = Hour(m_Date)
'                    udtSystemTime.wMinute = Minute(m_Date)
'                    udtSystemTime.wSecond = Second(m_Date)
'                    udtSystemTime.wMilliseconds = 0
'
'                    ' convert system time to local time
'                    SystemTimeToFileTime udtSystemTime, udtLocalTime
'                    ' convert local time to GMT
'                    LocalFileTimeToFileTime udtLocalTime, udtFileTime
'                    ' open the file to get the filehandle
'                    'Gerbing 10.09.2013
'                    lngHandle = CreateFileW(StrPtr(AppPath & "\fotos.mdb"), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)  'Gerbing 08.11.2005
'
'                    ' change date/time property of the file
'                    SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
'                    ' close the handle
'                    CloseHandle lngHandle
'                    End
'                End If
'            Else
'            '--------------------------------------------------------------------------------------------------
'                'vorher sicherstellen daß der Nutzer Administratorrechte besitzt, sonst kann ich nicht c:\windows\SysWOW64\msdmo.log erzeugen
'                'Shareware-version
'                'Datei = Dir(gblstrSystemDirectory & "\msdmo.log")                        'Gerbing 08.11.2005
'                'If Datei = "" Then
'                If file_path_exist(gblstrSystemDirectory & Chr(92) & Chr(109) & Chr(115) & Chr(100) & Chr(109) & Chr(111) & Chr(46) & Chr(108) & Chr(111) & Chr(103)) = False Then
'                    'Nur dann msdmo.log anlegen, wenn sie noch nicht vorhanden ist
'                    'sonst kann man mich austricksen durch Neuerstellen von fotos.ini
'                    On Error Resume Next
'                    Err = 0
'                    'Open gblstrSystemDirectory & "\msdmo.log" For Output As #DateiNummer 'Gerbing 08.11.2005
'                    Set oStream = Fso.OpenTextFile(gblstrSystemDirectory & Chr(92) & Chr(109) & Chr(115) & Chr(100) & Chr(109) & Chr(111) & Chr(46) & Chr(108) & Chr(111) & Chr(103), ForWriting, True, TristateFalse)
'                    If Err <> 0 Then
'                        Msg = "Error - 2305" & vbNewLine
'                        Msg = Msg & "Sie brauchen Administratorrechte und Schreibzugriff auf 'fotos.mdb' um die Sprache festzulegen" & vbNewLine
'                        Msg = Msg & "You need administrator rights and write access on 'fotos.mdb' for selecting the language"
'                        MsgBox Msg, vbCritical
'                        End
'                    End If
'                    On Error GoTo 0
'                    'Print #DateiNummer, "start-end"                                 'Gerbing 04.12.2011
'                    oStream.WriteLine "start-end"
'                    oStream.Close
'                    'Jetzt wird der Datei msdmo.log das Datum von heute - 100 verpaßt
'                    m_Date = Format(DateAdd("d", -100, Now), "DD-MM-YY")
'                    udtSystemTime.wYear = Year(m_Date)
'                    udtSystemTime.wMonth = Month(m_Date)
'                    udtSystemTime.wDay = Day(m_Date)
'                    udtSystemTime.wDayOfWeek = Weekday(m_Date) - 1
'                    udtSystemTime.wHour = Hour(m_Date)
'                    udtSystemTime.wMinute = Minute(m_Date)
'                    udtSystemTime.wSecond = Second(m_Date)
'                    udtSystemTime.wMilliseconds = 0
'
'                    ' convert system time to local time
'                    SystemTimeToFileTime udtSystemTime, udtLocalTime
'                    ' convert local time to GMT
'                    LocalFileTimeToFileTime udtLocalTime, udtFileTime
'                    ' open the file to get the filehandle
'                    'lngHandle = CreateFile(gblstrSystemDirectory & "\msdmo.log", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)  'Gerbing 08.11.2005
'                    'Gerbing 10.09.2013
'                    lngHandle = CreateFileW(StrPtr(gblstrSystemDirectory & Chr(92) & Chr(109) & Chr(115) & Chr(100) & Chr(109) & Chr(111) & Chr(46) & Chr(108) & Chr(111) & Chr(103)), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
'
'                    ' change date/time property of the file
'                    SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
'                    ' close the handle
'                    CloseHandle lngHandle
'                    gintDiffTage = 90
'                End If
'            End If
'        End If
        '---------------------------------------------------------------------------------------------------------------------
        'es gibt keine fotos.ini
        If gblnSQLServerVersion = False Then
            DBado.Close                                         'Gerbing 23.11.2017
            On Error Resume Next
            DBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
            DBado.mode = adModeReadWrite
            DBado.Open DBado.ConnectionString                   'Gerbing 23.11.2017
            If Err.Number <> 0 Then                                                             'Gerbing 01.09.2008
                Msg = "Errornumber=" & Err.Number & vbNewLine
                Msg = Msg & "Errortext=" & Err.Description & vbNewLine & vbNewLine
    
                Msg = Msg & "Sie brauchen Administratorrechte und Schreibzugriff auf 'fotos.mdb' um die Sprache festzulegen" & vbNewLine
                Msg = Msg & "You need administrator rights and write access on 'fotos.mdb' for selecting the language"
                MsgBox Msg
                End
            End If
            On Error GoTo 0                                                                     'Gerbing 01.09.2008
        End If
    Else
        '------------------------------------------------------------------------------------------------------------------
        'PublicLanguage <> "9"  'es gibt die fotos.ini
        If gblnSQLServerVersion = False Then
            On Error Resume Next
            Err.Number = 0
            DBado.Close                                                                         'Gerbing 23.11.2017
            DBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
            DBado.mode = adModeReadWrite
            DBado.Open DBado.ConnectionString                   'Gerbing 23.11.2017
            If Err.Number <> 0 Then                                                             'Gerbing 01.09.2008
                Msg = "Errornumber=" & Err.Number & vbNewLine
                Msg = Msg & "Errortext=" & Err.Description & vbNewLine & vbNewLine
    
                Msg = Msg & "Sie brauchen Schreibzugriff auf 'fotos.mdb' um die Sprache festzulegen" & vbNewLine
                Msg = Msg & "You need write access on 'fotos.mdb' for selecting the language"
                MsgBox Msg
                End
            End If
        End If
        On Error GoTo 0
        Frame1.Caption = LoadResString(1041 + Sprache)      'Gerbing 24.12.2007 'Sprache - Language
        Me.Caption = LoadResString(1119 + Sprache)          'Gerbing 24.12.2007 'GERBING Fotoalbum 15
        OptDeutsch.Caption = LoadResString(1049 + Sprache)    'Deutsch
        OptEnglish.Caption = LoadResString(1050 + Sprache)    'English
        If Sprache = 0 Then
            OptDeutsch.Value = True
            blnMerkeDeutsch = True
        Else
            OptEnglish.Value = True
            blnMerkeDeutsch = False
        End If
    End If
    Screen.MousePointer = vbDefault                         'Gerbing 07.11.2011
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gblnSQLServerVersion = True Then
        Call SpracheÄndernSQLServer
    Else
        Call SpracheÄndernAccess
    End If
End Sub

Private Sub SpracheÄndernSQLServer()
    Dim Msg As String
    Dim antwort As Long

    If PublicLanguage <> "9" Then                                                           'Gerbing 04.12.2011
        'nichts tun wenn die Sprache nicht wirklich gewechselt wurde
        If OptDeutsch.Value = True And blnMerkeDeutsch = True Then Exit Sub
        If OptEnglish.Value = True And blnMerkeDeutsch = False Then Exit Sub
    End If

    If PublicLanguage <> "9" Then                                                           'Gerbing 04.12.2011
        'msg = "Starten Sie nach dem Wechsel der Sprache das Programm neu"
        'msg = msg & "Diese Funktion sollten Sie nur benutzen, wenn Sie eine Sicherungskopie der Datei fotos.mdb besitzen"
        'msg = msg & "Wollen Sie wirklich die Sprache wechseln?"
        Msg = LoadResString(2131 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2257 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2256 + Sprache)
        antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then Exit Sub                                                     'Gerbing 09.02.2007
    Else
        'msg = "Starten Sie nach derAuswahl der Sprache das Programm neu"
        Msg = LoadResString(2331 + Sprache) & vbNewLine
        MsgBox Msg
    End If

    If OptDeutsch.Value = True Then
        Sprache = 3000
        'rename englische Spaltennamen in deutsche
        'zuerst die Spalten
        SQLR = "EXEC sp_rename 'Fotos.Marker', 'Merker', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.year', 'Jahr', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.location', 'Ort', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.country', 'Land', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.people', 'Personen', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.filename', 'Dateiname', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.BWC', 'SWF', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.comment', 'Kommentar', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.FilenameShort', 'DateinameKurz', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.FileDate', 'DDatum', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.WidthPixel', 'BreitePixel', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.HightPixel', 'HoehePixel', 'COLUMN'"
        DBado.Execute (SQLR)
        
        Call WriteGlL(0)     'Rückschreiben 0=deutsch in fotos.ini
        Call GlL
        If PublicLanguage = "1" Then
            Call VierUrsachenFürSchreibsperre
            Call Query.Beenden
            End
        Else
            Call Query.Beenden
            End
        End If
    Else
        'OptEnglish.Value = True
        Sprache = 0
        'rename deutsche Spaltennamen in englische
        'zuerst die Spalten
        SQLR = "EXEC sp_rename 'Fotos.Merker', 'Marker', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Jahr', 'year', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Ort', 'location', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Land', 'country', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Personen', 'people', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Dateiname', 'filename', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.SWF', 'BWC', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.Kommentar', 'comment', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.DateinameKurz', 'FilenameShort', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.DDatum', 'FileDate', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.BreitePixel', 'WidthPixel', 'COLUMN'"
        DBado.Execute (SQLR)
        SQLR = "EXEC sp_rename 'Fotos.HoehePixel', 'HightPixel', 'COLUMN'"
        DBado.Execute (SQLR)
        
        Call WriteGlL(1)     'Rückschreiben 1=english in fotos.ini
        Call GlL
        If PublicLanguage = "0" Then
            Call VierUrsachenFürSchreibsperre
            Call Query.Beenden
            End
        Else
            Call Query.Beenden
            End
        End If
    End If
End Sub

Private Sub SpracheÄndernAccess()
    Dim rc As Long
    Dim n As Long
    Dim Msg As String
    Dim antwort As Long
    Dim SQL As String
    Dim sSource As String                                                                   'Gerbing 23.11.2017
    Dim sDest As String
    Dim tbl As ADOX.Table                                                                   'Gerbing 23.11.2017
    Dim fld As ADOX.Column                                                                  'Gerbing 23.11.2017

    If PublicLanguage <> "9" Then                                                           'Gerbing 04.12.2011
        'nichts tun wenn die Sprache nicht wirklich gewechselt wurde
        If OptDeutsch.Value = True And blnMerkeDeutsch = True Then Exit Sub
        If OptEnglish.Value = True And blnMerkeDeutsch = False Then Exit Sub
    End If

    If PublicLanguage <> "9" Then                                                           'Gerbing 04.12.2011
        'msg = "Starten Sie nach dem Wechsel der Sprache das Programm neu"
        'msg = msg & "Diese Funktion sollten Sie nur benutzen, wenn Sie eine Sicherungskopie der Datei fotos.mdb besitzen"
        'msg = msg & "Wollen Sie wirklich die Sprache wechseln?"
        Msg = LoadResString(2131 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2257 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2256 + Sprache)
        antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then Exit Sub                                                     'Gerbing 09.02.2007
    Else
        'msg = "Starten Sie nach derAuswahl der Sprache das Programm neu"
        Msg = LoadResString(2331 + Sprache) & vbNewLine
        MsgBox Msg
    End If

    If OptDeutsch.Value = True Then
        '---------------------------------------------
        'nur wenn die Spalte filename existiert, wird die Sprache gewechselt
        SQL = "SELECT * From fotos WHERE not filename Is Null;"

        On Error Resume Next
        'On Error GoTo 0
        'On Error GoTo QUERYERR
        If rstsql Is Nothing Then
            Set rstsql = New ADODB.Recordset
        Else
            rstsql.Close
        End If
        Err.Number = 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        '-------------------------------------------------------------
        If Err.Number = 0 Then
            'Spalte filename existiert
            'rename englische Spaltennamen in deutsche                                      'Gerbing 23.11.2017
            On Error GoTo 0
            Set oCat = New ADOX.Catalog
            Set oCat.ActiveConnection = DBado
            Set tbl = oCat.Tables("Fotos")
            tbl.Columns("Marker").Name = "Merker"
            tbl.Columns("year").Name = "Jahr"
            'Situation bleibt unverändert
            tbl.Columns("location").Name = "Ort"
            tbl.Columns("country").Name = "Land"
            tbl.Columns("people").Name = "Personen"
            tbl.Columns("filename").Name = "Dateiname"
            tbl.Columns("BWC").Name = "SWF"
            tbl.Columns("comment").Name = "Kommentar"
            tbl.Columns("FilenameShort").Name = "DateinameKurz"
            tbl.Columns("FileDate").Name = "DDatum"
            tbl.Columns("WidthPixel").Name = "BreitePixel"
            tbl.Columns("HightPixel").Name = "HoehePixel"
            Sprache = 3000
        End If
        On Error Resume Next
        rstsql.Close
        DBado.Close                                                                     'Gerbing 23.11.2017
        DollarDBado.Close                                                               'Gerbing 23.11.2017
        '----------------------
        Err.Number = 0
        'Durch Komprimieren wird die Datenbank $fotos.mdb erzeugt
        On Error Resume Next
        'DBEngine.CompactDatabase gstrFotosMdbLocation & "\fotos.mdb", gstrFotosMdbLocation & "\$fotos.mdb" 'Gerbing 23.11.2017
        sSource = gstrFotosMdbLocation & "\fotos.mdb"
        sDest = gstrFotosMdbLocation & "\$fotos.mdb"
        rc = file_delete(gstrFotosMdbLocation & "\$fotos.mdb", , True)
        If CompactDB(sSource, sDest) Then
            'MsgBox "Compact complete"
            If file_path_exist(gstrFotosMdbLocation & "\$fotos.mdb") = True Then
                rc = file_delete(gstrFotosMdbLocation & "\fotos.mdb", , True)
                'rc = file_copy(Quellname, Zielname)                                             'Gerbing 18.10.2017
                rc = file_copy(gstrFotosMdbLocation & "\$fotos.mdb", gstrFotosMdbLocation & "\fotos.mdb") 'Gerbing 18.10.2017
            End If
            'MsgBox "Komprimieren der Datenbank wurde ausgeführt"
        Else
            'MsgBox "Komprimieren der Datenbank wurde versucht, aber konnte nicht ausgeführt werden"
        End If
        Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
        Call WriteGlL(0)     'Rückschreiben in fotos.ini
        Sprache = 0
        'End                                                                                        'Gerbing 02.09.2008
        Call GlL                                                                                    'Gerbing 02.09.2008
        If PublicLanguage = "1" Then                                                                'Gerbing 04.12.2011
            Call VierUrsachenFürSchreibsperre
            Call Query.Beenden
            End
        Else
            Call Query.Beenden
            End
        End If
    End If
'-----------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------
    If OptEnglish.Value = True Then
        Sprache = 3000
        '---------------------------------------------
        'nur wenn die Spalte Dateiname existiert, wird die Sprache gewechselt
        SQL = "SELECT * From fotos WHERE not Dateiname Is Null;"

        On Error Resume Next
        'On Error GoTo 0
        'On Error GoTo QUERYERR
        If rstsql Is Nothing Then
            Set rstsql = New ADODB.Recordset
        Else
            rstsql.Close
        End If
        Err.Number = 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        '-------------------------------------------------------------
        If Err.Number = 0 Then
            'Spalte Dateiname existiert
            'rename deutsche Spaltennamen in englische
            On Error GoTo 0
            Set oCat = New ADOX.Catalog
            Set oCat.ActiveConnection = DBado
            Set tbl = oCat.Tables("Fotos")
            tbl.Columns("Merker").Name = "Marker"
            tbl.Columns("Jahr").Name = "year"
            'Situation bleibt unverändert
            tbl.Columns("Ort").Name = "location"
            tbl.Columns("Land").Name = "country"
            tbl.Columns("Personen").Name = "people"
            tbl.Columns("Dateiname").Name = "filename"
            tbl.Columns("SWF").Name = "BWC"
            tbl.Columns("Kommentar").Name = "comment"
            tbl.Columns("DateinameKurz").Name = "FilenameShort"
            tbl.Columns("DDatum").Name = "FileDate"
            tbl.Columns("BreitePixel").Name = "WidthPixel"
            tbl.Columns("HoehePixel").Name = "HightPixel"
        End If
        On Error Resume Next
        rstsql.Close
        DBado.Close                                                                     'Gerbing 23.11.2017
        DollarDBado.Close                                                               'Gerbing 23.11.2017
        '----------------------
        Err.Number = 0
        'Durch Komprimieren wird die Datenbank $fotos.mdb erzeugt
        On Error Resume Next
        'DBEngine.CompactDatabase gstrFotosMdbLocation & "\fotos.mdb", gstrFotosMdbLocation & "\$fotos.mdb" 'Gerbing 23.11.2017
        sSource = gstrFotosMdbLocation & "\fotos.mdb"
        sDest = gstrFotosMdbLocation & "\$fotos.mdb"
        rc = file_delete(gstrFotosMdbLocation & "\$fotos.mdb", , True)
        If CompactDB(sSource, sDest) Then
            'MsgBox "Compact complete"
            If file_path_exist(gstrFotosMdbLocation & "\$fotos.mdb") = True Then
                rc = file_delete(gstrFotosMdbLocation & "\fotos.mdb", , True)
                'rc = file_copy(Quellname, Zielname)                                             'Gerbing 18.10.2017
                rc = file_copy(gstrFotosMdbLocation & "\$fotos.mdb", gstrFotosMdbLocation & "\fotos.mdb") 'Gerbing 18.10.2017
            End If
            'MsgBox "Komprimieren der Datenbank wurde ausgeführt"
        Else
            'MsgBox "Komprimieren der Datenbank wurde versucht, aber konnte nicht ausgeführt werden"
        End If

        Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
        Call WriteGlL(1)     'Rückschreiben in fotos.ini
        Call GlL                                                                            'Gerbing 02.09.2008
        If PublicLanguage = "0" Then                                                          'Gerbing 02.09.2008 'Gerbing 04.12.2011
            Call VierUrsachenFürSchreibsperre
            Call Query.Beenden
            End
        Else
            Call Query.Beenden
            End
        End If
    End If
End Sub
