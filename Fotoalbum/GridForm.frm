VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form DbGridForm 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   ClientHeight    =   6216
   ClientLeft      =   132
   ClientTop       =   1116
   ClientWidth     =   15528
   Icon            =   "GridForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6216
   ScaleWidth      =   15528
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnSpaltenbreitenSpeichern 
      Caption         =   "&Spaltenbreiten speichern"
      Height          =   492
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Sie können die Spaltenbreite mit der Maus durch Ziehen verändern. Diese Einstellung wird hiermit gespeichert."
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton btnMerkerspalteEinschalten 
      Caption         =   "&Merkerspalte in jedem Datensatz ein/ausschalten"
      Height          =   492
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Mit der Merkerspalte können Sie Fotos vormerken, die Sie exportieren oder löschen oder weiterselektieren wollen"
      Top             =   240
      Width           =   5292
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11160
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   6585
      Left            =   0
      ScaleHeight     =   6540
      ScaleWidth      =   15432
      TabIndex        =   2
      Top             =   0
      Width           =   15480
      Begin VB.CommandButton btnThumbnails 
         Caption         =   "Thumbnails"
         Height          =   492
         Left            =   11760
         TabIndex        =   9
         Top             =   2160
         Width           =   2772
      End
      Begin VB.PictureBox PictureThumb 
         AutoSize        =   -1  'True
         Height          =   750
         Left            =   20
         ScaleHeight     =   708
         ScaleWidth      =   948
         TabIndex        =   8
         Top             =   20
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton btnShowUsers 
         Caption         =   "Show users"
         Height          =   492
         Left            =   11760
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.CommandButton btnRefresh 
         Height          =   495
         Left            =   14160
         Picture         =   "GridForm.frx":038A
         Style           =   1  'Grafisch
         TabIndex        =   6
         ToolTipText     =   "Aktualisieren - nur sinnvoll in Multiuser-Umgebung"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton btnAktionAusführen 
         Caption         =   "&Aktion wählen und ausführen..."
         Height          =   492
         Left            =   10440
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   8040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   372
         Left            =   8040
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _ExtentX        =   3620
         _ExtentY        =   656
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DBGridNeu 
         Height          =   2652
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   7932
         _ExtentX        =   13991
         _ExtentY        =   4678
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image ImageThumb 
         Height          =   132
         Left            =   11160
         Top             =   1320
         Width           =   120
      End
   End
End
Attribute VB_Name = "DbGridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Im Design-Mode eingestellt
'Adodc1.Recordsource = "Fotos"
'Adodc1.CursorLocation = adUseClient
'Adodc1.Cursortype = adOpenStatic
'Adodc1.LockType = adLockOptimistic
'Adodc1.Mode = adModeUnknown
'DbGridNeu.DataSource = "Adodc1" / DbGridNeu ist ein DataGrid(MSDatGrd.ocx)
'DbgridNeu.AllowUpdate = False bei Aufruf ohne Commandline argument
'DbgridNeu.AllowUpdate = True bei Aufruf mit Commandline argument /WRITE

Option Explicit
    Dim SQL As String
    Dim msg As String
    Dim blnMarkiert As Boolean
    Public rsDataGrid As adodb.Recordset
    Dim ZuvorMsg As String
    Dim Bookmark As Variant                                     'Gerbing 20.12.2010
    Dim lngGewählteSpalte As Long                               'Gerbing 01.01.2014
    Public SQLneuHeadClick As String

Private Sub btnAktionAusführen_Click()
    Dim retcode As Long
    Dim epm As New EasyPopupMenu                                'Gerbing 09.02.2007
        
        If Form1.lngPointer Then                                'Gerbing 31.12.2012 und 04.09.2013
            retcode = GdipDisposeImage(Form1.lngPointer)
        End If
        If m_lngGraphics Then                                   'Gerbing 31.12.2012 und 04.09.2013
            If GdipDeleteGraphics(m_lngGraphics) Then _
                'MsgBox "Graphics object could not be deleted", vbCritical
            End If
        End If
        epm.AddMenuItem Combo1.List(0), MF_STRING, 1    '0 'Öffnen der mit 'xyz' verknüpften Anwendung für die aktuelle Datei
        epm.AddMenuItem Combo1.List(1), MF_STRING, 2    '1 'Öffne das Druckprogramm für die aktuelle Datei
        epm.AddMenuItem Combo1.List(2), MF_STRING, 3    '2 'Öffne das Fenster 'Neue Email senden'
        epm.AddMenuItem Combo1.List(3), MF_STRING, 4    '3 'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
        epm.AddMenuItem Combo1.List(4), MF_STRING, 5    '4 'Öffne RenamMdb für die aktuelle Datei                               'Gerbing 27.08.2012 08.10.2014
        epm.AddMenuItem Combo1.List(5), MF_STRING, 6    '5 'Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
        epm.AddMenuItem Combo1.List(6), MF_STRING, 7    '6 'Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort Gerbing 23.01.200
        If gblnProversion = True Then
            If gblnSQLServerConnected = False Then      'bei sql server gibt es kein Hyperlink
                epm.AddMenuItem Combo1.List(7), MF_STRING, 8 '7 'Gehe zum Hyperlink
            End If
        End If
        Call WeiterAnShellExecute(epm.TrackMenu(Me.hWnd))
        epm.DeleteMenu
    Set epm = Nothing
End Sub

Private Sub btnMerkerspalteEinschalten_Click()
    'wechselweise Merkerspalte ein/ausschalten
    Dim SQLalt As String
    Dim SQL As String
    Dim SQLMitte As String
    Dim pos As Long
    Dim Pos1 As Long
    
    If gblnSchreibgeschützt = True Then                                 'Gerbing 15.05.2006
        msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        msg = msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    '--------------------------------------------------------------------------------------
    'zuerst alle Merkerspalten in der ganzen Tabelle Fotos ausschalten  'Gerbing 26.07.2006
    'SQL = "UPDATE Fotos SET Fotos.Merker = 0;"
    SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 0;"
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    'SQLalt = Adodc1.RecordSource                                        'Gerbing 24.12.2005
    SQLalt = Query.SQL                                                  'Gerbing 28.08.2006
    pos = InStr(1, SQLalt, "WHERE", vbTextCompare)                      'Gerbing 26.07.2006
    Pos1 = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    SQLMitte = Mid(SQLalt, pos, Pos1 - pos) & ";"
    '--------------------------------------------------------------------------------------
    'dann die Merkerspalten bezüglich der Suchkriterien wieder einschalten
    If blnMarkiert = True Then
        'SQL = "UPDATE Fotos SET Fotos.Merker = 1;"
        SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 1 "
        SQL = SQL & SQLMitte                                            'Gerbing 26.07.2006
        blnMarkiert = False
    Else
        'SQL = "UPDATE Fotos SET Fotos.Merker = 0;"
        SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 0 "
        SQL = SQL & SQLMitte                                            'Gerbing 26.07.2006
        blnMarkiert = True
    End If

    On Error Resume Next
    DbGridForm.rsDataGrid.Close
    On Error GoTo 0
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    '------------------------------------------------------
    On Error Resume Next
    DbGridForm.rsDataGrid.Close
    On Error GoTo 0
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQLalt
            .ActiveConnection = DBsql
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQLalt
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    
    Set DbGridForm.Adodc1.Recordset = DbGridForm.rsDataGrid
    Set DbGridForm.DBGridNeu.DataSource = DbGridForm.rsDataGrid
    DbGridForm.DBGridNeu.ReBind

    Call SetSpaltenBreite                               'Gerbing 24.12.2005
    Me.MousePointer = vbNormal                                                              'Gerbing 29.07.2007
End Sub


Private Sub btnShowUsers_Click()
    Dim cn As New adodb.Connection
    Dim Rs As New adodb.Recordset
    Dim i, j As Long

    'das steht auch in der Datei fotos.ldb
    'diese gibts nur in einer Multiuser-Umgebung und wird von selbst gelöscht, wenn es nur noch einen Nutzer gibt
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open "Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set Rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.

    Debug.Print Rs.Fields(0).Name, "", Rs.Fields(1).Name, _
    "", Rs.Fields(2).Name, Rs.Fields(3).Name

    While Not Rs.EOF
        Debug.Print Rs.Fields(0), Rs.Fields(1), _
        Rs.Fields(2), Rs.Fields(3)
        Rs.MoveNext
    Wend

End Sub

Private Sub btnSpaltenbreitenSpeichern_Click()
    If gblnSchreibgeschützt = True Then                                 'Gerbing 15.05.2006
        msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        msg = msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If

    Call SpeichernSpaltenBreite
End Sub

Public Sub btnRefresh_Click()
    'Gerbing 20.12.2010
    'Gerbing 22.12.2010 Query.SQL war falsch richtig ist Adodc1.RecordSource
    Dim intLeftcol As Integer
    
    ExportForm.blnExportGestartet = True                                        'Gerbing 26.01.2015 damit nicht kurz das erste Bild aufflackert
    Me.MousePointer = vbHourglass
    intLeftcol = DBGridNeu.LeftCol
    Bookmark = DBGridNeu.Bookmark
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen adOpenStatic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = Adodc1.RecordSource
            .ActiveConnection = DBsql
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen adOpenDynamic
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = Adodc1.RecordSource
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    Set DbGridForm.Adodc1.Recordset = DbGridForm.rsDataGrid
    Set DbGridForm.DBGridNeu.DataSource = DbGridForm.rsDataGrid
    DbGridForm.DBGridNeu.ReBind
    DBGridNeu.Bookmark = Bookmark
    If DbGridForm.DBGridNeu.SelBookmarks.Count = 1 Then                         'Gerbing 30.11.2012
        DbGridForm.DBGridNeu.SelBookmarks.Remove 0                              'Gerbing 30.11.2012
    End If                                                                      'Gerbing 30.11.2012
    DbGridForm.DBGridNeu.SelBookmarks.Add DbGridForm.rsDataGrid.Bookmark        'Gerbing 30.11.2012
    Call SetSpaltenBreite
    'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
    On Error Resume Next
    DBGridNeu.Scroll intLeftcol, 0
    On Error GoTo 0
    Me.MousePointer = vbDefault
    ExportForm.blnExportGestartet = False                                       'Gerbing 26.01.2015
End Sub

Private Sub btnThumbnails_Click()
    If gblnWasHeadClick = True Then
        Unload frmThumbs
        gblnWasHeadClick = False
    End If
    On Error Resume Next
    frmThumbs.Show
    If gblnFrmThumbsLoaded = True Then
        frmThumbs.WindowState = vbMaximized
    End If
End Sub

Private Sub DBGridNeu_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim Merker                  'Gerbing 02.05.2006
    Dim Jahr As String
    Dim n As Long
    Dim strTemp As String
    Dim maximum As String
    Dim Pos1 As Long

    If ColIndex = 3 Then                                                            'Gerbing 16.10.2014
        'Ort
        If Gefundenexifdatetimeoriginal = True And gblnVollversion = True Then
            strTemp = oCat.Tables("fotos").Columns("ort").Properties(8)                 '8=ValidationRule=Len([Ort])<33
            If strTemp <> "" Then
                Pos1 = InStr(1, strTemp, "<", vbTextCompare)
                If Pos1 <> 0 Then
                    maximum = Mid(strTemp, Pos1 + 1, Len(strTemp) - Pos1)
                    If IsNumeric(maximum) Then
                        Trim (DBGridNeu.Columns(3).Text)
                        If Not Len(DBGridNeu.Columns(3).Text) < maximum Then
                            MsgBox "Gültigkeitsregel: " & oCat.Tables("fotos").Columns("ort").Properties(7)    '7=ValidationText
                            'DBGridNeu.Columns(3).Text = left(DBGridNeu.Columns(3).Text, 32)           'begrenzen auf maximal erlaubte Bytes
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    If ColIndex = 8 Then                                                            'Gerbing 16.10.2014
        'Kommentar
        If Gefundenexifdatetimeoriginal = True And gblnVollversion = True Then
            strTemp = oCat.Tables("fotos").Columns("Kommentar").Properties(8)           '8=ValidationRule=Len([Kommentar])<2001
            If strTemp <> "" Then
                Pos1 = InStr(1, strTemp, "<", vbTextCompare)
                If Pos1 <> 0 Then
                    maximum = Mid(strTemp, Pos1 + 1, Len(strTemp) - Pos1)
                    If IsNumeric(maximum) Then
                        Trim (DBGridNeu.Columns(8).Text)
                        If Not Len(DBGridNeu.Columns(8).Text) < maximum Then
                            MsgBox "Gültigkeitsregel: " & oCat.Tables("fotos").Columns("Kommentar").Properties(7)    '7=ValidationText
                            'DBGridNeu.Columns(8).Text = left(DBGridNeu.Columns(8).Text, 2000)           'begrenzen auf maximal erlaubte Bytes
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    If ColIndex = 11 Then                                                                   'Gerbing 31.12.2007
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(11).Text) Then
            'BreitePixel
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (Prüfen1), wenn Sie BreitePixel neu berechnen wollen."
            'MsgBox msg                                             'Gerbing 08.11.2005
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2005 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If ColIndex = 12 Then                                                                   'Gerbing 31.12.2007
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(12).Text) Then
            'HoehePixel
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (Prüfen1), wenn Sie HoehePixel neu berechnen wollen."
            'MsgBox msg                                             'Gerbing 08.11.2005
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2006 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If ColIndex = 13 Then                                                                   'Gerbing 31.12.2007
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(13).Text) Then
            'AudioFileExists                                                                'Gerbing 31.12.2007
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (PrüfenS), wenn Sie Differenzen zwischen Audio-Kommentaren und der Spalte 'AudioFileExist' reparieren wollen"
            'MsgBox msg                                                                     'Gerbing 31.12.2007
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2265 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If ColIndex = 14 Then                                                                   'Gerbing 16.04.2008
        If Trim(OldValue) <> Trim(DBGridNeu.Columns(14).Text) Then
            'IPTCPresent                                                                    'Gerbing 16.04.2008
    '        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
    '        msg = msg & "Benutzen Sie fotosmdb (PrüfenIPTC), wenn Sie das Feld IPTCPresent neu berechnen lassen wollen"
            'MsgBox msg                                                                     'Gerbing 31.12.2007
            MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(3146 + Sprache)
            Cancel = True
            Exit Sub
        End If
    End If
    '-----------------------------------------------------------------------------------------------
    'Kontrolle, ob in Spalte SWF erlaubter Inhalt steht                                     'Gerbing 31.12.2007
    If DBGridNeu.Columns(DBGridNeu.Col).Caption = LoadResString(1029 + Sprache) Then        '1029=SWF
        strTemp = UCase(DBGridNeu.Text)
        If Sprache = 0 Then
            'das ist deutsch
            Select Case strTemp
                Case "SW", "F", "FV", "SV", "BW", "C", "CV", "BV"
                
                Case Else
                    'MsgBox "Falscher Wert in Spalte SWF"
                    MsgBox LoadResString(2264 + Sprache)
                    Cancel = True
                    Exit Sub
            End Select
        Else
            'das ist english
            Select Case strTemp
                Case "SW", "F", "FV", "SV", "BW", "C", "CV", "BV"
                
                Case Else
                    'MsgBox "Falscher Wert in Spalte SWF"
                    MsgBox LoadResString(2264 + Sprache)
                    Cancel = True
                    Exit Sub
            End Select
        End If
    End If
    '-----------------------------------------------------------------------------------------------
    If DBGridNeu.Col = 0 Then 'Gerbing 21.03.2006                       'DBGridNeu.Col = 0 = Merker
        Merker = DBGridNeu.Columns(0).Value
        Select Case Merker
            Case "0"

            Case "1"

            Case "-1"

            Case Else
                'MsgBox "Sie dürfen in die Spalte Merker nur 0 oder 1 eintragen"
                MsgBox LoadResString(2001 + Sprache)            'Gerbing 08.11.2005
                DBGridNeu.Columns(0).Value = "0"                   'Gerbing 21.03.2006
                Cancel = True
                Exit Sub
        End Select
    End If
    '-----------------------------------------------------------------------------------------------
    If DBGridNeu.Col = 1 Then                           'Gerbing 14.05.2006 DBGridNeu.Col = 1 = Jahr
        Jahr = DBGridNeu.Columns(1).Text                'Gerbing 21.03.2006
        If Len(Jahr) <> 4 Then
            'MsgBox "Jahr muß eine 4-stellige Zahl sein"
            MsgBox LoadResString(2127 + Sprache)
            Cancel = True
            Exit Sub
        End If
        If Not IsNumeric(Jahr) Then
            'MsgBox "Jahr muß eine 4-stellige Zahl sein"
            MsgBox LoadResString(2127 + Sprache)
            Cancel = True
            Exit Sub
        End If
        'MsgBox "Klicken Sie nach Änderung des Feldes Jahr in die Nachbarspalte der gleichen Zeile" 'Gerbing 14.05.2006
        MsgBox LoadResString(2249 + Sprache)
    End If
    '-----------------------------------------------------------------------------------------------
    'Untersuchen ob die aktuelle Spalte eine Spalte mit dbHyperlinkField ist                'Gerbing 27.12.2007
    For n = 1 To HyperlinkFieldColumns.Count
        If HyperlinkFieldColumns(n) = DBGridNeu.Col Then
            If Not Right(DBGridNeu.Text, 1) = "#" Or Not Left(DBGridNeu.Text, 1) = "#" Then
                msg = LoadResString(1033 + Sprache)                                         '1033=Feldname
                msg = msg & "=" & DBGridNeu.Columns(DBGridNeu.Col).Caption & vbNewLine
                msg = msg & LoadResString(1034 + Sprache)                                   '1034=Feldinhalt
                msg = msg & "=" & DBGridNeu.Text & vbNewLine
                msg = msg & LoadResString(2263 + Sprache)   'Falsches Hyperlink-Format. Ein Hyperlink muss in # eingeschlossen sein. Beispielsweise #http://www.gerbingsoft.de#
                'MsgBox Msg
                MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
                Cancel = True
                Exit Sub
            End If
        End If
    Next n
    '-----------------------------------------------------------------------------------------------
    'Falls 'erster Treffer pro Jahr', dann Stichwortänderung in Tabelle Fotos nachziehen
    If Query.optNurErstenTreffer.Value = True Or Query.optErsterZufallstreffer.Value = True Then        'Gerbing 09.02.2013
'        msg = "Spalte=" & DBGridNeu.Columns(DBGridNeu.Col).Caption & vbNewLine
'        msg = msg & "Inhalt=" & DBGridNeu.Text & vbNewLine
'        msg = msg & "Dateiname=" & DBGridNeu.Columns(6).Text
'        MsgBox msg
        'SQL = "Select * FROM Fotos Where Dateiname=" & """" & DBGridNeu.Columns(6).Text  & """"
        SQL = "Select * FROM Fotos Where " & LoadResString(1028 + Sprache) & "='" & DBGridNeu.Columns(6).Text & "'"
        With rstsql
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        Do Until rstsql.EOF
            On Error Resume Next            'zur Fehlerabwehr wenn eine Spalte zB auf 2 Zeichen begrenzt ist
            rstsql.Fields(DBGridNeu.Columns(DBGridNeu.Col).Caption) = DBGridNeu.Text
            On Error GoTo 0
            rstsql.Update
            rstsql.MoveNext
        Loop
        rstsql.Close
        Call Query.FrageObNurErstenTreffer                                                  'Gerbing 26.08.2008
        'Hier wird 'Set DBGridNeu.Datasource' erneut ausgeführt weil dem Grid DBGridNeu eine Abfrage mit
        'Inner Join zugrunde liegt und ein solches Recordset cannot be updated
    End If
End Sub

Private Sub DBGridNeu_BeforeUpdate(Cancel As Integer)
    'MsgBox "DBGridNeu_BeforeUpdate"                                                         'Gerbing 15.10.2014
End Sub

Private Sub DbGridNeu_Change()
    Dim strTemp As String
    
    'Achtung eigenartiges Verhalten beim Debuggen
    'Beim Setzen des Haltepunktes auf zB die Zeile 'If Adodc1.Recordset.Fields(LoadResString(1028 + Sprache)) <> DBGridNeu.Columns(6).Text Then'
    'da geht er bei Ungleichheit zu End If
    
    'On Error GoTo BeiFehlerKeinePrüfung                     'Gerbing 22.01.2007
    If Adodc1.Recordset.Fields(LoadResString(1028 + Sprache)) <> DBGridNeu.Columns(6).Text Then
        'Dateiname
        DBGridNeu.Columns(6).Text = Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
'        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
'        msg = msg & "Benutzen Sie RenamMdb zur übereinstimmenden Änderung in der Datenbank und im Ordner."
        'MsgBox msg                                             'Gerbing 08.11.2005
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2003 + Sprache)
        Exit Sub
    End If
    If Adodc1.Recordset.Fields(LoadResString(1031 + Sprache)) <> DBGridNeu.Columns(9).Text Then
        'DateinameKurz
        DBGridNeu.Columns(9).Text = Adodc1.Recordset.Fields(LoadResString(1031 + Sprache))
'        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
'        msg = msg & "Benutzen Sie RenamMdb zur übereinstimmenden Änderung in der Datenbank und im Ordner."
        'MsgBox msg                                             'Gerbing 08.11.2005
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2003 + Sprache)
        Exit Sub
    End If
    If Trim(Adodc1.Recordset.Fields(LoadResString(1032 + Sprache))) <> Trim(DBGridNeu.Columns(10).Text) Then
        'DDatum
        DBGridNeu.Columns(10).Text = Adodc1.Recordset.Fields(LoadResString(1032 + Sprache))
'        msg = "Änderungen in diesem Feld sind verboten." & vbNewLine
'        msg = msg & "Benutzen Sie fotosmdb (Prüfen1), wenn Sie DDatum aktualisieren wollen."
        'MsgBox msg                                             'Gerbing 08.11.2005
        MsgBox LoadResString(2002 + Sprache) & vbNewLine & LoadResString(2004 + Sprache)
        Exit Sub
    End If
BeiFehlerKeinePrüfung:
    On Error GoTo 0
End Sub

Private Sub DbGridNeu_Click()
'    If gstrRowColChangeName = gstrFRODN Then Exit Sub
'    If Adodc1.Recordset.Fields(LoadResString(1028 + Sprache)) = gstrFRODN Then Exit Sub
End Sub

Private Sub DbGridNeu_DblClick()
    'es gibt irgendein Problem, wenn ich in der DbGridForm einen Doppelklick auf eine Spaltenüberschrift
    'ausführe. Das Image1 verrutscht außerhalb der Zentrierung. Es sieht so aus, als würde der Doppelklick
    'gleichzeitig als Verschiebeklick aufgefasst.
    'Ich erfinde den Schalter gblnDbGridFormDblClick            'Gerbing 06.06.2005
    gblnDbGridFormDblClick = True                               'Gerbing 06.06.2005
    DbGridForm.Hide
    Hilfebx.Hide                                                'Gerbing 16.09.2004
    Call Form1.BildAnzeigen
End Sub

Private Sub DBGridNeu_Error(ByVal DataError As Integer, Response As Integer)
    'MsgBox "DbGridNeu_Error " & DBGridNeu.ErrorText

    Response = 0
    Adodc1.Recordset.CancelUpdate
End Sub

Private Sub DbGridNeu_HeadClick(ByVal ColIndex As Integer)
    'In dieser Prozedur funktioniert keinerlei Mousepointer auf Hourglass setzen
    'Diese Prozedur kann nur richtig arbeiten, wenn es im SQL String eine ORDER BY Anweisung gibt
    'es gibt 4 Varianten wie ein SQL String aufgebaut wird. Mit allen 4 Varianten muss diese Prozedur
    'fertig werden.
    '1. Normale Suche nach Suchkriterien
    '2. Suchen Differenzen Jahr
    '3. Gespeicherte Abfrage
    '4. ErsterTreffer
    
    Dim SQL As String
    Dim SQLalt As String
    Dim pos As Long
    Dim Links As String
    Dim ColCaption As String
    Dim intLeftcol As Integer
    Dim n As Long
    
    'DoEvents 'auskommentiert Gerbing 10.09.2009
    intLeftcol = DBGridNeu.LeftCol                                      'Gerbing 13.03.2005
    ColCaption = DBGridNeu.Columns(ColIndex).Caption
    'SQLalt = Query.SQL                                                  'Gerbing 16.06.2005
    SQLalt = Adodc1.RecordSource                                         'Gerbing 24.12.2005
    '----------------------------------------------------------------------------------------
    pos = InStr(1, SQLalt, "DESC", vbTextCompare)
    If pos <> 0 Then
        SQL = " ORDER BY [" & ColCaption & "];"                        'Gerbing 22.02.2005
    Else
        SQL = " ORDER BY [" & ColCaption & "] DESC;"
    End If
    pos = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    If pos <> 0 Then                                                    'Gerbing 20.06.2006
        Links = Left(SQLalt, pos - 1)
    Else
        Links = SQLalt
        'wenn ein Semikolon am Ende steht, dann abschneiden
        pos = InStr(1, Links, ";")
        If pos <> 0 Then
            Links = Mid(Links, 1, pos - 1)
        End If
    End If
    SQLneuHeadClick = Links & SQL
    'Query.SQL = SQLneuHeadClick                                                  'Gerbing 16.06.2005
    
    On Error Resume Next
    DbGridForm.rsDataGrid.Close
    On Error GoTo 0
    
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQLneuHeadClick
            .ActiveConnection = DBsql
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen
        Set DbGridForm.rsDataGrid = New adodb.Recordset
        With DbGridForm.rsDataGrid
            .Source = SQLneuHeadClick
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    Set DbGridForm.Adodc1.Recordset = DbGridForm.rsDataGrid
    Set DbGridForm.DBGridNeu.DataSource = DbGridForm.rsDataGrid
    DbGridForm.DBGridNeu.ReBind
    
    Call SetSpaltenBreite
    'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
    On Error Resume Next
    DBGridNeu.Scroll intLeftcol, 0
    On Error GoTo 0
    DbGridForm.DBGridNeu.AllowUpdate = True
    gblnWasHeadClick = True
End Sub

Private Sub DBGridNeu_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeySubtract And Shift = vbCtrlMask Then 'vbKeySubtract ist das Minuszeichen auf dem Ziffernblock 'Gerbing 01.01.2014 04.02.2014
    If KeyCode = vbKeyMultiply And Shift = vbCtrlMask Then 'vbKeyMultiply ist das Multiplikationszeichen auf dem Ziffernblock 'Gerbing 01.01.2014 12.02.2014

        lngGewählteSpalte = DBGridNeu.Col
        'DBGridNeu.Text in die Zwischenablage kopieren
        ClipboardSetText Me.hWnd, DBGridNeu.Text                                        'Gerbing 04.02.2014
        KeyCode = 0                             'sonst wird in die Zelle 'c' eingetragen
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------
    If KeyCode = 40 Or KeyCode = 38 Then                                                'Gerbing 04.03.2013
        If DbGridForm.DBGridNeu.SelBookmarks.Count = 1 Then                             'Gerbing 04.03.2013
            DbGridForm.DBGridNeu.SelBookmarks.Remove 0                                  'Gerbing 04.03.2013
        End If                                                                          'Gerbing 04.03.2013
    End If
    If KeyCode = 40 Then 'Pfeil nach unten
        'damit wird die aktuelle Zeile schwarz
        DbGridForm.DBGridNeu.SelBookmarks.Add DbGridForm.rsDataGrid.Bookmark + 1        'Gerbing 04.03.2013
    End If
    If KeyCode = 38 Then 'Pfeil nach oben
        'damit wird die aktuelle Zeile schwarz
        DbGridForm.DBGridNeu.SelBookmarks.Add DbGridForm.rsDataGrid.Bookmark - 1        'Gerbing 04.03.2013
    End If
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10                      'Gerbing 21.09.2014
            Me.Hide
            'Tastatur-Eingabe weiterreichen
            '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                Call Form1.Form_KeyDown(KeyCode, Shift)
        Case vbKeyF5                                                                    'Gerbing 21.09.2014
            If Shift = vbShiftMask Then
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
    End Select
End Sub

Private Sub DBGridNeu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then                                              'Gerbing 26.11.2012
        Call Form1.Hilfebox
    End If
End Sub

Private Sub DBGridNeu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)    'Gerbing 01.01.2014
    Dim i As Long
    
    If Shift = vbShiftMask + vbCtrlMask Then
        'Kopieren mit Multiselect
        HandleDTGMultiSelect DBGridNeu, Y ', LastRow, dtgGDLookup.row
        If lngGewählteSpalte <> 0 Then
            Select Case lngGewählteSpalte
                Case 6, 9, 10, 11, 12, 13, 14                                                   'das sind die verbotenen Spalten
                    'MsgBox "verbotene Spalte"
                Case -1
                    'da ist keine Spalte gewählt worden                                         'Gerbing 13.04.2014
                Case Else
                    For i = 0 To DBGridNeu.SelBookmarks.Count - 1
                        gblnComeFromF2F3 = True                                                 'Gerbing 27.03.2014
                        DBGridNeu.Col = lngGewählteSpalte
                        DBGridNeu.FirstRow = DBGridNeu.SelBookmarks(i) - 1                      'DBGridNeu.Row muss sichtbar sein, sonst
                        If DBGridNeu.SelBookmarks(i) - DBGridNeu.FirstRow >= 0 Then
                            DBGridNeu.Row = DBGridNeu.SelBookmarks(i) - DBGridNeu.FirstRow      'Laufzeitfehler 6148 Ungültige Zeilennummer
                        End If
                        'DBGridNeu.Text = Clipboard.GetText(1)
                         'DBGridNeu.Text wieder aus der Zwischenablage auslesen
                        DBGridNeu.Text = ClipboardGetText(Me.hWnd)                              'Gerbing 04.02.2014
                    Next i
            End Select
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF8, vbKeyF10                              'Gerbing 26.09.2013
            Me.Hide
            'Tastatur-Eingabe weiterreichen
            '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                Call Form1.Form_KeyDown(KeyCode, Shift)
        Case vbKeyF5                                                                            'Gerbing 21.09.2014
            If Shift = vbShiftMask Then
                Me.Hide
                'Tastatur-Eingabe weiterreichen
                '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
                    Call Form1.Form_KeyDown(KeyCode, Shift)
            End If
    End Select
End Sub


Private Sub DbGridNeu_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Bei RowColChange stehen noch die Daten der Zeile zur Verfügung die den Fokus verloren hat
    'Bei SelChange stehen die Daten der Zeile zur Verfügung die den Fokus bekommen hat
    
    Dim DateinamenErweiterung As String
    Dim CellVal As String
    Dim intLänge As Integer
    Dim strTemp As String
    Dim pos As Long                                                                         'Gerbing 29.03.2015
    Dim MerkeIndex As Long
    Dim i As Long

    If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
        Exit Sub
    End If
    If gblnComeFromThumbs = True Then                                                       'Gerbing 29.03.2015
        Exit Sub
    End If
    On Error Resume Next                            'Gerbing 21.03.2006
    'Falls 'erster Treffer pro Jahr', kein Update weil wegen Inner Join cannot be updated   'Gerbing 26.08.2008
    If Query.optNurErstenTreffer.Value = False Then                                         'Gerbing 26.08.2008
        Adodc1.Recordset.Update                                                             'Gerbing 26.08.2008
    End If                                                                                  'Gerbing 26.08.2008
    'On Error GoTo 0                                 'Gerbing 21.03.2006
    On Error GoTo RowColChangeError                                                         'Gerbing 25.06.2013
    gstrRowColChangeName = Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 3, 1) = "." Then             'Gerbing 25.06.2006
        intLänge = 3
    End If
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 4, 1) = "." Then
        intLänge = 4
    End If
    If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 5, 1) = "." Then
        intLänge = 5
    End If
    DateinamenErweiterung = Right(gstrRowColChangeName, intLänge)
    DateinamenErweiterung = UCase(DateinamenErweiterung)            'Gerbing 26.11.2006
    'Gerbing 20.08.2006
    DbGridForm.Combo1.Clear                                         'Gerbing 20.08.2006
    strTemp = LoadResString(3111 + Sprache)
    strTemp = strTemp & DateinamenErweiterung
    strTemp = strTemp & LoadResString(3110 + Sprache)
    DbGridForm.Combo1.AddItem strTemp, 0                     '0 'Öffnen der mit '" ' verknüpften Anwendung für die aktuelle Datei
    DbGridForm.Combo1.AddItem LoadResString(3119 + Sprache), 1 'Öffne das Druckprogramm für die aktuelle Datei
    DbGridForm.Combo1.AddItem LoadResString(3120 + Sprache), 2 'Öffne das Fenster 'Neue Email senden'
    DbGridForm.Combo1.AddItem LoadResString(3121 + Sprache), 3 'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
    DbGridForm.Combo1.AddItem LoadResString(3145 + Sprache), 4 'Öffne RenamMdb für die aktuelle Datei                           'Gerbing 27.08.2012 08.10.2014
    DbGridForm.Combo1.AddItem LoadResString(3123 + Sprache), 5 'Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
    DbGridForm.Combo1.AddItem LoadResString(3124 + Sprache), 6 'Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort Gerbing 23.01.2007
    If gblnProversion = True Then
        DbGridForm.Combo1.AddItem LoadResString(3125 + Sprache), 7 'Gehe zum Hyperlink
    End If
    DbGridForm.Combo1.ListIndex = 0
    'test ob ich das aktuelle Bild als Thumbnail anzeigen sollte oder nicht
    'Gerbing 26.11.2006
    If ExportForm.blnExportGestartet = False Then                               'Gerbing 08.12.2006
        Select Case DateinamenErweiterung
            Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "TIFF", "WMF"
                'nur wenn es tatsächlich eine Bilddatei ist
                Call ThumbnailAnzeigen(gstrRowColChangeName)
            Case Else
                Call ThumbnailAnzeigen("")
        End Select
        '--------------------------------------------------------------------------------------
        If Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\") = gstrFRODN Then    'Gerbing 27.03.2014
            gblnComeFromF2F3 = True                                                             'Gerbing 27.03.2014
        End If                                                                                  'Gerbing 27.03.2014
        If gblnComeFromF2F3 = False And Form1.F6Continous = False Then                                      'Gerbing 27.03.2014
            Call Form1.BildAnzeigen                                                             'Gerbing 27.03.2014
            DbGridForm.Show                                                                     'Gerbing 27.03.2014
            If DbGridForm.DBGridNeu.SelBookmarks.Count = 1 Then                                 'Gerbing 27.03.2014
                DbGridForm.DBGridNeu.SelBookmarks.Remove 0                                      'Gerbing 27.03.2014
            End If                                                                              'Gerbing 27.03.2014
            DbGridForm.DBGridNeu.SelBookmarks.Add DbGridForm.rsDataGrid.Bookmark                'Gerbing 27.03.2014
        End If                                                                                  'Gerbing 27.03.2014
        'Falls es Thumbnails gibt, soll der zugehörige Thumbnail selektiert werden (blaue Umrandung)        'Gerbing 29.03.2015
        If gblnFrmThumbsLoaded = True Then
            'gstrFRODN enthält den vollen Dateiname
        
            For i = 0 To frmThumbs.optThumb.Count - 1
                pos = InStr(1, frmThumbs.Ulabel(i).Tag, gstrFRODN, vbTextCompare)
                If pos <> 0 Then
                    Exit For
                End If
            Next i
            MerkeIndex = i
            For i = 0 To frmThumbs.optThumb.Count - 1                         'etwa schon markierte ausschalten
                frmThumbs.optThumb(i).BackColor = vbButtonFace
            Next i
            frmThumbs.optThumb(MerkeIndex).BackColor = vbBlue                 'das gefundene wird blau
            frmThumbs.vsbSlide.Value = (frmThumbs.vsbSlide.Max \ frmThumbs.Koll.Count) * MerkeIndex

        
        End If
        gblnComeFromF2F3 = False                                                                'Gerbing 27.03.2014
    End If
    Exit Sub
RowColChangeError:
    msg = "Errornumber=" & Err.Number & vbNewLine
    msg = msg & Err.Description
    'MsgBox msg
    Resume Next
End Sub

Private Sub DbGridNeu_SelChange(Cancel As Integer)
    'Bei RowColChange stehen noch die Daten der Zeile zur Verfügung die den Fokus verloren hat
    'Bei SelChange stehen die Daten der Zeile zur Verfügung die den Fokus bekommen hat
    'SelChange kommt nur dran When the user selects a single row by clicking its record selector
End Sub
'
'Private Sub Form_Load()
'    Dim strTemp As String
'
'    Call AnpassenFontSize(Me)                                       'Gerbing 23.06.2011
'    Call AnpassenHeadFont(DbGridForm.DBGridNeu)                     'Gerbing 23.06.2011
'    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
'        Me.Top = Form1.Top
'        Me.Left = Form1.Left
'        Me.Width = Form1.Width
'    Else
'        DbGridForm.Top = 0                                          'Gerbing 16.09.2006
'        DbGridForm.Left = 0
'        DbGridForm.Width = Screen.Width
'    End If
'    Picture1.Width = Screen.Width                                   'Gerbing 16.09.2004
'    Picture1.Height = Screen.Height                                 'Gerbing 16.09.2004
'    DBGridNeu.Width = DbGridForm.Width - 250
'    If gblnF5Alt = True Then                                        'Gerbing 22.04.2014
'        DBGridNeu.Height = DbGridForm.Height - 440
'    Else
'        DBGridNeu.Height = DbGridForm.Height - 1280                 'Gerbing 05.12.2010
'    End If
'    DBGridNeu.RowHeight = 250                                       'Gerbing 29.03.2012
'    DBGridNeu.AllowRowSizing = False
'    '------------------------------
'
'    'Adodc1.Connect = "Access 2000;"
'
'    btnMerkerspalteEinschalten.ToolTipText = LoadResString(2504 + Sprache) 'Mit der Merkerspalte können Sie Fotos vormerken, die Sie exportieren oder löschen oder weiterselektieren wollen
'    btnSpaltenbreitenSpeichern.ToolTipText = LoadResString(2505 + Sprache) 'Sie können die Spaltenbreite mit der Maus durch Ziehen verändern. Diese Einstellung wird hiermit gespeichert.
'    btnRefresh.ToolTipText = LoadResString(1126 + Sprache)                  'Aktualisieren-nur sinnvoll in Multiuser-Umgebung       'Gerbing 20.12.2010
'    btnSpaltenbreitenSpeichern.Caption = LoadResString(3003 + Sprache) '&Spaltenbreiten speichern
'    btnMerkerspalteEinschalten.Caption = LoadResString(3004 + Sprache)  '&Merkerspalte in jedem Datensatz ein/ausschalten
'    btnAktionAusführen.Caption = LoadResString(3117 + Sprache)  '&Aktion wählen und ausführen...
'    'Gerbing 20.08.2006
'    Combo1.ToolTipText = LoadResString(3118 + Sprache) 'Wählen Sie eine Aktion aus, die mit der aktuellen Datei ausgeführt werden soll
'    'combo1.AddItem "Öffnen der mit 'jpg' verknüpften Anwendung für die aktuelle Datei"
'    strTemp = LoadResString(3111 + Sprache) '&Öffnen der mit '"
'    strTemp = strTemp & "JPG"
'    strTemp = strTemp & LoadResString(3110 + Sprache)  ' verknüpften Anwendung für die aktuelle Datei
'    Combo1.AddItem strTemp, 0
'    Combo1.AddItem LoadResString(3119 + Sprache), 1 'Öffne das Druckprogramm für die aktuelle Datei
'    Combo1.AddItem LoadResString(3120 + Sprache), 2 'Öffne das Fenster 'Neue Email senden'
'    Combo1.AddItem LoadResString(3121 + Sprache), 3 'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
'    Combo1.AddItem LoadResString(3145 + Sprache), 4 'Öffne RenamMdb für die aktuelle Datei                              'Gerbing 27.08.2012 08.10.2014
'    Combo1.AddItem LoadResString(3123 + Sprache), 5 'Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
'    Combo1.AddItem LoadResString(3124 + Sprache), 6 'Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort Gerbing 23.01.2007
'    If gblnProversion = True Then
'        DbGridForm.Combo1.AddItem LoadResString(3125 + Sprache), 7 'Gehe zum Hyperlink
'    End If
'    Combo1.ListIndex = 0
'End Sub
'
'Private Sub Form_Paint()
'    Dim msg As String
'
'    On Error Resume Next                                            'Gerbing 15.11.2012
'    DbGridForm.Show                                                 'Gerbing 06.11.2012
'    'DbGridForm.Caption = "  Zum Auswählen eines Bildes Doppel-klicken Sie in die gewünschte Zeile"
'    DbGridForm.Caption = LoadResString(1010 + Sprache)              'Gerbing 08.11.2005
'    On Error Resume Next
'    If Query.CheckDifferenzen.Value = 0 Then                        'Gerbing 09.02.2005
'        'DbGridForm.Caption = "Bildanzahl=" & Query.RecordCount & DbGridForm.Caption 'Gerbing 16.06.2005
'        DbGridForm.Caption = LoadResString(1011 + Sprache) & Query.RecordCount & DbGridForm.Caption 'Gerbing 08.11.2005
'    End If
'    If Err = 91 Then    'Objektvariable oder With-Blockvariable nicht festgelegt
'        msg = "Es wurde kein einziger Datensatz gefunden." & NL
'        msg = msg & "Mit der F8-Taste können Sie die Suche wiederholen"
'        'MsgBox msg                                                 'Gerbing 08.11.2005
'        MsgBox LoadResString(2007 + Sprache) & NL & LoadResString(2008 + Sprache)
'        Exit Sub
'    End If
'End Sub
'
'Private Sub Form_Resize()
'    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
'        DbGridForm.Width = Form1.Width
'    Else
'        DbGridForm.Width = Screen.Width
'    End If
'    DBGridNeu.Width = DbGridForm.Width - 250
'    On Error Resume Next
'    If gblnF5Alt = True Then                                        'Gerbing 22.04.2014
'        DBGridNeu.Height = DbGridForm.Height - 440
'    Else
'        DBGridNeu.Height = DbGridForm.Height - 1280                 'Gerbing 05.12.2010
'    End If
'    On Error GoTo 0
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    'Query.Hide     'Gerbing 15.11.2012     'Gerbing 06.07.2003
'    Me.Hide
'    If gblnComefromVideo = True Then                                            'Gerbing 16.06.2012
'        frmVideo.Show
'    End If
'    Cancel = True       'ich will kein Unload
'End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then                                              'Gerbing 26.11.2012
        Call Form1.Hilfebox
    End If
End Sub

Public Sub WeiterAnShellExecute(ID)                                             'Gerbing 09.02.2007
   'Der echte Dateiname steht in gstrFRODN oder in gstrRowColChangeName
    Dim RetVal As Long
    Dim DateinamenErweiterung As String
    Dim intLänge As Integer
    Dim ErrorText As String
    Dim strLinks As String
    Dim strRechts As String
    Dim pos As Long
    Dim Pos1 As Long
    Dim antwort As Long
    Dim KeyCode As Integer
    Dim Shift As Integer
    Dim strTemp As String
    Dim AppId
    Dim cmdline As String                                                    'Gerbing 07.11.2011
   
    If ID = 0 Then Exit Sub     'Bei ID = 0 wurde keine Aktion ausgewählt           'Gerbing 09.02.2007
    ID = ID - 1
    gstrFRODN = Replace(gstrRowColChangeName, "+:\", gstrFotosMdbLocation & "\")                         'Gerbing 04.01.2006
    If gstrFRODN = "" Then                                                              'Gerbing 09.07.2008
        'wenn es noch kein Ereignis DbGridNeu_RowColChange gab, dann ist gstrRowColChangeName und damit gstrFRODN leer
        gstrFRODN = Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
        gstrFRODN = Replace(gstrFRODN, "+:\", gstrFotosMdbLocation & "\")
    End If
'   Werte für Combo1.Listindex                                                              'Gerbing 20.08.2006
'   0 = Öffnen der mit 'jpg' verknüpften Anwendung für die aktuelle Datei
'   1 = Öffne das Druckprogramm für die aktuelle Datei
'   2 = Öffne das Fenster 'Neue Email senden'
'   3 = Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
'   4 = Öffne RenamMdb für die aktuelle Datei                                               'Gerbing 27.08.2012 08.10.2014
'   5 = Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
'   6 = Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort                   'Gerbing 23.01.2007
    'If gblnProversion = True Then
    '   7 = Gehe zum Hyperlink                                                              'Gerbing 04.02.2007
    'end if
    
    Select Case ID
    'Select Case Combo1.ListIndex
        Case 0
            '0 = Öffnen der mit 'jpg' verknüpften Anwendung für die aktuelle Datei
            'RetVal = RunShellExecute(Me.hWnd, "open", gstrFRODN, vbNull, vbNull, 1)
            'RetVal = RunShellExecute(Me.hWnd, "open", gstrFRODN, vbNullString, vbNullString, 1)    'Gerbing 31.12.2007
            RetVal = RunShellExecute(Me.hWnd, vbNullString, gstrFRODN, vbNullString, vbNullString, 1)    'Gerbing 18.01.2014
            If RetVal <= 32 Then
                If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 3, 1) = "." Then           'Gerbing 25.06.2006
                    intLänge = 3
                End If
                If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 4, 1) = "." Then
                    intLänge = 4
                End If
                If Mid(gstrRowColChangeName, Len(gstrRowColChangeName) - 5, 1) = "." Then
                    intLänge = 5
                End If
                DateinamenErweiterung = Right(gstrRowColChangeName, intLänge)
                ErrorText = GetShellError(RetVal)           'Gerbing 20.08.2008
                msg = "Errortext=" & ErrorText & vbNewLine
                msg = msg & "Errornr=" & RetVal & vbNewLine & vbNewLine
                
                'Msg = "Der Dateiname lautet nach Ersetzen von +:\ folgendermaßen:" & vbNewLine
                msg = msg & LoadResString(1375 + Sprache) & vbNewLine
                msg = msg & gstrFRODN & vbNewLine
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
                'MsgBox Msg
                MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
            End If
        Case 1
            '1 = Öffne das Druckprogramm für die aktuelle Datei
'            'RetVal = RunShellExecute(Me.hWnd, "print", gstrFRODN, vbNull, vbNull, 1)
'            RetVal = RunShellExecute(Me.hWnd, "print", gstrFRODN, vbNullString, vbNullString, 1)   'Gerbing 31.12.2007
'            If RetVal <= 32 Then
'                MsgBox LoadResString(3122 + Sprache) 'Es wurde kein geeignetes Druckprogramm gefunden um diese Datei auszudrucken
'            End If
            Shell ("rundll32.exe SHELL32,OpenAs_RunDLL " & gstrFRODN)                                'Gerbing 07.08.2013
        Case 2
            '2 = Öffne das Fenster 'Neue Email senden'
            'RetVal = RunShellExecute(Me.hWnd, "open", "mailto:xxx@yyy.zzz?", vbNull, vbNull, 1)
            RetVal = RunShellExecute(Me.hWnd, "open", "mailto:xxx@yyy.zzz?", vbNullString, vbNullString, 1) 'Gerbing 31.12.2007
        Case 3
            'Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist     Gerbing 12.11..2007
            'Hierbei gibt es einen Fehler wenn im Dateiname ein Komma
            'enthalten ist -> "Der Pfad '...Teil hinter dem Komma' ist nicht vorhanden oder weist auf kein
            'Verzeichnis
            'Man muss den Dateinamen in doppelte Hochkomma einschließen
            'RetVal = RunShellExecute(Me.hWnd, "open", "explorer.exe", "/e,/select," & """" & gstrFRODN & """", vbNull, 1) 'Gerbing 12.11.2007
            RetVal = RunShellExecute(Me.hWnd, "open", "explorer.exe", "/e,/select," & """" & gstrFRODN & """", vbNullString, 1) 'Gerbing 31.12.2007
        Case 4
            '4 = Öffne RenamMdb für die aktuelle Datei                                              'Gerbing 27.08.2012 08.10.2014
            'If Dir(AppPath & "\RenamMdb.exe") = "" Then
            If file_path_exist(AppPath & "\RenamMdb.exe") = False Then
                'msg = "RenamMdb konnte nicht gestartet werden." & vbNewLine
                msg = LoadResString(2169 + Sprache) & vbNewLine
                'msg = msg & "RenamMdb.exe muss im gleichen Ordner stehen wie fotos.exe"
                msg = msg & LoadResString(2170 + Sprache)
                MsgBox msg
                Exit Sub
            End If
            'CommandLine aufbauen mit access
                'RowColChangeName=...;                                                  'Gerbing 09.10.2014
                'fotosmdblocation=...;
            
            'CommandLine aufbauen mit sql server
                'RowColChangeName=...;                                                  'Gerbing 09.10.2014
                'sqlservername=...;
                'datenbankname=...;
                'WindowsAuthentication=0; heißt nein
                'WindowsAuthentication=1; heißt ja
                'username=...;
                'Password=...;
                'StandortFotos=...;
        
            'CommandLine aufbauen mit access
            If gblnSQLServerVersion = False Then
                If gstrRowColChangeName = "" Then                                       'Gerbing 09.10.2014
                    gstrRowColChangeName = Adodc1.Recordset.Fields(LoadResString(1028 + Sprache))
                End If
                cmdline = "RowColChangeName=" & gstrRowColChangeName & ";"
                If gstrFotosMdbLocation <> "" Then
                    cmdline = cmdline & "fotosmdblocation=" & gstrFotosMdbLocation & ";"
                End If
                AppId = Shell(AppPath & "\RenamMdb.exe " & cmdline, vbNormalFocus)
                AppActivate AppId
            Else
            'CommandLine aufbauen mit sql server
                cmdline = "sqlservername=" & PublicSQLServer & ";"
                cmdline = cmdline & "datenbankname=" & PublicSQLDatabase & ";"
                cmdline = cmdline & "WindowsAuthentication=" & PublicWindowsAuthentication & ";"
                If PublicWindowsAuthentication = "0" Then
                    cmdline = cmdline & "username=" & PublicSQLServerUserName & ";"
                    cmdline = cmdline & "Password=" & PublicSQLServerPassword & ";"
                End If
                cmdline = cmdline & "StandortFotos=" & PublicLocationFotos & ";"
                cmdline = cmdline & "gstrRowColChangeName=" & gstrRowColChangeName & ";"
                AppId = Shell(AppPath & "\RenamMdb.exe" & " " & cmdline, vbNormalFocus)
                AppActivate AppId
            End If
'            'Öffne RenamMdb für die aktuelle Datei                                      'Gerbing 27.08.2012 08.10.2014
'            Unload DbGridForm
'            Unload Hilfebx
'            Unload KommentarForm
'            Unload Query
'            'Unload QueryJedesFeld
'            Unload MP
'            End
        Case 5                                      'Gerbing 28.08.2006
            '5 = Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
        '    SQL = " SELECT *"
        '    SQL = SQL & " FROM " & "Fotos"
        '    SQL = SQL & " WHERE Merker<>0
        '    SQL = SQL & " ORDER BY Dateiname" & ";"
            SQL = " SELECT *"
            SQL = SQL & " FROM Fotos"
            SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
            SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
            With rstsql
                .Source = SQL
                .ActiveConnection = DBsql
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If rstsql.EOF Then
                rstsql.Close
                MsgBox LoadResString(3061 + Sprache) 'Es gibt keine mit Merkerspalte markierten Sätze
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------
            SQL = Query.SQL
            pos = InStr(1, SQL, "WHERE", vbTextCompare)
            Pos1 = InStr(1, SQL, "ORDER BY", vbTextCompare)
            strLinks = Left(SQL, Pos1 - 1)
            strRechts = Right(SQL, Len(SQL) - Pos1 + 1)
            If pos <> 0 Then
                'WHERE gibt es schon
                'AND Merker = 1 hinzufügen
                strLinks = Replace(strLinks, "WHERE", "WHERE (", , , vbTextCompare)         'Gerbing 19.01.2007
                SQL = strLinks & ")"
                'SQL = SQL & " AND " & LoadResString(2524 + Sprache) & "=1 " & strRechts
                SQL = SQL & " AND " & LoadResString(2524 + Sprache) & "<>0 " & strRechts    'Gerbing 26.07.2012
            Else
                'WHERE gibt es noch nicht
                'WHERE Merker<>0 hinzufügen
                SQL = strLinks & " WHERE " & LoadResString(2524 + Sprache) & "<>0 " & strRechts
            End If
            On Error Resume Next
            DbGridForm.rsDataGrid.Close
            On Error GoTo 0
            If gblnSchreibgeschützt = True Then
            ' Recordset erstellen und öffnen
                Set DbGridForm.rsDataGrid = New adodb.Recordset
                With DbGridForm.rsDataGrid
                    .Source = SQL
                    .ActiveConnection = DBsql
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
            Else
                ' Recordset erstellen und öffnen
                Set DbGridForm.rsDataGrid = New adodb.Recordset
                With DbGridForm.rsDataGrid
                    .Source = SQL
                    .ActiveConnection = DBsql
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
            End If
            Set DbGridForm.Adodc1.Recordset = DbGridForm.rsDataGrid
            Set DbGridForm.DBGridNeu.DataSource = DbGridForm.rsDataGrid
            DbGridForm.DBGridNeu.ReBind
            Call SetSpaltenBreite
        Case 6
            '6 = Löschen markierte Dateien(Merkerspalte) in Datenbank und Standort          'Gerbing 23.01.2007
        '    SQL = " SELECT *"
        '    SQL = SQL & " FROM " & "Fotos"
        '    SQL = SQL & " WHERE Merker<>0
        '    SQL = SQL & " ORDER BY Dateiname" & ";"
            SQL = " SELECT *"
            SQL = SQL & " FROM Fotos"
            SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0"
            SQL = SQL & " ORDER BY " & LoadResString(1028 + Sprache) & ";"
            With rstsql
                .Source = SQL
                .ActiveConnection = DBsql
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If rstsql.EOF Then
                rstsql.Close
                MsgBox LoadResString(3061 + Sprache) 'Es gibt keine mit Merkerspalte markierten Sätze
                Exit Sub
            End If
            'msg = "Anzahl markierte Dateien = " & rst1.RecordCount                         'Gerbing 25.06.2008
            'msg = msg & "Wollen Sie wirklich alle mit der Merkerspalte markierten Dateien aus der Datenbank und an ihrem Standort löschen?" & vbnewline
            'msg = msg & "Sie gelangen anschließend in das Fenster zum Angeben der Suchkriterien."
            msg = LoadResString(2274 + Sprache) & rstsql.RecordCount & vbNewLine              'Gerbing 25.06.2008
            msg = msg & LoadResString(1523 + Sprache) & vbNewLine
            msg = msg & LoadResString(1524 + Sprache)
            'antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
            antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
            If antwort = vbNo Then
                Exit Sub
            End If
            Do Until rstsql.EOF
                Call LöschenInDatenbankUndStandort(rstsql.Fields(LoadResString(1028 + Sprache)), rstsql) '1028=Dateiname
                rstsql.MoveNext
            Loop
            rstsql.Close
            'so tun als wäre F8 gedrückt worden
            KeyCode = vbKeyF8
            Shift = 0
            'Tastatur-Eingabe weiterreichen
            Sleep (3000)
            Call Form1.Form_KeyDown(KeyCode, Shift)
        Case 7                                                              'Gerbing 04.02.2007
            '7 = Gehe zum Hyperlink
            If Right(DBGridNeu.Text, 1) = "#" And Left(DBGridNeu.Text, 1) = "#" Then
                If Mid(DBGridNeu.Text, 2, 3) = "+:\" Then
                    strTemp = Replace(DBGridNeu.Text, "+:\", gstrFotosMdbLocation & "\")                'Gerbing 07.11.2011
                    strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                Else
                    strTemp = Mid(DBGridNeu.Text, 2, Len(DBGridNeu.Text) - 2)
                End If
                RetVal = RunShellExecute(Me.hWnd, "Open", strTemp, "", gstrFotosMdbLocation, 1)
                If RetVal <= 32 Then
                    MsgBox LoadResString(3129 + Sprache) 'Es wurde kein geeigneter Browser gefunden um diese URL zu öffnen
                End If
            Else
'                msg = "Die gewählte Aktion erfordert, dass das aktive Feld einen Hyperlink enthält," & vbNewLine
'                msg = msg & "im Format #Hyperlink#"
'                msg = msg & "Inhalt des aktiven Feldes=" & DBGridNeu.Text
                msg = LoadResString(3126 + Sprache) & vbNewLine
                msg = msg & LoadResString(3127 + Sprache) & vbNewLine
                msg = msg & LoadResString(3128 + Sprache) & DBGridNeu.Text
                'MsgBox Msg
                MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
                Exit Sub
            End If
    End Select
End Sub

Function GetShellError(lErrorCode As Long) As String
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

Public Sub SetSpaltenBreite()
    Dim ColWidth As Long
    Dim ColCaption As String
    Dim n As Long
    Dim Werte() As Long
    Dim AnzahlStandardfelder As Long
    
    AnzahlStandardfelder = 13
    ReDim Werte(1 + AnzahlStandardfelder + ND.ListNutzerdefinierteFelder.ListItems.Count)
    SQL = "SELECT SpaltenBreite.* FROM SpaltenBreite;"
    'SQL = "SELECT " & LoadResString(2525 + Sprache) & ".* FROM " & LoadResString(2525 + Sprache) & ";" 'Gerbing 08.11.2005
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number = 3078 Then
'        msg = "Seit Version 10.0.0.0 gibt es in der Datenbank fotos.mdb eine Tabelle SpaltenBreite," & NL
        msg = LoadResString(2014 + Sprache) & NL
'        msg = msg & "wo Änderungen des Nutzers an den Spaltenbreiten im Fenster Datenbank-Übersicht (siehe F5)" & NL
        msg = msg & LoadResString(2015 + Sprache) & NL
'        msg = msg & "eingetragen werden." & NL & NL
        msg = msg & LoadResString(2016 + Sprache) & NL & NL
'        msg = msg & "Diese Tabelle wurde nicht gefunden."
        msg = msg & LoadResString(2017 + Sprache)
        MsgBox msg                                     'Gerbing 08.11.2005
        rstsql.Close
        Exit Sub
    End If
    If rstsql.EOF Then
'        msg = "Seit Version 10.0.0.0 gibt es in der Datenbank fotos.mdb eine Tabelle SpaltenBreite," & NL
        msg = LoadResString(2014 + Sprache) & NL
'        msg = msg & "wo Änderungen des Nutzers an den Spaltenbreiten im Fenster Datenbank-Übersicht (siehe F5)" & NL
        msg = msg & LoadResString(2015 + Sprache) & NL
'        msg = msg & "eingetragen werden." & NL & NL
        msg = msg & LoadResString(2016 + Sprache) & NL & NL
'        msg = msg & "Diese Tabelle wurde nicht gefunden."
        msg = msg & LoadResString(2017 + Sprache)
        MsgBox msg                                     'Gerbing 08.11.2005
        rstsql.Close
        Exit Sub
    End If
    On Error GoTo 0

    n = 0
    ColCaption = DBGridNeu.Columns(0).Caption
'        If ColCaption = LoadResString(2524 + Sprache) Then                  'merker
'            n = 1
'        End If
    Do Until rstsql.EOF
        If n = DBGridNeu.Columns.Count Then Exit Do
        DBGridNeu.Columns(n).Width = rstsql.Fields("Spaltenbreite")
        n = n + 1
        rstsql.MoveNext
    Loop
    rstsql.Close
End Sub

Public Sub SpeichernSpaltenBreite()
    Dim n As Long
    Dim ColWidth As Long
    
    SQL = "SELECT SpaltenBreite.* FROM SpaltenBreite;"
    'SQL = "SELECT " & LoadResString(2525 + Sprache) & ".* FROM " & LoadResString(2525 + Sprache) & ";" 'Gerbing 08.11.2005
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With

    'Bei jedem Speichern der Spaltenbreiten wird der bisherige Inhalt der Tabelle Spaltenbreite zuerst
    'gelöscht, dann werden neue Einträge gemacht
    On Error GoTo 0
    
    If gblnSQLServerVersion = True Then
        'beim SQL Server muss es heißen 'Delete from table
        SQL = "DELETE From Spaltenbreite"
        'SQL = "DELETE FROM " & LoadResString(2525 + Sprache)
    Else
        SQL = "DELETE * From Spaltenbreite"
        'SQL = "DELETE * FROM " & LoadResString(2525 + Sprache)          '2525=Spaltenbreite
    End If
    DBsql.Execute (SQL)
    For n = 0 To DBGridNeu.Columns.Count - 1                        'es wird ab Spalte Merker gespeichert
        rstsql.AddNew
        ColWidth = DBGridNeu.Columns(n).Width
        If DBGridNeu.Columns(n).Visible = False Then ColWidth = 0
        rstsql.Fields("Spaltenbreite") = ColWidth
        rstsql.Update
    Next n
    rstsql.Close
End Sub

Private Sub ThumbnailAnzeigen(EchterStandort)
    Dim Bildbreite As Long
    Dim Bildhöhe As Long
    Dim ImageThumbTop As Long
    Dim ImageThumbLeft As Long
    Dim BHV As Double   'BreitenHöhenVerhältnis
    Dim strTemp As String

    strTemp = Replace(EchterStandort, "+:\", gstrFotosMdbLocation & "\")        'Gerbing 11.04.2005
    Bildbreite = 1000                                                                       'Gerbing 05.12.2010
    Bildhöhe = 750                                                                          'Gerbing 05.12.2010
    ImageThumbTop = PictureThumb.Top
    ImageThumbLeft = PictureThumb.Left

    On Error GoTo 0
    Err = 0
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    ImageThumb.Visible = False
    PictureThumb.Visible = False
    PictureThumb.AutoSize = True
    'On Error Resume Next                                                                   'Gerbing 04.03.2013
    Call LoadPictureWDbGridForm(strTemp)                                                    'Gerbing 29.03.2015
    If Err.Number <> 0 Then
        Me.MousePointer = vbNormal                                                          'Gerbing 29.07.2007
        'Call BildFehler(strTemp)
        Exit Sub
    End If
    On Error GoTo 0
    '-----------------------------------------------------------------------------------------------
    'Untersuchung, ob das Bild größer ist als die Bildbreite/höhe und dessen Konsequenzen
    BHV = PictureThumb.Width / PictureThumb.Height

    If PictureThumb.Width > Bildbreite Or PictureThumb.Height > Bildhöhe Then
    'wenn das Bild größer ist als Bildbreite/Bildhöhe wird es verkleinert
        ImageThumb.Stretch = True
        ImageThumb.Picture = PictureThumb.Picture
        Select Case BHV
            Case 1.33 To 1.34
                'das Breitenverhältnis ist 4/3 = 1.33
                ImageThumb.Top = ImageThumbTop
                ImageThumb.Left = ImageThumbLeft
                ImageThumb.Width = Bildbreite
                ImageThumb.Height = Bildhöhe
            Case Is < 1.33
                'das Bild ist zu hoch und zu schmal
                ImageThumb.Top = ImageThumbTop
                ImageThumb.Left = ImageThumbLeft
                ImageThumb.Height = Bildhöhe
                ImageThumb.Width = Bildhöhe * BHV
            Case Else
                'das Bild ist zu niedrig und zu breit
                ImageThumb.Top = ImageThumbTop
                ImageThumb.Left = ImageThumbLeft
                ImageThumb.Width = Bildbreite
                ImageThumb.Height = Bildbreite / BHV
        End Select
    Else
    'wenn das Bild nicht größer ist als Bildbreite/Bildhöhe
        'Bild in links oben im Bildbereich anordnen
        ImageThumb.Stretch = True
        ImageThumb.Picture = PictureThumb.Picture
        ImageThumb.Top = ImageThumbTop
        ImageThumb.Left = ImageThumbLeft
        ImageThumb.Width = PictureThumb.Width
        ImageThumb.Height = PictureThumb.Height
    End If
    ImageThumb.Visible = True
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
End Sub

Private Sub LöschenInDatenbankUndStandort(Dateiname As String, rst1 As adodb.Recordset)
    Dim antwort As Long
    Dim strTemp As String
    Dim DateinameFoto As String
    Dim temp As String
    Dim temp1 As String
    Dim rc As Boolean
    
    'On Error Resume Next
    strTemp = Replace(Dateiname, "+:\", gstrFotosMdbLocation & "\")
    'Falls es einen zugehörigen Audio-Kommentar gibt, wird dieser zuerst gelöscht    'Gerbing 12.04.2006
    DateinameFoto = ErmittleDateiname(strTemp)
    If file_path_exist(DateinameFoto & ".mp3") = True Then
        temp = DateinameFoto & ".mp3"
    End If
    If file_path_exist(DateinameFoto & ".wav") = True Then
        temp = DateinameFoto & ".wav"
    End If
    If temp <> "" Then                                                              'Gerbing 04.09.2013
        rc = file_delete(temp, , True)                                              'Gerbing 04.09.2013
    End If
    rc = file_delete(strTemp, , True)                                       'Gerbing 04.09.2013
    If rc = False Then Exit Sub
    '----------------------------------------------------------------------------
    'Löschen aus der Datenbank
LöschenAusDerDatenbank:
    On Error Resume Next
    If gblnSchreibgeschützt = False Then
        rst1.Delete
        If Err.Number <> 0 Then                     'Gerbing 10.02.2007
            msg = "Error number=" & Err.Number & vbNewLine
            msg = msg & "Error text=" & Err.Description & vbNewLine
            If Err.Number = 3218 Then               'Datensatz ist momentan gesperrt
                'msg = msg & "Wiederholen Sie den Löschversuch zu einem späteren Zeitpunkt"
                msg = msg & LoadResString(2326 + Sprache)
            End If
            MsgBox msg
        End If
    End If
    On Error GoTo 0
End Sub

Private Function ErmittleDateiname(Dateiname As String) As String
    Dim pos As Long
    Dim start As Long
    Dim MeinDateiname As String

    'Der Dateiname wird ermittelt durch Suchen ab rechtem Rand bis zum Punkt
    start = Len(Dateiname) - 2
    Do
        pos = InStr(start, Dateiname, ".")
        If pos <> 0 Then
            MeinDateiname = Mid(Dateiname, 1, pos - 1)
            Exit Do
        End If
        start = start - 1
    Loop
    ErmittleDateiname = MeinDateiname
End Function
