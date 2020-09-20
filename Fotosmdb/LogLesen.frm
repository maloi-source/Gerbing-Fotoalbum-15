VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form LogLesen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Datei Pruef.log"
   ClientHeight    =   3876
   ClientLeft      =   -12
   ClientTop       =   276
   ClientWidth     =   7812
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogLesen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3876
   ScaleWidth      =   7812
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnPr�fen1�berfl�ssigeL�schen 
      Caption         =   "Pr�fen1 - Fehlerhafte oder nichtvorhandene Dateien l�schen..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7572
   End
   Begin VB.CommandButton btnGefundeneAufnehmen 
      Caption         =   "Pr�fen3 - Gefundene Dateien in die Datenbank &aufnehmen..."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   7572
   End
   Begin VB.CommandButton btnPr�fen2Dateienverschieben 
      Caption         =   "Pr�fen2 - Dateien in den richtigen Jahres-Ordner &verschieben"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Die Spalte 'Jahr' hat Priorit�t �ber die Spalte 'Dateiname'"
      Top             =   1560
      Width           =   7572
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "Ab&brechen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   7572
   End
   Begin VB.CommandButton btn�berfl�ssigeL�schen 
      Caption         =   "Pr�fen3 - Die gefundenen Dateien sind �berfl�ssig -> &l�schen..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Sie k�nnen ausw�hlen, welche der �berfl�ssigen Dateien gel�scht werden sollen"
      Top             =   2160
      Width           =   7572
   End
   Begin EditCtlsLibUCtl.TextBox TxtU 
      Height          =   732
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7572
      _cx             =   13356
      _cy             =   1291
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   3
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -1
      DisabledEvents  =   3075
      DisabledForeColor=   -1
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   -1  'True
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   3
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "LogLesen.frx":038A
      Text            =   "LogLesen.frx":03B2
   End
End
Attribute VB_Name = "LogLesen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'General Function Declarations
    Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
    Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
    Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
    Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function SendMessageStringA Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function SendMessageStringW Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
     
    'General Type Declarations
    Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
    End Type
     
    'General Variable Declarations
    Private Const OBJ_FONT As Long = 6
     
    Private Const TM_PLAINTEXT As Long = 1
     
    Private Const WM_USER As Long = &H400
    Private Const EM_SETTEXTMODE As Long = (WM_USER + 89)
     
    Private Const ES_MULTILINE As Long = &H4
     
    Private Const WM_SETFONT As Long = &H30
    Private Const WM_SETTEXT As Long = &HC
     
    Private Const WS_CHILD As Long = &H40000000
    Private Const WS_EX_CLIENTEDGE As Long = &H200&
    Private Const WS_VISIBLE As Long = &H10000000
    Private Const WS_VSCROLL As Long = &H200000
    Private Const WS_HSCROLL As Long = &H100000
     
    Private hFont As Long
    Private hRich As Long
    Private hWndRich As Long
     
    Private TempStr As String
     
    Private WinVer As OSVERSIONINFO

Private Sub btnAbbrechen_Click()
    Call Form1.btnReset_Click                                       'Gerbing 03.11.2013
    Unload Me
End Sub

Private Sub btnGefundeneAufnehmen_Click()
    Dim i As Long
    
    NachPr�fen3Aufnehmen.Image1 = LoadPicture("")
    If NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count <> 0 Then                                   'Gerbing 26.10.2013
        For i = 1 To NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count
            NachPr�fen3Aufnehmen.lstZus�tzlicheDateien.ListItems.Add NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Item(i)
            TxtU.Text = NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Item(i)
            DoEvents
        Next i
    End If
    NachPr�fen3Aufnehmen.Show 1
    'Unload Me                                                       'Gerbing 15.08.2005
End Sub

Private Sub btnPr�fen1�berfl�ssigeL�schen_Click()
    Dim Msg As String
    
    If gblnSchreibgesch�tzt = True Then                             'Gerbing 23.01.2007
        'msg = "Bei einer schreibgesch�tzten Datenbank ist diese Funktion nicht m�glich"
        Msg = LoadResString(2421 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
    NachPr�fen1L�schen.Show 1                                       'Gerbing 16.01.2006
    Unload Me
End Sub

Private Sub btnPr�fen2Dateienverschieben_Click()
    Dim SQL As String
    Dim Fotodatei As String
    Dim JahresZahl As String
    Dim VordemJahr As String
    Dim NachDemJahr As String
    Dim start As Long
    Dim Pos As Long
    Dim JahrInFilename As String
    Dim temp As String
    Dim Msg As String
    Dim SoundDatei As String
    Dim rc As Boolean
        
    'Nach Pr�fen2 l��t sich die Ungleichheit zwischen Jahreszahlen im Feld 'Jahr' und im Feld
    'Dateiname' korrigieren. Dabei hat das Feld 'Jahr' Priorit�t. Bei Ungleichheit wird die Datei
    'aus dem falschen Ordner in den Ordner mit der richtigen Jahreszahl verschoben. Im Feld
    'Dateiname' mu� eine Korrektur vorgenommen werden.

    'Schon vor der Korrektur wird das Formular geschlossen und nach der Korrektur wird der Nutzer aufgefordert
    'mit Pr�fen2 erneut zu kontrollieren, ob jetzt �bereinstimmung zwischen Feld 'Jahr'
    'und Feld 'Dateiname' herrscht
    
    Me.Hide
    If gblnSQLServerVersion = True Then
        'CharIndex hat andere Parameterreihenfolge als InStr
        'SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(jahr,Dateiname,1)=0;"
        SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(" & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & ")=0;" 'Gerbing 08.11.2005
    Else
        'SQL = "SELECT Fotos.* From Fotos WHERE instr(1,Dateiname, jahr)=0;"         'Gerbing 17.09.2004
        SQL = "SELECT Fotos.* From Fotos WHERE instr(1," & LoadResString(1028 + Sprache) & ", " & LoadResString(1023 + Sprache) & ")=0;"     'Gerbing 17.09.2004
    End If
    Set Form1.rstsql = New ADODB.Recordset
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With

    If Not Form1.rstsql.EOF Then                                                    'Gerbing 10.05.2013
        Form1.rstsql.MoveFirst
        Screen.MousePointer = vbHourglass
        Do Until Form1.rstsql.EOF
            'Fotodatei = Form1.rstsql("Dateiname")
            Fotodatei = Form1.rstsql(LoadResString(1028 + Sprache))
            Fotodatei = Replace(Fotodatei, "+:\", AppPath & "\")                   'Gerbing 11.04.2005
            'JahresZahl = Form1.rstsql("Jahr")
            JahresZahl = Form1.rstsql(LoadResString(1023 + Sprache))
            start = 1
            Do
                Pos = InStr(start, Fotodatei, "\") 'hintersten \ suchen
                If Pos = 0 Then Exit Do
                start = Pos + 1
            Loop
            JahrInFilename = Mid(Fotodatei, start - 5, 4)                             'Gerbing 29.12.2011
            NachDemJahr = Mid(Fotodatei, start, Len(Fotodatei) - start + 1)
            If JahresZahl <> JahrInFilename Then
                Form1.gstrNeuerName = AppPath & "\" & JahresZahl & "\" & NachDemJahr   'Gerbing 20.07.2005
                'Die �nderung wird nur gemacht, wenn kein neues Verzeichnis angelegt werden mu�
                'd.h. wenn das Verzeichnis existiert
                'temp = Dir(AppPath & "\" & JahresZahl, vbDirectory)                    'Gerbing 20.07.2005
                If file_path_exist(AppPath & "\" & JahresZahl) = True Then
                'If temp <> "" Then
                    'Wenn temp = "" dann ist es ein nichtexistierender Zielordner
                    'gstrNeuerName mu� in eine Rename-Operation einflie�en
                    '26.04.2004 Es k�nnte aber sein, da� es im Zielordner bereits eine Datei mit dem gleichen
                    'Namen gibt
                    'Name Fotodatei As Form1.gstrNeuerName    'rename altername As neuername
                    If file_path_exist(Form1.gstrNeuerName) = True Then
                    'If rc = True Then
                        'Fehler beim Umnennen kann an dieser Stelle nur Duplikatfehler sein
                        'Formular anbieten wo der doppelt vorkommende Name gezeigt wird und darunter eine
                        'Zeile zum Ausw�hlen eines anderen Namens
                        'solange wie der andere Name auch wieder ein Duplikat ist, geht es nicht weiter
                        'oder der Nutzer w�hlt 'Abbrechen'          'Gerbing 20.07.2005
                        Screen.MousePointer = vbDefault
                        DuplikatName.Show 1
                        Screen.MousePointer = vbHourglass
                        If gblnAbbrechen = True Then GoTo Movenext  'Gerbing 20.07.2005
                        'Name Fotodatei As Form1.gstrNeuerName    'rename altername As neuername
                        rc = NameAs(Fotodatei, Form1.gstrNeuerName)     'rename altername, neuername      'Gerbing 04.03.2013
                    Else
                        rc = NameAs(Fotodatei, Form1.gstrNeuerName)     'rename altername, neuername      'Gerbing 04.03.2013
                    End If
                    '-------------------------------------------------------------------------------------------------------------------
                    'eine weitere Rename-Operation ist n�tig, wenn es eine gleichnamige Sounddatei WAV oder MP3 gibt 'Gerbing 11.11.2010
                    SoundDatei = left(Fotodatei, Len(Fotodatei) - 3) & "WAV"
                    'Msg = Dir(SoundDatei)
                    If file_path_exist(SoundDatei) = True Then
                    'If Msg <> "" Then
                        temp = left(Form1.gstrNeuerName, Len(Form1.gstrNeuerName) - 3) & "WAV"
                        'Name SoundDatei As temp    'rename altername, neuername
                        rc = NameAs(SoundDatei, temp)     'rename altername As neuername      'Gerbing 04.03.2013
                    End If
                    SoundDatei = left(Fotodatei, Len(Fotodatei) - 3) & "MP3"
                    'Msg = Dir(SoundDatei)
                    If file_path_exist(SoundDatei) = True Then
                    'If Msg <> "" Then
                        temp = left(Form1.gstrNeuerName, Len(Form1.gstrNeuerName) - 3) & "MP3"
                        'Name SoundDatei As temp    'rename altername As neuername
                        rc = NameAs(SoundDatei, temp)     'rename altername, neuername      'Gerbing 04.03.2013
                    End If
                    '---------------------------------------------------------------------------
                    'NeuerName mu� in den Recordset eingetragen werden
                    Form1.DBGridNeu.AllowUpdate = True
                    'Form1.rstsql.Edit
                    temp = Replace(Form1.gstrNeuerName, AppPath, "+:")             'Gerbing 11.04.2005
                    'Form1.rstsql.Fields("Dateiname") = temp                'Gerbing 11.04.2005
                    Form1.rstsql.Fields(LoadResString(1028 + Sprache)) = temp                'Gerbing 11.04.2005
                    Form1.rstsql.Update
                    Form1.DBGridNeu.AllowUpdate = False
                End If
            End If
Movenext:
            Form1.rstsql.Movenext
        Loop
        Screen.MousePointer = vbDefault
    '    Msg = "Wiederholen Sie jetzt die Funktion Pr�fen2, zur Kontrolle, ob alle Korrekturen gemacht worden sind." & vbNewLine
    '    Msg = Msg & "Wenn es den Ordner mit der geforderten Jahreszahl nicht gibt, werden die Korrekturen nicht gemacht."
        Msg = LoadResString(1415 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(1416 + Sprache)
        MsgBox Msg
    End If
    Unload Me
    Call Form1.btnReset_Click                                                        'Gerbing 11.04.2005
End Sub

Private Sub btn�berfl�ssigeL�schen_Click()
    Dim i As Long
    
    If NachPr�fen3L�schen.KollZus�tzlicheDateien.Count <> 0 Then                                   'Gerbing 26.10.2013
        For i = 1 To NachPr�fen3L�schen.KollZus�tzlicheDateien.Count
            NachPr�fen3L�schen.lstZus�tzlicheDateien.ListItems.Add NachPr�fen3L�schen.KollZus�tzlicheDateien.Item(i)
            TxtU.Text = NachPr�fen3L�schen.KollZus�tzlicheDateien.Item(i)
            DoEvents
        Next i
    End If
    NachPr�fen3L�schen.Show 1
End Sub

Private Sub Form_Load()
    Dim LogFileName As String
    Dim Msg As String
    Dim myStream As TextStream
    Dim sLine As String
    
    Call AnpassenNutzerWunsch(Me)                                               'Gerbing 11.03.2017
    Me.Caption = LoadResString(1337 + Sprache)   'Datei Pruef.log auswerten
    btnPr�fen2Dateienverschieben.Caption = LoadResString(1338 + Sprache)            'Pr�fen2 - Dateien in den richtigen Jahres-Ordner &verschieben
    btn�berfl�ssigeL�schen.Caption = LoadResString(1339 + Sprache)      'Pr�fen3 - Die gefundenen Dateien sind �berfl�ssig -> &l�schen...
    btnAbbrechen.Caption = LoadResString(1325 + Sprache)                'Abbru&ch
    btnPr�fen2Dateienverschieben.ToolTipText = LoadResString(1430 + Sprache)    'Die Spalte 'Jahr' hat Priorit�t �ber die Spalte 'Dateiname'
    btn�berfl�ssigeL�schen.ToolTipText = LoadResString(1431 + Sprache)          'Sie k�nnen ausw�hlen, welche der �berfl�ssigen Dateien gel�scht werden sollen
    btnPr�fen1�berfl�ssigeL�schen.Caption = LoadResString(1460 + Sprache)       'Pr�fen1 - L�sche Datens�tze mit nichtvorhandenen Dateien...
    btnPr�fen1�berfl�ssigeL�schen.ToolTipText = LoadResString(1461 + Sprache)   'Hiermit k�nnen Sie Datens�tze aus der Datenbank entfernen, die zu nicht/nicht mehr existierenden Dateien geh�ren und nach Pr�fen1 gefunden wurden
    btnGefundeneAufnehmen.Caption = LoadResString(1340 + Sprache)       'Pr�fen3 - Gefundene Dateien in die Datenbank &aufnehmen...
    
    Me.top = 0
    Me.left = 0
    'If Form1.Pr�fenNummer <> "Pr�fen3" Then
    If Form1.Pr�fenNummer <> LoadResString(1459 + Sprache) Then
        btn�berfl�ssigeL�schen.Visible = False
        btnGefundeneAufnehmen.Visible = False
    End If
    'If Form1.Pr�fenNummer <> "Pr�fen2" Then
    If Form1.Pr�fenNummer <> LoadResString(1444 + Sprache) Then
        btnPr�fen2Dateienverschieben.Visible = False
    End If
    'If Form1.Pr�fenNummer <> "Pr�fen1" Then                            'Gerbing 16.01.2006
    If Form1.Pr�fenNummer <> LoadResString(1443 + Sprache) Then
        btnPr�fen1�berfl�ssigeL�schen.Visible = False
    End If
    LogFileName = PruefLogFile
    On Error GoTo Fehler
    TxtU.Text = ""
    Screen.MousePointer = vbHourglass
    If (myStream Is Nothing) Then
        ' Open the file for reading.
            Set myStream = PruefFso.OpenTextFile(LogFileName, 1, False, -1)     'unicode
        If (Not myStream Is Nothing) Then
            With myStream
                Do Until myStream.AtEndOfStream
                    sLine = myStream.ReadLine
                    'TxtU.Text = TxtU.Text & sLine & vbNewLine
                Form1.txtArbeitsfortschrittU.Text = sLine
                DoEvents                                                'Gerbing 25.10.2013
                Loop
                .Close
            End With
            Set myStream = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Fehler:
    'Msg = LogFileName & " kann nicht ge�ffnet werden"
    Msg = LogFileName & " " & LoadResString(1372 + Sprache)
    MsgBox Msg
    Unload Me
    Exit Sub
End Sub

