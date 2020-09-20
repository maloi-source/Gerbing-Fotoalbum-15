VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form ImportForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Importieren mit Drag&Drop"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8868
   Icon            =   "ImportForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   8868
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtDragDropDatenbank 
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   8652
      _cx             =   15261
      _cy             =   656
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
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
      HAlignment      =   1
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "ImportForm.frx":038A
      Text            =   "ImportForm.frx":03AA
   End
   Begin VB.ListBox lstFensterTitel 
      Height          =   240
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drag&&Drop Name der Exportdatenbank"
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6852
   End
   Begin VB.Label lblArbeitsfortschritt 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   8532
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2892
   End
End
Attribute VB_Name = "ImportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Kontrolle As String
    Dim AbbrechenGedrückt As Boolean
    Dim OKGedrückt As Boolean
    Dim Stil As Long
    Dim antwort As Long
    Dim msg As String
    Dim Prüfdir As String
    Dim MsgBoxMussKommen As Boolean
    Dim DatensatzNr As Long
    Dim blnAbbruch As Boolean
    
    Private Const CF_OEMTEXT = 7
    Private Const CF_UNICODETEXT = 13


Private Sub btnAbbrechen_Click()
    blnAbbruch = True
    Unload Me
End Sub

Private Sub btnHilfe_Click()
    Dim RetVal As Long
    Dim CHMFile As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = AppPath & "\Help\Deutsch\fotos.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            msg = CHMFile & vbNewLine
            msg = msg & LoadResString(2544 + Sprache) & vbNewLine
            msg = msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    Else
        CHMFile = AppPath & "\Help\English\fotos.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            msg = CHMFile & vbNewLine
            msg = msg & LoadResString(2544 + Sprache) & vbNewLine
            msg = msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub UnterverzeichnisErzeugen(Parm, Quellname, Zielname)
    'Parm = 0 bei Quell-Ordner und Ziel-Ordner benutzen
    'Parm = 1 bei sofort aus anderer Anwendung importieren
    
    Dim QuellAnteil As String
    Dim ZielAnteil As String
    Dim pos, start As Long
    Dim Verz As String
    Dim rc As Boolean
    
    If Parm = 0 Then
'        QuellAnteil = Right(Quellname, Len(Quellname) - Len(DirQuell.Path) - 1)    'Quell-Ordner abschneiden
'        Start = 1
'        ZielAnteil = DirZiel.Path
    Else
        'Left(txtDragDropDatenbank.text, Len(txtDragDropDatenbank.text) - 11)
        start = 1
        QuellAnteil = Right(Zielname, Len(Zielname) - Len(gstrFotosMdbLocation) - 1)
        ZielAnteil = gstrFotosMdbLocation
    End If
    Do
        pos = InStr(start, QuellAnteil, "\", vbTextCompare)
        If pos = 0 Then Exit Do
        Verz = Left(QuellAnteil, pos - 1)
        QuellAnteil = Right(QuellAnteil, Len(QuellAnteil) - pos)
        ZielAnteil = ZielAnteil & "\" & Verz
        On Error Resume Next
        Err.Number = 0
        MkDir ZielAnteil    'Kein Fehler wenn das Verzeichnis in einem
                            'vorhergehenden Aufruf schon angelegt wurde
    Loop
    On Error Resume Next
    'On Error GoTo 0
    Err.Number = 0
    'Jetzt das mißlungene FileCopy wiederholen
    'Bereits vorhanden Dateien werden nicht überschrieben
'    Kontrolle = Dir(Zielname)
'    If Kontrolle <> "" Then
    If file_path_exist(Zielname) = True Then
            'msg = "Sie wollen eine Datei importieren, deren Namen es schon gibt:" & NL
            msg = LoadResString(2095 + Sprache) & NL
            msg = msg & Zielname & NL
            'msg = msg & "Das Importieren dieser Datei wird nicht ausgeführt." & NL
            msg = msg & LoadResString(2096 + Sprache) & NL
            'msg = msg & "Sie sollten vorher absichern, dass keine doppelten Dateinamen entstehen." & NL & NL
            msg = msg & LoadResString(2097 + Sprache) & NL & NL
        Stil = vbOKCancel + vbDefaultButton1  ' Schaltflächen
        'antwort = MsgBox(Msg, Stil)
        antwort = MessageBoxW(0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), Stil)
        AbbrechenGedrückt = False
        If antwort = vbCancel Then     ' Benutzer hat "Abbrechen" gedrückt
            AbbrechenGedrückt = True
        End If
        Exit Sub
    End If
    Err.Number = 0
'    FileCopy Quellname, Zielname
'    If Err.Number <> 0 Then
    rc = file_copy(Quellname, Zielname)
    If rc = False Then
        If Err.Number = 76 Then
            'msg = "Kein Schreibzugriff bei FileCopy" & NL                          'Gerbing 26.09.2007
            msg = LoadResString(2327 + Sprache) & NL
            msg = msg & Quellname & "," & Zielname & NL
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Else
            'msg = "Programmierfehler bei FileCopy" & NL
            msg = LoadResString(2034 + Sprache) & NL
            msg = msg & Quellname & "," & Zielname & NL
            msg = msg & "Errorcode=" & Err.Number & NL                              'Gerbing 26.09.2007
            msg = msg & "Errortext=" & Err.Description
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        End If
        End
    End If
    Call AudioDateiMitkopieren(Quellname, Zielname)                                 'Gerbing 29.07.2007
End Sub


Private Sub btnStart_Click()
    'wenn jetzt noch keine zweite Anwendung (d.h. keine zweite Datenbank fotos.mdb) geöffnet ist, muss
    'das Einschalten dieser Option verhindert werden
    Dim CurrWnd As Long
    Dim TitelLength As Long
    Dim FensterTitel As String
    Dim X As Long
    Dim msg As String
    Const GW_HWNDFIRST = 0
    Const GW_HWNDNEXT = 2

    If gblnSQLServerVersion = False Then
        If gblnSchreibgeschützt = True Then                     'Gerbing 26.01.2006
            'schreibgeschützt
            msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
            'Msg= msg & "Die Datenbank ist schreibgeschützt, Import mit Drag&Drop ist nicht möglich"
            msg = msg & LoadResString(2223 + Sprache)
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
            Exit Sub
        End If
    End If
    
    If Query.chkFensterGrößeÄnderbar.Value = 0 Then
        MsgBox LoadResString(3066 + Sprache) 'Wenn Sie 'sofort in andere geöffnete Anwendung exportieren' wollen, darf die Anwendung nicht den gesamten Bildschirm ausfüllen. Beide Anwendungen müssen mit 'Fenstergröße änderbar' gestartet werden, weil für das Auslösen des Kopierens(Exportierens) Drag&Drop zwischen Export-Fenster und Import-Fenster benutzt wird.
        Exit Sub
    End If

    lstFensterTitel.Clear
    CurrWnd = GetWindow(Form1.hWnd, GW_HWNDFIRST)
    While CurrWnd <> 0
    'In dieser Schleife werden gefüllt
    'lstFensterTitel enthält Application Titles
        TitelLength = GetWindowTextLength(CurrWnd)
        FensterTitel = Space$(TitelLength + 1)
        TitelLength = GetWindowText(CurrWnd, FensterTitel, TitelLength + 1)
        If TitelLength > 0 Then
            If Left(FensterTitel, Len(LoadResString(1001 + Sprache))) = LoadResString(1001 + Sprache) Then  '1001=FotoAlbum-
                lstFensterTitel.AddItem FensterTitel
            End If
        End If
        CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
        X = DoEvents()
    Wend
    If lstFensterTitel.ListCount < 3 Then                                                               'Gerbing 28.03.2014
        msg = LoadResString(3056 + Sprache) & vbNewLine 'es muß eine zweite Anwendung (Datenbank fotos.mdb) geöffnet sein. Dort wird nach erfolgreichem 'Export vorbereiten' Drag&Drop gestartet.
        msg = msg & LoadResString(3059 + Sprache)   'Beide Anwendungen müssen mit 'Fenstergröße änderbar' gestartet werden, weil für das Auslösen des Kopierens(Exportierens) Drag&Drop zwischen Export-Fenster und Import-Fenster benutzt wird.
        MsgBox msg
        Exit Sub
    End If
    Label7.Visible = True
    txtDragDropDatenbank.Visible = True

End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
        Me.Top = Form1.Top                                          'Gerbing 06.12.2006
        Me.Left = Form1.Left
    End If

    Me.Caption = LoadResString(1076 + Sprache)        'Importieren mit Drag&Drop" Gerbing 31.01.2006
    Label3.Caption = LoadResString(1079 + Sprache)    'Arbeitsfortschritt:
    btnHilfe.Caption = LoadResString(3018 + Sprache) '&Hilfe
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)   '&Abbrechen
    Label7.Caption = LoadResString(3060 + Sprache) 'Drag&&Drop Name der Exportdatenbank
    txtDragDropDatenbank.tooltipText = LoadResString(3069 + Sprache) 'Ziehen Sie mit Drag&Drop den Namen der Exportdatenbank aus der anderen Anwendung in dieses Feld
    
    On Error Resume Next
    Label7.Visible = False
    txtDragDropDatenbank.Visible = False
End Sub


Private Sub KontrolleUndFileCopy(Parm As Long, Quellname As String, Zielname As String)
    Dim Erg As Long
    Dim rc As Boolean

    On Error Resume Next
    'Bereits vorhanden Dateien werden nicht überschrieben
'    Kontrolle = Dir(Zielname)
'    If Kontrolle <> "" Then
    If file_path_exist(Zielname) = True Then
        Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
        'msg = "Sie wollen eine Datei importieren, deren Namen es schon gibt:" & NL
        msg = LoadResString(2095 + Sprache) & NL
        msg = msg & Zielname & NL
        'msg = msg & "Das Importieren dieser Datei wird nicht ausgeführt." & NL
        msg = msg & LoadResString(2096 + Sprache) & NL
        'msg = msg & "Sie sollten vorher absichern, dass keine doppelten Dateinamen entstehen." & NL & NL
        msg = msg & LoadResString(2097 + Sprache) & NL & NL
        
        'msg = msg & "Wenn Sie OK wählen, erhalten Sie diesen Hinweis für jede betroffene Datei." & NL
        msg = msg & LoadResString(2098 + Sprache) & NL
        'msg = msg & "wenn Sie Abbrechen wählen, wird der Import-Vorgang abgebrochen."
        msg = msg & LoadResString(2099 + Sprache)
        Stil = vbOKCancel + vbDefaultButton1  ' Schaltflächen
        'antwort = MsgBox(Msg, Stil)
        antwort = MessageBoxW(0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), Stil)
        AbbrechenGedrückt = False
        If antwort = vbCancel Then     ' Benutzer hat "Abbrechen" gedrückt
            AbbrechenGedrückt = True
        Else
            OKGedrückt = True
        End If
        Me.MousePointer = vbHourglass                                                       'Gerbing 29.07.2007
    Else
'        FileCopy Quellname, Zielname
'        If Err <> 0 Then
        rc = file_copy(Quellname, Zielname)
        If rc = False Then
            'Wenn das Kopieren mit Fehler reagiert, muß ich ein neues Unterverzeichnis anlegen
            Call UnterverzeichnisErzeugen(Parm, Quellname, Zielname)
        Else
            Call AudioDateiMitkopieren(Quellname, Zielname)                                 'Gerbing 29.07.2007                                                      'Gerbing 26.09.2007
        End If
    End If
    On Error GoTo 0
    DatensatzNr = DatensatzNr + 1
    'lblArbeitsfortschritt.Caption = "DatensatzNr." & DatensatzNr
    lblArbeitsfortschritt.Caption = LoadResString(1008 + Sprache) & DatensatzNr
    Erg = DatensatzNr Mod 20
    If Erg = 0 Then
        'DoEvents
        lblArbeitsfortschritt.Refresh                                                       'Gerbing 29.06.2011
    End If
End Sub

Private Sub txtDragDropDatenbank_OLEDragDrop(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Quellname As String
    Dim Zielname As String
    Dim SQL As String
    
    Dim dbs1 As ADODB.Connection                                                            'Gerbing 23.11.2017
    Dim dbs2 As ADODB.Connection                                                            'Gerbing 23.11.2017
    Dim rst1 As ADODB.Recordset                                                            'Gerbing 23.11.2017
    Dim rst2 As ADODB.Recordset
    Dim AddErr As Long
    Dim UpdErr As Long
    Dim Erg As Long
    Dim OhneVorspann As String

    blnAbbruch = False
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    If Data.GetFormat(vbCFText) = True Then
        txtDragDropDatenbank.Text = Data.GetData(vbCFText)
    End If
    If Data.GetFormat(CF_OEMTEXT) = True Then
        txtDragDropDatenbank.Text = Data.GetData(CF_OEMTEXT)
    End If
    If Data.GetFormat(CF_UNICODETEXT) = True Then
        txtDragDropDatenbank.Text = Data.GetData(CF_UNICODETEXT)
    End If
        
    'das ist die fremde Export-Datenbank, dort steht, was in die eigene Import-Datenbank
    'und in den eigenen gstrFotosmdblocation kopiert werden soll
    'für jeden Satz der fremden Export-Datenbank
    '1. die Datei kopieren
    '2. den Datensatz kopieren
    Set dbs1 = CreateObject("ADODB.Connection")                                         'Gerbing 23.11.2017
    dbs1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & txtDragDropDatenbank.Text
    dbs1.mode = adModeReadWrite          'adModeRead=1=Read-only.    'adModeReadWrite=3=Read/write.
    dbs1.Open dbs1.ConnectionString
    SQL = " SELECT Fotos.* FROM Fotos"
    Set rst1 = New ADODB.Recordset
    On Error Resume Next
    rst1.Close
    On Error GoTo 0
    With rst1
        .Source = SQL
        .ActiveConnection = dbs1                                               'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number <> 0 Then
        'msg = "Fehler beim Zugrif auf Datei " & txtDragDropDatenbank.text & NL
        msg = LoadResString(2029 + Sprache) & txtDragDropDatenbank.Text & NL
        'msg = msg & "Fehler beim Zugriff auf Tabelle Fotos" & NL
        msg = msg & LoadResString(2030 + Sprache) & NL
        msg = msg & "Errortext=" & Err.Description & NL
        msg = msg & "Errornumber=" & Err.Number & NL
        If Err.Number = 3078 Then
            msg = msg & NL
            msg = msg & LoadResString(2226 + Sprache)  '"Sie müssen gewährleisten, dass die Export-Datenbank und die Import-Datenbank mit derselben Sprache arbeiten"
        End If
        'MsgBox Msg, vbCritical
        MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbCritical
        txtDragDropDatenbank.Text = ""
        Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
        Exit Sub
    End If
    On Error GoTo 0
    DatensatzNr = 1
    
    'Recordset rst2 enthält jeweils einen neu erzeugten Satz in gstrFotosmdblocation & \fotos.mdb, Tabelle Fotos
    Set dbs2 = CreateObject("ADODB.Connection")                                         'Gerbing 23.11.2017
    dbs2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\Fotos.mdb"
    dbs2.mode = adModeReadWrite          'adModeRead=1=Read-only.    'adModeReadWrite=3=Read/write.
    dbs2.Open dbs2.ConnectionString
    SQL = " SELECT Fotos.* FROM Fotos"
    Set rst2 = New ADODB.Recordset
    On Error Resume Next
    rst2.Close
    On Error GoTo 0
    With rst2
        .Source = SQL
        .ActiveConnection = dbs2                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do While Not rst1.EOF
        '------------------------------------------------------
        '1. die Datei kopieren
        On Error Resume Next                                                            'Gerbing 04.09.2012
        OhneVorspann = rst1.Fields(LoadResString(1028 + Sprache))
        If Err.Number <> 0 Then
            'msg = "Fehler beim Zugrif auf Datei " & txtDragDropDatenbank.text & NL
            msg = LoadResString(2029 + Sprache) & txtDragDropDatenbank.Text & NL
            'msg = msg & "Fehler beim Zugriff auf Tabelle Fotos" & NL
            msg = msg & LoadResString(2030 + Sprache) & NL
            msg = msg & "Errortext=" & Err.Description & NL
            msg = msg & "Errornumber=" & Err.Number & NL
            If Err.Number = 3265 Then
                msg = msg & NL
                msg = msg & LoadResString(2226 + Sprache)  '"Sie müssen gewährleisten, dass die Export-Datenbank und die Import-Datenbank mit derselben Sprache arbeiten"
            End If
            'MsgBox Msg, vbCritical
            MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbCritical
            txtDragDropDatenbank.Text = ""
            Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
            Exit Sub
        End If
        On Error GoTo 0                                                                     'Gerbing 04.09.2012
        OhneVorspann = Right(OhneVorspann, Len(OhneVorspann) - 2)
        Quellname = Left(txtDragDropDatenbank.Text, Len(txtDragDropDatenbank.Text) - 11) & OhneVorspann
        Zielname = gstrFotosMdbLocation & OhneVorspann
        OKGedrückt = False
        Call KontrolleUndFileCopy(1, Quellname, Zielname)
        If AbbrechenGedrückt = True Then Exit Do
        If OKGedrückt = True Then GoTo MoveNext
        '------------------------------------------------------
        '2. den Datensatz kopieren
        Err = 0
        rst2.AddNew                      ' Neuen Datensatz erstellen.
        AddErr = Err
        Err = 0
        On Error Resume Next
        rst2(LoadResString(1023 + Sprache)) = rst1.Fields(LoadResString(1023 + Sprache))   'jahr
        rst2(LoadResString(1024 + Sprache)) = rst1.Fields(LoadResString(1024 + Sprache))   'situation
        rst2(LoadResString(1025 + Sprache)) = rst1.Fields(LoadResString(1025 + Sprache))   'ort
        rst2(LoadResString(1026 + Sprache)) = rst1.Fields(LoadResString(1026 + Sprache))   'land
        rst2(LoadResString(1027 + Sprache)) = rst1.Fields(LoadResString(1027 + Sprache))   'personen
        rst2(LoadResString(1028 + Sprache)) = rst1.Fields(LoadResString(1028 + Sprache))   'dateiname
        rst2(LoadResString(1029 + Sprache)) = rst1.Fields(LoadResString(1029 + Sprache))   'SWF
        rst2(LoadResString(1030 + Sprache)) = rst1.Fields(LoadResString(1030 + Sprache))   'kommentar
        rst2(LoadResString(1031 + Sprache)) = rst1.Fields(LoadResString(1031 + Sprache))   'dateinamekurz
        rst2(LoadResString(1032 + Sprache)) = rst1.Fields(LoadResString(1032 + Sprache))   'ddatum
        rst2(LoadResString(1106 + Sprache)) = rst1.Fields(LoadResString(1106 + Sprache))   'breitepixel
        rst2(LoadResString(1107 + Sprache)) = rst1.Fields(LoadResString(1107 + Sprache))   'hoehepixel
        'Nutzerdefinierte Felder kann ich nicht übernehmen
        'weil ich nicht weiss ob die Inhalte selbst bei gleichlautenden Feldnamen zusammengehören
        
        rst2.Update                      ' Änderungen speichern.
        UpdErr = Err
        If Err <> 0 Then
            'msg = "Fehler beim Zugrif auf Datei " & gstrFotosmdblocation & "\Fotos.mdb" & NL
            msg = LoadResString(2029 + Sprache) & gstrFotosMdbLocation & "\Fotos.mdb" & NL
            'msg = msg & "Fehler beim Zugriff auf Tabelle Fotos" & NL
            msg = msg & LoadResString(2030 + Sprache) & NL
            'msg = msg & "Datensatz Nummer: " & rst2.RecordCount & NL
            msg = msg & LoadResString(2027 + Sprache) & rst2.RecordCount & NL
            msg = msg & "AddErr: " & AddErr & NL
            msg = msg & "UpdErr: " & UpdErr & NL
            msg = msg & "Errortext=" & Err.Description & NL
            msg = msg & "Errornumber=" & Err.Number & NL
            'MsgBox Msg, vbCritical
            MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbCritical
            txtDragDropDatenbank.Text = ""
            On Error GoTo 0
            Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
            Exit Sub
        End If
        On Error GoTo 0
        '----------------------------------------------------
MoveNext:
        rst1.MoveNext
        'DatensatzNr = DatensatzNr + 1
        'lblArbeitsfortschritt.Caption = "DatensatzNr." & DatensatzNr
        lblArbeitsfortschritt.Caption = LoadResString(1008 + Sprache) & DatensatzNr     'Gerbing 08.11.2005
        Erg = DatensatzNr Mod 20
        If Erg = 0 Then
            'DoEvents
            lblArbeitsfortschritt.Refresh                                                       'Gerbing 29.06.2011
        End If
        If blnAbbruch = True Then
            Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
            Exit Do
        End If
    Loop
    rst1.Close
    rst2.Close
    msg = LoadResString(2089 + Sprache) & vbNewLine         'Importier-Vorgang beendet Gerbing 02.08.2007
    msg = msg & LoadResString(2258 + Sprache)               'Neue Datensätze werden erst nach einer neuen Suche angezeigt
    MsgBox msg
    Unload Me                                                           'Gerbing 29.06.2011
End Sub

