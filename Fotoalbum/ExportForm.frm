VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form ExportForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kopieren(Exportieren) der aktuell selektierten Fotos"
   ClientHeight    =   10224
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   14304
   Icon            =   "ExportForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10224
   ScaleWidth      =   14304
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ListBox lstFensterTitel 
      Height          =   240
      Left            =   6360
      TabIndex        =   25
      Top             =   9240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   5292
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   14052
      Begin VB.OptionButton OptInAndereAnwendung 
         Caption         =   "Mit Drag&&Drop sofort in andere geöffnete Anwendung  (Datenbank fotos.mdb) exportieren"
         Height          =   372
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "es muß eine zweite Anwendung (Datenbank fotos.mdb) geöffnet sein. Dort wird nach erfolgtem Export der Import gestartet."
         Top             =   2040
         Width           =   12852
      End
      Begin VB.Frame Frame4 
         Height          =   1212
         Left            =   960
         TabIndex        =   30
         Top             =   2520
         Width           =   11772
         Begin VB.OptionButton OptAlleDateien 
            Caption         =   "Alle Dateien"
            Height          =   372
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   3015
         End
         Begin VB.OptionButton OptMerkerDateien 
            Caption         =   "mit der Merkerspalte markierte Dateien"
            Height          =   372
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   9012
         End
      End
      Begin VB.CommandButton btnExportVorbereiten 
         Caption         =   "&Export vorbereiten"
         Height          =   375
         Left            =   960
         TabIndex        =   29
         Top             =   3960
         Width           =   3972
      End
      Begin EditCtlsLibUCtl.TextBox txtDatenbank 
         Height          =   372
         Left            =   5280
         TabIndex        =   27
         Top             =   1440
         Width           =   8532
         _cx             =   15049
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
         HAlignment      =   0
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
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   0
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "ExportForm.frx":038A
         Text            =   "ExportForm.frx":03AA
      End
      Begin EditCtlsLibUCtl.TextBox txtZielVerzeichnis 
         Height          =   372
         Left            =   5280
         TabIndex        =   26
         Top             =   840
         Width           =   8532
         _cx             =   15049
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
         HAlignment      =   0
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
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   0
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "ExportForm.frx":03CA
         Text            =   "ExportForm.frx":03EA
      End
      Begin VB.OptionButton OptInsZielverzeichnis 
         Caption         =   "Für den Export das Zielverzeichnis benutzen"
         Height          =   372
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "die exportierten Fotos/Videos landen im Zielverzeichnis"
         Top             =   120
         Value           =   -1  'True
         Width           =   9612
      End
      Begin VB.CheckBox CheckFotos 
         Caption         =   "Kopieren der aktuellen Fotos"
         Height          =   372
         Left            =   480
         TabIndex        =   21
         Top             =   840
         Width           =   3855
      End
      Begin VB.CheckBox CheckDatenbank 
         Caption         =   "Datenbank aus den aktuellen Sätzen erzeugen"
         Height          =   492
         Left            =   480
         TabIndex        =   20
         Top             =   1320
         Width           =   3735
      End
      Begin EditCtlsLibUCtl.TextBox txtDragDropDatenbank 
         Height          =   372
         Left            =   5280
         TabIndex        =   28
         Top             =   4680
         Width           =   8532
         _cx             =   15049
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
         DisabledEvents  =   3074
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
         CueBanner       =   "ExportForm.frx":040A
         Text            =   "ExportForm.frx":042A
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label8 
         Caption         =   "Drag&&Drop Name der Exportdatenbank"
         Height          =   372
         Left            =   960
         TabIndex        =   34
         Top             =   4680
         Width           =   4332
      End
      Begin VB.Label Label2 
         Caption         =   "nach:"
         Height          =   372
         Left            =   4440
         TabIndex        =   23
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "in:"
         Height          =   336
         Left            =   4440
         TabIndex        =   22
         Top             =   1476
         Width           =   612
      End
   End
   Begin VB.TextBox txtDatenmenge 
      BackColor       =   &H00C0C0C0&
      Height          =   372
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Width           =   1692
   End
   Begin VB.Frame Frame1 
      Caption         =   "Einschränkungen nach Datum der Fotos"
      Height          =   2172
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   14052
      Begin VB.Frame Frame2 
         Height          =   1092
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   3972
         Begin VB.OptionButton optDatumNichtEinbeziehen 
            Caption         =   "nicht einbeziehen"
            Height          =   372
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Value           =   -1  'True
            Width           =   3732
         End
         Begin VB.OptionButton optDatumEinbeziehen 
            Caption         =   "in den Vergleich einbeziehen"
            Height          =   372
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3732
         End
      End
      Begin VB.CommandButton btnDatumBis 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   372
         Left            =   8400
         TabIndex        =   13
         Top             =   1440
         Width           =   372
      End
      Begin VB.CommandButton btnDatumVon 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   372
         Left            =   8400
         TabIndex        =   12
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox txtDatumBis 
         Height          =   372
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "beliebig"
         Top             =   1440
         Width           =   2412
      End
      Begin VB.TextBox txtDatumVon 
         Height          =   408
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "beliebig"
         Top             =   960
         Width           =   2412
      End
      Begin VB.Label Label6 
         Caption         =   "bis:"
         Height          =   372
         Left            =   5160
         TabIndex        =   9
         Top             =   1440
         Width           =   612
      End
      Begin VB.Label Label5 
         Caption         =   "von:"
         Height          =   372
         Left            =   5160
         TabIndex        =   8
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "Geändert am:"
         Height          =   372
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1932
      End
   End
   Begin VB.CommandButton btnMerkerSpalte 
      Caption         =   "Kopiere die in der &Merker-Spalte gewählten Fotos"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Das sind die Fotos, die nach Drücken der Taste F5 von Ihnen mit der Zahl 1 in der Merkerspalte vorgemerkt wurden."
      Top             =   9600
      Width           =   6012
   End
   Begin VB.CommandButton btnAbbrechen 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   8880
      Width           =   6012
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   9600
      Width           =   6012
   End
   Begin VB.CommandButton btnKopierealleFotos 
      Caption         =   "&Kopiere alle Fotos"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Das sind alle Fotos, die Sie im Fenster Suchkriterien durch die Eingabe Ihrer Suchbegriffe selektiert haben"
      Top             =   8880
      Width           =   6012
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "voraussichtliche Datenmenge (Bytes):"
      Height          =   372
      Left            =   120
      TabIndex        =   17
      Top             =   7920
      Width           =   4692
   End
   Begin VB.Label lblArbeitsfortschritt 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   5280
      TabIndex        =   4
      Top             =   8280
      Width           =   8892
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   372
      Left            =   5280
      TabIndex        =   3
      Top             =   7920
      Width           =   3972
   End
End
Attribute VB_Name = "ExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public ZielVerzeichnis As String
    Dim SQL As String
    Dim rst2 As ADODB.Recordset                                                         'Gerbing 29.12.2011
    Dim DatumVon As Date
    Dim DatumBis As Date
    Dim Gesamtmenge As Long
    Dim blnAbbruch As Boolean
    Dim MerkeCheckFotos As Integer
    Public blnExportGestartet As Boolean
    Private Const vbMsgBoxTopMost As Long = &H40000                                                         'Gerbing 11.03.2009
      
    Const GW_HWNDFIRST = 0                                                                                  'Gerbing 11.03.2009
    Const GW_HWNDLAST = 1
    Const GW_HWNDNEXT = 2
    Const GW_HWNDPREV = 3
    Const GW_OWNER = 4
    Const GW_CHILD = 5
    Const GW_MAX = 5
    
    Const GWL_STYLE = (-16)
    
    Const WS_VISIBLE = &H10000000
    Const WS_BORDER = &H800000
    
    Private suppressDefaultContextMenu As Boolean
    Private Const CF_OEMTEXT = 7
    Private Const CF_UNICODETEXT = 13

     
Private Sub btnAbbrechen_Click()
    blnAbbruch = True
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    Unload Me
End Sub

Private Sub btnDatumVon_Click()
    Dim DatumVon As Date
    Dim DatumBis As Date

    gstrDateChoose = txtDatumVon
    frmDateChoose.Top = 2000
    frmDateChoose.Left = 2000
    frmDateChoose.Show 1
    DatumVon = gstrDateChoose
    DatumBis = txtDatumBis
    If DatumBis < DatumVon Then
        'MsgBox "Das 'Datum bis' darf nicht früher sein als das 'Datum von'", , "Datums-Kontrolle"
        MsgBox LoadResString(2018 + Sprache), , LoadResString(2019 + Sprache)   'Gerbing 08.11.2005
        Exit Sub
    End If
    txtDatumVon = DatumVon
End Sub

Private Sub btnExportVorbereiten_Click()
    blnAbbruch = False
    blnExportGestartet = True                               'Gerbing 08.12.2006
    If OptInAndereAnwendung.Value = True Then
        Me.MousePointer = vbHourglass                                                       'Gerbing 29.07.2007
        Gesamtmenge = 0
        If txtDatumVon <> LoadResString(1110 + Sprache) Then  'Beliebig
            DatumVon = txtDatumVon
            DatumBis = txtDatumBis
        End If
        If OptAlleDateien.Value = True Then
            'alle gewählten Datensätze in $fotos.mdb eintragen
            'Jeden Satz des frmGridAndThumb.Adodc1.Recordset auswerten
            'alle Datensätze in $fotos.mdb eintragen
            Call DatenbankErzeugen(False)                                       'MerkerAuswerten=false
        Else
            'nur mit Merkerspalte versehenen Datensätze in $fotos.mdb eintragen
            Call DatenbankErzeugen(True)                                        'MerkerAuswerten=true
        End If
        Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
        If blnAbbruch = False Then
            'lblArbeitsfortschritt.Caption = "Fertig"                           'Gerbing 26.01.2006
            lblArbeitsfortschritt.Caption = LoadResString(1007 + Sprache)       'Gerbing 08.11.2005
            Label8.Visible = True
            txtDragDropDatenbank.Visible = True
            MsgBox LoadResString(3068 + Sprache) 'Sie können jetzt den Name der Exportdatenbank mit Drag&Drop in die andere Anwendung ziehen
        Else
            'lblArbeitsfortschritt.Caption = "Abgebrochen"
            lblArbeitsfortschritt.Caption = LoadResString(2220 + Sprache)        'Gerbing 26.01.2006
        End If
    End If
    blnExportGestartet = False
End Sub

Private Sub btnHilfe_Click()
    Dim RetVal As Long
    Dim CHMFile As String
    Dim Msg As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = AppPath & "\Help\Deutsch\fotos.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
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
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub btnKopierealleFotos_Click()
    Dim Mldg As String
    Dim Stil As Long
    Dim antwort As Long
    
    blnAbbruch = False
    If CheckDatenbank.Value = 0 And CheckFotos.Value = 0 Then
        'MsgBox "Sie müssen mindestens ein Häkchen setzen unter 'Für den Export das Zielverzeichnis benutzen"
        MsgBox LoadResString(2020 + Sprache)                          'Gerbing 08.11.2005
        Exit Sub
    End If
    If CheckDatenbank.Value = 1 And CheckFotos.Value = 0 Then
        'Nochmals rückfragen, ob wirklich nur die Datenbank erzeugt werden soll
        'Mldg = "Wollen Sie wirklich nur die Datenbank erzeugen ?"
        Mldg = LoadResString(2021 + Sprache)                         'Gerbing 08.11.2005
        Stil = vbYesNo + vbDefaultButton2
        antwort = MsgBox(Mldg, Stil)
        If antwort = vbNo Then
            Exit Sub
        End If
    End If
    '---------------------------------
    blnExportGestartet = True                                           'Gerbing 08.12.2006
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    'btnAbbrechen.Enabled = False
    CheckFotos.Enabled = False
    CheckDatenbank.Enabled = False
    optDatumEinbeziehen.Enabled = False
    optDatumNichtEinbeziehen.Enabled = False
    btnDatumVon.Enabled = False
    btnDatumBis.Enabled = False
    btnKopierealleFotos.Enabled = False
    btnMerkerSpalte.Enabled = False
    
    If CheckDatenbank.Value = 1 Then
        Gesamtmenge = 0
        If txtDatumVon <> LoadResString(1110 + Sprache) Then    'beliebig
            DatumVon = txtDatumVon
            DatumBis = txtDatumBis
        End If
        Call DatenbankErzeugen(False)                                  'MerkerAuswerten=false
    End If
    If CheckFotos.Value = 1 Then
        Gesamtmenge = 0
        If txtDatumVon <> LoadResString(1110 + Sprache) Then    'beliebig
            DatumVon = txtDatumVon
            DatumBis = txtDatumBis
        End If
        Call KopiereFotos(False)                                            'MerkerAuswerten=False
    End If
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    If blnAbbruch = True Then                                               'Gerbing 26.01.2006
        'lblArbeitsfortschritt.Caption = "abgebrochen"
        lblArbeitsfortschritt.Caption = LoadResString(2220 + Sprache)       'Gerbing 08.11.2005
        'MsgBox "Exportier-Vorgang abgebrochen"
        MsgBox LoadResString(2221 + Sprache)                                'Gerbing 08.11.2005
    Else
        'lblArbeitsfortschritt.Caption = "Fertig"
        lblArbeitsfortschritt.Caption = LoadResString(1007 + Sprache)       'Gerbing 08.11.2005
        'MsgBox "Exportier-Vorgang beendet"
        MsgBox LoadResString(2022 + Sprache)                                'Gerbing 08.11.2005
    End If
    blnExportGestartet = False
    Unload Me
End Sub

Private Sub btnMerkerSpalte_Click()
    'es sollen die mit der Merkerspalte ausgewählten Dateien kopiert werden
    Dim Mldg As String
    Dim Stil As Long
    Dim antwort As Long
    
    blnAbbruch = False
    If CheckDatenbank.Value = 0 And CheckFotos.Value = 0 Then
        'MsgBox "Sie müssen mindestens ein Häkchen setzen unter 'Für den Export das Zielverzeichnis benutzen"
        MsgBox LoadResString(2020 + Sprache)  'Gerbing 08.11.2005
        Exit Sub
    End If
    If CheckDatenbank.Value = 1 And CheckFotos.Value = 0 Then
        'Nochmals rückfragen, ob wirklich nur die Datenbank erzeugt werden soll
        'Mldg = "Wollen Sie wirklich nur die Datenbank erzeugen ?"
        Mldg = LoadResString(2021 + Sprache)  'Gerbing 08.11.2005
        Stil = vbYesNo + vbDefaultButton2
        antwort = MsgBox(Mldg, Stil)
        If antwort = vbNo Then
            Exit Sub
        End If
    End If
    '--------------------------------
    blnExportGestartet = True                                           'Gerbing 08.12.2006
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    'btnAbbrechen.Enabled = False
    CheckFotos.Enabled = False
    CheckDatenbank.Enabled = False
    optDatumEinbeziehen.Enabled = False
    optDatumNichtEinbeziehen.Enabled = False
    btnDatumVon.Enabled = False
    btnDatumBis.Enabled = False
    btnKopierealleFotos.Enabled = False
    btnMerkerSpalte.Enabled = False
    If CheckDatenbank.Value = 1 Then
        Gesamtmenge = 0
        If txtDatumVon <> LoadResString(1110 + Sprache) Then 'beliebig
            DatumVon = txtDatumVon
            DatumBis = txtDatumBis
        End If
        Call DatenbankErzeugen(True)                                   'MerkerAuswerten=true
    End If
    If CheckFotos.Value = 1 Then
        Gesamtmenge = 0
        If StrComp(txtDatumVon, LoadResString(1110 + Sprache), vbTextCompare) <> 0 Then  'beliebig  'Gerbing 08.12.2005
            DatumVon = txtDatumVon
            DatumBis = txtDatumBis
        End If
        Call KopiereFotos(True)                                           'MerkerAuswerten=True
    End If
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    If blnAbbruch = True Then                                               'Gerbing 26.01.2006
        'lblArbeitsfortschritt.Caption = "abgebrochen"
        lblArbeitsfortschritt.Caption = LoadResString(2220 + Sprache)       'Gerbing 08.11.2005
        'MsgBox "Exportier-Vorgang abgebrochen"
        MsgBox LoadResString(2221 + Sprache)                                'Gerbing 08.11.2005
    Else
        'lblArbeitsfortschritt.Caption = "Fertig"
        lblArbeitsfortschritt.Caption = LoadResString(1007 + Sprache)       'Gerbing 08.11.2005
        'MsgBox "Exportier-Vorgang beendet"
        MsgBox LoadResString(2022 + Sprache)                                'Gerbing 08.11.2005
    End If
    blnExportGestartet = False
    Unload Me
End Sub

Private Sub CheckFotos_Click()
    Dim Msg As String
    
    If CheckFotos.Value = 1 Then
        'ZielVerzeichnisForm verlangt eine Nutzer-Eingabe mit dem Ziel-Verzeichnis
        ZielVerzeichnisForm.Show 1                      'öffne ZielVerzeichnisForm
        If Form1.blnUnloadExportForm = True Then
            If MerkeCheckFotos = 0 Then
                CheckFotos.Value = 0
            End If
            Exit Sub
        End If
        txtZielVerzeichnis.Text = ZielVerzeichnis
        'If Dir(ZielVerzeichnis, vbDirectory) = "." Then
        If file_path_exist(ZielVerzeichnis) = False Then
            'Falls der Nutzer kein gültiges Verzeichnis wählt
            Msg = LoadResString(2539 + Sprache) & vbNewLine             'Dieses Zielverzeichnis gibt es nicht.
            Msg = Msg & ZielVerzeichnis
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum 15"), vbInformation
            txtZielVerzeichnis.Visible = False
            CheckFotos.Value = 0
        Else
            txtZielVerzeichnis.Visible = True
        End If
    Else
        txtZielVerzeichnis.Visible = False
    End If
    If CheckDatenbank.Value = 0 And CheckFotos.Value = 0 Then
        btnKopierealleFotos.Enabled = False
    Else
        btnKopierealleFotos.Enabled = True
    End If
End Sub

Private Sub CheckDatenbank_Click()
    Dim Msg As String
    
    If CheckDatenbank.Value = 1 Then
        If gblnSQLServerVersion = True Then
            txtDatenbank.Text = PublicSQLServer & " $" & PublicSQLDatabase
            txtDatenbank.Visible = True
        Else
            'abweisen, wenn $fotos.mdb schreibgeschützt ist
            If gblnSchreibgeschützt = True Then                                     'Gerbing 23.11.2017
                'schreibgeschützt
                Msg = gstrFotosMdbLocation & "\$Fotos.mdb" & vbNewLine
                'Msg= msg & "Die Datenbank ist schreibgeschützt, Export der Datenbanksätze ist nicht möglich"
                Msg = Msg & LoadResString(2224 + Sprache)
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
                CheckDatenbank.Value = 0
                On Error GoTo 0
                Exit Sub
            End If
            On Error GoTo 0
            txtDatenbank.Text = gstrFotosMdbLocation & "\$Fotos.mdb"
            txtDatenbank.Visible = True
        End If
    Else
        txtDatenbank.Visible = False
    End If
    If CheckDatenbank.Value = 0 And CheckFotos.Value = 0 Then
        btnKopierealleFotos.Enabled = False
    Else
        btnKopierealleFotos.Enabled = True
    End If
End Sub

Private Sub DatenbankErzeugen(MerkerAuswerten As Boolean)
    Dim Dollarrst As ADODB.Recordset
    Dim Msg As String
    Dim Einbeziehen As Boolean
    Dim DateTime As Date
    Dim n As Long
    Dim zähler As Long
    Dim SQL As String
    Dim DollarFieldsCount As Long
    Dim FieldsCount As Long
    Dim fs As New Scripting.FileSystemObject
    Dim f
    
    'Alle Sätze werden in die
    'Datenbank $fotos.mdb geschrieben
    
    '1. Datenbank $fotos.mdb Tabelle fotos Delete all records
    If gblnSQLServerVersion = True Then
        'beim SQL Server muss es heißen 'Delete from table
        SQL = "DELETE From Fotos"
    Else
        SQL = "DELETE * FROM Fotos"         '2529=Fotos
    End If
    DollarDBado.Execute (SQL)    '2529=fotos                                            'Gerbing 23.11.2017
    'Dollarrst nimmt die neuen Sätze auf
    SQL = "Select  * FROM Fotos "
    Set Dollarrst = New ADODB.Recordset
    With Dollarrst
        .Source = SQL
        .ActiveConnection = DollarDBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    DollarFieldsCount = Dollarrst.Fields.Count
    
    '2. Jeden Satz des frmGridAndThumb.Adodc1 nach $fotos.mdb übernehmen
        'Jeden Satz des rstsql nach $fotos.mdb übernehmen
        SQL = frmGridAndThumb.Adodc1.RecordSource                                        'Das hat der Nutzer im Formular Query formuliert
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        FieldsCount = rstsql.Fields.Count
        If DollarFieldsCount <> FieldsCount Then
'            msg = "Datenbank und $Datenbank müssen die gleiche Tabellenstruktur besitzen." & vbNewLine
'            msg = msg & "Erzeugen Sie die $Datenbank durch Kopieren der Datenbank."
            If gblnSQLServerVersion = True Then
                Msg = PublicSQLDatabase & " and " & "$" & PublicSQLDatabase & LoadResString(1833 + Sprache) & "." & vbNewLine
                Msg = Msg & LoadResString(1834 + Sprache) & "$" & PublicSQLDatabase & LoadResString(1835 + Sprache) & PublicSQLDatabase & "."
            Else
                Msg = "fotos.mdb and $fotos.mdb " & LoadResString(1833 + Sprache) & "." & vbNewLine
                Msg = Msg & LoadResString(1834 + Sprache) & "$fotos.mdb" & LoadResString(1835 + Sprache) & "fotos.mdb" & "."
            End If
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
            Exit Sub
        End If
        
        rstsql.MoveFirst
        zähler = 0
        Do Until rstsql.EOF            'Schleife bis Ende des Recordset
            'Neuen Datensatz erstellen in $fotos.mdb
            'Falls mit Datum-Einschränkung gearbeitet wird,
            'nur die Sätze mit zutreffendem Datei-Datum übernehmen
            '----------------------------------------------------
            'Das Datum 'Geändert am' in den Vergleich einbeziehen
            Einbeziehen = True
            If optDatumEinbeziehen.Value = True Then
                Call FRODateinameRstsql
                'DateTime = FileDateTime(gstrFRODN)
                Set f = fs.GetFile(gstrFRODN)                                  'Gerbing 04.03.2013
                DateTime = f.DateLastModified
                If DateTime < DatumVon Then
                    Einbeziehen = False
                End If
                'DatumVon und DatumBis geben nur den konkreten Tag an, die Uhrzeit wird mit 00:00:00 angenommen
                'darum 1 zum DatumBis addieren
                If DateTime > DatumBis + 1 Then
                    Einbeziehen = False
                End If
            End If
            If MerkerAuswerten = True Then
                'Merker' in den Vergleich einbeziehen                                               'Gerbing 29.06.2011
                If rstsql.Fields(LoadResString(2524 + Sprache)) = 0 Then
                    Einbeziehen = False
                End If
            End If
            If Einbeziehen = True Then
                Dollarrst.AddNew                      ' Neuen Datensatz erstellen.
                For n = 0 To rstsql.Fields.Count - 1
                    Dollarrst.Fields(n) = rstsql.Fields(n)
                Next n
                Dollarrst.Update                      ' Änderungen speichern.
            End If
            zähler = zähler + 1
            'lblArbeitsfortschritt.Caption = "Datensatznummer=" & Zähler                    'Gerbing 08.12.2006
            lblArbeitsfortschritt.Caption = LoadResString(1008 + Sprache) & zähler          'Gerbing 08.12.2006
            rstsql.MoveNext
            'DoEvents 'auskommentiert weil Blockierung auftritt                              'Gerbing 13.08.2009
            'stattdessen Control.Refresh
            txtDatenmenge.Refresh                                                           'Gerbing 29.06.2011
            lblArbeitsfortschritt.Refresh                                                   'Gerbing 29.06.2011
            If blnAbbruch = True Then
                Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
                Exit Do
            End If
        Loop
End Sub

Private Sub KopiereFotos(MerkerAuswerten As Boolean)
    Dim Quellname As String
    Dim QuellAnteil As String
    Dim Zielname As String
    Dim Msg As String
    Dim DatensatzNr, Erg As Long
    Dim DateTime As Date
    Dim Einbeziehen As Boolean
    Dim Einzelmenge As Long
    Dim SQL As String
    Dim fs As New Scripting.FileSystemObject
    Dim f
    Dim rc As Boolean

    If MerkerAuswerten = True Then
        'Selektiere alle Sätze in  Fotos, wo Spalte Merker=ja
        'die in "Dateiname" gefundene Datei nach ZielVerzeichnis kopieren
        'Nötige Unterverzeichnisse selbst erzeugen
    '    SQL = " SELECT *"
    '    SQL = SQL & " FROM " & "Fotos"
    '    SQL = SQL & " WHERE Merker<>0 & ";"
        SQL = " SELECT *"
        SQL = SQL & " FROM Fotos"
        SQL = SQL & " WHERE " & LoadResString(2524 + Sprache) & "<>0;" 'Gerbing 08.11.2005
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        If rstsql.EOF Then
            rstsql.Close
            MsgBox LoadResString(3061 + Sprache) 'Es gibt keine mit Merkerspalte markierten Sätze
            Exit Sub
        End If
    Else
        'Jeden Satz des rstsql auswerten
        'und die in "Dateiname" gefundene Datei nach ZielVerzeichnis kopieren
        'Nötige Unterverzeichnisse selbst erzeugen
        SQL = frmGridAndThumb.Adodc1.RecordSource
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                               'Gerbing 23.11.2017
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    rstsql.MoveFirst
    DatensatzNr = 1
    Do Until rstsql.EOF            'Schleife bis Ende des Recordset
        'Neuen Datensatz erstellen
        Call ReplaceFRODateiname
        Quellname = gstrFRODN
        QuellAnteil = Right(Quellname, Len(Quellname) - Len(gstrFotosMdbLocation))   'Vom Quellname gstrFotosmdblocation abschneiden    'Gerbing 27.01.2006
        Zielname = ZielVerzeichnis & QuellAnteil
        On Error Resume Next
        '----------------------------------------------------
        'Das Datum 'Geändert am' in den Vergleich einbeziehen
        Einbeziehen = True
        If optDatumEinbeziehen.Value = True Then
            'DateTime = FileDateTime(Quellname)
            Set f = fs.GetFile(Quellname)                                  'Gerbing 04.03.2013
            DateTime = f.DateLastModified
            If DateTime < DatumVon Then
                Einbeziehen = False
            End If
            'DatumVon und DatumBis geben nur den konkreten Tag an, die Uhrzeit wird mit 00:00:00 angenommen
            'darum 1 zum DatumBis addieren
            If DateTime > DatumBis + 1 Then
                Einbeziehen = False
            End If
        End If
        If Einbeziehen = True Then
'            FileCopy Quellname, Zielname
'            If Err <> 0 Then
            rc = file_copy(Quellname, Zielname)
            If rc = False Then
                'Wenn das Kopieren mit Fehler reagiert, muß ich ein neues Unterverzeichnis anlegen
                Call UnterverzeichnisErzeugen(Quellname, Zielname)
            Else
                Call AudioDateiMitkopieren(Quellname, Zielname)                             'Gerbing 29.07.2007
            End If
            'Datenmenge ausrechnen
            Einzelmenge = FileLen(Quellname)
            Gesamtmenge = Gesamtmenge + Einzelmenge
            txtDatenmenge = Format(Gesamtmenge, "###,###,###")
            'DoEvents 'auskommentiert weil Blockierung auftritt                              'Gerbing 13.08.2009
            'stattdessen Control.Refresh
            txtDatenmenge.Refresh                                                           'Gerbing 29.06.2011
            lblArbeitsfortschritt.Refresh                                                   'Gerbing 29.06.2011
        End If
        '----------------------------------------------------
        rstsql.MoveNext
        DatensatzNr = DatensatzNr + 1
        'lblArbeitsfortschritt.Caption = "DatensatzNr." & DatensatzNr
        lblArbeitsfortschritt.Caption = LoadResString(1008 + Sprache) & DatensatzNr       'Gerbing 08.11.2005
        Erg = DatensatzNr Mod 20
        If Erg = 0 Then
            'DoEvents hat zum Blockieren geführt darum auskommentiert                       'Gerbing 29.09.2009
            'stattdessen Control.Refresh                                                    'Gerbing 29.06.2011
        End If
        If blnAbbruch = True Then
            Me.MousePointer = vbDefault                                                     'Gerbing 29.07.2007
            Exit Do
        End If
    Loop
End Sub

Private Sub UnterverzeichnisErzeugen(Quellname, Zielname)
    Dim QuellAnteil As String
    Dim ZielAnteil As String
    Dim Pos, start As Long
    Dim Verz As String
    Dim Msg As String
    Dim rc As Boolean
    
    QuellAnteil = Right(Quellname, Len(Quellname) - Len(gstrFotosMdbLocation) - 1)  'Vom Quellname gstrFotosmdblocation abschneiden    'Gerbing 27.01.2006
    start = 1
    ZielAnteil = ZielVerzeichnis
    Do
        Pos = InStr(start, QuellAnteil, "\", vbTextCompare)
        If Pos = 0 Then Exit Do
        Verz = Left(QuellAnteil, Pos - 1)
        QuellAnteil = Right(QuellAnteil, Len(QuellAnteil) - Pos)
        ZielAnteil = ZielAnteil & "\" & Verz
        On Error Resume Next
        Err.Number = 0
        MkDir ZielAnteil    'Kein Fehler wenn das Verzeichnis in einem
                            'vorhergehenden Aufruf schon angelegt wurde
    Loop
    On Error Resume Next
    Err.Number = 0
    'Jetzt das mißlungene FileCopy wiederholen
    rc = file_copy(Quellname, Zielname)
    If rc = False Then
        If Err.Number = 76 Or Err.Number = 52 Then                                  'Gerbing 23.06.2011
            'msg = "Kein Schreibzugriff bei FileCopy" & NL                          'Gerbing 26.09.2007
            Msg = LoadResString(2327 + Sprache) & NL
            Msg = Msg & Quellname & "," & Zielname & NL
            Msg = Msg & "Errorcode=" & Err.Number & NL                              'Gerbing 26.09.2007
            Msg = Msg & "Errortext=" & Err.Description
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Else
            'msg = "Programmierfehler bei FileCopy" & NL
            Msg = LoadResString(2034 + Sprache) & NL
            Msg = Msg & Quellname & "," & Zielname & NL
            Msg = Msg & "Errorcode=" & Err.Number & NL                              'Gerbing 26.09.2007
            Msg = Msg & "Errortext=" & Err.Description
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        End If
        'End                                                                        'Gerbing 28.03.2014
    End If
    Call AudioDateiMitkopieren(Quellname, Zielname)                                 'Gerbing 29.07.2007
End Sub

Private Sub btnDatumBis_Click()
    Dim DatumVon As Date
    Dim DatumBis As Date

    gstrDateChoose = txtDatumBis
    frmDateChoose.Top = 2000
    frmDateChoose.Left = 2000
    frmDateChoose.Show 1
    DatumBis = gstrDateChoose
    DatumVon = txtDatumVon
    If DatumBis < DatumVon Then
        'MsgBox "Das 'Datum bis' darf nicht früher sein als das 'Datum von'", , "Datums-Kontrolle"
        MsgBox LoadResString(2018 + Sprache), , LoadResString(2019 + Sprache)   'Gerbing 08.11.2005
        Exit Sub
    End If
    txtDatumBis = DatumBis
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                           'Gerbing 11.03.2017
    Me.Top = Form1.Top                                      'Gerbing 06.12.2006
    Me.Left = Form1.Left
    Label1.Caption = LoadResString(1014 + Sprache)        'Arbeitsfortschritt:                'Gerbing 08.11.2005
    Label2.Caption = LoadResString(1015 + Sprache)        'nach:
    Label3.Caption = LoadResString(1016 + Sprache)        'in:
    Label4.Caption = LoadResString(1017 + Sprache)        'Geändert am:
    Label5.Caption = LoadResString(1018 + Sprache)        'von:
    Label6.Caption = LoadResString(1019 + Sprache)        'bis:
    Label7.Caption = LoadResString(1020 + Sprache)        'voraussichtliche Datenmenge (Bytes):
    Me.Caption = LoadResString(1021 + Sprache)            'Kopieren der aktuell selektierten Fotos"
    btnMerkerSpalte.tooltipText = LoadResString(2501 + Sprache) 'Das sind die Fotos, die nach Drücken der Taste F5 von Ihnen mit der Zahl 1 in der Merkerspalte vorgemerkt wurden.
    btnKopierealleFotos.tooltipText = LoadResString(2502 + Sprache)        'Das sind alle Fotos, die Sie im Fenster Suchkriterien durch die Eingabe Ihrer Suchbegriffe selektiert haben
    CheckFotos.Caption = LoadResString(3006 + Sprache) 'Kopieren der aktuellen Fotos
    CheckDatenbank.Caption = LoadResString(3007 + Sprache) 'Datenbank aus den aktuellen Sätzen erzeugen
    Frame1.Caption = LoadResString(3008 + Sprache) 'Einschränkungen nach Datum der Fotos
    optDatumEinbeziehen.Caption = LoadResString(3009 + Sprache) 'in den Vergleich einbeziehen
    optDatumNichtEinbeziehen.Caption = LoadResString(3010 + Sprache)     'nicht einbeziehen
    btnKopierealleFotos.Caption = LoadResString(3011 + Sprache) '&Kopiere alle Fotos
    btnMerkerSpalte.Caption = LoadResString(3012 + Sprache)    'Kopiere die in der &Merker-Spalte gewählten Fotos
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)   '&Abbrechen
    btnHilfe.Caption = LoadResString(3014 + Sprache)   '&Hilfe
    OptInsZielverzeichnis.Caption = LoadResString(3054 + Sprache) 'Für den Export das Zielverzeichnis benutzen
    OptInAndereAnwendung.Caption = LoadResString(3055 + Sprache) 'sofort in andere geöffnete Anwendung exportieren
    OptInAndereAnwendung.tooltipText = LoadResString(3072 + Sprache) 'es muß eine zweite Anwendung (Datenbank fotos.mdb) geöffnet sein. Dort liegt der Zielpunkt des Exportierens mit Drag&Drop.
    OptInsZielverzeichnis.tooltipText = LoadResString(3057 + Sprache) 'die exportierten Fotos/Videos landen im Zielverzeichnis
    Label8.Caption = LoadResString(3060 + Sprache) 'Drag&&Drop Name der Exportdatenbank
    btnExportVorbereiten.Caption = LoadResString(3067 + Sprache) '&Export vorbereiten
    txtDatumVon.Text = LoadResString(1110 + Sprache)   'beliebig
    txtDatumBis.Text = LoadResString(1110 + Sprache)   'beliebig
    OptAlleDateien.Caption = LoadResString(1112 + Sprache) 'Alle Dateien
    OptMerkerDateien.Caption = LoadResString(1113 + Sprache) 'mit der Merkerspalte markierte Dateien

    txtDatenbank.Visible = False
    Frame4.Visible = False
    txtDragDropDatenbank.Text = gstrFotosMdbLocation & "\$Fotos.mdb"
    txtDragDropDatenbank.Visible = False
    btnExportVorbereiten.Visible = False
    Label8.Visible = False
    If gblnSchreibgeschützt = True Then                     'Gerbing 26.01.2006
        btnMerkerSpalte.Enabled = False
    End If
    'Me.Show
    'CheckFotos.Value = 1
    MerkeCheckFotos = 0
    If gblnSQLServerVersion = True Then
        'FrameDragDrop.Visible = False                      'Gerbing 03.12.2014
    End If
End Sub

Private Sub optDatumEinbeziehen_Click()
    If optDatumEinbeziehen.Value = True Then
        txtDatumVon = Date
        txtDatumBis = Date
        btnDatumVon.Enabled = True
        btnDatumBis.Enabled = True
    End If
End Sub

Private Sub optDatumNichtEinbeziehen_Click()
    txtDatumVon = LoadResString(1110 + Sprache) 'beliebig
    txtDatumBis = LoadResString(1110 + Sprache) 'beliebig
    btnDatumVon.Enabled = False
    btnDatumBis.Enabled = False
End Sub

Private Sub OptInAndereAnwendung_Click()
    'wenn jetzt noch keine zweite Anwendung (d.h. keine zweite Datenbank fotos.mdb) geöffnet ist, muss
    'das Einschalten dieser Option verhindert werden
    Dim x As Long
    Dim Msg As String
    Dim hWnd As Long                                                                                           'Gerbing 11.03.2009

'    Set ws = DBEngine.Workspaces(0)
'    Set Dollardb = ws.OpenDatabase _
'        (gstrFotosMdbLocation & "\$Fotos.mdb", _
'        False, False, "MS Access;")
'
'    If Err.Number = 3051 Or Err.Number = 3045 Then      'Gerbing 20.04.2005
'        'schreibgeschützt
'        Msg = gstrFotosMdbLocation & "\$Fotos.mdb" & vbNewLine
'        'Msg= msg & "Die Datenbank ist schreibgeschützt, Export mit Drag&Drop ist nicht möglich"
'        Msg = Msg & LoadResString(2222 + Sprache)
'        'MsgBox Msg
'        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
'        OptInsZielverzeichnis.Value = True
'        On Error GoTo 0
'        Exit Sub
'    End If
    On Error GoTo 0
    
    If Query.chkFensterGrößeÄnderbar.Value = 0 Then
        MsgBox LoadResString(3066 + Sprache) 'Wenn Sie 'sofort in andere geöffnete Anwendung exportieren' wollen, darf die Anwendung nicht den gesamten Bildschirm ausfüllen. Beide Anwendungen müssen mit 'Fenstergröße änderbar' gestartet werden, weil für das Auslösen des Kopierens(Exportierens) Drag&Drop zwischen Export-Fenster und Import-Fenster benutzt wird.
        OptInsZielverzeichnis.Value = True
        Exit Sub
    End If
    
    lstFensterTitel.Clear
    'Auch der Desktop ist ein Fenster                                                                       'Gerbing 11.03.2009
    hWnd = GetDesktopWindow
    Call GetWindowInfo(hWnd)
    'Einstieg
    'hWnd = GetWindow(Me.hWnd, GW_HWNDFIRST)
    hWnd = GetWindow(Form1.hWnd, GW_HWNDFIRST)
    'Alle vorhandenen Fenster abklappern
    Do
    'In dieser Schleife werden gefüllt
    'lstFensterTitel enthält Application Titles
        Call GetWindowInfo(hWnd)
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        'X = DoEvents()                                                     'wichtig kein DoEvents machen   'Gerbing 11.03.2009
    Loop Until hWnd = 0                                                                                     'Gerbing 11.03.2009
    
    If lstFensterTitel.ListCount < 3 Then                                                                   'Gerbing 28.03.2014
        Msg = LoadResString(3072 + Sprache) & vbNewLine 'es muß eine zweite Anwendung (Datenbank fotos.mdb) geöffnet sein. Dort liegt der Zielpunkt des Exportierens mit Drag&Drop.
        Msg = Msg & LoadResString(3059 + Sprache)   'Beide Anwendungen müssen mit 'Fenstergröße änderbar' gestartet werden, weil für das Auslösen des Kopierens(Exportierens) Drag&Drop zwischen Export-Fenster und Import-Fenster benutzt wird.
        MsgBox Msg, vbMsgBoxSetForeground + vbMsgBoxTopMost                                                 'Gerbing 11.03.2009
        OptInsZielverzeichnis.Value = True
        OptInAndereAnwendung.Value = False                                                                  'Gerbing 28.03.2014
        Exit Sub
    End If
    
    CheckFotos.Visible = False
    CheckDatenbank.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    txtZielVerzeichnis.Visible = False
    txtDatenbank.Visible = False
'    Label8.Visible = True
'    txtDragDropDatenbank.Visible = True
    btnKopierealleFotos.Visible = False
    btnMerkerSpalte.Visible = False
    Frame4.Visible = True
    btnExportVorbereiten.Visible = True
End Sub

Private Sub GetWindowInfo(ByVal hWnd As Long)                                                                      'Gerbing 11.03.2009
  Dim parent As Long
  Dim Task As Long
  Dim result As Long
  Dim x As Long
  Dim style As Long
  Dim Title As String
  
    'Darstellung des Fensters
    style = GetWindowLong(hWnd, GWL_STYLE)
    style = style And (WS_VISIBLE Or WS_BORDER)
            
    'Title des Fenster auslesen
    result = GetWindowTextLength(hWnd) + 1
    Title = Space$(result)
    result = GetWindowText(hWnd, Title, result)
    Title = Left$(Title, Len(Title) - 1)
    
    'In Abhängigkeit der Optionen die Ausgabe erstellen
    If Title <> "" Then
        If result > 0 Then
            If Left(Title, Len(LoadResString(1001 + Sprache))) = LoadResString(1001 + Sprache) Then
                'ob "FotoAlbum-" im Fenstertitel steht
                lstFensterTitel.AddItem Title
            End If
        End If
      'Elternfenster ermitteln
      parent = hWnd
      Do
        parent = GetParent(parent)
      Loop Until parent = 0
      'Task Id ermitteln
      result = GetWindowThreadProcessId(hWnd, Task)
    End If
End Sub

Private Sub OptInsZielverzeichnis_Click()
    CheckFotos.Visible = True
    CheckDatenbank.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    txtZielVerzeichnis.Visible = True
    If CheckDatenbank.Value = 1 Then
        txtDatenbank.Visible = True
    End If
    Label8.Visible = False
    txtDragDropDatenbank.Visible = False
    btnKopierealleFotos.Visible = True
    btnMerkerSpalte.Visible = True
    Frame4.Visible = False
    btnExportVorbereiten.Visible = False
End Sub

Sub ReplaceFRODateiname()
    gstrFRODN = Replace(rstsql(LoadResString(1028 + Sprache)), "+:\", gstrFotosMdbLocation & "\") '1028=Dateiname  'Gerbing 07.11.2011
End Sub

Private Sub txtDragDropDatenbank_AbortedDrag()
    txtDragDropDatenbank.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, -1
End Sub

Private Sub txtDragDropDatenbank_BeginDrag(ByVal firstChar As Long, ByVal lastChar As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    txtDragDropDatenbank.OLEDrag , , , firstChar, lastChar
End Sub

Private Sub txtDragDropDatenbank_BeginRDrag(ByVal firstChar As Long, ByVal lastChar As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    txtDragDropDatenbank.OLEDrag , , , firstChar, lastChar
End Sub

Private Sub txtDragDropDatenbank_ContextMenu(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  ' we don't want the default menu during right-button drag'n'drop
  showDefaultMenu = Not suppressDefaultContextMenu
End Sub

Private Sub txtDragDropDatenbank_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    txtDragDropDatenbank.OLEDrag
End Sub

Private Sub txtDragDropDatenbank_OLEDragMouseMove(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, effect As EditCtlsLibUCtl.OLEDropEffectConstants, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
    Dim newCharIndex As Long
    Dim newInsertMarkRelativePosition As InsertMarkPositionConstants
    
    txtDragDropDatenbank.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, newCharIndex
    txtDragDropDatenbank.SetInsertMarkPosition newInsertMarkRelativePosition, newCharIndex
    
    effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
    If Shift And vbShiftMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeMove
    If Shift And vbCtrlMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeCopy
    If Shift And vbAltMask Then effect = EditCtlsLibUCtl.OLEDropEffectConstants.odeLink
End Sub

Private Sub txtDragDropDatenbank_OLESetData(ByVal Data As EditCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
    Dim firstChar As Long
    Dim lastChar As Long
    
    Select Case formatID
      Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
        txtDragDropDatenbank.GetDraggedTextRange firstChar, lastChar
        Data.SetData formatID, txtDragDropDatenbank.Text
    End Select
End Sub

Private Sub txtDragDropDatenbank_OLEStartDrag(ByVal Data As EditCtlsLibUCtl.IOLEDataObject)
    Data.SetData vbCFText
    Data.SetData CF_OEMTEXT
    Data.SetData CF_UNICODETEXT
End Sub

