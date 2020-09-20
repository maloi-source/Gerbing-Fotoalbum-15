VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form AboutForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Version"
   ClientHeight    =   8160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11796
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11796
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameAllgemein 
      Height          =   2172
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   11412
      Begin EditCtlsLibUCtl.TextBox txtInstallationsordner 
         Height          =   372
         Left            =   2520
         TabIndex        =   30
         Top             =   1680
         Width           =   8172
         _cx             =   14414
         _cy             =   656
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   0
         AutoScrolling   =   2
         BackColor       =   -2147483633
         BorderStyle     =   0
         CancelIMECompositionOnSetFocus=   0   'False
         CharacterConversion=   0
         CompleteIMECompositionOnKillFocus=   0   'False
         DisabledBackColor=   -2147483633
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
         ReadOnly        =   -1  'True
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
         CueBanner       =   "AboutForm.frx":038A
         Text            =   "AboutForm.frx":03AA
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "App.Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5892
      End
      Begin VB.Label lblAbout 
         Caption         =   "Copyright (C) 1997-2020 by gerbing software"
         Height          =   348
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   5892
      End
      Begin VB.Label Label1 
         Caption         =   "http://www.gerbingsoft.de"
         Height          =   372
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   4212
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version "
         Height          =   348
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   3888
      End
      Begin VB.Label lblInstallationsordner 
         Caption         =   "Installationsordner:"
         Height          =   372
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   2412
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4800
      TabIndex        =   1
      Top             =   7680
      Width           =   2220
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   240
      Picture         =   "AboutForm.frx":03CA
      ScaleHeight     =   263.118
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   263.118
      TabIndex        =   0
      Top             =   120
      Width           =   432
   End
   Begin VB.Frame FrameSQLServer 
      Height          =   3972
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   11412
      Begin EditCtlsLibUCtl.TextBox txtSQLServerName 
         Height          =   372
         Left            =   3720
         TabIndex        =   22
         Top             =   1200
         Width           =   7572
         _cx             =   13356
         _cy             =   656
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   0
         AutoScrolling   =   2
         BackColor       =   -2147483633
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
         Enabled         =   0   'False
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
         CueBanner       =   "AboutForm.frx":100C
         Text            =   "AboutForm.frx":102C
      End
      Begin EditCtlsLibUCtl.TextBox txtDatabaseName 
         Height          =   372
         Left            =   3720
         TabIndex        =   23
         Top             =   1680
         Width           =   7572
         _cx             =   13356
         _cy             =   656
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   0
         AutoScrolling   =   2
         BackColor       =   -2147483633
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
         Enabled         =   0   'False
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
         CueBanner       =   "AboutForm.frx":105C
         Text            =   "AboutForm.frx":107C
      End
      Begin EditCtlsLibUCtl.TextBox txtLoggedInAs 
         Height          =   372
         Left            =   3720
         TabIndex        =   21
         Top             =   2640
         Width           =   7572
         _cx             =   13356
         _cy             =   656
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   0
         AutoScrolling   =   2
         BackColor       =   -2147483633
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
         Enabled         =   0   'False
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
         CueBanner       =   "AboutForm.frx":10AC
         Text            =   "AboutForm.frx":10CC
      End
      Begin EditCtlsLibUCtl.TextBox txtStandortFotos 
         Height          =   372
         Left            =   3720
         TabIndex        =   20
         Top             =   720
         Width           =   7572
         _cx             =   13356
         _cy             =   656
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   0
         AutoScrolling   =   2
         BackColor       =   -2147483633
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
         Enabled         =   0   'False
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
         CueBanner       =   "AboutForm.frx":10FC
         Text            =   "AboutForm.frx":111C
      End
      Begin VB.CheckBox chkWindowsAuthentication 
         Enabled         =   0   'False
         Height          =   372
         Left            =   3720
         TabIndex        =   14
         Top             =   2160
         Width           =   372
      End
      Begin VB.Label lblAnzahlEingeloggteUser 
         Caption         =   "anzahl"
         Height          =   372
         Left            =   3720
         TabIndex        =   19
         Top             =   3600
         Width           =   852
      End
      Begin VB.Label lbllblAnzahlEingeloggteUser 
         Caption         =   "Anzahl eingeloggte user:"
         Height          =   372
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   3492
      End
      Begin VB.Label lblAnzahlDerLizenzen 
         Caption         =   "anzahl"
         Height          =   372
         Left            =   3720
         TabIndex        =   17
         Top             =   3120
         Width           =   852
      End
      Begin VB.Label lbllblAnzahlDerLizenzen 
         Caption         =   "Anzahl der Lizenzen:"
         Height          =   372
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   3492
      End
      Begin VB.Label lbllblLoggedInAs 
         Caption         =   "eingeloggt als:"
         Height          =   372
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   3492
      End
      Begin VB.Label lblWindowsAuthentication 
         Caption         =   "Windows Authentication:"
         Height          =   372
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   3612
      End
      Begin VB.Label lbllblDatabasename 
         Caption         =   "Database name:"
         Height          =   372
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   3492
      End
      Begin VB.Label lbllblSQLservername 
         Caption         =   "SQL server name:"
         Height          =   372
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   3492
      End
      Begin VB.Label lbllblStandortFotos 
         Caption         =   "Standort der Fotos/Videos:"
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3492
      End
      Begin VB.Label Label2 
         Caption         =   "MS SQL Server"
         Height          =   372
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   5172
      End
      Begin VB.Label lblSQLVerwendeteDatenbank 
         Caption         =   "verwendete Datenbank:"
         Height          =   372
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3492
      End
   End
   Begin VB.Frame FrameAccess 
      Height          =   2412
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   11412
      Begin VB.Label lblAccessVerwendeteDatenbank 
         Caption         =   "verwendete Datenbank:"
         Height          =   372
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3372
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "lblDescription"
         ForeColor       =   &H00000000&
         Height          =   336
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   10932
      End
      Begin VB.Label Label3 
         Caption         =   "MS Access"
         Height          =   372
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   7452
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   432
      Left            =   2880
      Picture         =   "AboutForm.frx":114C
      Top             =   120
      Width           =   6156
   End
   Begin VB.Label lblGeschichteDieserSoftware 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Geschichte dieser Software lesen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   7080
      Width           =   5412
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   7560
      Y2              =   7560
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
'        Me.Top = Form1.Top + Form1.Height \ 2 - Me.Height \ 2
'        Me.Left = Form1.Left + Form1.Width \ 2 - Me.Width \ 2
        Me.Top = Form1.Top
        Me.Left = Form1.Left
    End If
    
    ' Datei-Informationen ermitteln                                 'Gerbing 16.08.2013
    'MsgBox 1
    lblVersion.Caption = "Version " & GetFotosExeVersion            'Gerbing 16.08.2013
    
    'MsgBox 2
    lblTitle.Caption = glbAppTitle    'anstelle App.Title                                   'Gerbing 04.03.2013
    lblInstallationsordner.Caption = LoadResString(1128 + Sprache)                          'Installationsordner:
    txtInstallationsordner.Text = AppPath

    If gblnSQLServerVersion = False Then
        FrameSQLServer.Visible = False
        FrameAccess.Visible = True
        lblAccessVerwendeteDatenbank = LoadResString(1813 + Sprache)    'verwendete Datenbank
        If gblnVollversion = True Then
            lblDescription = "Professional-Version"         'Gerbing 13.10.2005
        Else
            lblDescription = "Shareware-Version"            'Gerbing 13.10.2005
        End If
    Else
        FrameAccess.Visible = False
        FrameSQLServer.Visible = True
        lblSQLVerwendeteDatenbank = LoadResString(1813 + Sprache)       'verwendete Datenbank
        lbllblStandortFotos = LoadResString(1812 + Sprache)             'lbl Standort der Fotos/Videos:
        txtStandortFotos = PublicLocationFotos
        lbllblSQLservername = LoadResString(1805 + Sprache)             'lbl SQL server name
        txtSQLServerName = PublicSQLServer
        lbllblDatabasename = LoadResString(1806 + Sprache)              'lbl datase name
        txtDatabaseName = PublicSQLDatabase
        If PublicWindowsAuthentication = "1" Then
            chkWindowsAuthentication.Value = 1
        End If
        lbllblLoggedInAs = LoadResString(1814 + Sprache)                'eingeloggt als:
        txtLoggedInAs = gstrLoggedInName
        lbllblAnzahlDerLizenzen = LoadResString(1815 + Sprache)         'allowed licenses:
        lblAnzahlDerLizenzen = gstrAllowedlicenses
        lbllblAnzahlEingeloggteUser = LoadResString(1816 + Sprache)     'Anzahl eingeloggte user:
        If gstrAllowedlicenses = 99 Then
            lblAnzahlEingeloggteUser = "unlimited"
        Else
            Set rstsql = New ADODB.Recordset
            With rstsql
                .Source = "select * from loggedinusers where (loggedin <> 0)"
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            lblAnzahlEingeloggteUser = rstsql.RecordCount
            rstsql.Close
        End If
    End If

    cmdOK.Caption = LoadResString(3001 + Sprache)           '&OK
    lblGeschichteDieserSoftware.Caption = LoadResString(1108 + Sprache) 'Geschichte dieser Software lesen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gblnComeFromAboutMenue = False Then              'Gerbing 19.09.2012
        gblnComeFromAboutMenue = False                  'Gerbing 19.09.2012
    End If
End Sub

Private Sub lblGeschichteDieserSoftware_Click()
    frmGeschichteDieserSoftware.Show 1
End Sub

