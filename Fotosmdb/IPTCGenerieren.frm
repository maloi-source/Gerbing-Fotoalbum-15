VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form IPTCGenerieren 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Datenbankfelder übertragen in EXIF/IPTC-Felder von JPG-Dateien"
   ClientHeight    =   9456
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   13932
   Icon            =   "IPTCGenerieren.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9456
   ScaleWidth      =   13932
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtExifToolOutput 
      Height          =   2052
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   68
      Top             =   120
      Width           =   6972
   End
   Begin VB.CheckBox chkSetIPTCPresent 
      BackColor       =   &H00C0C0C0&
      Caption         =   "set IPTCPresent = true"
      Height          =   372
      Left            =   240
      TabIndex        =   65
      Top             =   1200
      Value           =   1  'Aktiviert
      Width           =   6492
   End
   Begin EditCtlsLibUCtl.TextBox txtArbeitsfortschritt 
      Height          =   372
      Left            =   2520
      TabIndex        =   64
      Top             =   5760
      Width           =   11292
      _cx             =   19918
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
      CueBanner       =   "IPTCGenerieren.frx":038A
      Text            =   "IPTCGenerieren.frx":03AA
   End
   Begin VB.Frame Frame2 
      Height          =   2052
      Left            =   6840
      TabIndex        =   44
      Top             =   120
      Width           =   6972
      Begin VB.ComboBox cmbEinzelnesJahr 
         Height          =   288
         Left            =   360
         TabIndex        =   67
         Top             =   1080
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.OptionButton OptEinzelnesJahr 
         Caption         =   "Für ein einzelnes Jahr"
         Height          =   492
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   6612
      End
      Begin VB.OptionButton optNurFalse 
         Caption         =   "nur in JPG-Fotos mit IPTCPresent=False übertragen"
         Height          =   492
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Value           =   -1  'True
         Width           =   6732
      End
      Begin VB.OptionButton optInAlle 
         Caption         =   "in alle JPG-Fotos übertragen"
         Height          =   432
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   6732
      End
   End
   Begin VB.CheckBox chkDatumNichtAktualisieren 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datei-Datum (Geändert am) nicht aktualisieren"
      Height          =   372
      Left            =   240
      TabIndex        =   39
      Top             =   720
      Value           =   1  'Aktiviert
      Width           =   6492
   End
   Begin VB.Frame FrameStandardWerte 
      BackColor       =   &H0080C0FF&
      Caption         =   "Standard-Felder aus der Datenbank"
      Height          =   3372
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   13692
      Begin CBLCtlsLibUCtl.ComboBox cmbSWF 
         Height          =   288
         Left            =   2400
         TabIndex        =   53
         Top             =   2520
         Width           =   6492
         _cx             =   11451
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":03CA
         Text            =   "IPTCGenerieren.frx":03EA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbPersonen 
         Height          =   288
         Left            =   2400
         TabIndex        =   52
         Top             =   2160
         Width           =   6492
         _cx             =   11451
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":040A
         Text            =   "IPTCGenerieren.frx":042A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbLand 
         Height          =   288
         Left            =   2400
         TabIndex        =   51
         Top             =   1800
         Width           =   6492
         _cx             =   11451
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":044A
         Text            =   "IPTCGenerieren.frx":046A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbOrt 
         Height          =   288
         Left            =   2400
         TabIndex        =   50
         Top             =   1440
         Width           =   6492
         _cx             =   11451
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":048A
         Text            =   "IPTCGenerieren.frx":04AA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbSituation 
         Height          =   288
         Left            =   2400
         TabIndex        =   49
         Top             =   1080
         Width           =   6492
         _cx             =   11451
         _cy             =   508
         AcceptNumbersOnly=   -1  'True
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":04CA
         Text            =   "IPTCGenerieren.frx":04EA
      End
      Begin VB.ComboBox cmbSWFEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   48
         Top             =   2520
         Width           =   4000
      End
      Begin VB.ComboBox cmbJahr 
         Height          =   288
         Left            =   2400
         TabIndex        =   43
         Top             =   720
         Width           =   972
      End
      Begin VB.ComboBox cmbJahrEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtKommentar 
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Top             =   2880
         Width           =   6492
      End
      Begin VB.ComboBox cmbSituationEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   26
         Top             =   1080
         Width           =   4000
      End
      Begin VB.ComboBox cmbOrtEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   25
         Top             =   1440
         Width           =   4000
      End
      Begin VB.ComboBox cmbLandEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   24
         Top             =   1800
         Width           =   4000
      End
      Begin VB.ComboBox cmbPersonenEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   2160
         Width           =   4000
      End
      Begin VB.ComboBox cmbKommentarEx 
         Height          =   288
         Left            =   9500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   2880
         Width           =   4000
      End
      Begin VB.Label lblFeldinhalt1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Feldinhalt"
         Height          =   372
         Left            =   4440
         TabIndex        =   70
         Top             =   360
         Width           =   2292
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   11
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":050A
         Stretch         =   -1  'True
         Top             =   840
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   5
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":094C
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   4
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":0D8E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   3
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":11D0
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   2
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":1612
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   1
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":1A54
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   0
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":1E96
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   132
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "SWF:"
         Height          =   252
         Left            =   240
         TabIndex        =   47
         Top             =   2520
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Quell-Feld"
         Height          =   372
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Jahr:"
         Height          =   252
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   1932
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Situation:"
         Height          =   252
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   2052
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ort:"
         Height          =   252
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   2052
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Land:"
         Height          =   252
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Personen:"
         Height          =   252
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   2052
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kommentar:"
         Height          =   252
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   2052
      End
      Begin VB.Label lblExifIptc1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ziel EXIF/IPTC-Feld"
         Height          =   372
         Left            =   9500
         TabIndex        =   28
         ToolTipText     =   "IPTC-Felder, die Sie nicht auswählen, bleiben falls sie bisher existieren unverändert in der JPG-Datei erhalten"
         Top             =   360
         Width           =   4000
      End
   End
   Begin VB.Frame FrameNutzerDefiniert 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nutzerdefinierte Felder aus der Datenbank"
      Height          =   2652
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   13692
      Begin CBLCtlsLibUCtl.ComboBox Combo5 
         Height          =   288
         Left            =   4440
         TabIndex        =   58
         Top             =   2160
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":22D8
         Text            =   "IPTCGenerieren.frx":22F8
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo4 
         Height          =   288
         Left            =   4440
         TabIndex        =   57
         Top             =   1800
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2318
         Text            =   "IPTCGenerieren.frx":2338
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo3 
         Height          =   288
         Left            =   4440
         TabIndex        =   56
         Top             =   1440
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2358
         Text            =   "IPTCGenerieren.frx":2378
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo2 
         Height          =   288
         Left            =   4440
         TabIndex        =   55
         Top             =   1080
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2398
         Text            =   "IPTCGenerieren.frx":23B8
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo1 
         Height          =   288
         Left            =   4440
         TabIndex        =   54
         Top             =   720
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   0   'False
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":23D8
         Text            =   "IPTCGenerieren.frx":23F8
      End
      Begin VB.ComboBox cmbEx1 
         Height          =   288
         Left            =   9480
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   3972
      End
      Begin VB.ComboBox cmbEx2 
         Height          =   288
         Left            =   9480
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   3972
      End
      Begin VB.ComboBox cmbEx3 
         Height          =   288
         Left            =   9480
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   3972
      End
      Begin VB.ComboBox cmbEx4 
         Height          =   288
         Left            =   9480
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   3972
      End
      Begin VB.ComboBox cmbEx5 
         Height          =   288
         Left            =   9480
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   3972
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld1 
         Height          =   288
         Left            =   2280
         TabIndex        =   59
         Top             =   720
         Width           =   1932
         _cx             =   3408
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2418
         Text            =   "IPTCGenerieren.frx":2438
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld2 
         Height          =   288
         Left            =   2280
         TabIndex        =   60
         Top             =   1080
         Width           =   1932
         _cx             =   3408
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2458
         Text            =   "IPTCGenerieren.frx":2478
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld3 
         Height          =   288
         Left            =   2280
         TabIndex        =   61
         Top             =   1440
         Width           =   1932
         _cx             =   3408
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2498
         Text            =   "IPTCGenerieren.frx":24B8
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld4 
         Height          =   288
         Left            =   2280
         TabIndex        =   62
         Top             =   1800
         Width           =   1932
         _cx             =   3408
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":24D8
         Text            =   "IPTCGenerieren.frx":24F8
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld5 
         Height          =   288
         Left            =   2280
         TabIndex        =   63
         Top             =   2160
         Width           =   1932
         _cx             =   3408
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   267503
         DontRedraw      =   0   'False
         DoOEMConversion =   0   'False
         DragDropDownTime=   -1
         DropDownKey     =   0
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         IntegralHeight  =   -1  'True
         ItemHeight      =   -1
         ListAlwaysShowVerticalScrollBar=   0   'False
         ListBackColor   =   -2147483643
         ListDragScrollTimeBase=   -1
         ListForeColor   =   -2147483640
         ListHeight      =   -1
         ListInsertMarkColor=   0
         ListScrollableWidth=   0
         ListWidth       =   0
         Locale          =   1024
         MaxTextLength   =   -1
         MinVisibleItems =   30
         MousePointer    =   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         SelectionFieldHeight=   -1
         Sorted          =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "IPTCGenerieren.frx":2518
         Text            =   "IPTCGenerieren.frx":2538
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Quell-Feld"
         Height          =   372
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1332
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   10
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":2558
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   9
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":299A
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   8
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":2DDC
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   7
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":321E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   6
         Left            =   9120
         Picture         =   "IPTCGenerieren.frx":3660
         Stretch         =   -1  'True
         Top             =   840
         Width           =   132
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4280
         TabIndex        =   20
         Top             =   720
         Width           =   252
      End
      Begin VB.Label lblFeldname1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4280
         TabIndex        =   18
         Top             =   1080
         Width           =   252
      End
      Begin VB.Label lblFeldname2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2052
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4280
         TabIndex        =   16
         Top             =   1440
         Width           =   252
      End
      Begin VB.Label lblFeldname3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2052
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4280
         TabIndex        =   14
         Top             =   1800
         Width           =   252
      End
      Begin VB.Label lblFeldname4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4280
         TabIndex        =   12
         Top             =   2160
         Width           =   252
      End
      Begin VB.Label lblFeldname5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   2052
      End
      Begin VB.Label lblFeldinhalt2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldinhalt"
         Height          =   372
         Left            =   4440
         TabIndex        =   10
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label lblExifIptc2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ziel EXIF/IPTC-Feld"
         Height          =   372
         Left            =   9480
         TabIndex        =   9
         Top             =   360
         Width           =   4092
      End
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "S&tart"
      Default         =   -1  'True
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1692
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   492
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1692
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   492
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label lblNochÜbertragen 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   12840
      TabIndex        =   69
      Top             =   6240
      Width           =   972
   End
   Begin VB.Label lblNochZuÜbertragen 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   12840
      TabIndex        =   38
      Top             =   5760
      Width           =   972
   End
   Begin VB.Label lblNochZuVerarb 
      BackColor       =   &H00C0C0C0&
      Caption         =   "noch vorzubereiten:"
      Height          =   372
      Left            =   9600
      TabIndex        =   37
      Top             =   6240
      Width           =   3132
   End
   Begin VB.Label lblSchonÜbertragen 
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   6240
      TabIndex        =   36
      Top             =   6240
      Width           =   972
   End
   Begin VB.Label lblSchonVerarbeitet 
      BackColor       =   &H00C0C0C0&
      Caption         =   "für exiftool vorbereitet:"
      Height          =   372
      Left            =   2520
      TabIndex        =   35
      Top             =   6240
      Width           =   3612
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   372
      Left            =   120
      TabIndex        =   34
      Top             =   4560
      Width           =   2292
   End
End
Attribute VB_Name = "IPTCGenerieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim MyAppPath As String
    Dim NL As String
    
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias _
            "RtlMoveMemory" (lpTo As Any, lpFrom As Any, _
            ByVal lLen As Long)
           
    Dim b() As Byte                             'Array zum Einlesen/Ausgeben der JPG-Datei
    
    Dim Filename As String
    Dim blnAbbrechenGeklickt As Boolean

    
    '21 IPTC-Felder, die mit Gasanov IPTC Ocx angezeigt werden können
    Dim QuellFeldFürObjectName As String
    Dim QuellFeldFürUrgency As String
    Dim QuellFeldFürCategory As String
    Dim QuellFeldFürSupplementalCategories As String
    Dim QuellFeldFürKeywords As String
    Dim QuellFeldFürSpecialInstructions As String
    Dim QuellFeldFürDateCreated As String
    Dim QuellFeldFürByline As String
    Dim QuellFeldFürBylineTitle As String
    Dim QuellFeldFürCity As String
    Dim QuellFeldFürState As String
    Dim QuellFeldFürCountry As String
    Dim QuellFeldFürOriginalTransmissionReference As String
    Dim QuellFeldFürHeadline As String
    Dim QuellFeldFürCredits As String
    Dim QuellFeldFürSource As String
    Dim QuellFeldFürCaption As String
    Dim QuellFeldFürCaptionWriter As String
    Dim QuellFeldFürTimeCreated As String
    Dim QuellFeldFürCopyright As String
    Dim QuellFeldFürOriginatingProgram As String
    '--------------------------------------------
    'jetzt die restlichen 8
    '--------------------------------------------
    Dim QuellFeldFürReleaseDate As String
    Dim QuellFeldFürReleaseTime As String
    Dim QuellFeldFürObjectcycle As String
    Dim QuellFeldFürLocationCode As String
    Dim QuellFeldFürSubLocation As String
    Dim QuellFeldFürProgramVersion As String
    Dim QuellFeldFürEditStatus As String
    Dim QuellFeldFürJobID As String
    '--------------------------------------------
    'jetzt die 5 EXIF-Felder                                            'Gerbing 16.11.2015
    '--------------------------------------------
    Dim QuellFeldFürXPTitle As String
    Dim QuellFeldFürXPSubject As String
    Dim QuellFeldFürXPKeywords As String
    Dim QuellFeldFürXPComment As String
    Dim QuellFeldFürXPAuthor As String
    Dim QuellFeldFürGPSLatitude As String                               'Gerbing 02.10.2019
    Dim QuellFeldFürGPSLongitude As String                              'Gerbing 02.10.2019
    
    'wird für Datum/Uhrzeit Retten und wiederherstellen gebraucht------------------------------------------------Start
    'Gerbing 23.06.2009
    Dim MeinDatum As Date
    Dim DatumJetzt As Date
    
    Private Type FILETIME
      dwLowDateTime As Long
      dwHighDateTime As Long
    End Type
    
    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type

    Private Const GENERIC_WRITE = &H40000000
    Private Const OPEN_EXISTING = 3
    Private Const FILE_SHARE_READ = &H1
    Private Const FILE_SHARE_WRITE = &H2
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    
    Dim lngHandle As Long
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME
        'wird für Datum/Uhrzeit Retten und wiederherstellen gebraucht------------------------------------------------End
        
    'Gerbing 16.11.2015--------------------------------------------------------------------------------------------Start
    Private Const INFINITE = -1&
    Dim blnArgFileErzeugt As Boolean
    Dim ArgfileTxtFile As String
    
    Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
    
    Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
    End Type

    Private Type PROCESS_INFORMATION
          hProcess As Long
          hThread As Long
          dwProcessId As Long
          dwThreadId As Long
    End Type
    
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
          hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
    Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
          lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
          lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
          ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
          ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
          lpStartupInfo As STARTUPINFO, lpProcessInformation As _
          PROCESS_INFORMATION) As Long
    
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
          (ByVal hProcess As Long, lpExitCode As Long) As Long

    Private Const NORMAL_PRIORITY_CLASS = &H20&
    Dim strArgfile As String
    Dim strHTausend As String
    Dim strSammeln As String
    Dim blnExiftoolExists As Boolean
    Dim blnDefaultFieldsNotEmpty As Boolean                                         'Gerbing 27.10.2016


    'Purpose: Unicode aware MkDir
    Public Function MkDir(ByVal lpPathName As String, Optional ByVal lpSecurityAttributes As Long = 0) As Boolean
       MkDir = CreateDirectoryW(StrPtr("\\?\" & lpPathName), ByVal lpSecurityAttributes) <> 0
    End Function
'---Gerbing 16.11.2015----------------------------------------------------------------------------------------------End

Private Sub btnAbbrechen_Click()
    blnAbbrechenGeklickt = True
    btnStart.Enabled = True                                                         'Gerbing 04.01.2020
    Unload Me
End Sub

Private Sub btnHilfe_Click()
    Dim retval As Long
    Dim CHMFile As String
    Dim Msg As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = AppPath & "\Help\Deutsch\fotosmdb.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING fotosmdb"), vbInformation
            Exit Sub
        Else
            retval = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If retval <= 32 Then
                Call HelpFileErrorMsg(retval, CHMFile)
            End If
        End If
    Else
        CHMFile = AppPath & "\Help\English\fotosmdb.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING fotosmdb"), vbInformation
            Exit Sub
        Else
            retval = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If retval <= 32 Then
                Call HelpFileErrorMsg(retval, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub btnStart_Click()
    Dim SQL As String
    Dim Msg As String
    Dim antwort As Long
    Dim Dateiname As String
    Dim DateinamenErweiterung As String
    Dim n As Long
    Dim RetteFileDateTime As Date
    Dim rc As Integer
    Dim zählerSchon As Long
    Dim zählerNoch As Long
    Dim ArgNew As String
    Dim rstDefFields As ADODB.Recordset                                     'Gerbing 27.10.2016
    Dim rstUserDef As ADODB.Recordset                                       'Gerbing 12.10.2019
    Dim blnErrorsAufgesammelt As Boolean
    Dim strErrSammlung As String
        
    btnStart.Enabled = False                                                'Gerbing 04.01.2020
    blnAbbrechenGeklickt = False
    'es ist verboten beim Frame Standardfelder alle IPTC-Ziele leer zu lassen
    If cmbJahrEx.Text = "" And cmbSituationEx.Text = "" And cmbOrtEx.Text = "" And cmbLandEx.Text = "" And cmbPersonenEx.Text = "" And cmbSWFEx.Text = "" And cmbKommentarEx.Text = "" Then 'Gerbing 04.01.2009
        'msg = "You did not define target EXIF/IPTC fields"
        Msg = LoadResString(1531 + Sprache)
        MsgBox Msg
        btnStart.Enabled = True                                             'Gerbing 04.01.2020
        Exit Sub
    End If
    '---------------------------------------------------------
    'Bei nutzerdefinierten Feldern auf korrekte Angaben prüfen      'Gerbing 14.02.2005
    If FrameNutzerDefiniert.Visible = True Then
        'Kontrolle der Benutzerauswahl
        If cmbEx1.Text <> "" And cmbFeld1.Text = "" Then
            'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
            MsgBox LoadResString(1528 + Sprache)
            cmbFeld1.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbEx2.Text <> "" And cmbFeld2.Text = "" Then
            'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
            MsgBox LoadResString(1528 + Sprache)
            cmbFeld2.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbEx3.Text <> "" And cmbFeld3.Text = "" Then
            'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
            MsgBox LoadResString(1528 + Sprache)
            cmbFeld3.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbEx4.Text <> "" And cmbFeld4.Text = "" Then
            'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
            MsgBox LoadResString(1528 + Sprache)
            cmbFeld4.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbEx5.Text <> "" And cmbFeld5.Text = "" Then
            'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
            MsgBox LoadResString(1528 + Sprache)
            cmbFeld5.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        '-----------------------------------------------
        If cmbFeld1.Text <> "" And cmbEx1.Text = "" Then
            'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch ein EXIF/IPTC-Feld auswählen"
            MsgBox LoadResString(1529 + Sprache)
            cmbEx1.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbFeld2.Text <> "" And cmbEx2.Text = "" Then
            'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch ein EXIF/IPTC-Feld auswählen"
            MsgBox LoadResString(1529 + Sprache)
            cmbEx2.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbFeld3.Text <> "" And cmbEx3.Text = "" Then
            'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch ein EXIF/IPTC-Feld auswählen"
            MsgBox LoadResString(1529 + Sprache)
            cmbEx3.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbFeld4.Text <> "" And cmbEx4.Text = "" Then
            'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch ein EXIF/IPTC-Feld auswählen"
            MsgBox LoadResString(1529 + Sprache)
            cmbEx4.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
        If cmbFeld5.Text <> "" And cmbEx5.Text = "" Then
            'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch ein EXIF/IPTC-Feld auswählen"
            MsgBox LoadResString(1529 + Sprache)
            cmbEx5.SetFocus
            btnStart.Enabled = True                                             'Gerbing 04.01.2020
            Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------
    If optInAlle.Value = True Then                          'Gerbing 04.02.2008
        'ich will einen Recordset mit sämtlichen Fotos in aufsteigender Reihenfolge des Dateinamens
        'SQL = "Select * from Fotos ORDER BY Dateiname"
        SQL = "Select * from Fotos ORDER BY " & LoadResString(1028 + Sprache)
    Else
        If OptEinzelnesJahr.Value = True Then
            If cmbEinzelnesJahr.Text = "" Then
                'msg= "Sie haben unter" & " '" & "Für ein einzelnes Jahr" & "' " & "keine Jahreszahl ausgewählt"
                Msg = LoadResString(2340 + Sprache) & " '" & LoadResString(2287 + Sprache) & "' " & LoadResString(2341 + Sprache)
                MsgBox Msg, , "FotosMdb"
                btnStart.Enabled = True                                             'Gerbing 04.01.2020
                Exit Sub
            Else
                'ich will einen Recordset mit sämtlichen Fotos in aufsteigender Reihenfolge des Dateinamens für ein einzelnes Jahr
                'SQL = "Select * from Fotos where Jahr =  cmbeinzelnesJahr.Text ORDER BY Dateiname"
                SQL = "Select * from Fotos where " & LoadResString(1023 + Sprache) & "=" & cmbEinzelnesJahr.Text & " ORDER BY " & LoadResString(1028 + Sprache)  '1023=jahr
            End If
        Else
            'ich will einen Recordset mit Fotos where IPTCPresent = False in aufsteigender Reihenfolge des Dateinamens
            'SQL = "Select * from Fotos where IPTCPresent=False ORDER BY Dateiname"
            SQL = "Select * from Fotos where IPTCPresent=0 ORDER BY " & LoadResString(1028 + Sprache)
        End If
    End If
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        '.CursorLocation = adUseServer  'bei adUseServer kommt kein RecordCount             'Gerbing 16.11.2015
        .Open
    End With
    If Form1.rstsql.RecordCount <> 0 Then                                                   'Gerbing 04.01.2009
        Form1.rstsql.MoveFirst
    End If
    zählerSchon = 0
    zählerNoch = Form1.rstsql.RecordCount
    blnErrorsAufgesammelt = False
    '--------------------------------------------------------------------------------------
    'Schleife durch alle Sätze der Datenbank                                    'Gerbing 21.11.2016
    Do
        If Form1.rstsql.EOF Then
            If strHTausend = "" Then
                txtArbeitsfortschritt.Text = LoadResString(1130 + Sprache)          'nothing to do Gerbing 08.04.2017
            End If
            Exit Do
        End If
        strSammeln = ""
        strHTausend = ""
        strArgfile = ""
        Call DeleteExiftoolFiles                                                'Gerbing 16.11.2015
        ExiftoolNichtBenutzbar = False                                          'Gerbing 16.11.2015

        'Schleife für jeweils 100 Dateinamen die nach argfile.txt wandern, bei jedem neuen Dateiname wächst argfile.txt
        'Inhalt von argfile.txt = ""
        'bei jedem neuen Dateiname wächst argfile.txt
        'am Ende der 100 (oder weniger falls unter 100) steht '-execute' und '-stay_open False'
        'dann wird argfile.txt geclosed
        'dann Starte 'exiftool.bat' mit dem Inhalt 'exiftool -stay_open true -@ argfile.txt 2> exiftoolerrors.log'
        'dann arbeitet exiftool und liefert seine StdOut Ausgaben nach txtExifToolOutput
        'dann muss ich vermutlich die Fehler von exiftoolerrors.log aufsammeln
        'dann beginnt die Schleife erneut
        Do                                                                      'Gerbing 21.11.2016
            If Form1.rstsql.EOF Then
                Exit Do
            End If
            'wenn es ein JPG-Foto ist werden die Datenbank-Felder in die EXIF/IPTC-Felder übertragen
            Dateiname = Form1.rstsql.Fields(LoadResString(1028 + Sprache))   '1028=Dateiname
            Dateiname = Replace(Dateiname, "+:\", MyAppPath & "\")
            'MessageBoxW 0, StrPtr(Dateiname), StrPtr("Dateiname mit MyAppPath"), vbInformation
            DateinamenErweiterung = UCase(Right(Dateiname, 3))
            If DateinamenErweiterung = "JPG" Then
                'SchreibenExifIptc erledigt die Arbeit für EXIF/IPTC-Felder
                strHTausend = SchreibenExifIptc(strArgfile, Form1.rstsql, Dateiname)
                If strHTausend = "no exiftool" Then
                    Form1.rstsql.Close
                    btnStart.Enabled = True                                             'Gerbing 04.01.2020
                    Exit Sub
                End If
                strHTausend = strHTausend & "-execute" & vbNewLine    '-execute muss am Ende jeder Datei stehen und startet deren Bearbeitung
                strHTausend = strHTausend & "#" & vbNewLine
                
                If Len(strHTausend) > 100000 Then                       'String concatenation wird sonst
                    strSammeln = strSammeln & strHTausend               'immer langsamer
                    strHTausend = ""
                    strArgfile = ""
                Else
                    strArgfile = strHTausend
                End If
                
                If rc = 0 Then
                    blnArgFileErzeugt = True
                    If chkSetIPTCPresent.Value = 1 Then
                        'Feld IPTCPresent eintragen                                                 'Gerbing 04.02.2008
                         On Error Resume Next    'Für den Fall dass der Nutzer eine schreibgeschützte fotos.mdb benutzt
                        'Form1.rstsql.Edit
                        Form1.rstsql.Fields("IPTCPresent") = True
                        Form1.rstsql.Update
                        On Error GoTo 0
                    End If
                Else
                    blnArgFileErzeugt = False
                    Exit Do
                End If
            End If
            txtArbeitsfortschritt.Text = Dateiname
            zählerSchon = zählerSchon + 1
            zählerNoch = zählerNoch - 1
            lblSchonÜbertragen.Caption = zählerSchon
            lblNochÜbertragen.Caption = zählerNoch
            DoEvents
            If blnAbbrechenGeklickt = True Then Exit Do
            If zählerSchon = 100 Or zählerNoch = 0 Then
                If strHTausend <> "" Then
                    'argfile.txt schreiben als UTF-8 Datei
                    'Nach dem Schreiben argfile.txt wird exiftool gestartet
                    strHTausend = strHTausend & "-stay_open" & vbNewLine
                    strHTausend = strHTausend & "false" & vbNewLine
                    strSammeln = strSammeln & strHTausend
                    WriteUTF8File AppPath & "\argfile.txt", strSammeln
                    '--------------------------------------------------
                    txtExifToolOutput.Visible = True                                                            'Gerbing 21.11.2016
                    rc = StarteExifTool
                    If rc = 1 Then                          'rc = 1 Fehler beim Start von exiftool
                        'Unload Me
                        btnStart.Enabled = True                                             'Gerbing 04.01.2020
                        Exit Sub
                    End If
                    '--------------------------------------------------
                    'Aufsammeln exiftoolerrors.log
                    'Wenn exiftoolerrors.log nicht leer ist, muss ich sie aufsammeln
                    If FileLen(AppPath & "\exiftoolerrors.log") <> 0 Then
                        Call ErrorsAufsammeln(AppPath & "\exiftoolerrors.log", strErrSammlung)
                        blnErrorsAufgesammelt = True
                    End If
                    Exit Do
                Else
                    txtArbeitsfortschritt.Text = LoadResString(1130 + Sprache)              'nothing to do Gerbing 08.04.2017
                End If
            End If
            If Form1.rstsql.EOF Then
                Exit Do
            Else
                Form1.rstsql.Movenext                                                                       'Gerbing 03.01.2009
            End If
        Loop
        '--------------------------------------------------------------------------------------------------------------
        'hier gehts weiter am Ende der 100-er Schleife
        If Form1.rstsql.EOF Then
            Exit Do
        Else
            Form1.rstsql.Movenext                                                                       'Gerbing 03.01.2009
        End If
    Loop
    'Ende der Schleife durch die Datenbank
    '------------------------------------------------------------------------------------------------------------------
    txtExifToolOutput.Visible = False                                                                   'Gerbing 21.11.2016
    If blnAbbrechenGeklickt = False Then
        Form1.rstsql.Close
    End If
    Screen.MousePointer = vbDefault              'Gerbing 17.03.2005
    '-----------------------------------------------
    If blnAbbrechenGeklickt = False Then
        'Wenn exiftoolerrors.log nicht leer gewesen ist, muss eine Msgbox kommen
        If blnErrorsAufgesammelt = True Then
            'in strErrSammlung stehen die aufgesammelten exiftoolerror.log Zeilen
            'die müssen wieder in eine Datei exiftoolerrors.log geschrieben werden
            Call ErrorsNeuSchreiben(AppPath & "\exiftoolerrors.log", strErrSammlung)
            'Msg = "Beim Datenbankfelder übertragen in EXIF/IPTC-Felder von JPG-Fotos (Batch-Modus) sind Fehler aufgetreten." & NL
            'Msg = Msg & "Sie finden die Fehler-Nachrichten in"
            Msg = LoadResString(1539 + Sprache) & NL
            Msg = Msg & LoadResString(1502 + Sprache) & NL
            Msg = Msg & AppPath & "\exiftoolerrors.log"
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
        End If
        If FileLen(AppPath & "\exiftoolerrors.log") = 64 Then
            Msg = "Wenn der Fehler lautet " & "Can't locate PAR.pm in @INC (@INC contains: .) at -e line 860." & vbNewLine
            Msg = Msg & "dann haben Sie im AppPath mindestens ein unicode Zeichen." & vbNewLine
            Msg = Msg & "exiftool.exe kann nicht arbeiten, wenn im AppPath unicode Zeichen stehen."
            
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
        End If
    End If
    '-------------------------------------------------------------------------------------------
    'nur wenn die Tabelle DefaultFields existiert                             'Gerbing 27.10.2016
    'sie kann leer oder nicht leer sein
    On Error Resume Next
    SQL = "select * From DefaultFields;"
    Set rstDefFields = New ADODB.Recordset
    With rstDefFields
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number = 0 Then
        'hier existiert die Tabelle DefaultFields
        If rstDefFields.EOF Then
            'hier ist die Tabelle DefaultFields leer
            'hier brauche ich AddNew
            If gblnSchreibgeschützt = False Then
                rstDefFields.AddNew
                rstDefFields.Fields("SituationSource") = cmbSituationEx.Text
                rstDefFields.Fields("LocationSource") = cmbOrtEx.Text
                rstDefFields.Fields("CountrySource") = cmbLandEx.Text
                rstDefFields.Fields("PeopleSource") = cmbPersonenEx.Text
                rstDefFields.Fields("BWCSource") = cmbSWFEx.Text
                rstDefFields.Fields("CommentSource") = cmbKommentarEx.Text
                rstDefFields.Update
            End If
            rstDefFields.Close
        Else                                                                        'Gerbing 27.10.2016
            'hier ist die Tabelle DefaultFields nicht leer
            If rstDefFields.RecordCount = 1 Then
                If gblnSchreibgeschützt = False Then
                    rstDefFields.Fields("SituationSource") = cmbSituationEx.Text
                    rstDefFields.Fields("LocationSource") = cmbOrtEx.Text
                    rstDefFields.Fields("CountrySource") = cmbLandEx.Text
                    rstDefFields.Fields("PeopleSource") = cmbPersonenEx.Text
                    rstDefFields.Fields("BWCSource") = cmbSWFEx.Text
                    rstDefFields.Fields("CommentSource") = cmbKommentarEx.Text
                    rstDefFields.Update
                End If
            End If
            rstDefFields.Close
        End If
    End If
    On Error GoTo 0
    '---------------------------------End Gerbing 27.10.2016------------------------------------
    If FrameNutzerDefiniert.Visible = True Then                                     'Gerbing 12.10.2019
        'nur wenn nutzerdefinierte Felder existieren
        'nur wenn die Tabelle UserDefined existiert
        'und nur wenn die Tabelle UserDefined leer ist. Leer ist sie, wenn vorher FrameNutzerDefiniert rechtsgeklickt wurde
        On Error Resume Next
        SQL = "select * From UserDefined;"
        Set rstUserDef = New ADODB.Recordset
        With rstUserDef
            .Source = SQL
            .ActiveConnection = Form1.DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        If ERR.Number = 0 Then
            On Error GoTo 0
            'hier existiert die Tabelle UserDefined
            If rstUserDef.EOF Then
                'hier ist die Tabelle UserDefined leer
                If gblnSchreibgeschützt = False Then
                    If cmbFeld1.Text <> "" Or cmbFeld2.Text <> "" Or cmbFeld3.Text <> "" Or cmbFeld4.Text <> "" Or cmbFeld5.Text <> "" Then
                        rstUserDef.AddNew
                        rstUserDef.Fields("FieldName1") = cmbFeld1.Text
                        rstUserDef.Fields("SourceField1") = cmbEx1.Text
                        rstUserDef.Fields("FieldName2") = cmbFeld2.Text
                        rstUserDef.Fields("SourceField2") = cmbEx2.Text
                        rstUserDef.Fields("FieldName3") = cmbFeld3.Text
                        rstUserDef.Fields("SourceField3") = cmbEx3.Text
                        rstUserDef.Fields("FieldName4") = cmbFeld4.Text
                        rstUserDef.Fields("SourceField4") = cmbEx4.Text
                        rstUserDef.Fields("FieldName5") = cmbFeld5.Text
                        rstUserDef.Fields("SourceField5") = cmbEx5.Text
                        rstUserDef.Update
                    End If
                End If
                rstUserDef.Close
            End If
        End If
    End If
    'Unload Me
    btnStart.Enabled = True                                             'Gerbing 04.01.2020
End Sub

Private Sub cmbEx1_Click()
    If cmbEx1.Text <> "" Then
        If cmbEx1.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbSWFEx.Text Then                                                 'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        If cmbEx1.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbEx1)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Ex1 die Quelle sein
        Call QuelleZuZielZurückNehmen(cmbFeld1.Text)                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(cmbFeld1.Text, cmbEx1.Text)   'der Datenbank-Feldname steht in cmbFeld1, das EXIF/IPTC-Feld in cmbEx1.Text
    End If
End Sub

Private Sub cmbEx1_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbEx1.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbEx1.ListIndex = -1
        cmbEx1.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbEx2_Click()
    If cmbEx2.Text <> "" Then
        If cmbEx2.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbSWFEx.Text Then                                                 'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        If cmbEx2.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbEx2)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll cmbFeld2 die Quelle sein
        Call QuelleZuZielZurückNehmen(cmbFeld2.Text)                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(cmbFeld2.Text, cmbEx2.Text)   'der Datenbank-Feldname steht in cmbFeld2, das EXIF/IPTC-Feld in cmbEx2.Text
    End If
End Sub


Private Sub cmbEx2_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbEx2.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbEx2.ListIndex = -1
        cmbEx2.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbEx3_Click()
    If cmbEx3.Text <> "" Then
            If cmbEx3.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbSWFEx.Text Then                                                 'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        If cmbEx3.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbEx3)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll cmbEx3 die Quelle sein
        Call QuelleZuZielZurückNehmen(cmbFeld3.Text)                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(cmbFeld3.Text, cmbEx3.Text)   'der Datenbank-Feldname steht in cmbFeld3, das EXIF/IPTC-Feld in cmbEx3.text
    End If
End Sub

Private Sub cmbEx3_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbEx3.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbEx3.ListIndex = -1
        cmbEx3.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbEx4_Click()
    If cmbEx4.Text <> "" Then
        If cmbEx4.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbSWFEx.Text Then                                                 'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        If cmbEx4.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbEx4)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll cmbFeld4 die Quelle sein
        Call QuelleZuZielZurückNehmen(cmbFeld4.Text)                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(cmbFeld4.Text, cmbEx4.Text)   'der Datenbank-Feldname steht in cmbFeld4, das EXIF/IPTC-Feld in cmbEx4.Text
    End If
End Sub

Private Sub cmbEx4_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbEx4.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbEx4.ListIndex = -1
        cmbEx4.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbEx5_Click()
    If cmbEx5.Text <> "" Then
        If cmbEx5.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbSWFEx.Text Then                                                 'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        If cmbEx5.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbEx5)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll cmbFeld5 die Quelle sein
        Call QuelleZuZielZurückNehmen(cmbFeld5.Text)                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(cmbFeld5.Text, cmbEx5.Text)   'der Datenbank-Feldname steht in cmbFeld5, das EXIF/IPTC-Feld in cmbEx5.Text
    End If
End Sub

Private Sub cmbEx5_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbEx5.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbEx5.ListIndex = -1
        cmbEx5.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbJahrEx_Click()                                   'Gerbing 23.01.2008
    If cmbJahrEx.Text <> "" Then
        If cmbJahrEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbSWFEx.Text Then                                              'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        If cmbJahrEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbJahrEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Jahr die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1023 + Sprache))                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1023 + Sprache), cmbJahrEx.Text)   '1023=jahr, das EXIF/IPTC-Feld steht in cmbJahrEx.Text
    End If
End Sub

Private Sub cmbKommentarEx_Click()
    If cmbKommentarEx.Text <> "" Then
        If cmbKommentarEx.Text = cmbJahr.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbSWFEx.Text Then                                         'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        If cmbKommentarEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbKommentarEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Kommentar die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1030 + Sprache))                                 'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1030 + Sprache), cmbKommentarEx.Text)   '1030=Kommentar, das EXIF/IPTC-Feld steht in cmbKommentarEx.Text
    End If
End Sub

Private Sub cmbKommentarEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                     'Gerbing 16.11.2015
    
    RetteText = cmbKommentarEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbKommentarEx.ListIndex = -1
        cmbKommentarEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbLandEx_Click()
    If cmbLandEx.Text <> "" Then
        If cmbLandEx.Text = cmbJahrEx.Text Then                 'Gerbing 23.01.2008
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbSWFEx.Text Then                                              'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        If cmbLandEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbLandEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Land die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1026 + Sprache))                              'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1026 + Sprache), cmbLandEx.Text)   '1026=land, das EXIF/IPTC-Feld steht in cmbLandEx.Text
    End If
End Sub

Private Sub cmbLandEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbLandEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbLandEx.ListIndex = -1
        cmbLandEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbOrtEx_Click()
    If cmbOrtEx.Text <> "" Then
        If cmbOrtEx.Text = cmbJahrEx.Text Then                  'Gerbing 23.01.2008
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbSWFEx.Text Then                                               'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        If cmbOrtEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbOrtEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Ort die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1025 + Sprache))                             'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1025 + Sprache), cmbOrtEx.Text)   '1025=ort, das EXIF/IPTC-Feld steht in cmbOrtEx.Text
    End If
End Sub

Private Sub cmbOrtEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbOrtEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbOrtEx.ListIndex = -1
        cmbOrtEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbPersonenEx_Click()
    If cmbPersonenEx.Text <> "" Then
        If cmbPersonenEx.Text = cmbJahrEx.Text Then             'Gerbing 23.01.2008
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbSWFEx.Text Then                                          'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        If cmbPersonenEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbPersonenEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Personen die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1027 + Sprache))                              'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1027 + Sprache), cmbPersonenEx.Text)   '1027=personen, das EXIF/IPTC-Feld steht in cmbPersonenEx.Text
    End If
End Sub

Private Sub cmbPersonenEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbPersonenEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbPersonenEx.ListIndex = -1
        cmbPersonenEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub


Private Sub cmbSituationEx_Click()
    If cmbSituationEx.Text <> "" Then
        If cmbSituationEx.Text = cmbJahrEx.Text Then                'Gerbing 23.01.2008
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbPersonenEx.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbSWFEx.Text Then                                         'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        If cmbSituationEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbSituationEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll Situation die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1024 + Sprache))                                  'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1024 + Sprache), cmbSituationEx.Text)   '1024=situation, das EXIF/IPTC-Feld steht in cmbSituationEx.Text
    End If
End Sub

Private Sub cmbSituationEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbSituationEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbSituationEx.ListIndex = -1
        cmbSituationEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbjahrEx_KeyPress(KeyAscii As Integer)                 'Gerbing 23.01.2008
    Dim RetteText As String                                         'Gerbing 16.11.2015
    
    RetteText = cmbJahrEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbJahrEx.ListIndex = -1
        cmbJahrEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub cmbSWFEx_Click()                                                               'Gerbing 04.01.2009
    If cmbSWFEx.Text <> "" Then
        If cmbSWFEx.Text = cmbJahrEx.Text Then             'Gerbing 23.01.2008
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbSituationEx.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbOrtEx.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbLandEx.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbPersonenEx.Text Then                                          'Gerbing 04.01.2009
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbKommentarEx.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbEx1.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbEx2.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbEx3.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbEx4.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        If cmbSWFEx.Text = cmbEx5.Text Then
            Call ErrorDoppeltesZiel(cmbSWFEx)
            Exit Sub
        End If
        'Für welches Ziel-IPTC- oder EXIF-Feld soll SWF die Quelle sein
        Call QuelleZuZielZurückNehmen(LoadResString(1029 + Sprache))                               'Gerbing 16.11.2015
        Call QuelleZuZielZuordnen(LoadResString(1029 + Sprache), cmbSWFEx.Text)   '1029=SWF, das EXIF/IPTC-Feld steht in cmbSWFEx.Text
    End If
End Sub

Private Sub cmbSWFEx_KeyPress(KeyAscii As Integer)
    Dim RetteText As String
    
    RetteText = cmbSWFEx.Text
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        cmbSWFEx.ListIndex = -1
        cmbSWFEx.ListIndex = -1
        Call QuelleZuZielKeyPress(RetteText)
    End If
End Sub

Private Sub Combo1_GotFocus()                                       'Gerbing 02.10.2019
    Call cmbFeld1_LostFocus
End Sub

Private Sub Combo2_GotFocus()                                       'Gerbing 02.10.2019
    Call cmbFeld2_LostFocus
End Sub

Private Sub Combo3_GotFocus()                                       'Gerbing 02.10.2019
    Call cmbFeld3_LostFocus
End Sub

Private Sub Combo4_GotFocus()                                       'Gerbing 02.10.2019
    Call cmbFeld4_LostFocus
End Sub

Private Sub Combo5_GotFocus()                                       'Gerbing 02.10.2019
    Call cmbFeld5_LostFocus
End Sub

Private Sub Form_Load()
    Dim Feldname As String
    Dim Gefunden As Boolean
    Dim n As Long
    Dim SQL As String
    Dim rc As Long
    Dim Msg As String
    Dim antwort As Long
    Dim rstDefFields As ADODB.Recordset                             'Gerbing 03.08.2016

    Form1.FehlerGefunden = False                                    'Gerbing 14.10.2014
    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    NL = vbNewLine
    Me.Caption = LoadResString(1527 + Sprache)
    txtExifToolOutput.Visible = False                               'Gerbing 21.11.2016
    btnStart.Caption = LoadResString(3101 + Sprache)        'S&tart
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        '&Abbrechen
    FrameStandardWerte.Caption = LoadResString(1543 + Sprache)  'Standard-Felder aus der Datenbank
    Label1.Caption = LoadResString(1023 + Sprache) & ":"        'Jahr
    Label2.Caption = LoadResString(1550 + Sprache)              'Quell-Feld
    Label3.Caption = LoadResString(1024 + Sprache) & ":"      'Situation:
    Label4.Caption = LoadResString(1025 + Sprache) & ":"       'Ort:
    Label5.Caption = LoadResString(1026 + Sprache) & ":"       'Land:
    Label6.Caption = LoadResString(1027 + Sprache) & ":"       'Personen:
    Label7.Caption = LoadResString(1030 + Sprache) & ":"        'Kommentar:
    chkDatumNichtAktualisieren.Caption = LoadResString(1545 + Sprache) 'Datei-Datum (Geändert am) nicht aktualisieren
    chkSetIPTCPresent.Caption = LoadResString(1537 + Sprache)           'IPTCPresent auf True setzen                'Gerbing 16.11.2015
    FrameNutzerDefiniert.Caption = LoadResString(1544 + Sprache) 'Nutzerdefinierte Felder aus der Datenbank
    FrameNutzerDefiniert.Visible = True
    'ob Professional Version wird bei SQL Server Version nicht überprüft
    If gblnSQLServerVersion = False Then
        If gblnProversion = False Then
            FrameNutzerDefiniert.Visible = False
        Else
            FrameNutzerDefiniert.Visible = True
        End If
    End If
    lblFeldname1.Caption = LoadResString(1033 + Sprache)        'Feldname
    lblFeldname2.Caption = LoadResString(1033 + Sprache)        'Feldname
    lblFeldname3.Caption = LoadResString(1033 + Sprache)        'Feldname
    lblFeldname4.Caption = LoadResString(1033 + Sprache)        'Feldname
    lblFeldname5.Caption = LoadResString(1033 + Sprache)        'Feldname
    Label17.Caption = LoadResString(1014 + Sprache)             'Arbeitsfortschritt:
'--------------------------------------------------------------------------------------------
'    FrameNutzerDefiniert.ToolTipText = LoadResString(1433 + Sprache)    'Irrtümlich eingetragene Feldnamen entfernen Sie mit der Return-Taste oder Entf-Taste im numerischen Tastenfeld
    lblExifIptc1.ToolTipText = LoadResString(1549 + Sprache)        'EXIF/IPTC-Felder, die Sie nicht auswählen, bleiben falls sie bisher existieren unverändert in der JPG-Datei erhalten
    chkDatumNichtAktualisieren.ToolTipText = LoadResString(2556 + Sprache) 'sonst bekommen die Fotos das heutige Datum
    chkSetIPTCPresent.ToolTipText = LoadResString(2557 + Sprache) 'IPTCPresent=true ist das Kennzeichen, dass in dieses Foto schon einmal exportiert worden ist. Es wird ignoriert bei Export in die ganze Datenbank oder für ein einzelnes Jahr
    Label2.ToolTipText = LoadResString(2558 + Sprache)          'Der Feld-Inhalt darf für jedes Foto anders sein. Hier sehen Sie nur Beispiels-Inhalte
    Label12.ToolTipText = LoadResString(2558 + Sprache)         'Der Feld-Inhalt darf für jedes Foto anders sein. Hier sehen Sie nur Beispiels-Inhalte
    FrameStandardWerte.ToolTipText = LoadResString(2564 + Sprache)      'Mit Rechtsklick können Sie die Feld-Zuordnung zurücksetzen 'Gerbing 12.10.2019
    FrameNutzerDefiniert.ToolTipText = LoadResString(2564 + Sprache)    'Mit Rechtsklick können Sie die Feld-Zuordnung zurücksetzen 'Gerbing 12.10.2019
'--------------------------------------------------------------------------------------------
    btnHilfe.Caption = LoadResString(3014 + Sprache)                '&Hilfe
    lblExifIptc1.Caption = LoadResString(1530 + Sprache)            'Ziel EXIF/IPTC-Feld
    lblExifIptc2.Caption = LoadResString(1530 + Sprache)            'Ziel EXIF/IPTC-Feld
    lblFeldinhalt1.Caption = LoadResString(1480 + Sprache)          'Feldinhalts-Beispiele
    lblFeldinhalt2.Caption = LoadResString(1480 + Sprache)          'Feldinhalts-Beispiele          'Gerbing 02.10.2019
    lblSchonVerarbeitet.Caption = LoadResString(1546 + Sprache)     'Fotos schon übertragen:
    lblNochZuVerarb.Caption = LoadResString(1547 + Sprache)         'Fotos noch zu übertragen
    optInAlle.Caption = LoadResString(2271 + Sprache)               'Für die ganze Datenbank
    OptEinzelnesJahr.Caption = LoadResString(2287 + Sprache)        'Für ein einzelnes Jahr
    optNurFalse.Caption = LoadResString(2272 + Sprache)             'Für IPTCPresent=False
'------------------------------------------------------------------------------------------
    'jetzt wird untersucht, ob es nutzerdefinierte Felder gibt
    'die haben dann andere Feldnamen als die Namen der Standardfelder
    Screen.MousePointer = vbHourglass                                                               'Gerbing 14.10.2014
    SQL = "select * from fotos"
    On Error Resume Next
    Form1.rstsql.Close
    On Error GoTo 0
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    For n = 0 To Form1.rstsql.Fields.Count - 1
        Feldname = Form1.rstsql.Fields(n).Name
        Feldname = LCase(Feldname)
        Select Case Feldname
            'AudioFileExists gehört nicht zu den nutzerdefinierten Feldern      'Gerbing 14.05.2006
            'IPTCPresent gehört nicht zu den nutzerdefinierten Feldern          'Gerbing 04.02.2008
            'Case "merker", "jahr", "situation", "ort", "land", "personen", "dateiname", "swf", "kommentar", "dateinamekurz", "ddatum", "breitepixel", "hoehepixel", "audiofileexists", "iptcpresent"
            Case LCase(LoadResString(2524 + Sprache)), LCase(LoadResString(1023 + Sprache)), LCase(LoadResString(1024 + Sprache)), LCase(LoadResString(1025 + Sprache)), LCase(LoadResString(1026 + Sprache)), LCase(LoadResString(1027 + Sprache)), _
                LCase(LoadResString(1028 + Sprache)), LCase(LoadResString(1029 + Sprache)), LCase(LoadResString(1030 + Sprache)), LCase(LoadResString(1031 + Sprache)), LCase(LoadResString(1032 + Sprache)), LCase(LoadResString(1106 + Sprache)), LCase(LoadResString(1107 + Sprache)), "audiofileexists", "iptcpresent"
                Gefunden = True
            Case Else
                Gefunden = False
        End Select
        If Gefunden = False Then
            cmbFeld1.ComboItems.Add Form1.rstsql.Fields(n).Name
            cmbFeld2.ComboItems.Add Form1.rstsql.Fields(n).Name
            cmbFeld3.ComboItems.Add Form1.rstsql.Fields(n).Name
            cmbFeld4.ComboItems.Add Form1.rstsql.Fields(n).Name
            cmbFeld5.ComboItems.Add Form1.rstsql.Fields(n).Name
        End If
    Next n
    Form1.rstsql.Close
    '-----------------------------------------------------------------------------------------
    'für alle Felder Jahr Situation Ort Land Personen
    'wird eine Combobox mit den schon vorhandenen Werten angeboten  Gerbing 10.06.2005
    'SQL = "SELECT DISTINCT Fotos.Jahr From Fotos ORDER BY Jahr;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1023 + Sprache) & " From Fotos ORDER BY " & LoadResString(1023 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(form1.rstsql.Fields("Jahr")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1023 + Sprache))) Then
            'cmbSituation.AddItem form1.rstsql.Fields("Jahr")
            cmbJahr.AddItem Form1.rstsql.Fields(LoadResString(1023 + Sprache))
            cmbEinzelnesJahr.AddItem Form1.rstsql.Fields(LoadResString(1023 + Sprache))             'Gerbing 16.11.2015
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    'SQL = "SELECT DISTINCT Fotos.Situation From Fotos WHERE ((Not (Fotos.Situation)='')) ORDER BY Situation;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1024 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1024 + Sprache) & ")='')) ORDER BY " & LoadResString(1024 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("Situation")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1024 + Sprache))) Then
            'cmbSituation.AddItem Form1.rstsql.Fields("Situation")
            cmbSituation.ComboItems.Add Form1.rstsql.Fields(LoadResString(1024 + Sprache))
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    'SQL = "SELECT DISTINCT Fotos.Ort From Fotos WHERE ((Not (Fotos.Ort)='')) ORDER BY Ort;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1025 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1025 + Sprache) & ")='')) ORDER BY " & LoadResString(1025 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("Ort")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1025 + Sprache))) Then
            'cmbOrt.AddItem Form1.rstsql.Fields("Ort")
            cmbOrt.ComboItems.Add Form1.rstsql.Fields(LoadResString(1025 + Sprache))
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    'SQL = "SELECT DISTINCT Fotos.Land From Fotos WHERE ((Not (Fotos.Land)='')) ORDER BY Land;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1026 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1026 + Sprache) & ")='')) ORDER BY " & LoadResString(1026 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("Land")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1026 + Sprache))) Then
            'cmbLand.AddItem Form1.rstsql.Fields("Land")
            cmbLand.ComboItems.Add Form1.rstsql.Fields(LoadResString(1026 + Sprache))
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    'SQL = "SELECT DISTINCT Fotos.Personen From Fotos WHERE ((Not (Fotos.Personen)='')) ORDER BY Personen;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1027 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1027 + Sprache) & ")='')) ORDER BY " & LoadResString(1027 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("Personen")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1027 + Sprache))) Then
            'cmbPersonen.AddItem Form1.rstsql.Fields("Personen")
            cmbPersonen.ComboItems.Add Form1.rstsql.Fields(LoadResString(1027 + Sprache))
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    'SQL = "SELECT DISTINCT Fotos.SWF From Fotos WHERE ((Not (Fotos.SWF)='')) ORDER BY SWF;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1029 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1029 + Sprache) & ")='')) ORDER BY " & LoadResString(1029 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("SWF")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1029 + Sprache))) Then
            'cmbPersonen.AddItem Form1.rstsql.Fields("SWF")
            cmbSWF.ComboItems.Add Form1.rstsql.Fields(LoadResString(1029 + Sprache))
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    '--------------------------------
    'alle Comboboxen füllen, die auswählbare EXIF/IPTC-Felder anzeigen sollen
    Call AddExifFelder(cmbJahrEx)                       'Gerbing 23.01.2008
    Call AddExifFelder(cmbSituationEx)
    Call AddExifFelder(cmbOrtEx)
    Call AddExifFelder(cmbLandEx)
    Call AddExifFelder(cmbPersonenEx)
    Call AddExifFelder(cmbSWFEx)                                                            'Gerbing 04.01.2009
    Call AddExifFelder(cmbKommentarEx)
    Call AddExifFelder(cmbEx1)
    Call AddExifFelder(cmbEx2)
    Call AddExifFelder(cmbEx3)
    Call AddExifFelder(cmbEx4)
    Call AddExifFelder(cmbEx5)
    'für die Standarddatenbankfelder gibt es einige EXIF/IPTC-Felder, die ich voreinstelle (Eigenbedarf)
    cmbSituationEx.ListIndex = 11 + 5 + 2       '11=Headline        '5 addieren Gerbing 16.11.2015 2 addieren Gerbing 02.10.2019
    cmbOrtEx.ListIndex = 5 + 5 + 2              '5=City             '5 addieren Gerbing 16.11.2015 2 addieren Gerbing 02.10.2019
    cmbLandEx.ListIndex = 7 + 5 + 2             '7=Country          '5 addieren Gerbing 16.11.2015 2 addieren Gerbing 02.10.2019
    cmbKommentarEx.ListIndex = 3 + 5 + 2        '3=Caption-Abstract '5 addieren Gerbing 16.11.2015 2 addieren Gerbing 02.10.2019
    '----------------------------------------------------------------------------------------------------------
    'Tabelle DefaultFields auswerten                                                        'Gerbing 03.08.2016
    'wenn in der Tabelle DefaultFields etwas steht,                                         'Gerbing 03.08.2016
    'dann überschreibe ich die voreingestellten EXIF/IPTC-Felder
    On Error Resume Next
    SQL = "select * From DefaultFields;"
    Set rstDefFields = New ADODB.Recordset
    With rstDefFields
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number = 0 Then
        'hier existiert die Tabelle DefaultFields
        If Not rstDefFields.EOF Then
            'hier ist die Tabelle DefaultFields nicht leer
            blnDefaultFieldsNotEmpty = True                                                 'Gerbing 27.10.2016
            cmbSituationEx.Text = rstDefFields.Fields("SituationSource")
            cmbOrtEx.Text = rstDefFields.Fields("LocationSource")
            cmbLandEx.Text = rstDefFields.Fields("CountrySource")
            cmbPersonenEx.Text = rstDefFields.Fields("PeopleSource")
            cmbSWFEx.Text = rstDefFields.Fields("BWCSource")
            cmbKommentarEx.Text = rstDefFields.Fields("CommentSource")
        End If
    End If
    rstDefFields.Close                                                                      'Gerbing 02.10.2019
    'Tabelle DefaultFields auswerten End                                                    'Gerbing 03.08.2016
    '----------------------------------------------------------------------------------------------------------
    Call UserdefinedAuswerten                                                               'Gerbing 12.10.2019
    'Tabelle UserDefined auswerten End
    '----------------------------------------------------------------------------------------------------------
    If cmbFeld1.ComboItems.Count = 0 Then
        FrameNutzerDefiniert.Visible = False
        Me.Height = 8000
    End If
    Me.top = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbFeld1_LostFocus()
    Dim SQL As String
    
    If cmbFeld1.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
    Combo1.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    'Alle Werte nach Combo1 stellen die zu diesem Feld in der Datenbank gefunden werden
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld1.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld1.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld1.Text & "];"
    On Error Resume Next
    Form1.rstsql.Close
    'On Error GoTo 0                                        'Gerbing 02.10.2019
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld1.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            Combo1.ComboItems.Add Form1.rstsql.Fields(0)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    If Combo1.ComboItems.Count <> 0 Then
        'Combo1.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld2_LostFocus()
    Dim SQL As String
    
    If cmbFeld2.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
    Combo2.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    'Alle Werte nach Combo2 stellen die zu diesem Feld in der Datenbank gefunden werden
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld2.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld2.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld2.Text & "];"
    On Error Resume Next
    Form1.rstsql.Close
    'On Error GoTo 0                                        'Gerbing 02.10.2019
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld2.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            Combo2.ComboItems.Add Form1.rstsql.Fields(0)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    If Combo2.ComboItems.Count <> 0 Then
        'Combo2.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld3_LostFocus()
    Dim SQL As String
    
    If cmbFeld3.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
    Combo3.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    'Alle Werte nach Combo3 stellen die zu diesem Feld in der Datenbank gefunden werden
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld3.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld3.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld3.Text & "];"
    On Error Resume Next
    Form1.rstsql.Close
    'On Error GoTo 0                                        'Gerbing 02.10.2019
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld3.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            Combo3.ComboItems.Add Form1.rstsql.Fields(0)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    If Combo3.ComboItems.Count <> 0 Then
        'Combo3.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld4_LostFocus()
    Dim SQL As String
    
    If cmbFeld4.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
    Combo4.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    'Alle Werte nach Combo4 stellen die zu diesem Feld in der Datenbank gefunden werden
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld4.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld4.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld4.Text & "];"
    On Error Resume Next
    Form1.rstsql.Close
    'On Error GoTo 0                                        'Gerbing 02.10.2019
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld4.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            Combo4.ComboItems.Add Form1.rstsql.Fields(0)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    If Combo4.ComboItems.Count <> 0 Then
        'Combo4.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld5_LostFocus()
    Dim SQL As String
    
    If cmbFeld5.Text = "" Then Exit Sub                          'Gerbing 10.06.2005
    Combo5.ComboItems.RemoveAll
    'Alle Werte nach Combo5 stellen die zu diesem Feld in der Datenbank gefunden werden
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld5.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld5.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld5.Text & "];"
    On Error Resume Next
    Form1.rstsql.Close
    'On Error GoTo 0                                        'Gerbing 02.10.2019
    With Form1.rstsql
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld5.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            Combo5.ComboItems.Add Form1.rstsql.Fields(0)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close
    If Combo5.ComboItems.Count <> 0 Then
        'Combo5.ListIndex = 0
    End If
End Sub

Private Sub AddExifFelder(Control)
    'Comboboxen füllen mit auswählbaren EXIF/IPTC-Feldern
    Control.AddItem "IPTC-Object Name"
    Control.AddItem "IPTC-Byline"
    Control.AddItem "IPTC-Byline title"
    Control.AddItem "IPTC-Caption-Abstract"
    Control.AddItem "IPTC-Caption writer"
    Control.AddItem "IPTC-Copyright notice"
    Control.AddItem "IPTC-Special instructions"
    Control.AddItem "IPTC-Urgency"
    Control.AddItem "IPTC-Date created"
    Control.AddItem "IPTC-Time created"
    Control.AddItem "IPTC-City"
    Control.AddItem "IPTC-Province/State"
    Control.AddItem "IPTC-Country"
    Control.AddItem "IPTC-Source"
    Control.AddItem "IPTC-Headline"
    Control.AddItem "IPTC-Transmission reference"
    Control.AddItem "IPTC-Category"
    Control.AddItem "IPTC-Supplemental categories"
    Control.AddItem "IPTC-Keywords"
    Control.AddItem "IPTC-Credits"                          'Gerbing 10.01.2008
    Control.AddItem "IPTC-Originating Program"              'Gerbing 10.01.2008
    '-----------------------------------------
    'jetzt die restlichen 8 EXIF/IPTC-Felder                     'Gerbing 04.03.2013
    '-----------------------------------------
    Control.AddItem "IPTC-Release date"
    Control.AddItem "IPTC-Release time"
    Control.AddItem "IPTC-Object cycle"
    Control.AddItem "IPTC-Location code"
    Control.AddItem "IPTC-Sublocation"
    Control.AddItem "IPTC-Program version"
    Control.AddItem "IPTC-Edit status"
    Control.AddItem "IPTC-JobID"
    '-----------------------------------------
    'jetzt die 5 EXIF-Felder                                'Gerbing 16.11.2015
    '-----------------------------------------
    Control.AddItem "EXIF-XPTitle"
    Control.AddItem "EXIF-XPSubject"
    Control.AddItem "EXIF-XPKeywords"
    Control.AddItem "EXIF-XPComment"
    Control.AddItem "EXIF-XPAuthor"
    '-----------------------------------------
    'jetzt die 5 EXIF-GPS-Felder                            'Gerbing 02.10.2019
    '-----------------------------------------
    Control.AddItem "EXIF-GPSLatitude"
    Control.AddItem "EXIF-GPSLongitude"
End Sub

Private Function LängeAusrechnen(ss, pos) As Long
    'es sind generell 2 Bytes lange Felder
    Dim strLen1 As String           'Byte1
    Dim strlen2 As String           'Byte2
    Dim lngLen1 As Long
    Dim lnglen2 As Long
    
    strLen1 = Mid(ss, pos, 1)
    strlen2 = Mid(ss, pos + 1, 1)
    lngLen1 = Asc(strLen1)
    lnglen2 = Asc(strlen2)
    lngLen1 = lngLen1 * 256         'das erste Byte mit 256 multiplizieren
    lngLen1 = lngLen1 + lnglen2
    LängeAusrechnen = lngLen1
End Function

Private Sub ErrorDoppeltesZiel(Control)
    'MsgBox "Sie dürfen ein Ziel EXIF/IPTC-Feld nicht mehrfach angeben"
    MsgBox LoadResString(1532 + Sprache)
    Control.ListIndex = -1
End Sub

Private Sub QuelleZuZielZurückNehmen(Quellfeld)
    'Quellfeld enthält zB 'situation'
    
    If QuellFeldFürObjectName = Quellfeld Then QuellFeldFürObjectName = ""
    If QuellFeldFürByline = Quellfeld Then QuellFeldFürByline = ""
    If QuellFeldFürBylineTitle = Quellfeld Then QuellFeldFürBylineTitle = ""
    If QuellFeldFürCaption = Quellfeld Then QuellFeldFürCaption = ""


    If QuellFeldFürCaptionWriter = Quellfeld Then QuellFeldFürCaptionWriter = ""
    If QuellFeldFürCopyright = Quellfeld Then QuellFeldFürCopyright = ""
    If QuellFeldFürSpecialInstructions = Quellfeld Then QuellFeldFürSpecialInstructions = ""
    If QuellFeldFürUrgency = Quellfeld Then QuellFeldFürUrgency = ""
    If QuellFeldFürDateCreated = Quellfeld Then QuellFeldFürDateCreated = ""
    If QuellFeldFürTimeCreated = Quellfeld Then QuellFeldFürTimeCreated = ""
    If QuellFeldFürCity = Quellfeld Then QuellFeldFürCity = ""
    If QuellFeldFürState = Quellfeld Then QuellFeldFürState = ""
    If QuellFeldFürCountry = Quellfeld Then QuellFeldFürCountry = ""
    If QuellFeldFürSource = Quellfeld Then QuellFeldFürSource = ""
    If QuellFeldFürHeadline = Quellfeld Then QuellFeldFürHeadline = ""
    If QuellFeldFürOriginalTransmissionReference = Quellfeld Then QuellFeldFürOriginalTransmissionReference = ""
    If QuellFeldFürCategory = Quellfeld Then QuellFeldFürCategory = ""
    If QuellFeldFürSupplementalCategories = Quellfeld Then QuellFeldFürSupplementalCategories = ""
    If QuellFeldFürKeywords = Quellfeld Then QuellFeldFürKeywords = ""
    If QuellFeldFürCredits = Quellfeld Then QuellFeldFürCredits = ""
    If QuellFeldFürOriginatingProgram = Quellfeld Then QuellFeldFürOriginatingProgram = ""
'    '---------------------------------------------------------------------------------------
'    'jetzt die restlichen 8
'    '---------------------------------------------------------------------------------------
    If QuellFeldFürReleaseDate = Quellfeld Then QuellFeldFürReleaseDate = ""
    If QuellFeldFürReleaseTime = Quellfeld Then QuellFeldFürReleaseTime = ""
    If QuellFeldFürObjectcycle = Quellfeld Then QuellFeldFürObjectcycle = ""
    If QuellFeldFürLocationCode = Quellfeld Then QuellFeldFürLocationCode = ""
    If QuellFeldFürSubLocation = Quellfeld Then QuellFeldFürSubLocation = ""
    If QuellFeldFürProgramVersion = Quellfeld Then QuellFeldFürProgramVersion = ""
    If QuellFeldFürEditStatus = Quellfeld Then QuellFeldFürEditStatus = ""
    If QuellFeldFürJobID = Quellfeld Then QuellFeldFürJobID = ""
'    '---------------------------------------------------------------------------------------
'    'jetzt die 5 EXIF-Felder                                                                'Gerbing 16.11.2015
'    '---------------------------------------------------------------------------------------
    If QuellFeldFürXPTitle = Quellfeld Then QuellFeldFürXPTitle = ""
    If QuellFeldFürXPSubject = Quellfeld Then QuellFeldFürXPSubject = ""
    If QuellFeldFürXPKeywords = Quellfeld Then QuellFeldFürXPKeywords = ""
    If QuellFeldFürXPComment = Quellfeld Then QuellFeldFürXPComment = ""
    If QuellFeldFürXPAuthor = Quellfeld Then QuellFeldFürXPAuthor = ""
'    '---------------------------------------------------------------------------------------
'    'jetzt die 2 EXIF-GPS-Felder                                                           'Gerbing 02.10.2019
'    '---------------------------------------------------------------------------------------
    If QuellFeldFürGPSLatitude = Quellfeld Then QuellFeldFürGPSLatitude = ""
    If QuellFeldFürGPSLongitude = Quellfeld Then QuellFeldFürGPSLongitude = ""
End Sub

Private Sub QuelleZuZielZuordnen(Quellfeld, Zielfeld)
    'Quellfeld zB 'situation'  Zielfeld zB 'IPTC-Caption'
    
    If Zielfeld = "IPTC-Object Name" Then QuellFeldFürObjectName = Quellfeld
    If Zielfeld = "IPTC-Byline" Then QuellFeldFürByline = Quellfeld
    If Zielfeld = "IPTC-Byline title" Then QuellFeldFürBylineTitle = Quellfeld
    If Zielfeld = "IPTC-Caption-Abstract" Then QuellFeldFürCaption = Quellfeld
    If Zielfeld = "IPTC-Caption writer" Then QuellFeldFürCaptionWriter = Quellfeld
    If Zielfeld = "IPTC-Copyright notice" Then QuellFeldFürCopyright = Quellfeld
    If Zielfeld = "IPTC-Special instructions" Then QuellFeldFürSpecialInstructions = Quellfeld
    If Zielfeld = "IPTC-Urgency" Then QuellFeldFürUrgency = Quellfeld
    If Zielfeld = "IPTC-Date created" Then QuellFeldFürDateCreated = Quellfeld
    If Zielfeld = "IPTC-Time created" Then QuellFeldFürTimeCreated = Quellfeld
    If Zielfeld = "IPTC-City" Then QuellFeldFürCity = Quellfeld
    If Zielfeld = "IPTC-Province/State" Then QuellFeldFürState = Quellfeld
    If Zielfeld = "IPTC-Country" Then QuellFeldFürCountry = Quellfeld
    If Zielfeld = "IPTC-Source" Then QuellFeldFürSource = Quellfeld
    If Zielfeld = "IPTC-Headline" Then QuellFeldFürHeadline = Quellfeld
    If Zielfeld = "IPTC-Transmission reference" Then QuellFeldFürOriginalTransmissionReference = Quellfeld
    If Zielfeld = "IPTC-Category" Then QuellFeldFürCategory = Quellfeld
    If Zielfeld = "IPTC-Supplemental categories" Then QuellFeldFürSupplementalCategories = Quellfeld
    If Zielfeld = "IPTC-Keywords" Then QuellFeldFürKeywords = Quellfeld
    If Zielfeld = "IPTC-Credits" Then QuellFeldFürCredits = Quellfeld
    If Zielfeld = "IPTC-Originating Program" Then QuellFeldFürOriginatingProgram = Quellfeld
    '---------------------------------------------------------------------------------------
    'jetzt die restlichen 8
    '---------------------------------------------------------------------------------------
    If Zielfeld = "IPTC-Release date" Then QuellFeldFürReleaseDate = Quellfeld
    If Zielfeld = "IPTC-Release time" Then QuellFeldFürReleaseTime = Quellfeld
    If Zielfeld = "IPTC-Object cycle" Then QuellFeldFürObjectcycle = Quellfeld
    If Zielfeld = "IPTC-Location code" Then QuellFeldFürLocationCode = Quellfeld
    If Zielfeld = "IPTC-Sublocation" Then QuellFeldFürSubLocation = Quellfeld
    If Zielfeld = "IPTC-Program version" Then QuellFeldFürProgramVersion = Quellfeld
    If Zielfeld = "IPTC-Edit status" Then QuellFeldFürEditStatus = Quellfeld
    If Zielfeld = "IPTC-JobID" Then QuellFeldFürJobID = Quellfeld
    '---------------------------------------------------------------------------------------
    'jetzt die 5 EXIF-Felder                                                                'Gerbing 16.11.2015
    '---------------------------------------------------------------------------------------
    If Zielfeld = "EXIF-XPTitle" Then QuellFeldFürXPTitle = Quellfeld
    If Zielfeld = "EXIF-XPSubject" Then QuellFeldFürXPSubject = Quellfeld
    If Zielfeld = "EXIF-XPKeywords" Then QuellFeldFürXPKeywords = Quellfeld
    If Zielfeld = "EXIF-XPComment" Then QuellFeldFürXPComment = Quellfeld
    If Zielfeld = "EXIF-XPAuthor" Then QuellFeldFürXPAuthor = Quellfeld
    '---------------------------------------------------------------------------------------
    'jetzt die 5 EXIF-GPS-Felder                                                            'Gerbing 02.10.2019
    '---------------------------------------------------------------------------------------
'    If Zielfeld = "EXIF-GPSLatitude" Then QuellFeldFürGPSLatitude = Quellfeld
'    If Zielfeld = "EXIF-GPSLongitude" Then QuellFeldFürGPSLongitude = Quellfeld
    If Zielfeld = "EXIF-GPSLatitude" Then QuellFeldFürGPSLatitude = "GPSLatitude"           'Gerbing 03.01.2020
    If Zielfeld = "EXIF-GPSLongitude" Then QuellFeldFürGPSLongitude = "GPSLongitude"        'Gerbing 03.01.2020

End Sub

Private Sub QuelleZuZielKeyPress(Zielfeld)
    'Zielfeld enthält zB 'IPTC-Caption
    
    If Zielfeld = "IPTC-Object Name" Then QuellFeldFürObjectName = ""
    If Zielfeld = "IPTC-Byline" Then QuellFeldFürByline = ""
    If Zielfeld = "IPTC-Byline title" Then QuellFeldFürBylineTitle = ""
    If Zielfeld = "IPTC-Caption" Then QuellFeldFürCaption = ""
    If Zielfeld = "IPTC-Caption writer" Then QuellFeldFürCaptionWriter = ""
    If Zielfeld = "IPTC-Copyright notice" Then QuellFeldFürCopyright = ""
    If Zielfeld = "IPTC-Special instructions" Then QuellFeldFürSpecialInstructions = ""
    If Zielfeld = "IPTC-Urgency" Then QuellFeldFürUrgency = ""
    If Zielfeld = "IPTC-Date created" Then QuellFeldFürDateCreated = ""
    If Zielfeld = "IPTC-Time created" Then QuellFeldFürTimeCreated = ""
    If Zielfeld = "IPTC-City" Then QuellFeldFürCity = ""
    If Zielfeld = "IPTC-Province/State" Then QuellFeldFürState = ""
    If Zielfeld = "IPTC-Country" Then QuellFeldFürCountry = ""
    If Zielfeld = "IPTC-Source" Then QuellFeldFürSource = ""
    If Zielfeld = "IPTC-Headline" Then QuellFeldFürHeadline = ""
    If Zielfeld = "IPTC-Transmission reference" Then QuellFeldFürOriginalTransmissionReference = ""
    If Zielfeld = "IPTC-Category" Then QuellFeldFürCategory = ""
    If Zielfeld = "IPTC-Supplemental categories" Then QuellFeldFürSupplementalCategories = ""
    If Zielfeld = "IPTC-Keywords" Then QuellFeldFürKeywords = ""
    If Zielfeld = "IPTC-Credits" Then QuellFeldFürCredits = ""
    If Zielfeld = "IPTC-Originating Program" Then QuellFeldFürOriginatingProgram = ""
    '---------------------------------------------------------------------------------------
    'jetzt die restlichen 8
    '---------------------------------------------------------------------------------------
    If Zielfeld = "IPTC-Release date" Then QuellFeldFürReleaseDate = ""
    If Zielfeld = "IPTC-Release time" Then QuellFeldFürReleaseTime = ""
    If Zielfeld = "IPTC-Object cycle" Then QuellFeldFürObjectcycle = ""
    If Zielfeld = "IPTC-Location code" Then QuellFeldFürLocationCode = ""
    If Zielfeld = "IPTC-Sublocation" Then QuellFeldFürSubLocation = ""
    If Zielfeld = "IPTC-Program version" Then QuellFeldFürProgramVersion = ""
    If Zielfeld = "IPTC-Edit status" Then QuellFeldFürEditStatus = ""
    If Zielfeld = "IPTC-JobID" Then QuellFeldFürJobID = ""
    '---------------------------------------------------------------------------------------
    'jetzt die 5 EXIF-Felder                                                                'Gerbing 16.11.2015
    '---------------------------------------------------------------------------------------
    If Zielfeld = "EXIF-XPTitle" Then QuellFeldFürXPTitle = ""
    If Zielfeld = "EXIF-XPSubject" Then QuellFeldFürXPSubject = ""
    If Zielfeld = "EXIF-XPKeywords" Then QuellFeldFürXPKeywords = ""
    If Zielfeld = "EXIF-XPComment" Then QuellFeldFürXPComment = ""
    If Zielfeld = "EXIF-XPAuthor" Then QuellFeldFürXPAuthor = ""
    '---------------------------------------------------------------------------------------
    'jetzt die 5 EXIF-Felder                                                                'Gerbing 16.11.2015
    '---------------------------------------------------------------------------------------
    If Zielfeld = "EXIF-GPSLatitude" Then QuellFeldFürGPSLatitude = ""
    If Zielfeld = "EXIF-GPSLongitude" Then QuellFeldFürGPSLongitude = ""
End Sub



Private Sub SetOriginalDateTime(Dateiname As String)
    udtSystemTime.wYear = Year(MeinDatum)
    udtSystemTime.wMonth = Month(MeinDatum)
    udtSystemTime.wDay = Day(MeinDatum)
    udtSystemTime.wDayOfWeek = Weekday(MeinDatum) - 1
    udtSystemTime.wHour = Hour(MeinDatum)
    udtSystemTime.wMinute = Minute(MeinDatum)
    udtSystemTime.wSecond = Second(MeinDatum)
    udtSystemTime.wMilliseconds = 0

    ' convert system time to local time
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    ' convert local time to GMT
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    ' open the file to get the filehandle
    'Gerbing 10.09.2013
    lngHandle = CreateFileW(StrPtr(Dateiname), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)  'Gerbing 10.09.2013
    ' change date/time property of the file
    SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
    ' close the handle
    CloseHandle lngHandle
End Sub


Private Function SchreibenExifIptc(ArgOld As String, rst As ADODB.Recordset, Dateiname As String) As String               'Gerbing 16.11.2015
    'rc = 0 kein Fehler
    'rc = 1 Fehler
    Dim strExiftool As String
    Dim Msg As String
    Dim ArgNew As String
    
    If ExiftoolNichtBenutzbar = True Then
        SchreibenExifIptc = 1
        Exit Function
    End If
    
    strExiftool = AppPath & "\exiftool.exe"
    If blnExiftoolExists = True Then
        '
    Else
        If Not file_exist(strExiftool) Then
            Msg = LoadResString(2339 + Sprache) & AppPath
            'MsgBox "Das Exportieren nach IPTC- oder EXIF/XP...-Feldern erfordert exiftool.exe von Phil Harvey im Ordner " & AppPath
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
            ExiftoolNichtBenutzbar = True
            SchreibenExifIptc = 1
            SchreibenExifIptc = "no exiftool"
            Exit Function
        Else
            blnExiftoolExists = True
        End If
    End If
    'Mögliche Fehler stehen in exiftoolerrors.log
    
    ArgNew = ExifIptcSegmentNeuErzeugen(ArgOld, rst, Dateiname)
    SchreibenExifIptc = ArgNew
End Function

Private Function ExifIptcSegmentNeuErzeugen(ArgOld As String, rst As ADODB.Recordset, Dateiname As String) As String         'Gerbing 16.11.2015
     Dim rc As Long
     Dim Msg As String
     Dim ArgNew As String
       
    'Die strings zB ObjectName müssen von exiftool verstanden werden können, d.h. ich muss sie genauso bezeichnen wie exiftool
    'siehe tag name documentation for a complete list of available tag names
    'Phil Harvey internet http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/index.html
     If QuellFeldFürObjectName <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürObjectName, "ObjectName", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürByline <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürByline, "By-line", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürBylineTitle <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürBylineTitle, "By-lineTitle", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCaption <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCaption, "Caption-Abstract", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCaptionWriter <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCaptionWriter, "Writer-Editor", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCopyright <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCopyright, "CopyrightNotice", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürSpecialInstructions <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürSpecialInstructions, "SpecialInstructions", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürUrgency <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürUrgency, "Urgency", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürDateCreated <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürDateCreated, "DateCreated", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürTimeCreated <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürTimeCreated, "TimeCreated", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCity <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCity, "City", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürState <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürState, "Province-State", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCountry <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCountry, "Country-PrimaryLocationName", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürSource <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürSource, "Source", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürHeadline <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürHeadline, "Headline", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürOriginalTransmissionReference <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürOriginalTransmissionReference, "OriginalTransmissionReference", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCategory <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCategory, "Category", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürSupplementalCategories <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürSupplementalCategories, "SupplementalCategories", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürKeywords <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürKeywords, "Keywords", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürCredits <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürCredits, "Credit", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürOriginatingProgram <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürOriginatingProgram, "OriginatingProgram", Dateiname)
        ArgOld = ArgNew
     End If
    '---------------------------------------------------------------------------------------
    'jetzt die restlichen 8
    '---------------------------------------------------------------------------------------
     If QuellFeldFürReleaseDate <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürReleaseDate, "ReleaseDate", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürReleaseTime <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürReleaseTime, "ReleaseTime", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürObjectcycle <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürObjectcycle, "ObjectCycle", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürLocationCode <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürLocationCode, "Country-PrimaryLocationCode", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürSubLocation <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürSubLocation, "Sub-location", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürProgramVersion <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürProgramVersion, "ProgramVersion", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürEditStatus <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürEditStatus, "EditStatus", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürJobID <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürJobID, "JobID", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürXPTitle <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürXPTitle, "XPTitle", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürXPSubject <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürXPSubject, "XPSubject", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürXPKeywords <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürXPKeywords, "XPKeywords", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürXPComment <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürXPComment, "XPComment", Dateiname)
        ArgOld = ArgNew
     End If
     If QuellFeldFürXPAuthor <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürXPAuthor, "XPAuthor", Dateiname)
        ArgOld = ArgNew
     End If
     '----------------------------------------------------------------------------------------
     'Gerbing 02.10.2019
     If QuellFeldFürGPSLatitude <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürGPSLatitude, "GPSLatitude", Dateiname)
        ArgOld = ArgNew
        If left(QuellFeldFürGPSLatitude, 1) = "-" Then                              '- ist Südhalbkugel
            ArgNew = EXIFEditierenGPS(ArgOld, rst, "S", "GPSLatitudeRef", Dateiname)
        Else
            ArgNew = EXIFEditierenGPS(ArgOld, rst, "N", "GPSLatitudeRef", Dateiname)
        End If
        ArgOld = ArgNew
     End If
     If QuellFeldFürGPSLongitude <> "" Then
        ArgNew = EXIFEditieren(ArgOld, rst, QuellFeldFürGPSLongitude, "GPSLongitude", Dateiname)
        ArgOld = ArgNew
        If left(QuellFeldFürGPSLongitude, 1) = "-" Then                              '- ist WestHemisphäre
            ArgNew = EXIFEditierenGPS(ArgOld, rst, "W", "GPSLongitudeRef", Dateiname)
        Else
            ArgNew = EXIFEditierenGPS(ArgOld, rst, "E", "GPSLongitudeRef", Dateiname)
        End If
        ArgOld = ArgNew
     End If
     
'    ArgNew = ArgNew & "-execute" & vbNewLine
'    ArgNew = ArgNew & "#" & vbNewLine
        
    ExifIptcSegmentNeuErzeugen = ArgNew
End Function

Private Function EXIFEditieren(ArgOld As String, rst As ADODB.Recordset, QuellF As String, ZielF As String, Dateiname As String) As String 'Gerbing 16.11.2015
    Dim Msg As String
    Dim strUniFile As String
    Dim arg As String
    'rc = 0 kein Fehler
    'rc = 1 Fehler

    'jetzt schreibe ich die für jedes foto benötigten parameter die datei argfile.txt als UTF-8 Datei
    
    'Beispiel
    '-preserve
    '-overwrite_original
    '-charset
    'FileName=UTF8
    '-charset
    'IPTC=UTF8
    '-headline=Sport
    '?????.jpg
    
    If Not IsNull(rst.Fields(QuellF)) Then
        Msg = rst.Fields(QuellF)
        If Msg <> "" Then
            If ZielF = "GPSLatitude" Or ZielF = "GPSLongitude" Then                             'Gerbing 02.10.2019
                Msg = Replace(Msg, ",", ".")
            End If
            
            arg = ""
            If chkDatumNichtAktualisieren.Value = 1 Then
                arg = arg & "-preserve" & vbNewLine       '-preserve läßt das Änderungsdatum wie es ist
            End If
            arg = arg & "-overwrite_original" & vbNewLine '-overwrite_original verhindert eine Rettekopie des Fotos mit Dateiname & '_original'
            arg = arg & "-charset" & vbNewLine            '-charset Filename=xyz ermöglicht unicode file name
            arg = arg & "FileName=UTF8" & vbNewLine
            arg = arg & "-charset" & vbNewLine            '-charset IPTC=xyz ermöglicht unicode im EXIF/IPTC-Feld
            arg = arg & "IPTC=UTF8" & vbNewLine
            'Wenn in Msg vbNewLine vorkommen müssen diese entfernt werden, sonst meldet exiftoolerrors.log zB
            'Error: File not found - 1 Kloster-Thor 51 Klostervorstadt
            'Error: File not found - 2 Johannis-Thor 52 Kloster-Mühle
            Msg = Replace(Msg, vbNewLine, "", , , vbBinaryCompare)
            arg = arg & "-" & ZielF & "=" & Msg & vbNewLine
            arg = arg & Dateiname & vbNewLine
        End If
    End If
    EXIFEditieren = ArgOld & arg

End Function

Private Function EXIFEditierenGPS(ArgOld As String, rst As ADODB.Recordset, QuellF As String, ZielF As String, Dateiname As String) As String 'Gerbing 02.10.2019
    Dim Msg As String
    Dim strUniFile As String
    Dim arg As String
    'rc = 0 kein Fehler
    'rc = 1 Fehler

    'jetzt schreibe ich die für jedes foto benötigten parameter die datei argfile.txt als UTF-8 Datei
    
    'Beispiel
    '-preserve
    '-overwrite_original
    '-charset
    'FileName=UTF8
    '-charset
    'IPTC=UTF8
    '-GPSLatitudeRef=N
    '?????.jpg
    
    Msg = QuellF
    If Msg <> "" Then
        arg = ""
        If chkDatumNichtAktualisieren.Value = 1 Then
            arg = arg & "-preserve" & vbNewLine       '-preserve läßt das Änderungsdatum wie es ist
        End If
        arg = arg & "-overwrite_original" & vbNewLine '-overwrite_original verhindert eine Rettekopie des Fotos mit Dateiname & '_original'
        arg = arg & "-charset" & vbNewLine            '-charset Filename=xyz ermöglicht unicode file name
        arg = arg & "FileName=UTF8" & vbNewLine
        arg = arg & "-charset" & vbNewLine            '-charset IPTC=xyz ermöglicht unicode im EXIF/IPTC-Feld
        arg = arg & "IPTC=UTF8" & vbNewLine
        'Wenn in Msg vbNewLine vorkommen müssen diese entfernt werden, sonst meldet exiftoolerrors.log zB
        'Error: File not found - 1 Kloster-Thor 51 Klostervorstadt
        'Error: File not found - 2 Johannis-Thor 52 Kloster-Mühle
        Msg = Replace(Msg, vbNewLine, "", , , vbBinaryCompare)
        arg = arg & "-" & ZielF & "=" & Msg & vbNewLine
        arg = arg & Dateiname & vbNewLine
    End If
    EXIFEditierenGPS = ArgOld & arg

End Function

Private Function StarteExifTool()                                            'Gerbing 16.11.2015
    'rc = 0 kein Fehler
    'rc = 1 mit Fehler
    Dim retval As Long
    Dim strBat As String
    Dim exiftoolbat As String
    
    'Ich will über auftretende Fehler informiert werden, zB wenn eins der Fotos geöffnet ist kann es nicht überschrieben werden
    'Das Erzeugen von Fehlermitteilungen in exiftoolerrors.log geht nur richtig, wenn ich dazu eine Batch-Datei aufrufe
    'diese ruft exiftool auf und übergibt die Parameterdatei argfile.txt und schreibt auftretende Fehler nach exiftoolerrors.log
    'exiftool -@ argfile.txt 2> exiftoolerrors.log
    
    Screen.MousePointer = vbDefault
    Screen.MousePointer = vbHourglass
    strBat = "exiftool -stay_open true -@ argfile.txt 2> exiftoolerrors.log"
    WriteUTF8File AppPath & "\exiftool.bat", strBat
    exiftoolbat = "exiftool.bat"

    
    retval = ExecAndCapture("cmd.exe " & "/C" & """" & "exiftool.bat" & """", txtExifToolOutput)
    If retval = 1 Then
        StarteExifTool = 1
    Else
        StarteExifTool = 0
    End If
    Screen.MousePointer = vbHourglass
End Function

Private Function WriteUTF8File(ByRef Filename As String, ByRef Contents As String)  'Gerbing 16.11.2015
    'rc=0 fehlerfrei
    'rc=1 mit Fehler
    
    'zum Schreiben der AppPath & "\exiftool.bat" und AppPath & "\argfile.txt"
    
    Dim FileNumber As Integer
    Dim hHandle As Long
    Dim BytesToWrite As Long
    Dim BytesWritten As Long
    Dim fn As Long
    Dim Msg As String
    Dim antwort As Long
    Dim Buffer() As Byte
    
    hHandle = GetFileHandle(Filename, False)   'false=write
    If hHandle <> INVALID_HANDLE_VALUE Then
        If hHandle Then
            Buffer = ConvertToUTF8(Contents)
            BytesToWrite = UBound(Buffer) + 1
            fn = WriteFile(hHandle, Buffer(0), BytesToWrite, BytesWritten, ByVal 0&)
            CloseHandle hHandle
        End If
    Else
        'Msg = "Fehler beim Zugrif auf Datei " & FileName & vbNewLine               'Gerbing 30.12.2012
        Msg = LoadResString(2029 + Sprache) & " " & Filename & vbNewLine            'Gerbing 30.12.2012
        'Msg = Msg & "Möglicherweise haben Sie keine Administratorrechte." & vbNewLine  'Gerbing 30.12.2012
        Msg = Msg & LoadResString(1556 + Sprache) & vbNewLine                       'Gerbing 30.12.2012
        Msg = Msg & "Errornumber=" & ERR.Number & vbNewLine                         'Gerbing 30.12.2012
        Msg = Msg & "Errortext=" & ERR.Description & vbNewLine & vbNewLine          'Gerbing 30.12.2012
        'Msg = Msg & "Möchten Sie EXIF/IPTC... abbrechen?"                          'Gerbing 30.10.2012
        Msg = Msg & LoadResString(1553 + Sprache)                                   'Gerbing 30.10.2012
        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)                           'Gerbing 30.10.2012
        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
        If antwort = vbYes Then                                                     'Gerbing 30.10.2012
            blnAbbrechenGeklickt = True                                             'Gerbing 30.10.2012
        End If                                                                      'Gerbing 30.10.2012
        WriteUTF8File = 1                       'Fehler
        Exit Function
    End If
    WriteUTF8File = 0                           'kein Fehler
End Function

Private Function ConvertToUTF8(ByRef Source As String) As Byte()                    'Gerbing 16.11.2015
    Dim Length As Long
    Dim Pointer As Long
    Dim Size As Long
    Dim Buffer() As Byte
    
    Length = Len(Source)
    Pointer = StrPtr(Source)
    Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
    ReDim Buffer(0 To Size - 1)
    
    WideCharToMultiByte CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), _
        Size, 0, 0
    ConvertToUTF8 = Buffer
End Function

Private Sub DeleteExiftoolFiles()                                                           'Gerbing 16.11.2015
    Dim rc As Long
    
    rc = file_delete(AppPath & "\argfile.txt", False, True) 'ohne Papierkorb, silent
    'rc = file_delete(AppPath & "\exiftoolerrors.log", False, True) 'ohne Papierkorb, silent
End Sub

Private Function RemoveNulls(OriginalString As String) As String
    Dim pos As Long
    pos = InStr(OriginalString, Chr$(0))
    If pos > 1 Then
        RemoveNulls = Mid$(OriginalString, 1, pos - 1)
    Else
        RemoveNulls = OriginalString
    End If
End Function

Public Function ExecCmd(cmdline As String)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    Dim Msg As String
    
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
    ' Start the shelled application:
    ret = CreateProcessA(vbNullString, cmdline, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
    If ret = 0 Then
        Msg = cmdline & NL
        Msg = Msg & "could not start the shelled application"
        MsgBox Msg
        ret = 99
        ExecCmd = ret
        Exit Function
    End If

    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, ret)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = ret
End Function

Private Sub Form_Unload(Cancel As Integer)
    blnAbbrechenGeklickt = True
End Sub

Private Sub FrameStandardWerte_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'Gerbing 27.10.2016
    Dim antwort As Long
    Dim rstDefFields As ADODB.Recordset                                             'Gerbing 03.08.2016
    Dim SQL As String
    
    If Button = vbRightButton Then                                                  'Gerbing 03.08.2016
        'If blnDefaultFieldsNotEmpty = True Then
            'antwort = MsgBox("Wollen Sie abbrechen und die Feld-Zuordnung zurücksetzen?", vbYesNo)
            antwort = MsgBox(LoadResString(2469 + Sprache), vbYesNo)
            If antwort = vbYes Then
                On Error Resume Next
                If gblnSQLServerVersion = True Then
                    'beim SQL Server muss es heißen 'Delete from table
                    SQL = "DELETE From DefaultFields"
                Else
                    SQL = "DELETE * FROM DefaultFields"
                End If
                Set rstDefFields = New ADODB.Recordset
                With rstDefFields
                    .Source = SQL
                    .ActiveConnection = Form1.DBsql
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                rstDefFields.Close
                On Error GoTo 0
                'Wiederherstellen Standard-Werte
                'für die Standarddatenbankfelder gibt es einige IPTC-Felder, die ich voreinstelle (Eigenbedarf)
'                cmbJahrEx.ListIndex = 9                                '9=Date created
'                cmbSituationEx.ListIndex = 11 + 5                       '11=Headline            '5 addieren Gerbing 16.11.2015
'                cmbOrtEx.ListIndex = 5 + 5                              '5=City                 '5 addieren Gerbing 16.11.2015
'                cmbLandEx.ListIndex = 7 + 5                             '7=Country              '5 addieren Gerbing 16.11.2015
'                cmbKommentarEx.ListIndex = 3 + 5                        '3=Caption-Abstract     '5 addieren Gerbing 16.11.2015

                cmbJahrEx.ListIndex = -1                                'Gerbing 27.10.2016
                cmbSituationEx.ListIndex = -1
                cmbOrtEx.ListIndex = -1
                cmbLandEx.ListIndex = -1
                cmbPersonenEx.ListIndex = -1
                cmbKommentarEx.ListIndex = -1
                cmbSWFEx.ListIndex = -1
            End If
        'End If
    End If
End Sub

Private Sub FrameNutzerDefiniert_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'Gerbing 12.10.2019
    Dim antwort As Long
    Dim rstUsrFields As ADODB.Recordset
    Dim SQL As String
    
    If Button = vbRightButton Then
        'If blnDefaultFieldsNotEmpty = True Then
            'antwort = MsgBox("Wollen Sie die Feld-Zuordnung zurücksetzen?", vbYesNo)
            antwort = MsgBox(LoadResString(2469 + Sprache), vbYesNo)
            If antwort = vbYes Then
                On Error Resume Next
                If gblnSQLServerVersion = True Then
                    'beim SQL Server muss es heißen 'Delete from table
                    SQL = "DELETE From UserDefined"
                Else
                    SQL = "DELETE * FROM UserDefined"
                End If
                Set rstUsrFields = New ADODB.Recordset
                With rstUsrFields
                    .Source = SQL
                    .ActiveConnection = Form1.DBsql
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                rstUsrFields.Close
                On Error GoTo 0
                cmbFeld1.Text = ""          'Gerbing 12.10.2019 Text ist schreibgeschützt bei DropDown-Liste -> DropDown-Kombinationsfeld benutzen
                cmbEx1.Text = ""
                cmbFeld2.Text = ""
                cmbEx2.Text = ""
                cmbFeld3.Text = ""
                cmbEx3.Text = ""
                cmbFeld4.Text = ""
                cmbEx4.Text = ""
                cmbFeld5.Text = ""
                cmbEx5.Text = ""
            End If
        'End If
    End If
End Sub

Private Sub OptEinzelnesJahr_Click()                                'Gerbing 16.11.2015
    If OptEinzelnesJahr.Value = True Then
        cmbEinzelnesJahr.Visible = True
        chkSetIPTCPresent.Visible = False
    End If
End Sub

Private Sub optInAlle_Click()                                       'Gerbing 16.11.2015
    If optInAlle.Value = True Then
        cmbEinzelnesJahr.Visible = False
        chkSetIPTCPresent.Visible = False
    End If
End Sub

Private Sub optNurFalse_Click()                                     'Gerbing 16.11.2015
    If optNurFalse = True Then
        cmbEinzelnesJahr.Visible = False
        chkSetIPTCPresent.Visible = True
    End If
End Sub

Private Sub ErrorsAufsammeln(errfile As String, strErrSammlung As String)   'Gerbing 21.11.2016
    Dim strAll As String
    Dim strConverted

    strAll = ReadFile(errfile)
    strAll = FromUTF8String(Mid(strAll, 1))
    strErrSammlung = strErrSammlung & strAll
End Sub

Private Sub ErrorsNeuSchreiben(errfile As String, strErrSammlung As String) 'Gerbing 21.11.2016
    Dim Msg As String
    On Error Resume Next
    ERR = 0
    'object.CreateTextFile(filename[, overwrite[, unicode]])
    Set oStream = PruefFso.CreateTextFile(errfile, True, True)
    If ERR <> 0 And ERR <> 55 Then                                                          'Gerbing 23.06.2011
        'Msg = "Die Datei " & errfile & " kann nicht geöffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & errfile & " " & LoadResString(1372 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number
        MsgBox Msg, vbCritical
    End If
    On Error GoTo 0
    oStream.WriteLine strErrSammlung
    oStream.Close
End Sub

Function FromUTF8String(ByVal S As String) As String                        'Gerbing 21.11.2016
   Dim i As Integer, b(2) As Byte
   
   i = 1
   S = S & Chr(0) & Chr(0)
   Do While i <= Len(S) - 2
      b(0) = Asc(Mid(S, i, 1))
      b(1) = Asc(Mid(S, i + 1, 1))
      b(2) = Asc(Mid(S, i + 2, 1))
      If (b(0) And &HE0) = &HE0 Then
         FromUTF8String = FromUTF8String & ChrW((b(0) And &HF) * CLng(&H1000) + (b(1) And &H3F) * CLng(&H40) + (b(2) And &H3F))
         i = i + 3
      ElseIf (b(0) And &HC0) = &HC0 Then
         FromUTF8String = FromUTF8String & ChrW((b(0) And &H1F) * &H40 + (b(1) And &H3F))
         i = i + 2
      Else
         FromUTF8String = FromUTF8String & Chr(b(0))
         i = i + 1
      End If
   Loop
End Function

Public Function ReadFile(ByVal strFileName As String) As String         'Gerbing 21.11.2016
   Dim intHandle As Integer
   
   intHandle = FreeFile
   Open strFileName For Input As #intHandle
   ReadFile = Input(LOF(intHandle), #intHandle)
   Close #intHandle
End Function

Private Sub UserdefinedAuswerten()
    Dim rstUserDef As ADODB.Recordset                                                       'Gerbing 02.10.2019
    Dim SQL As String
    Dim n As Integer
    
    'Tabelle UserDefined auswerten                                                          'Gerbing 02.10.2019
    'wenn in der Tabelle UserDefined etwas steht,
    'dann überschreibe ich die voreingestellten EXIF/IPTC-Felder
    On Error Resume Next
    SQL = "select * From UserDefined;"
    Set rstUserDef = New ADODB.Recordset
    With rstUserDef
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number = 0 Then
        'hier existiert die Tabelle UserDefined
        If Not rstUserDef.EOF Then
            'hier ist die Tabelle UserDefined nicht leer
            'Wenn einer der bereits eingetragenen ComboItems mit dem Wert in rstUserDef.Fields("FieldName1") übereinstimmt
            'dann muss ich diesen Item selektieren
            For n = 0 To cmbFeld1.ComboItems.Count - 1
                If cmbFeld1.ComboItems(n).Text = rstUserDef.Fields("FieldName1") Then
                    cmbFeld1.SelectItemByText (rstUserDef.Fields("FieldName1"))
                    Exit For
                End If
            Next n
            cmbEx1.Text = rstUserDef.Fields("SourceField1")
            Call cmbEx1_Click                                                           'Gerbing 03.01.2020
            For n = 0 To cmbFeld2.ComboItems.Count - 1
                If cmbFeld2.ComboItems(n).Text = rstUserDef.Fields("FieldName2") Then
                    cmbFeld2.SelectItemByText (rstUserDef.Fields("FieldName2"))
                    Exit For
                End If
            Next n
            cmbEx2.Text = rstUserDef.Fields("SourceField2")
            Call cmbEx2_Click                                                           'Gerbing 03.01.2020
            For n = 0 To cmbFeld3.ComboItems.Count - 1
                If cmbFeld3.ComboItems(n).Text = rstUserDef.Fields("FieldName3") Then
                    cmbFeld3.SelectItemByText (rstUserDef.Fields("FieldName3"))
                    Exit For
                End If
            Next n
            cmbEx3.Text = rstUserDef.Fields("SourceField3")
            Call cmbEx3_Click                                                           'Gerbing 03.01.2020
            For n = 0 To cmbFeld4.ComboItems.Count - 1
                If cmbFeld4.ComboItems(n).Text = rstUserDef.Fields("FieldName4") Then
                    cmbFeld4.SelectItemByText (rstUserDef.Fields("FieldName4"))
                    Exit For
                End If
            Next n
            cmbEx4.Text = rstUserDef.Fields("SourceField4")
            Call cmbEx4_Click                                                           'Gerbing 03.01.2020
            For n = 0 To cmbFeld5.ComboItems.Count - 1
                If cmbFeld5.ComboItems(n).Text = rstUserDef.Fields("FieldName5") Then
                    cmbFeld5.SelectItemByText (rstUserDef.Fields("FieldName5"))
                    Exit For
                End If
            Next n
            cmbEx5.Text = rstUserDef.Fields("SourceField5")
            Call cmbEx5_Click                                                           'Gerbing 03.01.2020
        End If
    End If
    rstUserDef.Close
End Sub
