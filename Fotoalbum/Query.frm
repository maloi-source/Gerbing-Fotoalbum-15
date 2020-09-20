VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Query 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GERBING Fotoalbum - Formulieren Sie die Such-Kriterien"
   ClientHeight    =   10200
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   11268
   Icon            =   "Query.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   11268
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox CheckWeitereFilterAktiv 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   5520
      TabIndex        =   25
      ToolTipText     =   "Zum Deaktivieren klicken Sie auf weitere Filter..."
      Top             =   6360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.ListBox lstFensterTitel 
      Height          =   240
      Left            =   8760
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox CheckUseAudioComments 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Audio-Kommentare benutzen"
      Height          =   372
      Left            =   1560
      TabIndex        =   44
      ToolTipText     =   "zu einer Foto-Datei kann eine gleichnamige Audio-Datei aufgenommen oder abgespielt werden"
      Top             =   840
      Width           =   5052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   2040
      Top             =   9240
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   11172
   End
   Begin VB.CheckBox CheckGespeicherteAbfragen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "gespeicherte Abfragen benutzen"
      Height          =   312
      Left            =   1560
      TabIndex        =   38
      Top             =   480
      Width           =   5892
   End
   Begin VB.CheckBox chkFensterGrößeÄnderbar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fenstergröße änderbar"
      Height          =   372
      Left            =   8520
      TabIndex        =   37
      ToolTipText     =   $"Query.frx":038A
      Top             =   7320
      Width           =   3012
   End
   Begin VB.CheckBox CheckNutzerdefinierteFelder 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suche nach nutzerdefinierten Feldern ist aktiv"
      Enabled         =   0   'False
      Height          =   372
      Left            =   5520
      TabIndex        =   36
      ToolTipText     =   "Zum Deaktivieren klicken Sie auf nutzerdefinierte Felder..."
      Top             =   6840
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton btnNutzerdefinierteFelder 
      Caption         =   "&nutzerdefinierte Felder..."
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      ToolTipText     =   "Klicken Sie hier, wenn Sie die Suche nach nutzerdefinierten Feldern erweitern wollen"
      Top             =   6840
      Width           =   3852
   End
   Begin VB.CheckBox CheckDifferenzen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "error check for differences in year and filename"
      Height          =   312
      Left            =   1560
      TabIndex        =   24
      ToolTipText     =   "Nur sinnvoll, wenn Jahresordner im Dateiname vorkommen. Es wird nach den Differenzen gesucht."
      Top             =   120
      Width           =   6972
   End
   Begin VB.CommandButton btnMehrerePersonen 
      Caption         =   "&Weitere Filter..."
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      ToolTipText     =   "Sie können die Suche auf weitere Personen/andere Sortierung/Dateidatum erweitern"
      Top             =   6360
      Width           =   3852
   End
   Begin VB.CommandButton btnRefresh 
      Height          =   495
      Left            =   840
      Picture         =   "Query.frx":043C
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Reset - Alle Felder auf die Standardwerte einstellen"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnTürZu 
      Height          =   495
      Left            =   120
      Picture         =   "Query.frx":0586
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Beenden"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox CheckSucheJedesFeld 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suche Begriff in jedem Feld"
      Height          =   372
      Left            =   1560
      TabIndex        =   20
      ToolTipText     =   "Jedes Standardfeld wird nach diesem Begriff durchsucht. Wenn Sie nach einem Datum suchen, benutzen Sie 'Weitere Filter...'"
      Top             =   7320
      Width           =   3852
   End
   Begin VB.Timer TimerSetFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   6960
   End
   Begin VB.CheckBox CheckSQL 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SQL nachbearbeiten"
      Height          =   372
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Nur sinnvoll, wenn Sie die SQL-Sprache beherrschen"
      Top             =   7320
      Width           =   2892
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Fotos finden"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Gefunden werden nur die Fotos, bei denen die gewählten Suchbegriffe in den gewählten Felder stehen"
      Top             =   8280
      Width           =   3852
   End
   Begin VB.Frame Frame2 
      Height          =   4692
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10932
      Begin VB.ComboBox TFileType 
         Height          =   288
         Left            =   7080
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   3000
         Width           =   1332
      End
      Begin VB.TextBox TJahr 
         Height          =   288
         Left            =   1440
         TabIndex        =   46
         Top             =   600
         Width           =   1332
      End
      Begin VB.CheckBox CheckAudioFileExists 
         Caption         =   "nur Fotos mit Audio-Kommentar finden"
         Height          =   372
         Left            =   5760
         TabIndex        =   43
         Top             =   4200
         Visible         =   0   'False
         Width           =   5052
      End
      Begin VB.Frame FrameJahrErweiterung 
         Caption         =   "Trefferauswahl"
         Height          =   828
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   5292
         Begin VB.OptionButton optErsterZufallstreffer 
            Caption         =   "ein Zufallstreffer pro Jahr"
            Height          =   252
            Left            =   1560
            TabIndex        =   47
            Top             =   480
            Width           =   3612
         End
         Begin VB.OptionButton optAlleTreffer 
            Caption         =   "Alle"
            Height          =   372
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Alle Treffer pro Jahr (=Standard)"
            Top             =   240
            Value           =   -1  'True
            Width           =   1332
         End
         Begin VB.OptionButton optNurErstenTreffer 
            Caption         =   "erster Treffer pro Jahr"
            Height          =   372
            Left            =   1560
            TabIndex        =   33
            ToolTipText     =   "Nur den ersten Treffer pro Jahr. So können Sie beispielsweise die jährliche Entwicklung eines Kindes verfolgen."
            Top             =   120
            Width           =   3612
         End
      End
      Begin VB.Frame Frame7 
         Height          =   528
         Left            =   8520
         TabIndex        =   29
         Top             =   480
         Width           =   2244
         Begin VB.OptionButton JOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1080
            TabIndex        =   31
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton JUnd 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   30
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox CheckVollesWort 
         Caption         =   "Person nur als volles Wort finden"
         Height          =   372
         Left            =   1440
         TabIndex        =   27
         ToolTipText     =   "Wenn Sie wirklich zB nur Ina finden wollen und nicht auch Martina, Bei mehreren Personen ist diese Funktion nicht möglich"
         Top             =   4200
         Width           =   4572
      End
      Begin VB.Frame Frame6 
         Height          =   528
         Left            =   8520
         TabIndex        =   12
         Top             =   1080
         Width           =   2244
         Begin VB.OptionButton SOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1080
            TabIndex        =   14
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton SUnd 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   528
         Left            =   8520
         TabIndex        =   9
         Top             =   1680
         Width           =   2244
         Begin VB.OptionButton OOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1080
            TabIndex        =   11
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton OUnd 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   10
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   528
         Left            =   8520
         TabIndex        =   6
         Top             =   2280
         Width           =   2244
         Begin VB.OptionButton LOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1080
            TabIndex        =   8
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton LUnd 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   528
         Left            =   8520
         TabIndex        =   3
         Top             =   2880
         Width           =   2244
         Begin VB.OptionButton SWFOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1080
            TabIndex        =   5
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton SWFUnd 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   4
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox TSWF 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   19
         Text            =   "TSWF"
         Top             =   3000
         Width           =   1215
      End
      Begin CBLCtlsLibUCtl.ComboBox TSituation 
         Height          =   288
         Left            =   1440
         TabIndex        =   48
         Top             =   1200
         Width           =   6972
         _cx             =   12298
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5349
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "Query.frx":0BF8
         Text            =   "Query.frx":0C18
      End
      Begin CBLCtlsLibUCtl.ComboBox TPersonen 
         Height          =   288
         Left            =   1440
         TabIndex        =   51
         Top             =   3600
         Width           =   6972
         _cx             =   12298
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5349
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "Query.frx":0C3A
         Text            =   "Query.frx":0C5A
      End
      Begin CBLCtlsLibUCtl.ComboBox TLand 
         Height          =   288
         Left            =   1440
         TabIndex        =   50
         Top             =   2400
         Width           =   6972
         _cx             =   12298
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5349
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "Query.frx":0C7C
         Text            =   "Query.frx":0C9C
      End
      Begin CBLCtlsLibUCtl.ComboBox TOrt 
         Height          =   288
         Left            =   1440
         TabIndex        =   49
         Top             =   1800
         Width           =   6972
         _cx             =   12298
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5349
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "Query.frx":0CBE
         Text            =   "Query.frx":0CDE
      End
      Begin VB.Label LFileType 
         Caption         =   "Dateityp:"
         Height          =   252
         Left            =   5760
         TabIndex        =   59
         Top             =   3000
         Width           =   1212
      End
      Begin VB.Label LJahr 
         Caption         =   "Jahr:"
         Height          =   372
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label LPersonen 
         Caption         =   "Personen:"
         Height          =   372
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   1212
      End
      Begin VB.Label LSituation 
         Caption         =   "Situation:"
         Height          =   372
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label LOrt 
         Caption         =   "Ort:"
         Height          =   372
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1092
      End
      Begin VB.Label LLand 
         Caption         =   "Land:"
         Height          =   372
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1092
      End
      Begin VB.Label LSWF 
         Caption         =   "SW/F:"
         Height          =   372
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "SW=Schwarz/Weiss-Foto F=Farbfoto SV=Schwarz/Weiss-Video FV=Farbvideo"
         Top             =   3000
         Width           =   1092
      End
   End
   Begin CBLCtlsLibUCtl.ComboBox TBegriff 
      Height          =   288
      Left            =   1560
      TabIndex        =   52
      Top             =   7800
      Visible         =   0   'False
      Width           =   9612
      _cx             =   16954
      _cy             =   508
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   267501
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
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   0   'False
      CueBanner       =   "Query.frx":0D00
      Text            =   "Query.frx":0D20
   End
   Begin EditCtlsLibUCtl.TextBox SQLText 
      Height          =   1332
      Left            =   120
      TabIndex        =   53
      Top             =   8760
      Visible         =   0   'False
      Width           =   11052
      _cx             =   19494
      _cy             =   2350
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
      DisabledEvents  =   3073
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
      CueBanner       =   "Query.frx":0D40
      Text            =   "Query.frx":0D60
   End
   Begin VB.Frame FrameGespeicherteAbfragen 
      Height          =   5775
      Left            =   120
      TabIndex        =   39
      Top             =   1320
      Width           =   11052
      Begin CBLCtlsLibUCtl.ListBox ListGespeicherteAbfragen 
         Height          =   2412
         Left            =   1080
         TabIndex        =   54
         Top             =   600
         Width           =   8652
         _cx             =   15261
         _cy             =   4254
         AllowDragDrop   =   0   'False
         AllowItemSelection=   -1  'True
         AlwaysShowVerticalScrollBar=   0   'False
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   0
         ColumnWidth     =   -1
         DisabledEvents  =   1048809
         DontRedraw      =   0   'False
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
         HasStrings      =   -1  'True
         HoverTime       =   -1
         IMEMode         =   -1
         InsertMarkColor =   0
         InsertMarkStyle =   1
         IntegralHeight  =   0   'False
         ItemHeight      =   -1
         Locale          =   1024
         MousePointer    =   0
         MultiColumn     =   0   'False
         MultiSelect     =   0
         OLEDragImageStyle=   0
         OwnerDrawItems  =   0
         ProcessContextMenuKeys=   -1  'True
         ProcessTabs     =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         ScrollableWidth =   0
         Sorted          =   0   'False
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         ToolTips        =   0
         UseSystemFont   =   0   'False
         VirtualMode     =   0   'False
      End
      Begin EditCtlsLibUCtl.TextBox txtSQLGespeicherteAbfrage 
         Height          =   2052
         Left            =   1080
         TabIndex        =   55
         Top             =   3480
         Width           =   8652
         _cx             =   15261
         _cy             =   3619
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
         CueBanner       =   "Query.frx":0D80
         Text            =   "Query.frx":0DA0
      End
      Begin VB.Label Label2 
         Caption         =   "SQL-Darstellung der gespeicherten Abfrage"
         Height          =   372
         Left            =   1080
         TabIndex        =   41
         Top             =   3120
         Width           =   7452
      End
      Begin VB.Label Label1 
         Caption         =   "gespeicherte Abfragen in fotos.mdb"
         Height          =   372
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   5412
      End
   End
   Begin VB.Label lblWeitereFilterAktiv 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WeiterFilterAktiv"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   5900
      TabIndex        =   56
      Top             =   6420
      Visible         =   0   'False
      Width           =   5532
   End
   Begin VB.Label lblNutzerdefinierteFelder 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suche nach nutzerdefinierten Feldern ist aktiv"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   5900
      TabIndex        =   57
      Top             =   6900
      Visible         =   0   'False
      Width           =   5532
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDiashow 
         Caption         =   "Diashow starten"
      End
      Begin VB.Menu mnuFotosmdb 
         Caption         =   "&Fotosmdb starten"
      End
      Begin VB.Menu mnuRenamMdb 
         Caption         =   "&RenamMdb starten"
      End
   End
   Begin VB.Menu mnuEinstellungen 
      Caption         =   "&Einstellungen"
   End
   Begin VB.Menu mnuSprache 
      Caption         =   "&Sprache"
   End
   Begin VB.Menu mnuReset 
      Caption         =   "&Reset"
   End
   Begin VB.Menu mnuResetAll 
      Caption         =   "Reset&All"
   End
   Begin VB.Menu mnuSpaltenbreite 
      Caption         =   "Spaltenbreite"
   End
   Begin VB.Menu mnuHilfe 
      Caption         =   "&Hilfe"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Über..."
   End
   Begin VB.Menu mnuBeenden 
      Caption         =   "&Beenden"
   End
End
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public SQL As String
    Public RecordCount As Long
    Public SucheInJedemFeld As Boolean
    Public OKGewählt As Boolean
    Dim Msg As String
    Dim SQLBearbeitenZähler As Integer
    Dim FDF As String            'FDF = Formatiertes Datums Feld
    'Public adocn As ADODB.Connection
    Public enumCursorOrt As ADODB.CursorLocationEnum
    Public adoRs As ADODB.Recordset                     'Gerbing 04.01.2006
    Dim blnComeFromMsgbox As Boolean
    
    Dim SQLJahr As String
    Dim Plus1 As String
    Dim Plus As String
    Dim Vergleich As String
    Dim pos1 As Integer
    Dim JahrVon As String
    Dim JahrBis As String
    Dim Von As String
    Dim Bis As String
    Public SQLJahresZahl As String
    Dim blnExitChange As Boolean
    Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, _
        ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
                 
    Private Const WM_SETTEXT = &HC
    Dim strTemp As String
    Dim blnNewSelectedNothing As Boolean
    Dim blnMsgCheckUseAudioComments As Boolean
    Dim blnMsgCheckGespeicherteAbfragen As Boolean

Private Sub btnMehrerePersonen_Click()
    Me.MousePointer = vbHourglass                       'Gerbing 25.09.2007
    CheckVollesWort.Value = 0       '0 = deaktiviert
    MP.Show 1
    Me.MousePointer = vbDefault                         'Gerbing 25.09.2007
End Sub

Private Sub btnNutzerdefinierteFelder_Click()
    Dim Msg As String
    Dim n As Long
    
    'Shareware user bekommen einen Hinweis auf Professional Version         'Gerbing 15.05.2014
    If Not (gblnVollversion = True And gblnProversion = True) Then
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    'Abweisen, wenn es keine nutzerdefinierten Felder gibt
    If ND.ListNutzerdefinierteFelder.ListItems.Count = 0 Then
        'msg = "Sie können erst dann nutzerdefinierte Felder in die Suche einbeziehen," & vbNewLine
        Msg = LoadResString(2113 + Sprache) & vbNewLine                     'Gerbing 08.11.2005
        'msg = msg & "nachdem Sie nutzerdefinierte Felder angelegt haben." & vbNewLine
        Msg = Msg & LoadResString(2114 + Sprache) & vbNewLine
        'msg = msg & "Lesen Sie in der Hilfe, wie nutzerdefinierte Felder angelegt werden"
        Msg = Msg & LoadResString(2115 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
    '-------------------------------------------------------------------------------
    'Formular ND öffnen, wenn die Suche mit diesen Feldern erweitert werden soll
    If ND.cmbFeld1.ComboItems.Count = 0 Then
        'ND.ListNutzerdefinierteFelder.ListIndex = 0
        For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
            ND.cmbFeld1.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
        Next n
    End If
    If ND.cmbFeld2.ComboItems.Count = 0 Then
        'ND.ListNutzerdefinierteFelder.ListIndex = 0
        For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
            ND.cmbFeld2.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
        Next n
    End If
    If ND.cmbFeld3.ComboItems.Count = 0 Then
        'ND.ListNutzerdefinierteFelder.ListIndex = 0
        For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
            ND.cmbFeld3.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
        Next n
    End If
    If ND.cmbFeld4.ComboItems.Count = 0 Then
        'ND.ListNutzerdefinierteFelder.ListIndex = 0
        For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
            ND.cmbFeld4.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
        Next n
    End If
    If ND.cmbFeld5.ComboItems.Count = 0 Then
        'ND.ListNutzerdefinierteFelder.ListIndex = 0
        For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
            ND.cmbFeld5.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
        Next n
    End If
    ND.Show 1
End Sub

Private Sub btnTürZu_Click()
    Call Beenden                   'Gerbing 24.06.2006
    End
End Sub

Private Sub CheckGespeicherteAbfragen_Click()
    Dim intLoop As Integer
    Dim SQL As String
    Dim tbl As String
        
    'Shareware user bekommen einen Hinweis auf Professional Version         'Gerbing 15.05.2014
    If Not (gblnVollversion = True And gblnProversion = True) Then
        If blnMsgCheckGespeicherteAbfragen = True Then
            blnMsgCheckGespeicherteAbfragen = False
            Exit Sub
        End If
        blnMsgCheckGespeicherteAbfragen = True
        CheckGespeicherteAbfragen.Value = 0
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    '-------------------------------------------
    ListGespeicherteAbfragen.ListItems.RemoveAll
    If CheckGespeicherteAbfragen.Value = 1 Then
        'Gespeicherte Abfragen benutzen - ja
        If optAlleTreffer.Value = False Then
            'erzwingen, dass nicht 'erster Treffer pro Jahr' eingeschaltet ist  'Gerbing 28.12.2005
            optAlleTreffer.Value = True
            'MsgBox "'erster Treffer pro Jahr'wird zurückgesetzt auf 'Alle'"
            MsgBox LoadResString(2200 + Sprache)
        End If
        CheckDifferenzen.Value = 0
        FrameGespeicherteAbfragen.Visible = True
        Frame2.Visible = False
        btnMehrerePersonen.Visible = False
        btnNutzerdefinierteFelder.Visible = False
        'CheckSQL.Visible = False                               'Gerbing 19.01.2007
        CheckSucheJedesFeld.Visible = False
        'chkFensterGrößeÄnderbar.Visible = False                'Gerbing 15.06.2008
        TBegriff.Visible = False
        SQLText.Visible = False
        'CheckSQL.Value = 0                                                                 'Gerbing 25.03.2009 auskommentiert
        CheckSQL.Visible = False                                                            'Gerbing 15.06.2008
        CheckSucheJedesFeld.Value = 0
        CheckNutzerdefinierteFelder.Visible = False
        lblNutzerdefinierteFelder.Visible = False                                           'Gerbing 25.06.2013
        CheckWeitereFilterAktiv.Visible = False                                             'Gerbing 05.02.2006
        lblWeitereFilterAktiv.Visible = False                                               'Gerbing 25.06.2013
        If gblnSQLServerVersion = True Then                                                'Gerbing 23.11.2017
            SQL = "select * from INFORMATION_SCHEMA.VIEWS"
            On Error Resume Next
            rstsql.Close
            On Error GoTo 0
            With rstsql
                .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
                .CursorType = adOpenForwardOnly
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Source = SQL
                .Open
            End With
            
            While Not rstsql.EOF
                ListGespeicherteAbfragen.ListItems.Add rstsql.Fields("TABLE_NAME")
                rstsql.MoveNext
            Wend
        Else
            Set rstsql = DBado.OpenSchema(adSchemaProcedures)                               'Gerbing 09.01.2018 05.07.2019
            While Not rstsql.EOF
                'Debug.Print rstsql!table_name
                ListGespeicherteAbfragen.ListItems.Add rstsql.Fields("PROCEDURE_NAME")
                rstsql.MoveNext
            Wend
            Set rstsql = DBado.OpenSchema(adSchemaViews)
            While Not rstsql.EOF
                ListGespeicherteAbfragen.ListItems.Add rstsql.Fields("TABLE_NAME")
                rstsql.MoveNext
            Wend                                                                            'Gerbing 05.07.2019
        End If
    Else
        'Gespeicherte Abfragen benutzen - nein
        FrameGespeicherteAbfragen.Visible = False
        SQLWurdeBearbeitet = False
        Frame2.Visible = True
        btnMehrerePersonen.Visible = True
        If CheckWeitereFilterAktiv.Value = 1 Then                                           'Gerbing 05.02.2006
            CheckWeitereFilterAktiv.Visible = True
            lblWeitereFilterAktiv.Visible = True                                            'Gerbing 25.06.2013
        End If
        btnNutzerdefinierteFelder.Visible = True
        CheckSQL.Visible = True
        CheckSucheJedesFeld.Visible = True
        'chkFensterGrößeÄnderbar.Visible = True             'Gerbing 15.06.2008
        txtSQLGespeicherteAbfrage = ""                      'Gerbing 20.06.2006
        Call RefreshWerte                                                                   'Gerbing 23.11.2008
    End If
End Sub

Private Sub CheckDifferenzen_Click()
    CheckUseAudioComments.Value = 0                                                 'Gerbing 06.05.2012
    If CheckDifferenzen.Value = 0 Then
        'SQLText = "Select * From Fotos ORDER BY Dateiname"
        SQLText = "Select * From Fotos ORDER BY " & LoadResString(1028 + Sprache)   'Gerbing 08.11.2005
        SQLWurdeBearbeitet = False
        Frame2.Visible = True
        btnMehrerePersonen.Visible = True
        btnNutzerdefinierteFelder.Visible = True
        #If Proversion Then
        If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
            If CheckNutzerdefinierteFelder.Value = 1 Then   'Gerbing 05.02.2006
                CheckNutzerdefinierteFelder.Visible = True
                lblNutzerdefinierteFelder.Visible = True                            'Gerbing 25.06.2013
            Else
                CheckNutzerdefinierteFelder.Visible = False
                lblNutzerdefinierteFelder.Visible = False                           'Gerbing
            End If
            If CheckWeitereFilterAktiv.Value = 1 Then
                CheckWeitereFilterAktiv.Visible = True
                lblWeitereFilterAktiv.Visible = True                                'Gerbing 25.06.2013
            Else
                CheckWeitereFilterAktiv.Visible = False
                lblWeitereFilterAktiv.Visible = False                               'Gerbing 25.06.2013
            End If
        End If
        #End If
        CheckSQL.Visible = True
        CheckSucheJedesFeld.Visible = True
        chkFensterGrößeÄnderbar.Visible = True
        'btnOK.Caption = "&Fotos finden"
        btnOK.Caption = LoadResString(1004 + Sprache)
    Else
        If optAlleTreffer.Value = False Then
            'erzwingen, dass nicht 'erster Treffer pro Jahr' eingeschaltet ist  'Gerbing 28.12.2005
            optAlleTreffer.Value = True
            'MsgBox "'erster Treffer pro Jahr'wird zurückgesetzt auf 'Alle'"
            MsgBox LoadResString(2200 + Sprache)
        End If
        CheckGespeicherteAbfragen.Value = 0
        Frame2.Visible = False
        btnMehrerePersonen.Visible = False
        btnNutzerdefinierteFelder.Visible = False
        CheckSQL.Visible = False
        CheckSucheJedesFeld.Visible = False
        chkFensterGrößeÄnderbar.Visible = False
        'btnOK.Caption = "&Differenzen finden"
        btnOK.Caption = LoadResString(1005 + Sprache)
        btnOK.Enabled = True
        TBegriff.Visible = False
        SQLText.Visible = False
        CheckSQL.Value = 0
        CheckSucheJedesFeld.Value = 0
        CheckNutzerdefinierteFelder.Visible = False
        lblNutzerdefinierteFelder.Visible = False                               'Gerbing 25.06.2013
        CheckWeitereFilterAktiv.Visible = False
        lblWeitereFilterAktiv.Visible = False                                   'Gerbing 25.06.2013
    End If
End Sub


Private Sub CheckSQL_Click()
    If CheckSQL.Value = 0 Then      '0 = deaktiviert
        SQLText.Visible = False
        SQLText = ""
        SQLWurdeBearbeitet = False
        btnOK.Enabled = True                                'Gerbing 10.03.2008
    Else
        SQLText.Visible = True
    End If
End Sub

Private Sub CheckSucheJedesFeld_Click()
    If CheckSucheJedesFeld.Value = 1 Then
        TBegriff.Visible = True
        TBegriff.SetFocus
    Else
        TBegriff.Visible = False
    End If
End Sub

Private Sub btnRefresh_Click()
    Call RefreshWerte
End Sub

Private Sub CheckUseAudioComments_Click()
    Dim RetVal As Long
    
    'Shareware user bekommen einen Hinweis auf Professional Version         'Gerbing 15.05.2014
    If Not (gblnVollversion = True And gblnProversion = True) Then
        If blnMsgCheckUseAudioComments = True Then
            blnMsgCheckUseAudioComments = False
            Exit Sub
        End If
        blnMsgCheckUseAudioComments = True
        CheckUseAudioComments.Value = 0
        Msg = LoadResString(2335 + Sprache) 'Für diese Funktion benötigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
'-------------------------------
    RetVal = GetVersionEx(OSInfo)
    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space(128)
    RetVal = GetVersionEx(OSInfo)
    If OSInfo.dwBuildNummer < 2600 Or OSInfo.dwMajorVersion < 5 Then
        If blnComeFromMsgbox = False Then
            'MsgBox "Die Funktion 'Audio-Kommentare benutzen' steht nur im Windows XP zur Verfügung" oder höher
            MsgBox LoadResString(2248 + Sprache)
            blnComeFromMsgbox = True
        Else
            blnComeFromMsgbox = False
        End If
        CheckUseAudioComments.Value = 0
        Exit Sub
    End If
    If CheckUseAudioComments.Value = 1 Then
        CheckAudioFileExists.Visible = True
        Load frmStartSoundAutomatisch                       'Gerbing 09.12.2009
    Else
        CheckAudioFileExists.Visible = False
        CheckAudioFileExists.Value = 0
        Unload frmStartSoundAutomatisch                     'Gerbing 09.12.2009
    End If
    Unload KommentarForm                                   'Gerbing 09.12.2009
End Sub

Private Sub CheckVollesWort_Click()
    If MP.AnzahlPersonen > 0 And CheckVollesWort.Value = 1 Then
        'MsgBox "Wenn mit 'Weitere Filter...' mehrere Personen gesucht werden, kann diese Funktion nicht benutzt werden."
        MsgBox LoadResString(2116 + Sprache)
        CheckVollesWort.Value = 0
    End If
End Sub

Private Sub chkFensterGrößeÄnderbar_Click()
    If chkFensterGrößeÄnderbar = 1 Then          '1=aktiviert
        'Fenstergröße änderbar ja
        Form1.WindowState = 2                      '0=normal                        'Gerbing 27.09.2017
        On Error Resume Next
        Form1.width = (screenWidth - 100) * Screen.TwipsPerPixelX                   'Gerbing 09.05.2012
        frmVideo.WindowState = 2                                                    'Gerbing 16.06.2012 27.09.2017
        frmVideo.width = Form1.width
        frmVideo.height = Form1.height
        glngSaveForm1Width = Form1.width
        glngSaveForm1Height = Form1.height
    Else
        'Fenstergröße änderbar nein
        Form1.WindowState = 2                      '2=maximiert
        frmVideo.WindowState = 2
        frmVideo.Caption = ""
    End If
    Form1.Hide                                  'Gerbing 30.09.2013
    Query.Show                                  'Gerbing 15.11.2012
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim ShiftDown As Boolean
'    Dim AltDown As Boolean
'    Dim CtrlDown As Boolean
'
'    ShiftDown = (Shift And vbShiftMask) > 0
'    AltDown = (Shift And vbAltMask) > 0
'    CtrlDown = (Shift And vbCtrlMask) > 0
'    Select Case KeyCode
'        Case vbKeyN
'            If CtrlDown And ShiftDown Then            'Strg+Num+N gleichzeitig                  'Gerbing 03.03.2012
'                gblnBildBeschreibung = True
'            End If
'        Case vbKeyM
'            If CtrlDown And ShiftDown Then            'Strg+Num+M gleichzeitig                  'Gerbing 03.03.2012
'                gblnBildBeschreibung = False
'            End If
'    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Beenden                                    'Gerbing 24.06.2006
    End
End Sub

Private Sub btnOK_Click()
    Dim pos As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim Plus2 As String
    Dim Plus3 As String
    Dim Plus4 As String
    Dim Plus5 As String
    Dim rc As Integer
    Dim DatumVon As Date
    Dim DatumBis As Date
    Dim TSplitSit() As String
    Dim TSplitOrt() As String
    Dim TSplitLand() As String
    Dim TSplitBeg() As String
    Dim n As Long
    Dim i As Long
    Dim strTemp As String
    Dim strVerk As String
    Dim gefunden As Boolean
    Dim errLoop As ADODB.Error
    Dim strLinks As String
    Dim strRechts As String
    Dim VPartyyyy As String
    Dim VPartmm As String
    Dim VPartdd As String
    Dim BPartyyyy As String
    Dim BPartmm As String
    Dim BPartdd As String

    Dim WndHnd As Long
    Dim CurrWnd As Long
    Dim TitelLength As Long
    Dim FensterTitel As String
    Dim X As Long
    Const GW_HWNDFIRST = 0
    Const GW_HWNDNEXT = 2
    Dim LinksOben As String                                                             'Gerbing 05.09.2016
    Dim LinksUnten As String
    Dim RechtsOben As String
    Dim RechtsUnten As String

    gblnComeFromF8 = False                                                              'Gerbing 06.06.2015
    If gblnSQLServerVersion = True Then
        If gblnWeiterMitLeererDatenbank = True Then
            Msg = LoadResString(1806 + Sprache) & " " & PublicSQLServer & " " & PublicSQLDatabase & LoadResString(1836 + Sprache) & vbNewLine
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End If
    Me.MousePointer = vbHourglass                           'Gerbing 06.08.2007
    If CheckSucheJedesFeld.Value = 1 Then
        If Sprache = 0 Then                                 'Gerbing 08.11.2005
            pos = InStr(1, TBegriff.Text, ".")                   'Gerbing 22.02.2005
            pos1 = InStr(pos + 1, TBegriff.Text, ".")
            pos2 = InStr(1, TBegriff.Text, ":")
            pos3 = InStr(pos2 + 1, TBegriff.Text, ":")
        Else
            pos = InStr(1, TBegriff.Text, "/")              'Gerbing 22.02.2005
            pos1 = InStr(pos + 1, TBegriff.Text, "/")
            pos2 = InStr(1, TBegriff.Text, ":")
            pos3 = InStr(pos2 + 1, TBegriff.Text, ":")
        End If
        If (pos <> 0 And pos1 <> 0) Or (pos2 <> 0 And pos3 <> 0) Then
            'msg = "Das Programm hat bei 'Suche Begriff in jedem Feld' mindestens zweimal die Zeichen . oder : erkannt" & NL
            Msg = LoadResString(2117 + Sprache) & NL
            'msg = msg & "Darum wird vermutet, dass Sie nach Datum oder Uhrzeit suchen" & NL
            Msg = Msg & LoadResString(2118 + Sprache) & NL
            'msg = msg & "Benutzen Sie für solche Zwecke 'Weitere Filter...'" & NL
            Msg = Msg & LoadResString(2119 + Sprache) & NL
            'msg = msg & "oder legen Sie nutzerdefinierte Felder vom Typ Datum/Uhrzeit an"
            Msg = Msg & LoadResString(2120 + Sprache)
            MsgBox Msg
            TBegriff.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        'Neuen Begriff nur in die Combobox aufnehmen, wenn er noch nicht aufgenommen ist Gerbing 31.03.2005
        If TBegriff.ComboItems.Count = 0 Then
            If TBegriff.Text <> "" Then                     'Gerbing 12.07.2016
                TBegriff.ComboItems.Add TBegriff.Text       'Gerbing 28.03.2005
            End If
        Else
            gefunden = False
            For n = 0 To TBegriff.ComboItems.Count - 1
                If StrComp(TBegriff.Text, TBegriff.ComboItems(n), vbTextCompare) = 0 Then gefunden = True
            Next n
            If gefunden = False Then
                TBegriff.ComboItems.Add TBegriff.Text       'Gerbing 28.03.2005
            End If
        End If
    End If
    '-------------------------------------------------------------------------------------------
    If Trim(TSituation.Text) = "" Then         'Gerbing 31.07.2007
        'msg = "Das Feld Ort darf nicht leer sein. Wenn Sie keinen Ort auswählen wollen, benutzen Sie '*' oder'Beliebig'. Wenn Sie Datensätze suchen, wo dieses Feld leer ist, benutzen Sie NULL"
        Msg = LoadResString(2121 + Sprache)
        MsgBox Msg
        TSituation.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(TOrt.Text) = "" Then               'Gerbing 31.07.2007
        'msg = "Das Feld Ort darf nicht leer sein. Wenn Sie keinen Ort auswählen wollen, benutzen Sie '*' oder'Beliebig'. Wenn Sie Datensätze suchen, wo dieses Feld leer ist, benutzen Sie NULL"
        Msg = LoadResString(2122 + Sprache)
        MsgBox Msg
        TOrt.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(TLand.Text) = "" Then              'Gerbing 31.07.2007
        'msg = "Das Feld Ort darf nicht leer sein. Wenn Sie keinen Ort auswählen wollen, benutzen Sie '*' oder'Beliebig'. Wenn Sie Datensätze suchen, wo dieses Feld leer ist, benutzen Sie NULL"
        Msg = LoadResString(2123 + Sprache)
        MsgBox Msg
        TLand.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(TPersonen.Text) = "" Then          'Gerbing 31.07.2007
        'msg = "Das Feld Ort darf nicht leer sein. Wenn Sie keinen Ort auswählen wollen, benutzen Sie '*' oder'Beliebig'. Wenn Sie Datensätze suchen, wo dieses Feld leer ist, benutzen Sie NULL"
        Msg = LoadResString(2124 + Sprache)
        MsgBox Msg
        TPersonen.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    If InStr(1, TSituation.Text, "%%%", vbTextCompare) <> 0 And InStr(1, TSituation.Text, "&&&", vbTextCompare) <> 0 Then
        'msg = "Sie dürfen %%% und &&& nicht gleichzeitig verwenden"
        Msg = LoadResString(2125 + Sprache)
        MsgBox Msg                  'Gerbing 15.11.2004
        TSituation.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If InStr(1, TOrt.Text, "%%%", vbTextCompare) <> 0 And InStr(1, TOrt.Text, "&&&", vbTextCompare) <> 0 Then
        'msg = "Sie dürfen %%% und &&& nicht gleichzeitig verwenden"
        Msg = LoadResString(2125 + Sprache)
        MsgBox Msg                  'Gerbing 15.11.2004
        TOrt.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If InStr(1, TLand.Text, "%%%", vbTextCompare) <> 0 And InStr(1, TLand.Text, "&&&", vbTextCompare) <> 0 Then
        'msg = "Sie dürfen %%% und &&& nicht gleichzeitig verwenden"
        Msg = LoadResString(2125 + Sprache)
        MsgBox Msg                  'Gerbing 15.11.2004
        TLand.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If InStr(1, TPersonen.Text, "%%%", vbTextCompare) <> 0 Or InStr(1, TPersonen.Text, "&&&", vbTextCompare) <> 0 Then
        'msg = "Benutzen Sie anstelle von %%% bzw &&& 'Weitere Filter'..."
        Msg = LoadResString(2126 + Sprache)
        MsgBox Msg                  'Gerbing 15.11.2004
        TPersonen.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If InStr(1, TBegriff.Text, "%%%", vbTextCompare) <> 0 And InStr(1, TBegriff.Text, "&&&", vbTextCompare) <> 0 Then
        'msg = "Sie dürfen %%% und &&& nicht gleichzeitig verwenden"
        Msg = LoadResString(2125 + Sprache)
        MsgBox Msg                  'Gerbing 15.11.2004
        TBegriff.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    If chkFensterGrößeÄnderbar = 0 Then                 '1=aktiviert           'Gerbing 13.04.2015
        ShowTitleBar False, True                        'taskbar unvisible, Foto
    End If
    
    If CheckDifferenzen.Value = 1 Then              'Gerbing 20.09.2004
        If gblnSQLServerVersion = True Then
            'beim sql server charindex verwenden
            'CharIndex hat andere Parameterreihenfolge als InStr
            'SQLText = "SELECT Fotos.* From Fotos WHERE CharIndex(jahr,Dateiname,1)=0 ORDER BY Jahr,Dateiname ASC;" 'Gerbing 16.06.2005
            SQLText = "SELECT Fotos.* From Fotos WHERE CharIndex(" & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & ")=0" & " ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC;" 'Gerbing 08.12.2005
        Else
            'SQLText = "SELECT Fotos.* From Fotos WHERE instr(1,Dateiname, jahr)=0 ORDER BY Jahr,Dateiname ASC;" 'Gerbing 16.06.2005
            SQLText = "SELECT Fotos.* From Fotos WHERE instr(1," & LoadResString(1028 + Sprache) & ", " & LoadResString(1023 + Sprache) & ")=0" & " ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC;" 'Gerbing 08.12.2005
        End If
        SQLWurdeBearbeitet = True
    End If
    
    If FrameGespeicherteAbfragen.Visible = True Then        'Gerbing 16.06.2005
        If txtSQLGespeicherteAbfrage = "" Then              'Gerbing 20.06.2006
            'MsgBox "Sie haben keine gespeicherte Abfrage ausgewählt"
            MsgBox LoadResString(2253 + Sprache)
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        SQLText = txtSQLGespeicherteAbfrage
        'in dem Abschnitt vor 'FROM könnten '*' zeichen auftauchen, die müssen bleiben
        pos = InStr(1, SQLText, "FROM", vbTextCompare)      'Gerbing 30.05.2006
        If pos <> 0 Then
            strLinks = Left(SQLText, pos - 1)
            strRechts = Right(SQLText, Len(SQLText) - pos + 1)
        End If
        'Replace von '*' zeichen durch '%' zeichen, nur wenn mindestens 2 auftreten
        'sonst replaced man zB having Count(*)
        pos = InStr(1, strRechts, "*")                      'Gerbing 20.06.2006
        pos1 = InStr(pos + 1, strRechts, "*")
        If pos <> 0 And pos1 <> 0 Then
            strRechts = Replace(strRechts, "*", "%")
        End If
        If strLinks <> "" Then
            SQLText = strLinks
        End If
        If strRechts <> "" Then
            SQLText = SQLText & strRechts
        End If
        'Wenn eine gespeicherte Abfrage keinen Abschnitt 'ORDER BY' enthält, wird ORDER BY Dateiname angefügt
        pos = InStr(1, SQLText, "ORDER BY", vbTextCompare)      'Gerbing 19.01.2007
        If pos = 0 Then
            'es gibt kein 'ORDER BY'
            pos1 = InStr(1, SQLText, ";", vbTextCompare)
            If pos1 <> 0 Then
                'es gibt ein Semikolon 'ORDER BY' wird anstelle des Semikolon angehängt
                'SQLtext = SQLText & " ORDER BY Jahr,Dateiname ASC"
                SQLText = Mid(SQLText, 1, pos1 - 1) & " ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC"
            Else
                'ORDER BY' wird ans Ende angehängt
                'SQLtext = SQLText & " ORDER BY Jahr,Dateiname ASC"
                SQLText = SQLText & " ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC"
            End If
        End If
        SQLWurdeBearbeitet = True
    End If
'------------------------------------------------------------------------------------------------------------
    Query.OKGewählt = True
    '----------------------------------------------
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                         'Gerbing 20.11.2008
        On Error Resume Next                                                                'Gerbing 20.11.2008
        Form1.Top = 0                                                                       'Gerbing 20.11.2008
        Form1.Left = 0
        screenWidth = GetDeviceCaps(Me.hDC, HORZRES)                                        'Gerbing 29.03.2012
        screenHeight = GetDeviceCaps(Me.hDC, VERTRES)                                       'Gerbing 29.03.2012
        Form1.SGVH = screenWidth / screenHeight                                             'Gerbing 24.09.2009
        'Form1.Width = ScreenWidth / 2 * Screen.TwipsPerPixelX                          'Gerbing 09.05.2012
        Form1.width = (screenWidth - 100) * Screen.TwipsPerPixelX                       'Gerbing 09.05.2012
        Form1.height = screenHeight * Screen.TwipsPerPixelY                                 'Gerbing 20.11.2008
        On Error GoTo SQLERR                                                                'Gerbing 20.11.2008
        lstFensterTitel.Clear
        CurrWnd = GetWindow(Form1.hWnd, GW_HWNDFIRST)
        While CurrWnd <> 0
        'In dieser Schleife werden gefüllt
        'lstFensterTitel enthält Application Titles
            TitelLength = GetWindowTextLength(CurrWnd)
            FensterTitel = Space$(TitelLength + 1)
            TitelLength = GetWindowText(CurrWnd, FensterTitel, TitelLength + 1)
            If TitelLength > 0 Then
                If Left(FensterTitel, Len(LoadResString(1001 + Sprache))) = LoadResString(1001 + Sprache) Then
                    'ob "FotoAlbum-" im Fenstertitel steht
                    lstFensterTitel.AddItem FensterTitel
                End If
            End If
            CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
            X = DoEvents()
        Wend
        If lstFensterTitel.ListCount > 1 Then
            On Error Resume Next                                                            'Gerbing 20.11.2008
            Form1.Left = screenWidth / 2
            On Error GoTo SQLERR                                                               'Gerbing 20.11.2008
        End If
    End If                                                                                  'Gerbing 20.11.2008
    '----------------------
    'QueryJedesFeld.AbbrechenGewählt = False
    
    If SQLWurdeBearbeitet = True Or SQLBearbeitenZähler > 0 Then        'Gerbing 09.06.2004
        SQL = SQLText
        If SQL <> "" Then                                               'Gerbing 07.11.2004
            GoTo SQLWurdeBearbeitet
        End If
    End If
        
    Plus1 = " And "
    Plus2 = " And "
    Plus3 = " And "
    Plus4 = " And "
    Plus5 = " And "
    If JUnd.Value = False Then Plus1 = " Or "
    If SUnd.Value = False Then Plus2 = " Or "
    If OUnd.Value = False Then Plus3 = " Or "
    If LUnd.Value = False Then Plus4 = " Or "
    If SWFUnd.Value = False Then Plus5 = " Or "
'Jahr--------------------------------------------------------------------------
    rc = JahreszahlPrüfen                                               'Gerbing 08.11.2012
    If rc <> 0 Then Exit Sub
'Situation--------------------------------------------------------------------
    TSituation.Text = RTrim(TSituation.Text)                                      'Gerbing 22.05.2004
    'TSituation = UCase(TSituation)
    If StrComp(TSituation.Text, "Null", vbTextCompare) = 0 Then
    'If TSituation = "NULL" Then
        'SQL = SQL & Plus & " (Situation Is Null OR Situation = "'")"   'Gerbing 06.07.2004
        SQL = SQL & Plus & " (" & LoadResString(1024 + Sprache) & " Is Null OR " & LoadResString(1024 + Sprache) & " = '')" 'Gerbing 15.02.2012
        Plus = Plus2
    Else
        'If TSituation <> UCase(LoadResString(1110 + Sprache)) And TSituation <> "*" Then    '1110=beliebig
        If StrComp(TSituation.Text, (LoadResString(1110 + Sprache)), vbTextCompare) <> 0 And TSituation.Text <> "*" Then '1110=beliebig
            TSplitSit = Split(TSituation.Text, "%%%", -1, vbTextCompare)        'Gerbing 15.11.2004
            If InStr(1, TSituation.Text, "%%%", vbTextCompare) <> 0 Then
                strVerk = " OR "
                TSplitSit = Split(TSituation.Text, "%%%", -1, vbTextCompare)        'Gerbing 15.11.2004
            End If
            If InStr(1, TSituation.Text, "&&&", vbTextCompare) <> 0 Then
                strVerk = " AND "
                TSplitSit = Split(TSituation.Text, "&&&", -1, vbTextCompare)        'Gerbing 15.11.2004
            End If
            If UBound(TSplitSit) > 0 Then
                If gblnProversion = False Then                                  'Gerbing 10.06.2005
                    'msg = "Eingeben von Suchkriterien (Profi) ist in der Shareware-Version nicht möglich." & NL
                    Msg = LoadResString(2129 + Sprache) & NL
                    Msg = Msg & TSituation.Text & NL
                    'msg = msg & "Diese Suchkriterien werden nicht gefunden."
                    Msg = Msg & LoadResString(2130 + Sprache)
                    'MsgBox Msg
                    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                    GoTo SituationLight
                End If
                #If Proversion Then
                If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
                    SQL = SQL & Plus & "("
                    For n = LBound(TSplitSit) To UBound(TSplitSit)
                        If n <> 0 Then
                            SQL = SQL & strVerk
                        End If
                        strTemp = Trim(TSplitSit(n))
                        'SQL = SQL & " Situation Like " & "'" & "%" & strTemp & "%" & "'"
                        If gblnSQLServerVersion = True Then
                            SQL = SQL & " " & LoadResString(1024 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'"
                        Else
                            SQL = SQL & " " & LoadResString(1024 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'"
                        End If
                    Next n
                    SQL = SQL & ")"
                End If
                #End If
            Else
SituationLight:
                'SQL = SQL & Plus & " Situation Like " & "'" & "%" & TSituation & "%" & "'"    'Gerbing 25.05.2004
                If gblnSQLServerVersion = True Then
                    SQL = SQL & Plus & " " & LoadResString(1024 + Sprache) & " Like " & "N'" & "%" & TSituation.Text & "%" & "'"    'Gerbing 08.11.2005
                Else
                    SQL = SQL & Plus & " " & LoadResString(1024 + Sprache) & " Like " & "'" & "%" & TSituation.Text & "%" & "'"    'Gerbing 08.11.2005
                End If
            End If
            Plus = Plus2
        End If
    End If
'Ort--------------------------------------------------------------------------
    TOrt.Text = RTrim(TOrt.Text)                                  'Gerbing 22.05.2004
    'TOrt.Text = UCase(TOrt.Text)
    If StrComp(TOrt.Text, "Null", vbTextCompare) = 0 Then                    'Gerbing 08.11.2012
    'If TOrt.Text = "NULL" Then
        SQL = SQL & Plus & " (" & LoadResString(1025 + Sprache) & " Is Null OR " & LoadResString(1025 + Sprache) & " = '')" 'Gerbing 15.02.2012
        Plus = Plus3
    Else
        'If TOrt.Text <> UCase(LoadResString(1110 + Sprache)) And TOrt.Text <> "*" Then    '1110=beliebig
        If StrComp(TOrt.Text, (LoadResString(1110 + Sprache)), vbTextCompare) <> 0 And TOrt.Text <> "*" Then '1110=beliebig
            TSplitOrt = Split(TOrt.Text, "%%%", -1, vbTextCompare)           'Gerbing 15.11.2004
            If InStr(1, TOrt.Text, "%%%", vbTextCompare) <> 0 Then
                strVerk = " OR "
                TSplitOrt = Split(TOrt.Text, "%%%", -1, vbTextCompare)       'Gerbing 15.11.2004
            End If
            If InStr(1, TOrt.Text, "&&&", vbTextCompare) <> 0 Then
                strVerk = " AND "
                TSplitOrt = Split(TOrt.Text, "&&&", -1, vbTextCompare)       'Gerbing 15.11.2004
            End If
            If UBound(TSplitOrt) > 0 Then
                If gblnProversion = False Then                          'Gerbing 10.06.2005
                    'msg = "Eingeben von Suchkriterien (Profi) ist in der Shareware-Version nicht möglich." & NL
                    Msg = LoadResString(2129 + Sprache) & NL
                    Msg = Msg & TOrt.Text & NL
                    'msg = msg & "Diese Suchkriterien werden nicht gefunden."
                    Msg = Msg & LoadResString(2130 + Sprache)
                    'MsgBox Msg
                    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                    GoTo OrtLight
                End If
                #If Proversion Then
                If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
                    SQL = SQL & Plus & "("
                    For n = LBound(TSplitOrt) To UBound(TSplitOrt)
                        If n <> 0 Then
                            SQL = SQL & strVerk
                        End If
                        strTemp = Trim(TSplitOrt(n))
                        'SQL = SQL & " Ort Like " & "'" & "%" & strTemp & "%" & "'"   'Gerbing 08.11.2005
                        If gblnSQLServerVersion = True Then
                            SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'"
                        Else
                            SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'"
                        End If
                    Next n
                    SQL = SQL & ")"
                End If
                #End If
            Else
OrtLight:
                'SQL = SQL & Plus & " Ort Like " & "'" & "%" & TOrt.Text & "%" & "'"    'Gerbing 25.05.2004
                If gblnSQLServerVersion = True Then
                    SQL = SQL & Plus & " " & LoadResString(1025 + Sprache) & " Like " & "N'" & "%" & TOrt.Text & "%" & "'"    'Gerbing 25.05.2004
                Else
                    SQL = SQL & Plus & " " & LoadResString(1025 + Sprache) & " Like " & "'" & "%" & TOrt.Text & "%" & "'"    'Gerbing 25.05.2004
                End If
            End If
            Plus = Plus3
        End If
    End If
'Land--------------------------------------------------------------------------
    TLand.Text = RTrim(TLand.Text)                                'Gerbing 22.05.2004
    'TLand.Text = UCase(TLand.Text)
    If StrComp(TLand.Text, "Null", vbTextCompare) = 0 Then
    'If TLand.Text = "NULL" Then
        'SQL = SQL & Plus & " (Land Is Null OR Land = "'")" 'Gerbing 06.07.2004
        SQL = SQL & Plus & " (" & LoadResString(1026 + Sprache) & " Is Null OR " & LoadResString(1026 + Sprache) & " = '')" 'Gerbing 15.02.2012
        Plus = Plus4
    Else
        'If TLand.Text <> UCase(LoadResString(1110 + Sprache)) And TLand.Text <> "*" Then  '1110=beliebig
        If StrComp(TLand.Text, (LoadResString(1110 + Sprache)), vbTextCompare) <> 0 And TLand.Text <> "*" Then  '1110=beliebig
            TSplitLand = Split(TLand.Text, "%%%", -1, vbTextCompare)        'Gerbing 15.11.2004
            If InStr(1, TLand.Text, "%%%", vbTextCompare) <> 0 Then
                strVerk = " OR "
                TSplitLand = Split(TLand.Text, "%%%", -1, vbTextCompare)        'Gerbing 15.11.2004
            End If
            If InStr(1, TLand.Text, "&&&", vbTextCompare) <> 0 Then
                strVerk = " AND "
                TSplitLand = Split(TLand.Text, "&&&", -1, vbTextCompare)        'Gerbing 15.11.2004
            End If
            If UBound(TSplitLand) > 0 Then
                If gblnProversion = False Then                              'Gerbing 10.06.2005
                    'msg = "Eingeben von Suchkriterien (Profi) ist in der Shareware-Version nicht möglich." & NL
                    Msg = LoadResString(2129 + Sprache) & NL
                    Msg = Msg & TLand.Text & NL
                    'msg = msg & "Diese Suchkriterien werden nicht gefunden."
                    Msg = Msg & LoadResString(2130 + Sprache)
                    'MsgBox Msg
                    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                    GoTo LandLight
                End If
                #If Proversion Then
                If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
                    SQL = SQL & Plus & "("
                    For n = LBound(TSplitLand) To UBound(TSplitLand)
                        If n <> 0 Then
                            SQL = SQL & strVerk
                        End If
                        strTemp = Trim(TSplitLand(n))
                        'SQL = SQL & " Land Like " & "'" & "%" & strTemp & "%" & "'"
                        If gblnSQLServerVersion = True Then
                            SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'"
                        Else
                            SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'"
                        End If
                    Next n
                    SQL = SQL & ")"
                End If
                #End If
            Else
LandLight:
                'SQL = SQL & Plus & " Land Like " & "'" & "%" & TLand.Text & "%" & "'"    'Gerbing 25.05.2004
                If gblnSQLServerVersion = True Then
                    SQL = SQL & Plus & " " & LoadResString(1026 + Sprache) & " Like " & "N'" & "%" & TLand.Text & "%" & "'"    'Gerbing 08.11.2005
                Else
                    SQL = SQL & Plus & " " & LoadResString(1026 + Sprache) & " Like " & "'" & "%" & TLand.Text & "%" & "'"    'Gerbing 08.11.2005
                End If
            End If
            Plus = Plus4
        End If
    End If
'SW/F--------------------------------------------------------------------------
    'If TSWF <> UCase(LoadResString(1110 + Sprache)) And TSWF <> "*" Then    '1110=beliebig
    If StrComp(TSWF, (LoadResString(1110 + Sprache)), vbTextCompare) <> 0 And TSWF <> "*" Then  '1110=beliebig
        'SQL = SQL & Plus & " SWF = " & "'" & TSWF & "'"
        SQL = SQL & Plus & " " & LoadResString(1029 + Sprache) & " = " & "'" & TSWF & "'"   'Gerbing 08.11.2005
        Plus = Plus5
    End If
    #If Proversion Then                                                         'Gerbing 12.04.2006
    If gblnVollversion = True And gblnProversion = True Then
        If CheckAudioFileExists.Value = 1 Then
            SQL = SQL & " AND AudioFileExists<>0 "
        End If
    End If
    #End If
    #If Proversion Then
    If gblnVollversion = True And gblnProversion = True Then                    'Gerbing 22.02.2006
'NutzerdefinierteFelder-----------------------------------------                'Gerbing 14.02.2005
        'Suche nach GEODaten innerhalb eines Rechtecks                          'Gerbing 05.09.2016
        'zB     gstrGEOStartPunkt=50.83517,12.81463         GPSLatitude,GPSLongitude   Breite,Länge
        '       gstrGEOEndPunkt=50.83017,12.82298
        If gstrGEOStartPunkt <> "" And gstrGEOEndPunkt <> "" Then               'Gerbing 12.10.2016
            SQL = SQL & " AND ("
            pos = InStr(1, gstrGEOStartPunkt, ",")
            pos2 = InStr(1, gstrGEOEndPunkt, ",")
            LinksOben = Mid(gstrGEOStartPunkt, 1, pos - 1)
            LinksUnten = Mid(gstrGEOEndPunkt, 1, pos2 - 1)
            RechtsOben = Mid(gstrGEOStartPunkt, pos + 1, Len(gstrGEOStartPunkt) - pos)
            RechtsUnten = Mid(gstrGEOEndPunkt, pos2 + 1, Len(gstrGEOEndPunkt) - pos2)
            SQL = SQL & "GPSLatitude <=" & LinksOben & " AND "
            SQL = SQL & "GPSLatitude >=" & LinksUnten & " AND "
            SQL = SQL & "GPSLongitude >=" & RechtsOben & " And "
            SQL = SQL & "GPSLongitude <=" & RechtsUnten
            SQL = SQL & ")"
        End If
        If ND.AnzahlFelder <> 0 Then
            SQL = SQL & " AND ("
            Do
                ND.Combo1.Text = RTrim(ND.Combo1.Text)
                ND.Combo1.Text = UCase(ND.Combo1.Text)
                If ND.Combo1.Text = "NULL" Then
                    'SQL = SQL & "[" & ND.cmbFeld1 & "]" & " Is Null OR " & "[" & ND.cmbFeld1 & "]" & "= "'""
                    SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & " Is Null "
                Else
                    'Die Funktion DatenTypDate gibt den Datentyp zurück         'Gerbing 04.02.2007
                    Select Case DatenTypDate(ND.cmbFeld1, ND.Combo1)
                        Case 135                                                'Datum bei sql server
                            SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & FDF
                        Case 8  '8=dbDate                                               'vorher dbDate Gerbing 23.11.2017
                            SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & FDF
                        Case 12         '12=Hyperlink
                            If ND.cmbVG1.Text = "=" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & " Like " & "'" & "%" & ND.Combo1.Text & "%" & "'"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & " Not Like " & "'" & "%" & ND.Combo1.Text & "%" & "'"
                            End If
                        Case 10, 202    '10=Text 202 Text bei SQL-Server        'Gerbing 25.10.2013
                            'strTemp = Replace(ND.Combo1.Text, ",", ".")        'Gerbing 02.09.2016 bei Text ist Komma richtig
                            strTemp = ND.Combo1.Text
                            If ND.cmbVG1.Text = "like" Then                                                         'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "'%" & strTemp & "%'"   'Gerbing 08.05.2019
                            Else                                                                                    'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "'" & strTemp & "'"     'Gerbing 08.05.2019
                            End If                                                                                  'Gerbing 08.05.2019
                        Case 1  '1=Boolean
                            If UCase(ND.Combo1) = "WAHR" Or UCase(ND.Combo1) = "TRUE" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & " True "
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & " False "
                            End If
                        Case 11             '11=Boolean bei SQL-Server                                      'Gerbing 25.10.2013
                            If ND.Combo1.Text = "0" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "0"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "1"
                            End If
                        Case 5, 6  '5=dbCurrency,6=money bei SQL-Server                                         'Gerbing 25.10.2013
                            'wenn in ND.Combo1 Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo1.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & strTemp
                        Case Else
                            'alles hier ist Zahlen
                            'wenn in ND.Combo1 Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo1.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & strTemp
                    End Select
                End If
                If ND.AnzahlFelder = 1 Then Exit Do
                '----------------------------------
                SQL = SQL & ND.Plus1
                ND.Combo2.Text = RTrim(ND.Combo2.Text)
                ND.Combo2.Text = UCase(ND.Combo2.Text)
                If ND.Combo2.Text = "NULL" Then
                    'SQL = SQL & "[" & ND.cmbFeld2 & "]" & " Is Null OR " & "[" & ND.cmbFeld2 & "]" & "= "'""
                    SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & " Is Null "
                Else
                    'Die Funktion DatenTypDate gibt den Datentyp zurück         'Gerbing 04.02.2007
                    Select Case DatenTypDate(ND.cmbFeld2.Text, ND.Combo2.Text)
                        Case 135                                                'Datum bei sql server
                            SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & FDF
                        Case 8  '8=dbDate
                            SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & FDF
                        Case 12         '12=Hyperlink
                            If ND.cmbVG2.Text = "=" Then
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & " Like " & "'" & "%" & ND.Combo2.Text & "%" & "'"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & " Not Like " & "'" & "%" & ND.Combo2.Text & "%" & "'"
                            End If
                        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 25.10.2013
                            'strTemp = Replace(ND.Combo2.Text, ",", ".")                                    'Gerbing 02.09.2016 bei Text ist Komma richtig
                            strTemp = ND.Combo2.Text
                            If ND.cmbVG2.Text = "like" Then                                                         'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & "'%" & strTemp & "%'"   'Gerbing 08.05.2019
                            Else                                                                                    'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & "'" & strTemp & "'"     'Gerbing 08.05.2019
                            End If                                                                                  'Gerbing 08.05.2019
                        Case 1  '1=dbBoolean
                            If UCase(ND.Combo2.Text) = "WAHR" Or UCase(ND.Combo2.Text) = "TRUE" Then
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & " True "
                            Else
                                SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & " False "
                            End If
                        Case 11             '11=Boolean bei SQL-Server                                      'Gerbing 25.10.2013
                            If ND.Combo1.Text = "0" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "0"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "1"
                            End If
                        Case 5, 6  '5=dbCurrency6=money bei SQL-Server                                         'Gerbing 25.10.2013
                            'wenn in ND.Combo2.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo2.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & strTemp
                        Case Else
                            'alles hier ist Zahlen
                            'wenn in ND.Combo2.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo2.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld2.Text & "]" & ND.cmbVG2.Text & strTemp
                    End Select
                End If
                If ND.AnzahlFelder = 2 Then Exit Do
                '----------------------------------
                SQL = SQL & ND.Plus2
                ND.Combo3.Text = RTrim(ND.Combo3.Text)
                ND.Combo3.Text = UCase(ND.Combo3.Text)
                If ND.Combo3.Text = "NULL" Then
                    'SQL = SQL & "[" & ND.cmbFeld3 & "]" & " Is Null OR " & "[" & ND.cmbFeld3 & "]" & "= "'""
                    SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & " Is Null "
                Else
                    'Die Funktion DatenTypDate gibt den Datentyp zurück         'Gerbing 04.02.2007
                    Select Case DatenTypDate(ND.cmbFeld3.Text, ND.Combo3.Text)
                        Case 135                                                'Datum bei sql server
                            SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & FDF
                        Case 8  '8=dbDate
                            SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & FDF
                        Case 12         '12=Hyperlink
                            If ND.cmbVG3.Text = "=" Then
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & " Like " & "'" & "%" & ND.Combo3.Text & "%" & "'"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & " Not Like " & "'" & "%" & ND.Combo3.Text & "%" & "'"
                            End If
                        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 25.10.2013
                            'strTemp = Replace(ND.Combo3.Text, ",", ".")                                    'Gerbing 02.09.2016 bei Text ist Komma richtig
                            strTemp = ND.Combo3.Text
                            If ND.cmbVG3.Text = "like" Then                                                         'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & "'%" & strTemp & "%'"   'Gerbing 08.05.2019
                            Else                                                                                    'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & "'" & strTemp & "'"     'Gerbing 08.05.2019
                            End If                                                                                  'Gerbing 08.05.2019
                        Case 1
                            If UCase(ND.Combo3.Text) = "WAHR" Or UCase(ND.Combo3.Text) = "TRUE" Then
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & " True "
                            Else
                                SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & " False "
                            End If
                        Case 11             '11=Boolean bei SQL-Server                                      'Gerbing 25.10.2013
                            If ND.Combo1.Text = "0" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "0"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "1"
                            End If
                        Case 5, 6  '5=dbCurrency6=money bei SQL-Server                                         'Gerbing 25.10.2013
                            'wenn in ND.Combo3.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo3.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & strTemp
                        Case Else
                            'alles hier ist Zahlen
                            'wenn in ND.Combo3.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo3.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld3.Text & "]" & ND.cmbVG3.Text & strTemp
                    End Select
                End If
                If ND.AnzahlFelder = 3 Then Exit Do
                '----------------------------------
                SQL = SQL & ND.Plus3
                ND.combo4.Text = RTrim(ND.combo4.Text)
                ND.combo4.Text = UCase(ND.combo4.Text)
                If ND.combo4.Text = "NULL" Then
                    'SQL = SQL & "[" & ND.cmbFeld4 & "]" & " Is Null OR " & "[" & ND.cmbFeld4 & "]" & "= "'""
                    SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & " Is Null "
                Else
                    'Die Funktion DatenTypDate gibt den Datentyp zurück         'Gerbing 04.02.2007
                    Select Case DatenTypDate(ND.cmbFeld4.Text, ND.combo4.Text)
                        Case 135                                                'Datum bei sql server
                            SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & FDF
                        Case 8  '8=dbDate
                            SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & FDF
                        Case 12         '12=Hyperlink
                            If ND.cmbVG4.Text = "=" Then
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & " Like " & "'" & "%" & ND.combo4.Text & "%" & "'"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & " Not Like " & "'" & "%" & ND.combo4.Text & "%" & "'"
                            End If
                        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 25.10.2013
                            'strTemp = Replace(ND.Combo4.Text, ",", ".")                                    'Gerbing 02.09.2016 bei Text ist Komma richtig
                            strTemp = ND.combo4.Text
                            If ND.cmbVG4.Text = "like" Then                                                         'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & "'%" & strTemp & "%'"   'Gerbing 08.05.2019
                            Else                                                                                    'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & "'" & strTemp & "'"     'Gerbing 08.05.2019
                            End If                                                                                  'Gerbing 08.05.2019
                        Case 1  '1=dbBoolean
                            If UCase(ND.combo4.Text) = "WAHR" Or UCase(ND.combo4.Text) = "TRUE" Then
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & " True "
                            Else
                                SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & " False "
                            End If
                        Case 11             '11=Boolean bei SQL-Server                                      'Gerbing 25.10.2013
                            If ND.Combo1.Text = "0" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "0"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "1"
                            End If
                        Case 5, 6  '5=dbCurrency,6=money bei SQL-Server                                         'Gerbing 25.10.2013
                            'wenn in ND.combo4.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.combo4.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & strTemp
                        Case Else
                            'alles hier ist Zahlen
                            'wenn in ND.combo4.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.combo4.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld4.Text & "]" & ND.cmbVG4.Text & strTemp
                    End Select
                End If
                If ND.AnzahlFelder = 4 Then Exit Do
                '----------------------------------
                SQL = SQL & ND.Plus4
                ND.Combo5.Text = RTrim(ND.Combo5.Text)
                ND.Combo5.Text = UCase(ND.Combo5.Text)
                If ND.Combo5.Text = "NULL" Then
                    'SQL = SQL & "[" & ND.cmbFeld5 & "]" & " Is Null OR " & "[" & ND.cmbFeld5 & "]" & "= "'""
                    SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & " Is Null "
                Else
                    'Die Funktion DatenTypDate gibt den Datentyp zurück         'Gerbing 04.02.2007
                    Select Case DatenTypDate(ND.cmbFeld5.Text, ND.Combo5.Text)
                        Case 135                                                'Datum bei sql server
                            SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & FDF
                        Case 8  '8=dbDate
                            SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & FDF
                        Case 12         '12=Hyperlink
                            If ND.cmbVG5.Text = "=" Then
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & " Like " & "'" & "%" & ND.Combo5.Text & "%" & "'"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & " Not Like " & "'" & "%" & ND.Combo5.Text & "%" & "'"
                            End If
                        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 25.10.2013
                            'strTemp = Replace(ND.Combo5.Text, ",", ".")                                    'Gerbing 02.09.2016 bei Text ist Komma richtig
                            strTemp = ND.Combo5.Text
                            If ND.cmbVG5.Text = "like" Then                                                         'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & "'%" & strTemp & "%'"   'Gerbing 08.05.2019
                            Else                                                                                    'Gerbing 08.05.2019
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & "'" & strTemp & "'"     'Gerbing 08.05.2019
                            End If                                                                                  'Gerbing 08.05.2019
                        Case 1  '1=dbBoolean
                            If UCase(ND.Combo5.Text) = "WAHR" Or UCase(ND.Combo5.Text) = "TRUE" Then
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & " True "
                            Else
                                SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & " False "
                            End If
                        Case 11             '11=Boolean bei SQL-Server                                      'Gerbing 25.10.2013
                            If ND.Combo1.Text = "0" Then
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "0"
                            Else
                                SQL = SQL & "[" & ND.cmbFeld1.Text & "]" & ND.cmbVG1.Text & "1"
                            End If
                        Case 5, 6  '5=dbCurrency,6=money bei SQL-Server                                         'Gerbing 25.10.2013
                            'wenn in ND.Combo5.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo5.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & strTemp
                        Case Else
                            'alles hier ist Zahlen
                            'wenn in ND.Combo5.text Komma auftritt muss ich Punkt daraus machen
                            strTemp = Replace(ND.Combo5.Text, ",", ".")
                            SQL = SQL & "[" & ND.cmbFeld5.Text & "]" & ND.cmbVG5.Text & strTemp
                    End Select
                End If
                If ND.AnzahlFelder = 5 Then Exit Do
            Loop
            SQL = SQL & ")"
        End If
    End If
    #End If
'Personen----------------------------------------------------------------------
    TPersonen.Text = RTrim(TPersonen.Text)                        'Gerbing 22.05.2004
    'TPersonen.Text = UCase(TPersonen.Text)
    If StrComp(TPersonen.Text, "Null", vbTextCompare) = 0 Then
    'If TPersonen.Text = "NULL" Then
        'SQL = SQL & Plus & " (Personen Is Null OR Personen = "'")" 'Gerbing 06.07.2004
        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Is Null OR " & LoadResString(1027 + Sprache) & " = '')" 'Gerbing 15.02.2012
        'Plus = Plus2
    Else
        'beim SQL-Server gibt es keine doppelten Hochkomma im sql string ( Nicht Like "%person1%" sondern Like '%person1%')
        'If TPersonen.Text = "MEHRERE PERSONEN" Then                                             'Gerbing 15.06.2008
        If TPersonen.Text = LoadResString(1124 + Sprache) Then
            Select Case MP.AnzahlPersonen
                Case 2
'                    SQL = SQL & Plus & " (Personen Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"    'Gerbing 25.05.2004
'                    SQL = SQL & MP.Plus1 & " Personen Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'" & ")"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson1.Text & "%" & "'"  'Gerbing 25.05.2004
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson2.Text & "%" & "'" & ")"
                    Else
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"  'Gerbing 25.05.2004
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'" & ")"
                    End If
                Case 3
'                    SQL = SQL & Plus & " (Personen Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
'                    SQL = SQL & MP.Plus1 & " Personen Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
'                    SQL = SQL & MP.Plus2 & " Personen Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'" & ")"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson3.Text & "%" & "'" & ")"
                    Else
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'" & ")"
                    End If
                Case 4
'                    SQL = SQL & Plus & " (Personen Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
'                    SQL = SQL & MP.Plus1 & " Personen Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
'                    SQL = SQL & MP.Plus2 & " Personen Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'"
'                    SQL = SQL & MP.Plus3 & " Personen Like " & "'" & "%" & MP.TPerson4.Text & "%" & "'" & ")"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson3.Text & "%" & "'"
                        SQL = SQL & MP.Plus3 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson4.Text & "%" & "'" & ")"
                    Else
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'"
                        SQL = SQL & MP.Plus3 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson4.Text & "%" & "'" & ")"
                    End If
                Case 5
'                    SQL = SQL & Plus & " (Personen Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
'                    SQL = SQL & MP.Plus1 & " Personen Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
'                    SQL = SQL & MP.Plus2 & " Personen Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'"
'                    SQL = SQL & MP.Plus3 & " Personen Like " & "'" & "%" & MP.TPerson4.Text & "%" & "'"
'                    SQL = SQL & MP.Plus4 & " Personen Like " & "'" & "%" & MP.TPerson5.Text & "%" & "'" & ")"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson3.Text & "%" & "'"
                        SQL = SQL & MP.Plus3 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson4.Text & "%" & "'"
                        SQL = SQL & MP.Plus4 & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & MP.TPerson5.Text & "%" & "'" & ")"
                    Else
                        SQL = SQL & Plus & " (" & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson1.Text & "%" & "'"
                        SQL = SQL & MP.Plus1 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson2.Text & "%" & "'"
                        SQL = SQL & MP.Plus2 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson3.Text & "%" & "'"
                        SQL = SQL & MP.Plus3 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson4.Text & "%" & "'"
                        SQL = SQL & MP.Plus4 & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & MP.TPerson5.Text & "%" & "'" & ")"
                    End If
            End Select
        Else
            'If TPersonen.Text <> UCase(LoadResString(1110 + Sprache)) And TPersonen.Text <> "*" Then  '1110=beliebig
            If StrComp(TPersonen.Text, (LoadResString(1110 + Sprache)), vbTextCompare) <> 0 And TPersonen.Text <> "*" Then '1110=beliebig
                If CheckVollesWort.Value = 1 Then       '1=aktiviert
                    SQL = SQL & Plus
'                    SQL = SQL & " (Personen Like " & "'" & "* " & TPersonen.Text & " *" & "'" & " Or "  '" * Elke * ""
'                    SQL = SQL & " Personen Like " & "'" & TPersonen.Text & " *" & "'" & " OR "           '"Elke *"
'                    SQL = SQL & " Personen = " & "'" & TPersonen.Text & "'" & " OR "               '"Elke"
'                    SQL = SQL & " Personen Like " & "'" & "* " & TPersonen.Text & "'" & ")"        '"* Elke"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & " (" & LoadResString(1027 + Sprache) & " Like " & "N'" & "% " & TPersonen.Text & " %" & "'" & " Or "  '" % Elke % ""
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & TPersonen.Text & " %" & "'" & " OR "           '"Elke %"
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " = " & "N'" & TPersonen.Text & "'" & " OR "               '"Elke"
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "% " & TPersonen.Text & "'" & ")"        '"% Elke"
                    Else
                        SQL = SQL & " (" & LoadResString(1027 + Sprache) & " Like " & "'" & "% " & TPersonen.Text & " %" & "'" & " Or "  '" % Elke % ""
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "'" & TPersonen.Text & " %" & "'" & " OR "           '"Elke %"
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " = " & "'" & TPersonen.Text & "'" & " OR "               '"Elke"
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "% " & TPersonen.Text & "'" & ")"        '"% Elke"
                    End If
                Else
                    'SQL = SQL & Plus & " Personen Like " & "'" & "%" & TPersonen.Text & "%" & "'"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & Plus & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & TPersonen.Text & "%" & "'"
                    Else
                        SQL = SQL & Plus & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & TPersonen.Text & "%" & "'"
                    End If
                End If
            End If
        End If
    End If
'Suche nach DDatum--------------------------------------------------------------- Gerbing 10.10.2004
    'zB (Fotos.DDatum)>#2/10/2004# And (Fotos.DDatum)<#4/24/2004#   10.02.2004 bis 24.04.2004
    If MP.optDatumEinbeziehen.Value = True Then
        DatumVon = MP.txtDatumVon
        DatumBis = MP.txtDatumBis
        SQL = SQL & " AND "
        If gblnSQLServerVersion = True Then
            'Beim sql server ist das Format yyyymmdd von der Landessprache unabhängig
            'm muss unbedingt 2-stellig sein und d muss unbedingt 2-stellig sein
            'beim 22.01.2012 kommt ohne Formatierung 2012122 beim Monat wird eine Null weggelassen
            VPartyyyy = DatePart("yyyy", DatumVon)
            VPartmm = DatePart("m", DatumVon)
            If Len(VPartmm) = 1 Then
                VPartmm = "0" & VPartmm
            End If
            VPartdd = DatePart("d", DatumVon)
            If Len(VPartdd) = 1 Then
                VPartdd = "0" & VPartdd
            End If
            BPartyyyy = DatePart("yyyy", DatumBis)
            BPartmm = DatePart("m", DatumBis)
            If Len(BPartmm) = 1 Then
                BPartmm = "0" & BPartmm
            End If
            BPartdd = DatePart("d", DatumBis)
            If Len(BPartdd) = 1 Then
                BPartdd = "0" & BPartdd
            End If
            Von = VPartyyyy & VPartmm & VPartdd
            Bis = BPartyyyy & BPartmm & BPartdd
            'beim sql server wird nicht # benutzt sondern '
            'SQL = SQL & "(DDatum>='" & Von & "' And DDatum<='" & Bis & "')"
            SQL = SQL & "(" & LoadResString(1032 + Sprache) & ">='" & Von & "' And " & LoadResString(1032 + Sprache) & "<='" & Bis & "')" 'Gerbing 29.12.2011
        Else
            Von = DatePart("m", DatumVon) & "/" & DatePart("d", DatumVon) & "/" & DatePart("yyyy", DatumVon)
            Bis = DatePart("m", DatumBis) & "/" & DatePart("d", DatumBis) & "/" & DatePart("yyyy", DatumBis)
            'SQL = SQL & "(DDatum>=#" & Von & "# And DDatum<=#" & Bis & "#)"
            SQL = SQL & "(" & LoadResString(1032 + Sprache) & ">=#" & Von & "# And " & LoadResString(1032 + Sprache) & "<=#" & Bis & "#)" 'Gerbing 08.11.2005
        End If
    End If
'Suche Begriff in jedem Feld aktiviert?------------------------------------------
    If CheckSucheJedesFeld.Value = 1 Then           'Suche Begriff in jedem Feld aktiviert?
        TSplitBeg = Split(TBegriff.Text, "%%%", -1, vbTextCompare)           'Gerbing 15.11.2004
        If InStr(1, TBegriff.Text, "%%%", vbTextCompare) <> 0 Then
            strVerk = " OR "
            TSplitBeg = Split(TBegriff.Text, "%%%", -1, vbTextCompare)       'Gerbing 15.11.2004
        End If
        If InStr(1, TBegriff.Text, "&&&", vbTextCompare) <> 0 Then
            strVerk = " AND "
            TSplitBeg = Split(TBegriff.Text, "&&&", -1, vbTextCompare)       'Gerbing 15.11.2004
        End If
        If UBound(TSplitBeg) > 0 Then
            If gblnProversion = False Then                              'Gerbing 10.06.2005
                    'msg = "Eingeben von Suchkriterien (Profi) ist in der Shareware-Version nicht möglich." & NL
                    Msg = LoadResString(2129 + Sprache) & NL
                Msg = Msg & TBegriff.Text & NL
                    'msg = msg & "Diese Suchkriterien werden nicht gefunden."
                    Msg = Msg & LoadResString(2130 + Sprache)
                'MsgBox Msg
                MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
                GoTo JedesFeldLight
            End If
            #If Proversion Then
            If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
                SQL = SQL & " AND " & "("
                For n = LBound(TSplitBeg) To UBound(TSplitBeg)
                    If n <> 0 Then
                        SQL = SQL & strVerk
                    End If
                    strTemp = Trim(TSplitBeg(n))
    '                SQL = SQL & " (Situation Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
    '                SQL = SQL & " Ort Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
    '                SQL = SQL & " Land Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
    '                SQL = SQL & " Personen Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
    '                SQL = SQL & " Dateiname Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
    '                SQL = SQL & " Kommentar Like " & "'" & "%" & strTemp & "%" & "'"
                    If gblnSQLServerVersion = True Then
                        SQL = SQL & " (" & LoadResString(1024 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1028 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1030 + Sprache) & " Like " & "N'" & "%" & strTemp & "%" & "'"
                    Else
                        SQL = SQL & " (" & LoadResString(1024 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1028 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'" & " OR "
                        SQL = SQL & " " & LoadResString(1030 + Sprache) & " Like " & "'" & "%" & strTemp & "%" & "'"
                    End If
    '                For i = 0 To ND.ListNutzerdefinierteFelder.ListCount - 1
    '                    SQL = SQL & " OR "
    '                SQL = SQL & "[" & ND.ListNutzerdefinierteFelder.List(i) & "]" & " Like " & "'" & "%" & TBegriff.Text & "%" & "'"
    '                Next i
                    SQL = SQL & ")"
                Next n
                SQL = SQL & ")"
            End If
            #End If
        Else
JedesFeldLight:
            TBegriff.Text = RTrim(TBegriff.Text)                  'Gerbing 22.05.2004
            SQL = SQL & " AND "
'            SQL = SQL & " (Situation Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
'            SQL = SQL & " Ort Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
'            SQL = SQL & " Land Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
'            SQL = SQL & " Personen Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
'            SQL = SQL & " Dateiname Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
'            SQL = SQL & " Kommentar Like " & "'" & "%" & TBegriff.Text & "%" & "'"
            If gblnSQLServerVersion = True Then
                SQL = SQL & " (" & LoadResString(1024 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1028 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1030 + Sprache) & " Like " & "N'" & "%" & TBegriff.Text & "%" & "'"
            Else
                SQL = SQL & " (" & LoadResString(1024 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1025 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1026 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1027 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1028 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'" & " OR "
                SQL = SQL & " " & LoadResString(1030 + Sprache) & " Like " & "'" & "%" & TBegriff.Text & "%" & "'"
            End If
'            For i = 0 To ND.ListNutzerdefinierteFelder.ListCount - 1
'                SQL = SQL & " OR "
'                SQL = SQL & "[" & ND.ListNutzerdefinierteFelder.List(i) & "]" & " Like " & "'" & "%" & TBegriff.Text & "%" & "'"
'            Next i
            SQL = SQL & ")"
        End If
    End If
'Einbeziehen Video-Filter                                                                               'Gerbing 12.06.2016
    If MP.txtBreite.Text <> "" Then
        'SQL = SQL & " AND " & "BreitePixel " & MP.cmbVG1.Text & MP.txtBreite.Text
        SQL = SQL & " AND " & LoadResString(1106 + Sprache) & MP.cmbVG1.Text & MP.txtBreite.Text
    End If
    If MP.txtHöhe.Text <> "" Then
        'SQL = SQL & " AND " & "HoehePixel " & MP.cmbVG2.Text & MP.txthöhe.Text
        If MP.Oder5.Value = True Then
            SQL = SQL & " OR " & LoadResString(1107 + Sprache) & MP.cmbVG2.Text & MP.txtHöhe.Text
        Else
            SQL = SQL & " AND " & LoadResString(1107 + Sprache) & MP.cmbVG2.Text & MP.txtHöhe.Text
        End If
    End If
    If MP.txtDauer.Text <> "" Then
        If MP.Oder6.Value = True Then
            SQL = SQL & " OR " & "VideoDuration " & MP.cmbVG3.Text & MP.txtDauer.Text
        Else
            SQL = SQL & " AND " & "VideoDuration " & MP.cmbVG3.Text & MP.txtDauer.Text
        End If
    End If
'SQL als letztes vor ORDER BY erweitern um eventuell verlangten Dateityp                                'Gerbing 25.03.2016
    If TFileType.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1028 + Sprache) & " like '%." & TFileType.Text & "'"
        'SQL = SQL & " AND " & Dateiname like '*." & TFileType.Text & "'"
    End If
'Sortierfolge nach DateinameKurz------------------------------------------------
    If MP.OptNurNachDateiname.Value = True Then                     'Gerbing 10.10.2004
        'SQL = SQL & " ORDER BY DateinameKurz ASC"
        SQL = SQL & " ORDER BY " & LoadResString(1031 + Sprache) & " ASC"
    Else
        'SQL = SQL & " ORDER BY Jahr,Dateiname ASC"
        SQL = SQL & " ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC"
    End If
'SQL Nachbearbeiten------------------------------------------------------------------------------
    If CheckSQL.Value = 1 Then      '1= aktiviert
        SQLBearbeitenZähler = SQLBearbeitenZähler + 1
        SQLText.Visible = True
        SQLText = SQL
        Me.MousePointer = vbDefault
        btnOK.Enabled = False                                                               'Gerbing 26.01.2008
        Exit Sub
    End If
'-----------------------------------------------------------------------
SQLWurdeBearbeitet:
    CheckSQL.Value = 0                      'Gerbing 09.06.2004
    SQLText.Text = ""                       'Gerbing 07.05.2010
    SQLBearbeitenZähler = 0                     'Gerbing 09.06.2004
    Me.MousePointer = vbHourglass   'zeige an, daß es eine Weile dauern kann                'Gerbing 29.07.2007
    On Error Resume Next
    Set adoRs = New ADODB.Recordset                             'Gerbing 23.11.2017
    With adoRs
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        'msg = "Fehler beim Ausführen der SQL-Anweisung." & NL
        Msg = LoadResString(2140 + Sprache) & NL
        Msg = Msg & SQL & vbNewLine
        'msg = msg & "Formulieren Sie die Suche neu." & NL & NL
        Msg = Msg & LoadResString(2141 + Sprache) & NL & NL
        
        Msg = Msg & "Errortext=" & Err.Description & NL
        Msg = Msg & "Errorcode=" & Err.Number & NL & NL
        
        'msg = msg & "Errortext und Errorcode sagen oft garnichts aus." & NL
        Msg = Msg & LoadResString(2142 + Sprache) & NL
        'msg = msg & "Sehen Sie sich die SQL-Anweisung genau an. Häufig liegt es daran," & NL
        Msg = Msg & LoadResString(2143 + Sprache) & NL
        'msg = msg & "dass Sie ein falsches Datenformat benutzt haben, zb Buchstaben wo Zahlen verlangt sind"
        Msg = Msg & LoadResString(2144 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Me.MousePointer = vbNormal                                                          'Gerbing 29.07.2007
'        CheckSQL.Value = 0                      'Gerbing 09.06.2004
'        SQLText.Text = ""                       'Gerbing 07.05.2010
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    '-------------------------------------------------------------------------------------------------------------
    'On Error GoTo 0
    DBado.Errors.Clear                                                                  'Gerbing 23.11.2017
    On Error GoTo SQLERR
    'frmGridAndThumb.rsDataGrid.Resync
    'frmGridAndThumb.rsDataGrid.Close
    
    If gblnSchreibgeschützt = True Then
        ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid                                                 'Gerbing 23.11.2017
            .Source = SQL
            .ActiveConnection = DBado
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Else
        ' Recordset erstellen und öffnen
        Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
        With frmGridAndThumb.rsDataGrid
            .Source = SQL
            .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
    Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
    frmGridAndThumb.DBGridNeu.ReBind
    frmGridAndThumb.DBGridNeu.AllowArrows = True
    frmGridAndThumb.DBGridNeu.TabAcrossSplits = True
    frmGridAndThumb.DBGridNeu.TabAction = dbgGridNavigation
    frmGridAndThumb.DBGridNeu.WrapCellPointer = True
    'Set frmGridAndThumb.DBGridNeu.DataSource = "Fotos"
    rc = FrageObNurErstenTreffer
    If rc = 0 Then                                          'Gerbing 01.01.2008
        'msg = "Mit diesen Such-Kriterien wurde kein einziger "
        Msg = "error MsDatgrd.ocx" & NL                                                 'Gerbing 27.09.2017
        Msg = Msg & LoadResString(2179 + Sprache)
        'msg = msg & "Datensatz gefunden." & NL
        Msg = Msg & LoadResString(2180 + Sprache) & NL
        'msg = msg & "Wiederholen Sie die Suche mit anderen Such-Kriterien"
        Msg = Msg & LoadResString(2181 + Sprache) & NL
        If IsUserAnAdmin = False Then                                                   'Gerbing 08.11.2015
            If Query.TSituation.ComboItems.Count = 3 Then
                Msg = Msg & NL
                Msg = Msg & "errornumber = " & Err.Number & NL
                Msg = Msg & "errortext = " & Err.Description & NL
            End If
        End If
        MsgBox Msg
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------
    If gstrCommandLine <> "/WRITE" Then
        frmGridAndThumb.DBGridNeu.AllowUpdate = True            'Gerbing 21.11.2007
        frmGridAndThumb.DBGridNeu.Columns(0).Locked = False      'in die Merker-Spalte darf man schreiben
        For n = 1 To frmGridAndThumb.DBGridNeu.Columns.Count - 1
            frmGridAndThumb.DBGridNeu.Columns(n).Locked = True
        Next n
    Else
        frmGridAndThumb.DBGridNeu.AllowUpdate = True
    End If
    If gblnSchreibgeschützt = True Then                     'Gerbing 04.01.2006
        frmGridAndThumb.DBGridNeu.AllowUpdate = False
    End If

    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    rc = PrüfeAnzahlSätze
    Me.MousePointer = vbDefault                                                             'Gerbing 25.06.2013
    If rc <> 0 Then
        Exit Sub
    End If
'    If Query.RecordCount < 100 Then                         'Gerbing 29.07.2006 26.09.2014
        ReDim BildPosList(Query.RecordCount)
        If Query.RecordCount > 0 Then
            For i = 0 To Query.RecordCount - 1
                BildPosList(i).Top = 0
                BildPosList(i).Left = 0
                BildPosList(i).ZoomPercent = 0
                Mid(BildPosList(i).Dateiname, 1, 1) = "?"
            Next i
        End If
'    End If
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
    Me.Hide
    'Query.TPersonen = "*"
    Form1.BackColor = &H0&              'Schwarz
    'Form1.Show                                                                             'Gerbing 29.04.2013
    Me.MousePointer = vbDefault
    If gblnComeFromF8 = True Then                                                           'Gerbing 20.06.2012
        Call Form1.BildAnzeigen
    End If
'---------------------------------------------------------------------------------------------------------------
    MP.Hide                                     'Gerbing 22.04.2005
    'frmGridAndThumb.DbGridNeu.Caption = "Such-Ergebnis in " & " " & PublicDatagridCaption
    frmGridAndThumb.DBGridNeu.Caption = LoadResString(1002 + Sprache) & " " & PublicDatagridCaption
    Form1.KommentarFensterEinblenden = False
    If Query.OKGewählt = True Then             'Gerbing 06.12.2005
        Form1.TimerNachFormLoad.Enabled = True
    End If
    gblnMouseSichtbar = True
    '-----------------------------------------------------------------
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then
        Form1.WindowState = 2                                                               'Gerbing 27.09.2017
'        Me.Width = Form1.Width
'        Me.Height = Form1.Height
        On Error Resume Next                                                                'Gerbing 10.06.2012
    'Form1.Show                                                                             'Gerbing 29.04.2013
        Form1.Picture1.width = Form1.width
        Form1.Picture1.height = Form1.height
    Else
        On Error Resume Next
        Form1.WindowState = 2       'Gerbing 17.10.2014 Aus 0 mache ich 2 'Gerbing 04.09.2012 '2=maximiert wenn die Taskleiste automatisch ausgeblendet wird, soll das Bild bis ganz unten hin gehen
        Form1.width = screenWidth * Screen.TwipsPerPixelX
        Form1.height = screenHeight * Screen.TwipsPerPixelY
        Form1.Picture1.width = Form1.width
        Form1.Picture1.height = Form1.height
        Form1.WindowState = 2
        'Form1.Show                                                                             'Gerbing 29.04.2013
    End If
    On Error GoTo 0
    Form1.Picture1.BackColor = &H0&
    udtData.GdiplusVersion = 1
    If GdiplusStartup(m_lngInstance, udtData, 0) Then
        MsgBox "GDI+ could not be initialized", vbCritical
        Exit Sub
    End If
    Exit Sub
    'Fortsetzung siehe TimerNachFormLoad_Timer
SQLERR:
    If DBado.Errors.Count > 0 Then
        For Each errLoop In DBado.Errors
            MsgBox "Fehler Nr.: " & errLoop.Number & vbCr & _
                errLoop.Description
        Next errLoop
    End If
    Resume Next
End Sub

Private Sub Form_Load()
    Dim antwort As Long
    Dim temp As Long
    Dim start As Long
    Dim pos As Long
    Dim Rechts As String
    Dim zähler As Long
    Dim AnzahlInFotos As Long
    Dim AnzahlÜbereinstimmung As Long
    Dim n As Long
    Dim Feldname As String
    Dim strTemp As String
    Dim DateinamenErweiterung As String
    Dim gefunden As Boolean
    Dim errLoop As Error
    Dim OsVersInfo As OSVERSIONINFO

    gblnComeFromF8 = False                              'Gerbing 20.06.2012
    glngQueryHwnd = Me.hWnd                             'Gerbing 20.06.2012
    Call AnpassenNutzerWunsch(Me)                       'Gerbing 11.03.2017
    If Me.Caption = "Shareware-Hinweis" Then            'Schnulli, damit cracken erschwert wird
        antwort = 12
    End If
    Me.Caption = LoadResString(1090 + Sprache)          'GERBING Fotoalbum - Formulieren Sie die Such-Kriterien Gerbing 08.11.2005
    If gblnSQLServerVersion = True Then
        Label1.Caption = LoadResString(1830 + Sprache)    'Views in der SQL-Server-Datenbank
    Else
        Label1.Caption = LoadResString(1091 + Sprache)    'gespeicherte Abfragen in fotos.mdb
    End If
    If gblnSQLServerVersion = True Then
        Label2.Caption = LoadResString(1832 + Sprache)    'SQL-Darstellung der SQL-Server view
    Else
        Label2.Caption = LoadResString(1092 + Sprache)    'SQL-Darstellung der gespeicherten Abfrage
    End If
    mnuTools.Caption = LoadResString(1036 + Sprache)      '&Tools
    mnuEinstellungen.Caption = LoadResString(1118 + Sprache)      '&Einstellungen   'Gerbing 21.08.2007
    mnuSprache.Caption = LoadResString(1038 + Sprache)    '&Sprache
    mnuReset.Caption = LoadResString(1037 + Sprache)      '&Reset
    mnuHilfe.Caption = LoadResString(1039 + Sprache)      '&Hilfe
    mnuAbout.Caption = LoadResString(1125 + Sprache)      '&Über...                 'Gerbing 14.09.2009
    mnuBeenden.Caption = LoadResString(1040 + Sprache)    '&Beenden
    mnuFotosmdb.Caption = LoadResString(1093 + Sprache)   '&Fotosmdb starten
    mnuRenammdb.Caption = LoadResString(1094 + Sprache)   '&RenamMdb starten
    mnuDiashow.Caption = LoadResString(3193 + Sprache)    'Diashow starten          'Gerbing 17.05.2018
    mnuSpaltenbreite.Caption = LoadResString(1844 + Sprache)      'Spaltenbreite
    LJahr.Caption = LoadResString(1023 + Sprache)           'Jahr:
    LLand.Caption = LoadResString(1026 + Sprache)           'Land:
    LOrt.Caption = LoadResString(1025 + Sprache)            'Ort:
    LPersonen.Caption = LoadResString(1027 + Sprache)       'Personen:
    LSituation.Caption = LoadResString(1024 + Sprache)      'Situation:
    LSWF.Caption = LoadResString(1095 + Sprache)            'SW/F:
    LFileType.Caption = LoadResString(3142 + Sprache)       'Dateityp:              'Gerbing 25.03.2016
    TFileType.tooltipText = LoadResString(3135 + Sprache)   'Sie können hier den von Ihnen gewünschten Dateityp(Dateinamen-Erweiterung) auswählen   'Gerbing 25.03.2016
    chkFensterGrößeÄnderbar.tooltipText = LoadResString(2508 + Sprache) 'Damit läßt sich das Programm 2x starten und bei jeder Instanz die Fenstergröße individuell einstellen. Ziehen Sie an den Fensterseitenkanten.
    CheckNutzerdefinierteFelder.tooltipText = LoadResString(2509 + Sprache) 'Zum Deaktivieren klicken Sie auf nutzerdefinierte Felder...
    lblNutzerdefinierteFelder.tooltipText = LoadResString(2509 + Sprache) 'Zum Deaktivieren klicken Sie auf nutzerdefinierte Felder...
    btnNutzerdefinierteFelder.tooltipText = LoadResString(2510 + Sprache) 'Klicken Sie hier, wenn Sie die Suche nach nutzerdefinierten Feldern erweitern wollen
    CheckWeitereFilterAktiv.tooltipText = LoadResString(2511 + Sprache) 'Zum Deaktivieren klicken Sie auf weitere Filter...
    lblWeitereFilterAktiv.tooltipText = LoadResString(2511 + Sprache) 'Zum Deaktivieren klicken Sie auf weitere Filter...
    CheckDifferenzen.tooltipText = LoadResString(2512 + Sprache) 'Nur sinnvoll, wenn Jahresordner im Dateiname vorkommen. Es wird nach den Differenzen gesucht.
    btnMehrerePersonen.tooltipText = LoadResString(2513 + Sprache) 'Sie können die Suche auf weitere Personen/andere Sortierung/Video-Filter/Dateidatum erweitern
    btnRefresh.tooltipText = LoadResString(2514 + Sprache) 'Reset - Alle Felder auf die Standardwerte einstellen
    btnTürZu.tooltipText = LoadResString(2515 + Sprache) 'Beenden
    CheckSucheJedesFeld.tooltipText = LoadResString(2516 + Sprache)  'Die Felder Situation Ort Land Personen Dateiname Kommentar werden nach diesem Begriff durchsucht. Wenn Sie nach einem Datum suchen, benutzen Sie 'Weitere Filter...'
    optAlleTreffer.tooltipText = LoadResString(2517 + Sprache) 'Alle Treffer pro Jahr (=Standard)
    optNurErstenTreffer.tooltipText = LoadResString(2518 + Sprache)  'Nur den ersten Treffer pro Jahr. So können Sie beispielsweise die jährliche Entwicklung eines Kindes verfolgen.
    optErsterZufallstreffer.tooltipText = LoadResString(2538 + Sprache) 'Nur den ersten Zufalls-Treffer pro Jahr. So können Sie beispielsweise die jährliche Entwicklung eines Kindes verfolgen.
    CheckVollesWort.tooltipText = LoadResString(2519 + Sprache)  'Wenn Sie wirklich zB nur Ina finden wollen und nicht auch Martina, Bei mehreren Personen ist diese Funktion nicht möglich
    CheckVollesWort.Caption = LoadResString(2530 + Sprache)    'Person nur als volles Wort finden
    LSWF.tooltipText = LoadResString(2520 + Sprache) 'SW=Schwarz/Weiss-Foto F=Farbfoto SV=Schwarz/Weiss-Video FV=Farbvideo
    CheckSQL.tooltipText = LoadResString(2521 + Sprache) 'Nur sinnvoll, wenn Sie die SQL-Sprache beherrschen
    btnOK.tooltipText = LoadResString(2522 + Sprache) 'Gefunden werden nur die Fotos, bei denen die gewählten Suchbegriffe in den gewählten Felder stehen
    CheckDifferenzen.Caption = LoadResString(3034 + Sprache) 'Fehlerkontrolle auf Differenzen in Jahr und Dateiname
    If gblnSQLServerVersion = True Then
        CheckGespeicherteAbfragen.Caption = LoadResString(1831 + Sprache)      'Views des SQL-Servers benutzen
    Else
        CheckGespeicherteAbfragen.Caption = LoadResString(3035 + Sprache)      'gespeicherte Abfragen benutzen
    End If
    FrameJahrErweiterung.Caption = LoadResString(3036 + Sprache) 'Trefferauswahl
    optAlleTreffer.Caption = LoadResString(3037 + Sprache)          'Alle
    optNurErstenTreffer.Caption = LoadResString(3038 + Sprache)     'erster Treffer pro Jahr
    optErsterZufallstreffer.Caption = LoadResString(3152 + Sprache) 'ein Zufallstreffer pro Jahr
    JUnd.Caption = LoadResString(3023 + Sprache) 'Und
    SUnd.Caption = LoadResString(3023 + Sprache) 'Und
    OUnd.Caption = LoadResString(3023 + Sprache) 'Und
    LUnd.Caption = LoadResString(3023 + Sprache) 'Und
    SWFUnd.Caption = LoadResString(3023 + Sprache) 'Und
    JOder.Caption = LoadResString(3024 + Sprache) 'Oder
    SOder.Caption = LoadResString(3024 + Sprache) 'Oder
    OOder.Caption = LoadResString(3024 + Sprache) 'Oder
    LOder.Caption = LoadResString(3024 + Sprache) 'Oder
    SWFOder.Caption = LoadResString(3024 + Sprache) 'Oder
    btnMehrerePersonen.Caption = LoadResString(3039 + Sprache)  '&Weitere Filter...
    'CheckWeitereFilterAktiv.Caption = LoadResString(3040 + Sprache) 'weitere Filter sind aktiv
    lblWeitereFilterAktiv.Caption = LoadResString(3040 + Sprache) 'weitere Filter sind aktiv
    btnNutzerdefinierteFelder.Caption = LoadResString(3041 + Sprache)   '&nutzerdefinierte Felder...
    CheckNutzerdefinierteFelder.Caption = LoadResString(3042 + Sprache)    'Suche nach nutzerdefinierten Feldern ist aktiv
    lblNutzerdefinierteFelder.Caption = LoadResString(3042 + Sprache)    'Suche nach nutzerdefinierten Feldern ist aktiv
    CheckSQL.Caption = LoadResString(3043 + Sprache)   'SQL nachbearbeiten
    CheckSucheJedesFeld.Caption = LoadResString(3044 + Sprache)    'Suche Begriff in jedem Feld
    chkFensterGrößeÄnderbar.Caption = LoadResString(3045 + Sprache)   'Fenstergröße änderbar
    btnOK.Caption = LoadResString(3046 + Sprache)      '&Fotos finden
    CheckAudioFileExists.Caption = LoadResString(2231 + Sprache)    'nur Fotos mit gesprochenem Kommentar finden
    CheckUseAudioComments.Caption = LoadResString(2232 + Sprache)   'Audio-Kommentare benutzen
    CheckUseAudioComments.tooltipText = LoadResString(2237 + Sprache)       'zu einer Foto-Datei kann eine gleichnamige Audio-Datei abgespielt werden
    
    SQLWurdeBearbeitet = False
    Frame2.Visible = True
    btnMehrerePersonen.Visible = True
    If gblnSQLServerVersion = True Then
        #If Not Proversion Then
           MsgBox "SQL Server verlangt Proversion"
           End                                                      'Gerbing 04.03.2012
        #End If
    End If
    #If Proversion Then
        If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
            FrameJahrErweiterung.Visible = True
        End If
    #Else
        CheckNutzerdefinierteFelder.Visible = False
        lblNutzerdefinierteFelder.Visible = False       'Gerbing 25.06.2013
        FrameJahrErweiterung.Visible = False
        CheckAudioFileExists.Visible = False            'Gerbing 12.04.2006
    #End If
    CheckSQL.Visible = True
    CheckSucheJedesFeld.Visible = True
    chkFensterGrößeÄnderbar.Visible = True
    chkFensterGrößeÄnderbar.Value = 1                   'Gerbing 04.12.2012
    
'    'Gerbing 04.09.2012 10.05.2019
'    'Ich kann vermeiden, für XP die Version 13.5.1 als letzte unterstützte Version auszuliefern, das Problem kommt
'    'wegen der Wirkungsweise von Form1.TimerRefresh
'    'Im XP legt sich bei eingeschaltetem Timer das GDIPlus Bild über alle anderen Fenster, auch die Fenster anderer Anwendungen.
'    'Wenn ich im XP arbeite, soll grundsätzlich chkFensterGrößeÄnderbar.Value = 1 sein
'    'und ich will erzwingen, daß Form1.Controlbox = True ist, dann kann der Nutzer selber sein gewünschtes Fenster wieder aktivieren
'    'In der Entwicklungsumgebung einstellen: Form1.ControlBox = True
'    OsVersInfo.dwOSVersionInfoSize = Len(OsVersInfo)
'    GetVersionEx1 OsVersInfo
'    If OsVersInfo.dwMajorVersion <= 5 And OsVersInfo.dwMinorVersion <= 2 Then       'dann ist es XP oder Server 2003 oder alles was älter ist
'        chkFensterGrößeÄnderbar.Value = 1
'        chkFensterGrößeÄnderbar.Enabled = False
'        ShowTitleBar True, True                                                             'taskbar visible, Foto
'    End If
    
    'btnOK.Caption = "&Fotos finden"
    btnOK.Caption = LoadResString(1004 + Sprache)
    Me.MousePointer = vbHourglass                   'Gerbing 16.09.2004                     'Gerbing 29.07.2007
    NL = vbNewLine
    SQLBearbeitenZähler = 0
    '-------------------------------------------------------------------------
'    'ADO wegen Benutzung von Adodc1                     'Gerbing 04.01.2006
    Set rstsql = New ADODB.Recordset
    With rstsql
        .Source = "SELECT * FROM Fotos"
        .ActiveConnection = DBado                       'Gerbing 23.11.2017
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    '-------------------------------------------------------------------------
    'If ERR.Number = 3051 Or ERR.Number = 3045 Then         'Gerbing 20.04.2005
    If gblnSQLServerVersion = False Then                    'Gerbing 25.11.2011
        If gblnSchreibgeschützt = True Then                 'Gerbing 23.11.2017
            'schreibgeschützt
            gblnSchreibgeschützt = True
            CheckDifferenzen.Visible = False                'Gerbing 20.09.2004
            btnOK.Enabled = False
            On Error GoTo 0
            Msg = gstrFotosMdbLocation & "\Fotos.mdb" & NL
            'msg = msg & "Die Datenbank ist schreibgeschützt. Sie kann nur im Lesemodus geöffnet werden." & NL
            Msg = Msg & LoadResString(2132 + Sprache) & NL
            'msg = msg & "Es gibt vier mögliche Ursachen für den Lesemodus:" & NL
            Msg = Msg & LoadResString(2133 + Sprache) & NL
            'msg = msg & "1. Das Dateiattribut 'Schreibgeschützt' ist gesetzt" & NL
            Msg = Msg & LoadResString(2134 + Sprache) & NL
            'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte für Ihren PC" & NL
            Msg = Msg & LoadResString(2135 + Sprache) & NL
            'msg = msg & "3. Sie arbeiten mit einer CD" & NL
            Msg = Msg & LoadResString(2136 + Sprache) & NL
            'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & NL & NL
            Msg = Msg & LoadResString(2137 + Sprache) & NL & NL
            
            'msg = msg & "Wollen Sie im Lesemodus weiterarbeiten?" & NL & NL
            Msg = Msg & LoadResString(2138 + Sprache) & NL & NL
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton1 + vbYesNo)
            If antwort = vbYes Then
                CheckDifferenzen.Visible = True                 'Gerbing 20.09.2004
                btnOK.Enabled = True
                'Datenbank im Lesemodus öffnen                  'Gerbing 23.11.2017
                DollarDBado.Close
                DollarDBado.mode = adModeRead                   'adModeRead=1=Read-only.    'adModeReadWrite=3=Read/write.
                DollarDBado.Open DollarDBado.ConnectionString
                DBado.Close
                DBado.mode = adModeRead
                DBado.Open DBado.ConnectionString
                SQL = "Select * FROM Fotos"
                Set adoRs = New ADODB.Recordset
                With adoRs
                    .ActiveConnection = DBado
                    .CursorType = adOpenStatic
                    .CursorLocation = adUseClient
                    .Source = SQL
                    '     .CacheSize = 2
                    .Open
                End With
                Set Adodc1.Recordset = adoRs
                Adodc1.mode = adModeRead
                Adodc1.CursorLocation = adUseClient
                Adodc1.CursorType = adOpenStatic
                temp = Adodc1.Recordset.RecordCount
            Else
                End
            End If
        Else
            'nicht schreibgeschützt
            If gblnSQLServerVersion = False And Err.Number <> 0 Then
                Msg = "Errorcode=" & Err.Number & vbNewLine
                Msg = Msg & "Errortext=" & Err.Description
                'MsgBox msg, , "Das Programm wird beendet"
                MsgBox Msg, , LoadResString(2139 + Sprache)
                End
            Else
                CheckDifferenzen.Visible = True                'Gerbing 20.09.2004
                btnOK.Enabled = True
                On Error GoTo 0
                SQL = "Select * FROM Fotos"
                Set adoRs = New ADODB.Recordset
                With adoRs
                    '.ActiveConnection = DBado                                                        'Gerbing 23.11.2017
                    .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
                    .CursorType = adOpenDynamic
                    '.CursorLocation = Query.enumCursorOrt
                    .Source = SQL
                    '     .CacheSize = 2
                    .Open
                End With
                Set Adodc1.Recordset = adoRs
                On Error Resume Next
                Adodc1.Recordset.MoveLast
                If Err.Number <> 0 Then
                    If gblnWeiterMitLeererDatenbank = False Then
                        Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
                        'Msg = "Die Datei " & gstrFotosMdbLocation & "\Fotos.mdb" & "  ist leer. Die einzige erlaubte Operation ist Import mit Drag&Drop oder Sie können eine neue Datenbank mit dem Tool FotosMdb erzeugen lassen." & vbnewline
                        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
                        Msg = LoadResString(2145 + Sprache) & " " & gstrFotosMdbLocation & "\Fotos.mdb" & " " & LoadResString(2146 + Sprache) & vbNewLine
                        Msg = Msg & LoadResString(2159 + Sprache)
                        'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton2 + vbYesNo)
                        If antwort = vbNo Then
    'Form1.Show                                                                             'Gerbing 29.04.2013
                            End
                        Else
                            Screen.MousePointer = vbNormal
                            gblnWeiterMitLeererDatenbank = True
                        End If
                    End If
                End If
                On Error GoTo 0
                temp = Adodc1.Recordset.RecordCount
            End If
        End If
    End If
    '----------------------------------------------
    SQLWurdeBearbeitet = False
    '---------------------------------------------
    Me.Show                                         'Gerbing 16.09.2004
    FrameGespeicherteAbfragen.Visible = False
    btnRefresh.Enabled = False
    btnMehrerePersonen.Enabled = False
    btnNutzerdefinierteFelder.Enabled = False
    btnOK.Enabled = False
    DoEvents
    On Error GoTo 0
    '-------------------------------------------------
    'Spalte Merker alle Sätze auf 0 setzen                                                  'Gerbing 25.06.2008
    'SQL = "UPDATE Fotos SET Fotos.Merker = False;"
    SQL = "UPDATE Fotos SET Fotos." & LoadResString(2524 + Sprache) & " = 0;"
    If gblnSchreibgeschützt = False Then
        On Error Resume Next
        rstsql.Close
        'On Error GoTo 0
        On Error Resume Next                                                                'Gerbing 04.09.2013
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                        'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    End If
    '---------------------------------------------
    '1.Wenn fotos.exe startet muss eine Abfrage gemacht werden, ob alle Felder Dateiname und
    'DateinameKurz übereinstimmen. Wenn nicht: Hinweis auf Ausführung von Prüfen1.
    'Die Abfrage muß genausoviel Sätze liefern wie Sätze in Tabelle Fotos sind.
    On Error Resume Next
    'SQL = "SELECT DateinameKurz From Fotos;"
'    SQL = "SELECT " & LoadResString(1031 + Sprache) & " From Fotos;"  'Gerbing 08.11.2005
    SQL = "SELECT " & LoadResString(1031 + Sprache) & " From Fotos;"  'Gerbing 08.11.2005
    Set adoRs = New ADODB.Recordset                             'Gerbing 23.11.2017
    With adoRs
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        'msg = "Seit Version 9.0.0.0 verlangt das Programm in der Tabelle Fotos die Spalten DateinameKurz und DDatum." & vbNewLine
        Msg = LoadResString(2147 + Sprache) & vbNewLine
        'msg = msg & "Diese Spalte ist/sind nicht vorhanden." & vbNewLine
        Msg = Msg & LoadResString(2148 + Sprache) & vbNewLine
        'msg = msg & "Nach einer Neuinstallation von Fotos.exe ab Version 9.0.0.0 finden Sie im Installationsverzeichnis" & vbNewLine
        Msg = Msg & LoadResString(2149 + Sprache) & vbNewLine
        'msg = msg & "eine Beispieldatenbank wo diese Spalten eingerichtet sind." & vbNewLine
        Msg = Msg & LoadResString(2150 + Sprache) & vbNewLine
        'msg = msg & "Passen Sie Ihre Datenbank entsprechend an."
        Msg = Msg & LoadResString(2151 + Sprache)
        'MsgBox msg, , "Das Programm wird beendet"
        MsgBox Msg, , LoadResString(2139 + Sprache)
        End
    End If
    On Error GoTo 0
    
    On Error Resume Next
    'SQL = "SELECT DDatum From Fotos;"
'    SQL = "SELECT " & LoadResString(1032 + Sprache) & " From Fotos;"
    SQL = "SELECT " & LoadResString(1032 + Sprache) & " From Fotos;"
    Set adoRs = New ADODB.Recordset                             'Gerbing 23.11.2017
    With adoRs
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        'msg = "Seit Version 9.0.0.0 verlangt das Programm in der Tabelle Fotos die Spalten DateinameKurz und DDatum." & vbNewLine
        Msg = LoadResString(2147 + Sprache) & vbNewLine
        'msg = msg & "Diese Spalte ist/sind nicht vorhanden." & vbNewLine
        Msg = Msg & LoadResString(2148 + Sprache) & vbNewLine
        'msg = msg & "Nach einer Neuinstallation von Fotos.exe ab Version 9.0.0.0 finden Sie im Installationsverzeichnis" & vbNewLine
        Msg = Msg & LoadResString(2149 + Sprache) & vbNewLine
        'msg = msg & "eine Beispieldatenbank wo diese Spalten eingerichtet sind." & vbNewLine
        Msg = Msg & LoadResString(2150 + Sprache) & vbNewLine
        'msg = msg & "Passen Sie Ihre Datenbank entsprechend an."
        Msg = Msg & LoadResString(2151 + Sprache)
        'MsgBox msg, , "Das Programm wird beendet"
        MsgBox Msg, , LoadResString(2139 + Sprache)
        End
    End If
    On Error GoTo 0

    If gblnWeiterMitLeererDatenbank = True Then                 'Gerbing 26.01.2006
        btnOK.Enabled = True
        Me.MousePointer = vbNormal
        Exit Sub
    End If

    SQL = "SELECT Fotos.* From Fotos"
    Set adoRs = New ADODB.Recordset                             'Gerbing 23.11.2017
    With adoRs
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Not adoRs.EOF Then
        adoRs.MoveLast
    End If
    AnzahlInFotos = adoRs.RecordCount
    'SQL = "SELECT Fotos.* From Fotos WHERE InStr(1,Dateiname,DateinameKurz)<>0;"
'    SQL = "SELECT Fotos.* From Fotos WHERE InStr(1," & LoadResString(1028 + Sprache) & "," & LoadResString(1031 + Sprache) & ")<>0;" 'Gerbing 08.11.2005
    If gblnSQLServerVersion = True Then
        'CharIndex hat andere Parameterreihenfolge als InStr
        'SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(1,DateinameKurz,Dateiname)<>0;"
        SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(" & LoadResString(1031 + Sprache) & "," & LoadResString(1028 + Sprache) & ")<>0;" 'Gerbing 08.11.2005
    Else
        SQL = "SELECT Fotos.* From Fotos WHERE InStr(1," & LoadResString(1028 + Sprache) & "," & LoadResString(1031 + Sprache) & ")<>0;" 'Gerbing 08.11.2005
    End If
    Set adoRs = New ADODB.Recordset                             'Gerbing 23.11.2017
    With adoRs
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Not adoRs.EOF Then
        adoRs.MoveLast
    End If
    AnzahlÜbereinstimmung = adoRs.RecordCount
    
    If AnzahlInFotos <> AnzahlÜbereinstimmung Then
        'msg = "Es gibt Datensätze wo Dateiname und DateinameKurz nicht übereinstimmt." & vbNewLine
        Msg = LoadResString(2152 + Sprache) & vbNewLine
        'msg = msg & "Sie müssen die Funktion Prüfen1 im Programm Fotosmdb.exe benutzen, um diesen Fehler zu korrigieren." & vbNewLine
        Msg = Msg & LoadResString(2153 + Sprache) & vbNewLine
        'msg = msg & "Das Programm wird beendet."
        Msg = Msg & LoadResString(2139 + Sprache)
        MsgBox Msg
        End
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------------------------------------
    '3-Einigkeit überprüfen Gerbing 11.04.2005
    'Feldname = rst.Fields("Dateiname")
    If Not adoRs.EOF Then
        Feldname = adoRs.Fields(LoadResString(1028 + Sprache))
    End If
    If Left(Feldname, 3) <> "+:\" Then
        'msg = "Seit Version 12.0.0.0 verlangt das Programm, dass in der Tabelle Fotos" & vbNewLine
        Msg = LoadResString(2154 + Sprache) & vbNewLine
        'msg = msg & "das Feld Dateiname generell mit den Zeichen +:\ beginnt" & vbNewLine
        Msg = Msg & LoadResString(2155 + Sprache) & vbNewLine
        'msg = msg & "Der String +:\ wird vom Programm durch gstrFotosMdbLocation ersetzt." & vbNewLine
        Msg = Msg & LoadResString(2156 + Sprache) & vbNewLine
        'msg = msg & "AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
        Msg = Msg & LoadResString(2157 + Sprache) & vbNewLine
        'msg = msg & "Diese Forderung wurde nicht eingehalten." & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(2158 + Sprache) & vbNewLine & vbNewLine
        
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = Msg & LoadResString(2159 + Sprache)
        antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then
            End
        End If
    End If
    Feldname = Replace(Feldname, "+:\", gstrFotosMdbLocation & "\")
    On Error Resume Next
    strTemp = ""
    'strTemp = Dir(Feldname)
    On Error GoTo 0
    'If strTemp = "" Then                                                    'wenn strTemp = "" bleibt, dann lag error.number 52 vor
    If file_path_exist(Feldname) = False Then
        If gblnSQLServerVersion = True Then
            If gblnWeiterMitAnderemFotosStandort = False Then
                'msg = Feldname & " existiert nicht." & vbNewLine
                Msg = Feldname & LoadResString(2162 + Sprache) & vbNewLine
                'msg = msg & "Prüfen Sie, ob die Angabe " & PublicLocationFotos  & " richtig ist" & vbNewLine & vbNewLine
                Msg = Msg & LoadResString(1820 + Sprache) & PublicLocationFotos & LoadResString(1821 + Sprache) & vbNewLine & vbNewLine
                    
                'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
                Msg = Msg & LoadResString(2159 + Sprache)
                'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
                antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton2 + vbYesNo)
                If antwort = vbNo Then
                    Call Beenden
                    End
                End If
            End If
        Else
            'msg = Feldname & " existiert nicht." & vbNewLine
            Msg = Feldname & LoadResString(2162 + Sprache) & vbNewLine
            'msg = "Datenbank und Fotos passen nicht zueinander" & vbNewLine
            Msg = Msg & LoadResString(2160 + Sprache) & vbNewLine
            'msg = msg & "Vermutlich benutzen Sie eine falsche Datenbank-Datei" & vbNewLine
            Msg = Msg & LoadResString(2161 + Sprache) & vbNewLine
            'msg = msg & "Benutzen Sie das Tool Fotosmdb um die Datenbank zu überprüfen" & vbNewLine & vbNewLine
            Msg = Msg & LoadResString(2163 + Sprache) & vbNewLine & vbNewLine
            
            'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
            Msg = Msg & LoadResString(2159 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton2 + vbYesNo)
            If antwort = vbNo Then
                Call Beenden
                End
            End If
        End If
    End If
    '-------------------------------------------------
    'die höchste gefundene Jahreszahl in die Combobox TJahr eintragen                                   'Gerbing 08.11.2012
    'Dadurch wird ComboBoxenFüllen ausgelöst
    'SQL = "SELECT MAX(Jahr)From fotos"
    SQL = "SELECT MAX(" & (LoadResString(1023 + Sprache)) & ")From fotos"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    TJahr.Text = adoRs.Fields.Item(0)   'jetzt kommt TJahr_Change dran
    
    '-------------------------------------------------
'    Set rst = db.OpenRecordset("fotos")
    SQL = "select * from Fotos"
    Set adoRs = New ADODB.Recordset
    With adoRs
        '.ActiveConnection = DBado                                                        'Gerbing 23.11.2017
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    'Feststellen ob und wieviel es nutzerdefinierte Felder gibt
    ND.ListNutzerdefinierteFelder.ListItems.RemoveAll
    ND.ListDatenTyp.Clear
'    If gblnSQLServerVersion = True Then
        For n = 0 To adoRs.Fields.Count - 1
            Feldname = adoRs.Fields(n).Name
            Feldname = LCase(Feldname)
            Select Case Feldname
                'AudioFileExists gehört nicht zu den nutzerdefinierten Feldern              'Gerbing 12.04.2006
                'IPTCPresent gehört nicht zu den nutzerdefinierten Feldern                  'Gerbing 04.02.2008
                'Case "merker", "jahr", "situation", "ort", "land", "personen", "dateiname", "swf", "kommentar", "dateinamekurz", "ddatum", "breitepixel", "hoehepixel", "audiofileexists", "iptcpresent"
                Case LCase(LoadResString(2524 + Sprache)), LCase(LoadResString(1023 + Sprache)), LCase(LoadResString(1024 + Sprache)), LCase(LoadResString(1025 + Sprache)), LCase(LoadResString(1026 + Sprache)), LCase(LoadResString(1027 + Sprache)), _
                    LCase(LoadResString(1028 + Sprache)), LCase(LoadResString(1029 + Sprache)), LCase(LoadResString(1030 + Sprache)), LCase(LoadResString(1031 + Sprache)), LCase(LoadResString(1032 + Sprache)), LCase(LoadResString(1106 + Sprache)), LCase(LoadResString(1107 + Sprache)), "audiofileexists", "iptcpresent"
                    gefunden = True
                Case "exifdatetimeoriginal"                                                 'Gerbing 16.10.2014
                    Gefundenexifdatetimeoriginal = True
                    gefunden = False                                                        'Gerbing 21.12.2015
                Case Else
                    gefunden = False
            End Select
            If gefunden = False Then
                ND.ListNutzerdefinierteFelder.ListItems.Add adoRs.Fields(n).Name
                ND.ListDatenTyp.AddItem adoRs.Fields(n).Type
                'Debug.Print adoRs.Fields(n).Name                                            'Gerbing 25.10.2013
                'Debug.Print adoRs.Fields(n).Type                                            'Gerbing 25.10.2013
            End If
        Next n
        adoRs.Close
'    Else
'        For n = 0 To rst.Fields.Count - 1
'            Feldname = rst.Fields(n).Name
'            Feldname = LCase(Feldname)
'            Select Case Feldname
'                'AudioFileExists gehört nicht zu den nutzerdefinierten Feldern              'Gerbing 12.04.2006
'                'IPTCPresent gehört nicht zu den nutzerdefinierten Feldern                  'Gerbing 04.02.2008
'                'Case "merker", "jahr", "situation", "ort", "land", "personen", "dateiname", "swf", "kommentar", "dateinamekurz", "ddatum", "breitepixel", "hoehepixel", "audiofileexists", "iptcpresent"
'                Case LCase(LoadResString(2524 + Sprache)), LCase(LoadResString(1023 + Sprache)), LCase(LoadResString(1024 + Sprache)), LCase(LoadResString(1025 + Sprache)), LCase(LoadResString(1026 + Sprache)), LCase(LoadResString(1027 + Sprache)), _
'                    LCase(LoadResString(1028 + Sprache)), LCase(LoadResString(1029 + Sprache)), LCase(LoadResString(1030 + Sprache)), LCase(LoadResString(1031 + Sprache)), LCase(LoadResString(1032 + Sprache)), LCase(LoadResString(1106 + Sprache)), LCase(LoadResString(1107 + Sprache)), "audiofileexists", "iptcpresent"
'                    gefunden = True
'                Case "exifdatetimeoriginal"                                                 'Gerbing 16.10.2014
'                    Gefundenexifdatetimeoriginal = True
'                    gefunden = False                                                        'Gerbing 21.12.2015
'                Case Else
'                    gefunden = False
'            End Select
'            If gefunden = False Then
'                ND.ListNutzerdefinierteFelder.ListItems.Add rst.Fields(n).Name
'                ND.ListDatenTyp.AddItem rst.Fields(n).Properties("type")
'            End If
'        Next n
'        rst.Close
'    End If
    '-------------------------------------------------
    gblnComefromVideo = True                                                            'Gerbing 08.11.2012
    'Form1.Show                                                                             'Gerbing 29.04.2013
    Screen.MousePointer = vbNormal                      'Gerbing 07.11.2011                     'Gerbing 29.07.2007
    If gblnSQLServerVersion = True Then                                                 'Gerbing 29.12.2011
        Call Form1.KontrolleManagement
    Else
        If gblnVollversion = False Then                         'Gerbing 13.10.2005
            Me.Show
            Form1.Hide                                          'Gerbing 23.10.2013
            Copy.Show 1
            If gintDiffTage > 545 Then
                'jetzt wird auf 5000 Fotos begrenzt
                If AnzahlInFotos > 5000 Then
                    'msg = "In der Shareware-Version" & vbNewLine
                    Msg = LoadResString(2165 + Sprache) & vbNewLine
                    'msg = msg & "ist die Anzahl der Datensätze auf 5000 begrenzt." & vbNewLine
                    Msg = Msg & LoadResString(2166 + Sprache) & vbNewLine
                    'msg = msg & "Das Programm wird beendet."
                    Msg = Msg & LoadResString(2139 + Sprache)
                    MsgBox Msg
                    End
                    Exit Sub
                End If
            End If
        End If
    End If
    TimerSetFocus.Enabled = True
    btnRefresh.Enabled = True                                                               'Gerbing 16.09.2004
    btnMehrerePersonen.Enabled = True
    btnNutzerdefinierteFelder.Enabled = True
    CheckDifferenzen.Visible = True                                                         'Gerbing 20.09.2004
    btnOK.Enabled = True
    Screen.MousePointer = vbNormal                      'Gerbing 07.11.2011
    Me.MousePointer = vbNormal                          'Gerbing 07.11.2011
    Exit Sub
Err_ADO:
    Msg = Err.Number & vbNewLine                                                    'Gerbing 10.10.2008
    Msg = Msg & Err.Description & vbNewLine             'Klasse nicht registriert
    If Err.Number = -2147221164 Or Err.Number = -2147418113 Then
        Msg = Msg & "Dieser Fehler tritt auf, wenn Sie GERBING Fotoalbum später installiert haben als ein Update mit XP SP3 (Service Pack 3)," & vbNewLine
        Msg = Msg & "und tritt nicht auf, wenn GERBING Fotoalbum bereits installiert war, als Sie das Update mit XP SP3 ausgeführt haben." & vbNewLine
        Msg = Msg & "Solange Microsoft keine Lösung dieses Fehlers bietet, haben Sie nur die Möglichkeit, die erforderliche Reihenfolge einzuhalten." & vbNewLine & vbNewLine
        
        Msg = Msg & "You get this error if you have first installed XP SP3 (service pack 3) and then GERBING Fotoalbum," & vbNewLine
        Msg = Msg & "and will not get this error if GERBING Fotoalbum was already installed before you installed XP SP3." & vbNewLine
        Msg = Msg & "As long as microsoft does not deliver a solution you must follow the required installation sequence."
    End If
    MsgBox Msg
    End
End Sub

Private Sub ListGespeicherteAbfragen_Click(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim pos As Long
    Dim pos1 As Long
    Dim Msg As String
    Dim SQL As String
    Dim strAbfrage As String                                            'Gerbing 09.01.2018
    Dim cat As ADOX.Catalog
    Dim cmd As ADODB.command
    
    Set cmd = New ADODB.command
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = DBado

'   Fehlerkorrektur bei gespeicherten Abfragen:                         'Gerbing 19.01.2007
'   Bisher kam es zum Fehler, wenn im SQL-Text zb 'Fotos.BreitePixel' formuliert wurde.
'   Richtig muss formuliert werden 'BreitePixel'
'   Das Programm ersetzt 'Fotos.' durch ""
    If gblnSQLServerVersion = True Then                                'Gerbing 23.11.2017
        'SQL = "select * from INFORMATION_SCHEMA.VIEWS where TABLE_NAME='" & ListGespeicherteAbfragen.List(ListGespeicherteAbfragen.ListIndex) & "'"
        SQL = "select * from INFORMATION_SCHEMA.VIEWS where TABLE_NAME='" & listItem & "'"
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .ActiveConnection = DBado
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Source = SQL
            .Open
        End With
        
        While Not rstsql.EOF
            txtSQLGespeicherteAbfrage = rstsql.Fields("VIEW_DEFINITION")
            rstsql.MoveNext
        Wend
        If txtSQLGespeicherteAbfrage <> "" Then
            pos = InStr(1, txtSQLGespeicherteAbfrage, "select", vbTextCompare)
            If pos <> 0 Then
                txtSQLGespeicherteAbfrage = Mid(txtSQLGespeicherteAbfrage, pos, Len(txtSQLGespeicherteAbfrage) - pos)
            End If
        End If
    Else
        'txtSQLGespeicherteAbfrage = DB.QueryDefs(ListGespeicherteAbfragen.ListIndex).SQL
        'txtSQLGespeicherteAbfrage = DB.QueryDefs(listItem).SQL
        strAbfrage = listItem
        If Sprache = 0 Then                                                             'Gerbing 05.07.2019
            On Error Resume Next
            Set cmd = cat.Views(strAbfrage).command                                     'Gerbing 09.01.2018
            Set cmd = cat.Procedures(strAbfrage).command                                'Gerbing 05.07.2019
            On Error GoTo 0
        Else
            'bei sprache = 3000 = english muss ich Procedures nehmen
            Set cmd = cat.Procedures(strAbfrage).command                                'Gerbing 05.07.2019
        End If
        txtSQLGespeicherteAbfrage = cmd.CommandText
    End If
    pos = InStr(1, txtSQLGespeicherteAbfrage, "Fotos.", vbTextCompare)
    If pos <> 0 Then
'        msg = "Sie benutzen im SQL-String Formulierungen der Art Tabellenname.Feldname." & vbNewLine
'        msg = msg & "Solche Formulierungen würden zu Laufzeitfehlern führen." & vbNewLine
'        msg = msg & "Das Programm wird diese Formulierungen jetzt entfernen."
        Msg = LoadResString(2315 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2316 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2317 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        txtSQLGespeicherteAbfrage = Replace(txtSQLGespeicherteAbfrage, "Fotos.", "", , , vbTextCompare)
    End If
    'Wenn im SQL-Text Select *, Feldname FROM ..... formuliert wird, werden die bezeichneten Feldnamen
    'als darzustellendes Feld betrachtet und stehen im Recordset vor den üblichen Standardfeldern.
    'Das bringt die gespeicherten Feldbreiten durcheinander
    'Der nutzer erhält eine Warnung
    pos = InStr(1, txtSQLGespeicherteAbfrage, "Select *", vbTextCompare)
    pos1 = InStr(1, txtSQLGespeicherteAbfrage, "FROM", vbTextCompare)
    If pos <> 0 And pos1 <> 0 Then
        If pos1 - pos > 10 Then
'            msg = "Vermeiden Sie im SQL-String Formulierungen wie Select *, Feldname FROM..." & vbNewLine
'            msg = msg & "Dadurch würden die gespeicherten Feldbreiten auf falsche Felder angewendet." & vbNewLine
'            msg = msg & "Formulieren Sie besser Select * FROM..."
'            msg = msg & "Das Programm wird diese Formulierungen jetzt entfernen."
            Msg = LoadResString(2318 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2319 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2320 + Sprache) & vbNewLine                       'Gerbing 07.11.2007
            Msg = Msg & LoadResString(2317 + Sprache)                                   'Gerbing 07.11.2007
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            txtSQLGespeicherteAbfrage = "Select * " & Right(txtSQLGespeicherteAbfrage, Len(txtSQLGespeicherteAbfrage) - pos1 + 1)   'Gerbing 07.11.2007
        End If
    End If
'   Warnung bei gespeicherten Abfragen:                                                 'Gerbing 07.11.2007
'   zwischen 'Select... und ...FROM' muss ein '*' stehen, sonst kann man den String 'Select * FROM...'
'   nicht herstellen. Es kommt ein Warnhinweis
    If pos = 0 Then
'        msg = "Die Formulierung 'Select * FROM' wurde nicht gefunden." & vbNewLine
'        msg = msg & "Möglicherweise könnten Sie unerwartete Ergebnisse erhalten."
        Msg = LoadResString(2328 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2329 + Sprache)
        MsgBox Msg
    End If
    txtSQLGespeicherteAbfrage = Replace(txtSQLGespeicherteAbfrage, """", "'")
End Sub

Private Sub mnuAbout_Click()
    gblnComeFromAboutMenue = True                   'Gerbing 19.09.2012
    AboutForm.Show 1                                'Gerbing 14.09.2009
End Sub

Private Sub mnuBeenden_Click()
    Call Beenden                                    'Gerbing 24.06.2006
    End
End Sub

Private Sub mnuDiashow_Click()                                                                      'Gerbing 27.01.2018
    Dim AppId
    Dim Msg As String
    
    'If Dir(AppPath & "\diashow.exe") = "" Then
    If file_path_exist(AppPath & "\diashow.exe") = False Then
        'msg = "diashow konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2561 + Sprache) & vbNewLine
        'msg = msg & "diashow.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2562 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    AppId = Shell(AppPath & "\diashow.exe", vbNormalFocus)
    AppActivate AppId
End Sub

Private Sub mnuEinstellungen_Click()
    Call AnpassenNutzerWunsch(Me)                       'Gerbing 11.03.2017
    Form1.Hide                                                                              'Gerbing 23.10.2013
    WertxForm.Show 1
End Sub

Private Sub mnuFotosmdb_Click()
    Dim AppId
    Dim Msg As String
    Dim cmdline As String
    
    'If Dir(AppPath & "\fotosmdb.exe") = "" Then
    If file_path_exist(AppPath & "\fotosmdb.exe") = False Then
        'msg = "Fotosmdb konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2167 + Sprache) & vbNewLine
        'msg = msg & "Fotosmdb.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2168 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    'CommandLine aufbauen mit access
        'fotosmdblocation=...;
    
    'CommandLine aufbauen mit sql server
        'sqlservername=...;
        'datenbankname=...;
        'WindowsAuthentication=0; heißt nein
        'WindowsAuthentication=1; heißt ja
        'username=...;
        'Password=...;
        'StandortFotos=...;

    'CommandLine aufbauen mit access
    If gblnSQLServerVersion = False Then
        If gstrFotosMdbLocation <> "" Then                                                          'Gerbing 07.11.2011
            AppId = Shell(AppPath & "\fotosmdb.exe" & " " & "fotosmdblocation=" & gstrFotosMdbLocation & ";", vbNormalFocus)
            AppActivate AppId
        Else
            AppId = Shell(AppPath & "\fotosmdb.exe", vbNormalFocus)
            AppActivate AppId
        End If
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
        AppId = Shell(AppPath & "\fotosmdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If
End Sub

Private Sub mnuHilfe_Click()
    Dim RetVal As Long
    Dim CHMFile As String

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

Private Sub mnuRenamMdb_Click()
    Dim AppId
    Dim Msg As String
    Dim cmdline As String
    
    'If Dir(AppPath & "\RenamMdb.exe") = "" Then
    If file_path_exist(AppPath & "\RenamMdb.exe") = False Then
        'msg = "RenamMdb konnte nicht gestartet werden." & vbNewLine
        Msg = LoadResString(2169 + Sprache) & vbNewLine
        'msg = msg & "RenamMdb.exe muss im gleichen Ordner stehen wie fotos.exe"
        Msg = Msg & LoadResString(2170 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    'CommandLine aufbauen mit access
        'fotosmdblocation=...;
    
    'CommandLine aufbauen mit sql server
        'sqlservername=...;
        'datenbankname=...;
        'WindowsAuthentication=0; heißt nein
        'WindowsAuthentication=1; heißt ja
        'username=...;
        'Password=...;
        'StandortFotos=...;

    'CommandLine aufbauen mit access
    If gblnSQLServerVersion = False Then
        If gstrFotosMdbLocation <> "" Then                                                          'Gerbing 07.11.2011
            AppId = Shell(AppPath & "\RenamMdb.exe" & " " & "fotosmdblocation=" & gstrFotosMdbLocation & ";", vbNormalFocus)
            AppActivate AppId
        Else
            AppId = Shell(AppPath & "\RenamMdb.exe", vbNormalFocus)
            AppActivate AppId
        End If
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
        AppId = Shell(AppPath & "\RenamMdb.exe" & " " & cmdline, vbNormalFocus)
        AppActivate AppId
    End If
End Sub

Private Sub mnuReset_Click()
    Call RefreshWerte
End Sub

Private Sub mnuResetAll_Click()                                                                     'Gerbing 10.01.2015
    Call RefreshWerte
    TJahr = "*"
End Sub

Private Sub mnuSpaltenbreite_Click()                                                                'Gerbing 19.04.2015
    Dim antwort As Long
    Dim Stil As Long
    Dim Msg As String
    Dim n As Long
    Dim RecordCount As Long
    
    If gblnSchreibgeschützt = True Then
        Msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        Msg = Msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
        Exit Sub
    End If
    
    'antwort = MsgBox("Wollen Sie alle Spalten auf die Standardbreite von 100 Pixel einstellen?", vbYesNo)
    Msg = LoadResString(2286 + Sprache)
    Stil = vbYesNo + vbDefaultButton2
    antwort = MsgBox(Msg, Stil)
    If antwort = vbNo Then
        Exit Sub
    End If

    SQL = "SELECT SpaltenBreite.* FROM SpaltenBreite;"
    'SQL = "SELECT " & LoadResString(2525 + Sprache) & ".* FROM " & LoadResString(2525 + Sprache) & ";"
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    RecordCount = rstsql.RecordCount
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
    DBado.Execute (SQL)                                                         'Gerbing 23.11.2017
    For n = 0 To RecordCount - 1                        'es wird ab Spalte Merker gespeichert
        rstsql.AddNew
        rstsql.Fields("Spaltenbreite") = 100
        rstsql.Update
    Next n
    rstsql.Close
End Sub

Private Sub mnuSprache_Click()
    'Dim cn As New ADODB.Connection
    Dim cn As ADODB.Connection
    
'Statt:                                     'Gerbing 29.11.2016
'    'Dim cn As New ADODB.Connection
'    Dim cn As ADODB.Connection
'
'
'  Set cn = New ADODB.Connection
'  Call cn.Open(sConnect)
'Das hier:
'  Set cn = CreateObject("ADODB.Connection")
'  cn.Open sConnect
    Dim Rs As New ADODB.Recordset
    Dim MultiUser As Boolean
    Dim NutzerName As String

    If gblnSQLServerVersion = False Then
        'Gerbing 18.02.2011------------------------------------------------------------------------------------------
        'Wer alles Nutzer der Datenbank ist, steht in einer Multi-Nutzer-Umgebung in der Datei fotos.ldb
        'diese gibts nur in einer Multiuser-Umgebung und wird von selbst gelöscht, wenn es nur noch einen Nutzer gibt
        
        Set cn = CreateObject("ADODB.Connection")                                       'Gerbing 23.11.2017
        cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & gstrFotosMdbLocation & "\fotos.mdb"
        cn.mode = adModeReadWrite
        cn.Open cn.ConnectionString
        Set Rs = cn.OpenSchema(adSchemaProviderSpecific, _
        , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
        
        'Output the list of all users in the current database.
    
    '    Debug.Print Rs.Fields(0).Name, "", Rs.Fields(1).Name, _
    '    "", Rs.Fields(2).Name, Rs.Fields(3).Name
    
        MultiUser = False
        Do
    '        Debug.Print Rs.Fields(0), Rs.Fields(1), _
    '        Rs.Fields(2), Rs.Fields(3)
            If Rs.EOF Then Exit Do
            NutzerName = Trim(Rs.Fields(0))
            Rs.MoveNext
            If Not Rs.EOF Then
                If NutzerName <> Trim(Rs.Fields(0)) Then
                    MultiUser = True
                    Exit Do
                End If
            End If
        Loop
        If MultiUser = True Then
            Rs.Close
            cn.Close
            Msg = gstrFotosMdbLocation & "\fotos.mdb" & vbNewLine
    '        msg = msg & "Sprache wechseln muss ausgeführt werden, wenn Sie der einzige Nutzer der Datenbank sind" & vbNewLine
    '        msg = msg & "Die Namen der anderen Nutzer finden Sie in der Datei " & gstrFotosMdbLocation & "\fotos.ldb"
            Msg = LoadResString(2282 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2280 + Sprache) & gstrFotosMdbLocation & "\fotos.ldb"
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Exit Sub
        End If
        Rs.Close
        cn.Close
    End If
'-----------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    If gblnSchreibgeschützt = True Then
        'schreibgeschützt
        Msg = gstrFotosMdbLocation & "\$Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        Msg = Msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-----------------------------
    'Beenden wenn Fotos.mdb schreibgeschützt ist
    If gblnSchreibgeschützt = True Then
        Msg = gstrFotosMdbLocation & "\Fotos.mdb" & vbNewLine
        'Msg= msg & "Die Datenbank ist schreibgeschützt, Änderungen sind nicht möglich"
        Msg = Msg & LoadResString(2210 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '----------------------
    frmSprache.Show 1
    
End Sub

Private Sub optAlleTreffer_Click()
    gblnFotosMitFET = False                             'Gerbing 16.02.2013
End Sub

Private Sub optNurErstenTreffer_Click()
    gblnFotosMitFET = False                             'Gerbing 16.02.2013
    If gblnSchreibgeschützt = True Then                 'Gerbing 11.04.2005
        'msg = "Bei einer schreibgeschützten Datenbank ist die Funktion 'erster Treffer pro Jahr' nicht möglich," & vbNewLine
        Msg = LoadResString(2171 + Sprache) & vbNewLine
        'msg = msg & "weil das Programm wegen des Schreibschutzes die Tabelle FET nicht" & vbNewLine
        Msg = Msg & LoadResString(2172 + Sprache) & vbNewLine
        'msg = msg & "anlegen kann."
        Msg = Msg & LoadResString(2173 + Sprache)
        'MsgBox msg, vbInformation, "Hinweis"
        MsgBox Msg, vbInformation, LoadResString(2174 + Sprache)
        Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
        optAlleTreffer.Value = True                     'Gerbing 26.06.2006
        Exit Sub
    End If
End Sub

Private Sub SQLText_Click(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnOK.Enabled = True                                                                    'Gerbing 26.01.2008
End Sub

Private Sub SQLText_TextChanged()
    SQLWurdeBearbeitet = True
End Sub

Private Sub TimerSetFocus_Timer()
    'Nur weil sich während Form_Load der Focus nicht auf TJahr setzen läßt
    TimerSetFocus.Enabled = False
    On Error Resume Next
    TJahr.SetFocus
    Screen.MousePointer = vbNormal                                                          'Gerbing 07.11.2011
End Sub

Public Function FrageObNurErstenTreffer()
    Dim rst1 As ADODB.Recordset
    'rc = 0 kein Ergebnis
    'rc = 1 es gibt Ergebnissätze
    FrageObNurErstenTreffer = 1
    #If Proversion Then
    If gblnVollversion = True And gblnProversion = True Then    'Gerbing 22.02.2006
        Dim JahrA As String
        Dim n As Long
        Dim i As Long
        Dim SQLR As String
        Dim SQLQ As String
        Dim zähler As Integer
        Dim pos As Long
        Dim Links As String
        Dim errLoop As Error
        Dim Bookmark As Long
        Dim lngRecordCount As Long
        Dim sngWert As Single
        Dim lngWert As Long
        Dim Obergrenze As Long
        Dim Untergrenze As Long
        'Gerbing 07.03.2005
        'Wenn optNurErstenTreffer.Value = True dann wird die Tabelle FET benutzt
        'zuerst wird diese Tabelle gelöscht mit DROP
        'dann wird diese Tabelle erzeugt mit den Suchkriterien des String SQL
        'dorthinein muss der Zusatz INTO FET
        'dorthinein muss GROUP BY Jahr ORDER BY Jahr;
        'dann kommt eine Abfrage mit zwei verknüpften Tabellen Fotos und FET
        If optNurErstenTreffer.Value = True Or optErsterZufallstreffer.Value = True Then
        'Datenbank Tabelle FET erzeugen, Fehler wenn Vorgang scheitert
            Bookmark = frmGridAndThumb.DBGridNeu.Bookmark                            'Gerbing 01.01.2011
            If gblnFotosMitFET = True Then GoTo FotosMitFET                     'Gerbing 16.02.2013
            On Error GoTo DropTableFET                                          'Gerbing 01.01.2008
            SQLR = "Drop Table FET"
            On Error Resume Next
            'DBsql.Execute SQLR, dbFailOnError
            DBado.Execute SQLR                                                  'Gerbing 23.11.2017
            '--------------
            On Error GoTo 0
            SQLR = SQL
            If gblnSQLServerVersion = True Then
                'SQL Server does not support FIRST
                'SQLR = Replace(SQL, "Select *", "SELECT MIN(Dateiname) AS FN INTO FET", 1, 1, vbTextCompare)
                SQLR = Replace(SQL, "Select *", "SELECT MIN(" & LoadResString(1028 + Sprache) & ") AS FN INTO FET", 1, 1, vbTextCompare)
                SQLR = Replace(SQLR, "*", "%")                              'Gerbing 08.05.2006
            Else
                'SQLR = Replace(SQL, "Select *", "SELECT First(Dateiname) AS FN INTO FET", 1, 1, vbTextCompare)
                SQLR = Replace(SQL, "Select *", "SELECT First(" & LoadResString(1028 + Sprache) & ") AS FN INTO FET", 1, 1, vbTextCompare)
                SQLR = Replace(SQLR, "*", "%")                              'Gerbing 08.05.2006
            End If
            'dorthinein muss GROUP BY Jahr ORDER BY Jahr;
            pos = InStr(1, SQLR, "ORDER BY", vbTextCompare)
            Links = Left(SQLR, pos - 1)
            SQLR = Links & " GROUP BY " & LoadResString(1023 + Sprache) & " ORDER BY " & LoadResString(1023 + Sprache) & ";"
            On Error GoTo SelectIntoFET                                          'Gerbing 01.01.2008
            DBado.Execute SQLR, 128 + 32                              'Gerbing 23.11.2017 dbFailOnError=128 dbConsistent=32
            '--------------
            'Wenn FET leer ist, kann ich hier aufhören                          'Gerbing 01.01.2008
            SQLQ = "Select * FROM FET;"
            On Error Resume Next
            rstsql.Close
            On Error GoTo 0
            With rstsql
                .Source = SQLQ
                .ActiveConnection = DBado                                       'Gerbing 23.11.2017
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If rstsql.RecordCount = 0 Then
                FrageObNurErstenTreffer = 0
                rstsql.Close
                Exit Function
                '---------------------------Exit Function---------------------
            End If
            rstsql.Close
            '---------------------------------------------------
            If optErsterZufallstreffer.Value = True Then                    'Gerbing 09.02.2013
                'Bei der Option 'ein Zufallstreffer pro Jahr' kann man die Suche beliebig oft wiederholen und bekommt immer ein neues
                'zufälliges Bild pro Jahr, vorausgesetzt es gibt pro Jahr mehr als ein Bild
                '0.1.Öffne Tabelle FET für das Einfügen neuer Sätze mit rst1
                '1.Suche in der Tabelle Fotos Select Count(Dateiname) und benutze den bisherigen SQL String
                '2.Das Ergebnis im Array TrefferProJahrList speichern
                '3.Suche in der Tabelle Fotos nach jedem Jahr aus TrefferProJahrList und benutze den bisherigen SQL String
                '3.1.Setze den Recordset mit Move auf eine Zufallsposition zwischen 1 und TrefferProJahrList(n).Anzahl
                    'und speichere diesen Dateiname mit AddNew in der Tabelle FET
                '0.1.Öffne Tabelle FET für das Einfügen neuer Sätze
                If gblnSQLServerVersion = True Then
                    'beim SQL Server muss es heißen 'Delete from table
                    SQLQ = "DELETE From FET"
                Else
                    SQLQ = "DELETE * FROM FET"
                End If
                On Error Resume Next
                Set rst1 = New ADODB.Recordset
                rst1.Close
                On Error GoTo 0
                With rst1
                    .Source = SQLQ
                    .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                SQLQ = "Select * FROM FET;"
                On Error Resume Next
                rst1.Close
                On Error GoTo 0
                With rst1
                    .Source = SQLQ
                    .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                '1.Suche in der Tabelle Fotos Select Count(Dateiname) und benutze den bisherigen SQL String
                SQLR = SQL
                'SQLR = Replace(SQL, "Select *", "SELECT Count(Dateiname) AS anz, jahr", 1, 1, vbTextCompare)
                SQLR = Replace(SQLR, "Select *", "SELECT Count(" & LoadResString(1028 + Sprache) & ") AS anz," & LoadResString(1023 + Sprache), 1, 1, vbTextCompare)
                SQLR = Replace(SQLR, "*", "%")                              'Gerbing 08.05.2006
                'dorthinein muss GROUP BY Jahr ORDER BY Jahr;
                pos = InStr(1, SQLR, "ORDER BY", vbTextCompare)
                Links = Left(SQLR, pos - 1)
                SQLR = Links & " GROUP BY " & LoadResString(1023 + Sprache) & " ORDER BY " & LoadResString(1023 + Sprache) & ";"
                On Error Resume Next
                rstsql.Close
                On Error GoTo 0
                With rstsql
                    .Source = SQLR
                    .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                '2.Das Ergebnis im Array TrefferProJahrList speichern
                ReDim TrefferProJahrList(rstsql.RecordCount)
                For n = 0 To rstsql.RecordCount - 1
                    TrefferProJahrList(n).Jahr = rstsql.Fields(LoadResString(1023 + Sprache))
                    TrefferProJahrList(n).Anzahl = rstsql.Fields("Anz")
                    rstsql.MoveNext
                Next n
                On Error Resume Next
                rstsql.Close
                On Error GoTo 0
                '3.Suche in der Tabelle Fotos nach jedem Jahr aus TrefferProJahrList und benutze den bisherigen SQL String
                For n = 0 To UBound(TrefferProJahrList) - 1
                    SQLR = SQL
                    'SQLR = Replace(SQL, sqljahreszahl, "Where Jahr = " & TrefferProJahrList(n).Jahr, 1, 1, vbTextCompare)
                    SQLR = Replace(SQLR, SQLJahresZahl, "WHERE " & LoadResString(1023 + Sprache) & "=" & TrefferProJahrList(n).Jahr, 1, 1, vbTextCompare)
                    'SQLR = Replace(SQLR, "*", "%")                             'Gerbing 08.05.2006
                    On Error GoTo 0
                    With rstsql
                        .Source = SQLR
                        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                        .CursorType = adOpenStatic
                        .LockType = adLockOptimistic
                        .CursorLocation = adUseClient
                        .Open
                    End With
                    '3.1.Setze den Recordset mit Move auf eine Zufallsposition zwischen 1 und TrefferProJahrList(n).Anzahl
                    'und speichere diesen Dateiname mit AddNew in der Tabelle FET
                    'Aus TrefferProJahrList(n).Anzahl eine Zufallszahl zwischen 1 und TrefferProJahrList(n).Anzahl bilden
                    If TrefferProJahrList(n).Anzahl > 1 Then
                        Obergrenze = TrefferProJahrList(n).Anzahl - 1
                        Untergrenze = 1
                        Randomize
                        sngWert = (Obergrenze - Untergrenze + 1) * Rnd + Untergrenze
                        lngWert = Round(sngWert)
                        If lngWert > Obergrenze Then
                            lngWert = Obergrenze
                        End If
                        If lngWert < 1 Then
                            lngWert = 1
                        End If
                        rstsql.Move lngWert
                    End If
                    rst1.AddNew
                    rst1.Fields("FN") = rstsql.Fields(LoadResString(1028 + Sprache))
                    rst1.Update
                    On Error Resume Next
                    rstsql.Close
                    On Error GoTo 0
                Next n
                On Error Resume Next
                rst1.Close
                On Error GoTo 0
            End If
FotosMitFET:
            On Error GoTo Errmsg
            '---------------------------------------------------
            'Abfrage mit zwei verknüpften Tabellen Fotos und FET
            
    '        SQLR = "SELECT Fotos.* FROM FET INNER JOIN Fotos ON FET.FN = Fotos.Dateiname ORDER BY Jahr;"
            SQLR = "SELECT Fotos.* FROM FET INNER JOIN Fotos ON FET.FN = Fotos." & LoadResString(1028 + Sprache) & " ORDER BY " & LoadResString(1023 + Sprache) & ";"
ResumeClose:
            frmGridAndThumb.rsDataGrid.Close
            If gblnSchreibgeschützt = True Then
                ' Recordset erstellen und öffnen adOpenStatic
                Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
                With frmGridAndThumb.rsDataGrid
                    .Source = SQLR
                    .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
            Else
                ' Recordset erstellen und öffnen adOpenDynamic
                Set frmGridAndThumb.rsDataGrid = New ADODB.Recordset
                With frmGridAndThumb.rsDataGrid
                    .Source = SQLR
                    .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
            End If
            'Das DbGridNeu soll das Ergebnis der Abfrage mit zwei verknüpften Tabellen Fotos und FET anzeigen
            frmGridAndThumb.rsDataGrid.MoveFirst
            Set frmGridAndThumb.Adodc1.Recordset = frmGridAndThumb.rsDataGrid
            Set frmGridAndThumb.DBGridNeu.DataSource = frmGridAndThumb.rsDataGrid
            frmGridAndThumb.DBGridNeu.ReBind
            frmGridAndThumb.DBGridNeu.Bookmark = Bookmark                'Gerbing 01.01.2011
            Call frmGridAndThumb.SetSpaltenBreite                        'Gerbing 01.01.2011
            gblnFotosMitFET = True                                  'Gerbing 16.02.2013
        End If
        FrageObNurErstenTreffer = 1
        Exit Function
Errmsg:
        Msg = "Errorcode: " & Err.Number & NL
        Msg = Msg & "Errortext: " & Err.Description
        Select Case Err.Number
            Case -2147217865
                Sleep (1000)
                Resume
            Case 3021                                               'Gerbing 24.06.2006
                'Resume Next                                        'Gerbing 01.01.2008 auskommentiert
                GoTo ResumeClose                                    'Gerbing 01.01.2008
            Case Else
                MsgBox Msg, vbCritical
                Resume Next
        End Select
    End If
    Exit Function
DropTableFET:
        'msg = "Tabelle kann nicht gelöscht werden." & NL
        Msg = LoadResString(2266 + Sprache) & NL
        Msg = Msg & "FET" & NL
        Msg = Msg & "Errorcode: " & Err.Number & NL
        Msg = Msg & "Errortext: " & Err.Description
        MsgBox Msg, vbCritical
        MsgBox Msg, vbCritical
        'End                                                        'Gerbing 26.08.2008
        Resume Next                                                 'Gerbing 26.08.2008
SelectIntoFET:
        'msg = "Tabelle kann nicht erzeugt werden." & NL
        Msg = LoadResString(2177 + Sprache) & NL
        Msg = Msg & "FET" & NL
        Msg = Msg & "Errorcode: " & Err.Number & NL
        Msg = Msg & "Errortext: " & Err.Description
        MsgBox Msg, vbCritical
        MsgBox Msg, vbCritical
        End
    #End If
End Function

Private Function PrüfeAnzahlSätze()
    'Teste, ob mindestens ein Satz gefunden wurde
    On Error Resume Next
    frmGridAndThumb.Adodc1.Recordset.MoveFirst
    If Err.Number <> 0 Then
        If CheckDifferenzen.Value = 0 Then
            'MsgBox "Function PrüfeAnzahlSätze " & vbNewLine & "Err.Number=" & Err.Number & vbNewLine & "Err.Description=" & Err.Description    'Gerbing 25.10.2015
            'msg = "Mit diesen Such-Kriterien wurde kein einziger "
            If Err.Number <> 3021 Then                                                      'Gerbing 27.09.2017
                Msg = "error MsDatgrd.ocx" & NL                                             'Gerbing 27.09.2017
            End If                                                                          'Gerbing 27.09.2017
            Msg = Msg & LoadResString(2179 + Sprache)
            'msg = msg & "Datensatz gefunden." & NL
            Msg = Msg & LoadResString(2180 + Sprache) & NL
            'msg = msg & "Wiederholen Sie die Suche mit anderen Such-Kriterien"
            Msg = Msg & LoadResString(2181 + Sprache)
            If IsUserAnAdmin = False Then                                                   'Gerbing 08.11.2015
                If Query.TSituation.ComboItems.Count = 3 Then
                    Msg = Msg & NL & NL
                    Msg = Msg & "errornumber = " & Err.Number & NL
                    Msg = Msg & "errortext = " & Err.Description & NL
                    Msg = Msg & NL
                End If
            End If
        Else
            'msg = "Es wurden keine Differenzen in Jahr und Dateiname gefunden"
            Msg = LoadResString(2182 + Sprache)
        End If
        If gblnWeiterMitLeererDatenbank = True Then             'Gerbing 26.01.2006
            PrüfeAnzahlSätze = 0
        Else
            MsgBox Msg
            Msg = ""                                            'Gerbing 02.10.2017
            PrüfeAnzahlSätze = 1
            Query.OKGewählt = False                             'Gerbing 20.04.2012
        End If
        Exit Function
    End If
'    frmGridAndThumb.Adodc1.Recordset.MoveLast                         'Gerbing 16.06.2005 'Gerbing 25.06.2013
'    frmGridAndThumb.Adodc1.Recordset.MoveFirst
    Query.RecordCount = frmGridAndThumb.Adodc1.Recordset.RecordCount
    PrüfeAnzahlSätze = 0
End Function

Private Function DatenTypDate(Feldname, Feldwert) As Integer
    Dim n As Long
    Dim Msg As String
    Dim Partyyyy As String
    Dim Partmm As String
    Dim Partdd As String
    
    'Gibt den Datentyp des Datenbankfeldes zurück                   'Gerbing 04.02.2007
    'und formatiert aus Feldwert zB 26.05.1944 das Feld FDF als String im Format #Monat/Tag/Jahr#
    'die Datentypen stehen in ND.ListDatenTyp
    'Die Listbox ND.ListNutzerdefinierteFelder(unsortiert) speichert die nutzerdefinierten Felder in der
    'Reihenfolge der Columns von links nach rechts
    'Die Listbox ND.ListDatenTyp(unsortiert) speichert den Datentyp der nutzerdefinierten Felder in derselben
    'Reihenfolge wie die Listbox ND.ListNutzerdefinierteFelder(unsortiert)
    
    On Error GoTo Fehler
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        If ND.ListNutzerdefinierteFelder.ListItems(n) = Feldname Then Exit For
    Next n
    If gblnSQLServerVersion = True Then
        If ND.ListDatenTyp.List(n) = 135 Then
            'ich unterscheide zwischen Feldern mit Datum und Feldern mit Uhrzeit
            If InStr(1, Feldwert, ".", vbTextCompare) <> 0 Then
                'Beim sql server ist das Format yyyymmdd von der Landessprache unabhängig
                'm muss unbedingt 2-stellig sein und d muss unbedingt 2-stellig sein
                'beim 22.01.2012 kommt ohne Formatierung 2012122 beim Monat wird eine Null weggelassen
                Partyyyy = DatePart("yyyy", Feldwert)
                Partmm = DatePart("m", Feldwert)
                If Len(Partmm) = 1 Then
                    Partmm = "0" & Partmm
                End If
                Partdd = DatePart("d", Feldwert)
                If Len(Partdd) = 1 Then
                    Partdd = "0" & Partdd
                End If
                FDF = "'" & Partyyyy & Partmm & Partdd & "'"
            Else
                FDF = "'" & Feldwert & "'"
            End If
        End If
    Else
        If ND.ListDatenTyp.List(n) = Date Then
            'ich unterscheide zwischen Feldern mit Datum und Feldern mit Uhrzeit
            If InStr(1, Feldwert, ".", vbTextCompare) <> 0 Then
                FDF = "#" & DatePart("m", Feldwert) & "/" & DatePart("d", Feldwert) & "/" & DatePart("yyyy", Feldwert) & "#"
            Else
                FDF = "#" & Feldwert & "#"
            End If
        End If
    End If
    DatenTypDate = ND.ListDatenTyp.List(n)                          'Gerbing 04.02.2007
    Exit Function
Fehler:
    'msg = "Fehler bei der Umwandlung einer Datums-Angabe" & vbNewLine
    Msg = LoadResString(2183 + Sprache) & vbNewLine
    'msg = msg & "Feldname=" & Feldname & vbNewLine
    Msg = Msg & LoadResString(2184 + Sprache) & Feldname & vbNewLine
    'msg = msg & "Feldinhalt=" & Feldwert
    Msg = Msg & LoadResString(2185 + Sprache) & Feldwert
    'MsgBox Msg
    MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
End Function

Private Sub RefreshWerte()
    TJahr.Text = ""
    '-------------------------------------------------
    'die höchste gefundene Jahreszahl in die Combobox TJahr eintragen                       'Gerbing 08.11.2012
    'Dadurch wird ComboBoxenFüllen ausgelöst
    'SQL = "SELECT MAX(Jahr)From fotos"
    SQL = "SELECT MAX(" & (LoadResString(1023 + Sprache)) & ")From fotos;"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                                           'Gerbing 23.1.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    TJahr.Text = adoRs.Fields.Item(0)   'jetzt kommt TJahr_Change dran                      'Gerbing 08.11.2012
    '-------------------------------------------------
    TFileType.Text = "*"                                                                    'Gerbing 10.04.2016
    TSituation.Text = "*"
    TOrt.Text = "*"
    TLand.Text = "*"
    TSWF = "*"
    TPersonen.Text = "*"
    TPersonen.Enabled = True                                        'Gerbing 04.03.2013
'    SQLText = "Select * From Fotos ORDER BY Dateiname"              'Gerbing 09.06.2004
    SQLText = "Select * From Fotos ORDER BY " & LoadResString(1028 + Sprache)               'Gerbing 09.06.2004
    SQL = ""                                                                                'Gerbing 23.11.2008
    SQLWurdeBearbeitet = False                                      'Gerbing 06.07.2004
    SQLBearbeitenZähler = 0                                         'Gerbing 12.11.2004
    CheckSucheJedesFeld.Value = 0                                   'Gerbing 12.11.2004
    optAlleTreffer.Value = 1                                        'Gerbing 29.03.2005
    
    SQLText.Visible = False
    CheckVollesWort.Value = 0
    CheckSQL.Value = 0
    JUnd.Value = 1
    SUnd.Value = 1
    OUnd.Value = 1
    LUnd.Value = 1
    SWFUnd.Value = 1
    MP.TPerson1.Text = ""
    MP.TPerson2.Text = ""
    MP.TPerson3.Text = ""
    MP.TPerson4.Text = ""
    MP.TPerson5.Text = ""
    MP.AnzahlPersonen = 0
    MP.optDatumEinbeziehen.Value = False
    MP.OptNachKomplettemDateiename.Value = True
    MP.OptNurNachDateiname.Value = False
    MP.optDatumNichtEinbeziehen.Value = True
    Query.CheckWeitereFilterAktiv.Value = 0
    CheckWeitereFilterAktiv.Visible = False
    lblWeitereFilterAktiv.Visible = False                           'Gerbing 25.06.2013
    TPersonen.Text = "*"                                 'Gerbing 08.11.2005
    #If Proversion Then
        ND.cmbFeld1.ComboItems.RemoveAll
        ND.cmbFeld2.ComboItems.RemoveAll
        ND.cmbFeld3.ComboItems.RemoveAll
        ND.cmbFeld4.ComboItems.RemoveAll
        ND.cmbFeld5.ComboItems.RemoveAll
        ND.cmbVG1.Clear
        ND.cmbVG2.Clear
        ND.cmbVG3.Clear
        ND.cmbVG4.Clear
        ND.cmbVG5.Clear
        ND.Combo1.ComboItems.RemoveAll
        ND.Combo2.ComboItems.RemoveAll
        ND.Combo3.ComboItems.RemoveAll
        ND.combo4.ComboItems.RemoveAll
        ND.Combo5.ComboItems.RemoveAll
        ND.AnzahlFelder = 0
    #End If
    Query.CheckNutzerdefinierteFelder.Value = 0
    Query.CheckNutzerdefinierteFelder.Visible = False
    Query.lblNutzerdefinierteFelder.Visible = False
    CheckAudioFileExists.Value = 0                  'Gerbing 12.04.2006
    CheckUseAudioComments.Value = 0                 'Gerbing 12.04.2006
End Sub

Public Sub Beenden()
    Dim strManagement As String
    Dim rc As Boolean
    Dim sSource As String                                                                           'Gerbing 23.11.2017
    Dim sDest As String

    gblnComeFromBeenden = True                                                                       'Gerbing 30.11.2018
    On Error Resume Next
    If gblnSQLServerVersion = False Then
'        'auskommentiert 28.05.2019
        If Left(gstrNetzwerkDir, 1) <> "\" Then
            'Bei fotos.mdb auf einem anderen Rechner im Netzwerk kein Komprimieren machen           'Gerbing 27.08.2017
            On Error Resume Next
            rstsql.Close
            DBado.Close
            'DBEngine.CompactDatabase gstrFotosMdbLocation & "\fotos.mdb", gstrFotosMdbLocation & "\Newfotos.mdb" 'Gerbing 23.11.2017
            sSource = gstrFotosMdbLocation & "\fotos.mdb"
            sDest = gstrFotosMdbLocation & "\Newfotos.mdb"
            rc = file_delete(gstrFotosMdbLocation & "\Newfotos.mdb", , True)
            If CompactDB(sSource, sDest) Then
                'MsgBox "Compact complete"
                If file_path_exist(gstrFotosMdbLocation & "\Newfotos.mdb") = True Then
                    rc = file_delete(gstrFotosMdbLocation & "\fotos.mdb", , True)
                    'rc = file_copy(Quellname, Zielname)                                             'Gerbing 18.10.2017
                    rc = file_copy(gstrFotosMdbLocation & "\Newfotos.mdb", gstrFotosMdbLocation & "\fotos.mdb") 'Gerbing 18.10.2017
                    rc = file_delete(gstrFotosMdbLocation & "\Newfotos.mdb", , True)
                End If
                'MsgBox "Komprimieren der Datenbank wurde ausgeführt"
            Else
                'MsgBox "Komprimieren der Datenbank wurde versucht, aber konnte nicht ausgeführt werden"
            End If
            On Error GoTo 0
       End If
'       'auskommentiert 28.05.2019
    Else
        'Alle user müssen sich Einloggen. Beim Programmende erfolgt das Ausloggen
        'Käufer mit unbegrenzten Lizenzen müssen sich nicht ausloggen
        If gstrAllowedlicenses <> 99 Then
            Set rstsql = New ADODB.Recordset
            With rstsql
                .Source = "select * from loggedinusers where (username = N'" & gstrLoggedInName & "')"
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            
            'strManagement soll zum Zeitpunkt des Login enthalten 'OUT&Datum&Uhrzeit' und verschlüsselt mit dem username
            strManagement = "OUT" & Now
            strManagement = Crypt(strManagement, rstsql.Fields("username"), True)
            rstsql("LoggedIn") = False
            rstsql("Management") = strManagement
            rstsql.Update
            rstsql.Close
            DBado.Close                                                                                 'Gerbing 23.11.2017
        End If
    End If
    Call Form1.MediaPlayerStop                                                                          'Gerbing 30.11.2016
'    If frmGridAndThumb.blnComeFromBtnMitThumbnailsClick = True Then                                     'Gerbing 09.06.2015
'        TerminateEXE "fotos.EXE"
'        End                                                                                             'Gerbing 15.09.2015
'    End If
    
    gblnComeFromButtonF8 = True                                                                         'Gerbing 29.03.2015 sonst Absturz am Ende von fotos.exe
    Unload frmGridAndThumb
    Unload Hilfebx
    Unload KommentarForm
    'Unload Query                       'Gerbing 27.11.2012 auskommentiert weil beim Sprache wechseln bleibt es hier hängen
    'Unload QueryJedesFeld
    Unload MP
    Set Fso = Nothing                                                                                   'Gerbing 13.04.2015
    GdiplusShutdown m_lngInstance                                                                       'Gerbing 13.04.2015
    Set Form1.EXF = Nothing                                                                             'Gerbing 13.04.2015
End Sub

Private Sub ComboBoxenFüllen()                                                                          'Gerbing 08.11.2012
    Dim DateinamenErweiterung As String
    Dim pos As Long
    Dim KollFileType As New Collection
    Dim i As Long
    Dim blnGefunden As Boolean
    
    TSituation.ComboItems.RemoveAll
    'TOrt.comboitems.Removeall
    TOrt.ComboItems.RemoveAll
    TLand.ComboItems.RemoveAll
    TSWF.Clear
    TPersonen.ComboItems.RemoveAll
    TSituation.Text = "*"
    TOrt.Text = "*"
    TLand.Text = "*"
    TSWF = "*"                                                                                          'Gerbing 08.11.2012
    TFileType = "*"                                                                                     'Gerbing 14.07.2016
    TPersonen.Text = "*"
    'Gerbing 22.11.2014                                                                                 'Gerbing 22.11.2014
    MP.TPerson1.Text = ""
    MP.TPerson2.Text = ""
    MP.TPerson3.Text = ""
    MP.TPerson4.Text = ""
    MP.TPerson5.Text = ""
    MP.AnzahlPersonen = 0
    MP.optDatumEinbeziehen.Value = False
    MP.OptNachKomplettemDateiename.Value = True
    MP.OptNurNachDateiname.Value = False
    MP.optDatumNichtEinbeziehen.Value = True
    Query.CheckWeitereFilterAktiv.Value = 0
    CheckWeitereFilterAktiv.Visible = False
    lblWeitereFilterAktiv.Visible = False                                                               'Gerbing 22.11.2014

    'SQL = "SELECT DISTINCT Fotos.Situation From Fotos " & SQLJahreszahl & " AND ((Not (Fotos.Situation)='')) ORDER BY Situation;" 'Gerbing 08.11.2012
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1024 + Sprache) & " From Fotos " & SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1024 + Sprache) & ")='')) ORDER BY " & LoadResString(1024 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset(LoadResString(1024 + Sprache))) Then            'situation
            TSituation.ComboItems.Add Adodc1.Recordset(LoadResString(1024 + Sprache))
        End If
        Adodc1.Recordset.MoveNext
        DoEvents                                        'Gerbing 16.09.2004
    Loop
    'SQL = "SELECT DISTINCT Fotos.Ort FROM Fotos " & SQLJahreszahl & " AND ((Not (Fotos.Ort)='')) ORDER BY Ort;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1025 + Sprache) & " FROM Fotos " & SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1025 + Sprache) & ")='')) ORDER BY " & LoadResString(1025 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset(LoadResString(1025 + Sprache))) Then
            TOrt.ComboItems.Add Adodc1.Recordset(LoadResString(1025 + Sprache))
        End If
        Adodc1.Recordset.MoveNext
        DoEvents                                        'Gerbing 16.09.2004
    Loop
    'SQL = "SELECT DISTINCT Fotos.Land FROM Fotos " & SQLJahreszahl & " AND ((Not (Fotos.Land)='')) ORDER BY Land;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1026 + Sprache) & " FROM Fotos " & SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1026 + Sprache) & ")='')) ORDER BY " & LoadResString(1026 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset(LoadResString(1026 + Sprache))) Then
            TLand.ComboItems.Add Adodc1.Recordset(LoadResString(1026 + Sprache))
        End If
        Adodc1.Recordset.MoveNext
        DoEvents                                        'Gerbing 16.09.2004
    Loop
'    TSWF.AddItem LoadResString(1110 + Sprache) 'BELIEBIG   'Gerbing 16.11.2004
    TSWF.AddItem "*"
    'SQL = "SELECT DISTINCT Fotos.SWF FROM Fotos " & SQLJahreszahl & " AND ((Not (Fotos.SWF)='')) ORDER BY SWF;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1029 + Sprache) & " FROM Fotos " & SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1029 + Sprache) & ")='')) ORDER BY " & LoadResString(1029 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset(LoadResString(1029 + Sprache))) Then
            TSWF.AddItem Adodc1.Recordset(LoadResString(1029 + Sprache))
        End If
        Adodc1.Recordset.MoveNext
        DoEvents                                        'Gerbing 16.09.2004
    Loop
    TSWF.ListIndex = 0                                  'Gerbing 16.11.2004
    '--------------------------------------------------------------------------------------------------------------
    'ich will alle Dateinamenerweiterungen in der Combobox TFileType auswählen lassen           'Gerbing 14.07.2016
    Set KollFileType = Nothing
    KollFileType.Add "*"
    'SQL = "SELECT Fotos.Dateiname FROM Fotos " & SQLJahreszahl & " ORDER BY Dateiname;"
    SQL = "SELECT Fotos." & LoadResString(1028 + Sprache) & " FROM Fotos " & SQLJahresZahl & " ORDER BY " & LoadResString(1028 + Sprache)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseServer       'geht schneller als aduseClient
        .Source = SQL
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Query.Adodc1.Recordset.EOF
        'Die Kollektion KollFileType soll nur Dateitypen aufnehmen, die bisher nicht drin stehen
        'zuerst suche ich die Position des hintersten Punktes im Dateiname
        blnGefunden = False
        DateinamenErweiterung = Query.Adodc1.Recordset(LoadResString(1028 + Sprache))
        pos = Len(DateinamenErweiterung) - 1
        'Pos = 0 passiert dann wenn der Dateiname keinen Punkt enthält
        If pos <> 0 Then                                                                    'Gerbing 20.04.2020
            Do
                If StrComp(Mid(DateinamenErweiterung, pos, 1), ".", vbTextCompare) = 0 Then
                    Exit Do
                Else
                    pos = pos - 1
                    If pos = 0 Then Exit Do
                End If
            Loop
        End If                                                                              'Gerbing 20.04.2020
        If pos <> 0 Then
            DateinamenErweiterung = Mid(DateinamenErweiterung, pos + 1, Len(DateinamenErweiterung) - pos + 1)
            For i = 1 To KollFileType.Count
                If StrComp(KollFileType.Item(i), DateinamenErweiterung, vbTextCompare) = 0 Then
                    blnGefunden = True
                    Exit For
                End If
            Next i
            If blnGefunden = False Then
                KollFileType.Add DateinamenErweiterung
            End If
        End If
        Adodc1.Recordset.MoveNext
        DoEvents
    Loop
    'jetzt die Kollektion in die Combobox stellen
    TFileType.Clear
    For i = 1 To KollFileType.Count
        TFileType.AddItem KollFileType.Item(i)
    Next i
    TFileType.ListIndex = 0                                                                     'Gerbing 14.07.2016
    '--------------------------------------------------------------------------------------------------------------
    'SQL = "SELECT DISTINCT Fotos.Personen FROM Fotos " & SQLJahreszahl & " AND ((Not (Fotos.Personen)='')) ORDER BY Personen;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1027 + Sprache) & " FROM Fotos " & SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1027 + Sprache) & ")='')) ORDER BY " & LoadResString(1027 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                             'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Adodc1.Recordset = adoRs
    Do Until Query.Adodc1.Recordset.EOF
        If Not IsNull(Query.Adodc1.Recordset(LoadResString(1027 + Sprache))) Then
            TPersonen.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
        End If
        Adodc1.Recordset.MoveNext
        DoEvents                                        'Gerbing 16.09.2004
    Loop
End Sub

Private Sub TJahr_Change()                                                  'Gerbing 08.11.2012
    Dim rc As Boolean
    
    If TJahr = "*" Then GoTo HierJahreszahlPrüfen                           'Gerbing 17.01.2013
    Select Case Len(TJahr)
        Case 9
            GoTo HierJahreszahlPrüfen
        Case 4
            If IsNumeric(TJahr) Then GoTo HierJahreszahlPrüfen
        Case 5
            If Left(TJahr, 1) = "<" Or Left(TJahr, 1) = ">" Then
                GoTo HierJahreszahlPrüfen
            End If
        Case Else
            Exit Sub
    End Select
    Exit Sub
HierJahreszahlPrüfen:
    rc = JahreszahlPrüfen
    If rc <> 0 Then Exit Sub
    Call ComboBoxenFüllen
End Sub

Private Function JahreszahlPrüfen() As Boolean                              'Gerbing 08.11.2012
    'rc = 0 wenn kein Fehler vorliegt
    'rc = -1 bei Fehler
    
    JahreszahlPrüfen = False
    SQLJahr = TJahr
    
    If SQLJahr = "*" Then
        'JahrVon = 0
        'JahrBis = höchste gefundene Jahreszahl                             'Gerbing 08.11.2012
        'vergleich = "between"
        'SQL = "SELECT MAX(Jahr)From fotos"
        SQL = "SELECT MAX(" & (LoadResString(1023 + Sprache)) & ")From fotos;"
        'Set rst = db.OpenRecordset(SQL)
        Set adoRs = New ADODB.Recordset
        With adoRs
            .ActiveConnection = DBado                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            '.CursorLocation = Query.enumCursorOrt
            .Source = SQL
            '     .CacheSize = 2
            .Open
        End With
        JahrVon = 0
        Vergleich = "between"
        JahrBis = adoRs.Fields.Item(0)
        GoTo SQLAufbauen
    End If
    
    'If SQLJahr = UCase(LoadResString(1110 + Sprache)) Or SQLJahr = "*" Then        '1110=beliebig
    If StrComp(SQLJahr, (LoadResString(1110 + Sprache)), vbTextCompare) = 0 Then      '1110=beliebig
        SQL = "Select * from Fotos "
        'SQL = SQL & " where Jahr >= 0"                                     'Gerbing 10.10.2004
        SQL = SQL & " where " & LoadResString(1023 + Sprache) & " >= 0"     'Gerbing 08.11.2005
        Plus = Plus1
    Else
        Vergleich = "="
        pos1 = InStr(1, SQLJahr, ">", vbTextCompare)              'suche >
        If pos1 <> 0 Then
            Vergleich = ">"
            If Len(SQLJahr) <> 5 Then
                'MsgBox "Jahr muß eine 4-stellige Zahl sein"
                MsgBox LoadResString(2127 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
            SQLJahr = Right(SQLJahr, 4)
        End If
        pos1 = InStr(1, SQLJahr, "<", vbTextCompare)              'suche <
        If pos1 <> 0 Then
            Vergleich = "<"
            If Len(SQLJahr) <> 5 Then
                'MsgBox "Jahr muß eine 4-stellige Zahl sein"
                MsgBox LoadResString(2127 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
            SQLJahr = Right(SQLJahr, 4)
        End If
        pos1 = InStr(1, SQLJahr, "-", vbTextCompare)  'wie 1990-1998'
        If pos1 = 5 Then                '- muß auf Position 5 stehen
            If Len(SQLJahr) <> 9 Then
                'MsgBox "Ein Zeitraum muß in folgender Form angegeben werden '1990-1998'"
                MsgBox LoadResString(2128 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
            Vergleich = "between"
            JahrVon = Left(SQLJahr, 4)
            JahrBis = Right(SQLJahr, 4)
        End If
        If Vergleich = "between" Then
            If Not IsNumeric(JahrVon) Then
                'MsgBox "Ein Zeitraum muß in folgender Form angegeben werden '1990-1998'"
                MsgBox LoadResString(2128 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
            If Not IsNumeric(JahrBis) Then
                'MsgBox "Ein Zeitraum muß in folgender Form angegeben werden '1990-1998'"
                MsgBox LoadResString(2128 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
            If JahrVon >= JahrBis Then
                'MsgBox "Ein Zeitraum muß in folgender Form angegeben werden '1990-1998'"
                MsgBox LoadResString(2128 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
        Else
            If Not IsNumeric(SQLJahr) Or Len(SQLJahr) <> 4 Then
                'MsgBox "Jahr muß eine 4-stellige Zahl sein"
                MsgBox LoadResString(2127 + Sprache)
                TJahr.SetFocus
                Me.MousePointer = vbDefault
                JahreszahlPrüfen = True
                Exit Function
            End If
        End If
SQLAufbauen:
        SQL = "Select * from Fotos "
        If Vergleich = "between" Then
            'SQL = SQL & "where Jahr " & Vergleich & " " & JahrVon & " And " & JahrBis
            SQL = SQL & "where " & LoadResString(1023 + Sprache) & " " & Vergleich & " " & JahrVon & " And " & JahrBis
            SQLJahresZahl = "where " & LoadResString(1023 + Sprache) & " " & Vergleich & " " & JahrVon & " And " & JahrBis
        Else
            'SQL = SQL & "where Jahr " & Vergleich & " " & SQLJahr
            SQL = SQL & "where " & LoadResString(1023 + Sprache) & " " & Vergleich & " " & SQLJahr
            SQLJahresZahl = "where " & LoadResString(1023 + Sprache) & " " & Vergleich & " " & SQLJahr
        End If
        Plus = Plus1
    End If
End Function

Private Sub TLand_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If TLand.Text <> "*" Then
        If TSituation.Text = "*" Then Call ComboSituation
        If TOrt.Text = "*" Then Call ComboOrt
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub

Private Sub TLand_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    If Not newSelectedItem Is Nothing Then
        TLand.Text = newSelectedItem
        If TSituation.Text = "*" Then Call ComboSituation
        If TOrt.Text = "*" Then Call ComboOrt
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub

Private Sub ComboOrt()                                                                                  'Gerbing 08.11.2012
    Dim strTemp As String
    Dim strDistinct As String
    
    strDistinct = LoadResString(1025 + Sprache)             '1025=ort
    SQL = "SELECT DISTINCT " & strDistinct & " From Fotos " & SQLJahresZahl
    If TSituation.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1024 + Sprache) & " Like '%" & TSituation.Text & "%'"      '1024=situation
    End If
    If TLand.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1026 + Sprache) & " Like '%" & TLand.Text & "%'"           '1026=land
    End If
    If TPersonen.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1027 + Sprache) & " Like '%" & TPersonen.Text & "%'"       '1027=personen
    End If
    If TOrt.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1025 + Sprache) & " Like '%" & TOrt.Text & "%'"             '1025=ort
    End If
    SQL = SQL & " ORDER BY " & strDistinct & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    On Error Resume Next
    With adoRs
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        Msg = "errornumber=" & Err.Number & vbNewLine
        Msg = Msg & "errortext=" & Err.Description
        MsgBox Msg
        Exit Sub
    End If
    On Error GoTo 0
    Set Adodc1.Recordset = adoRs
    If Not Adodc1.Recordset.EOF Then
        strTemp = TOrt.Text
        TOrt.ComboItems.RemoveAll
        blnExitChange = True
        TOrt.Text = strTemp
    End If
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset.Fields.Item(0)) And Adodc1.Recordset.Fields.Item(0) <> "" Then
            TOrt.ComboItems.Add Adodc1.Recordset.Fields.Item(0)
        End If
        Adodc1.Recordset.MoveNext
        DoEvents
    Loop
    blnExitChange = False
End Sub

Private Sub ComboLand()                                                                             'Gerbing 08.11.2012
    Dim strTemp As String
    Dim strDistinct As String
    
    strDistinct = LoadResString(1026 + Sprache)             '1026=land
    SQL = "SELECT DISTINCT " & strDistinct & " From Fotos " & SQLJahresZahl
    If TSituation.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1024 + Sprache) & " Like '%" & TSituation.Text & "%'"   '1024=situation
    End If
    If TOrt.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1025 + Sprache) & " Like '%" & TOrt.Text & "%'"         '1025=ort
    End If
    If TPersonen.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1027 + Sprache) & " Like '%" & TPersonen.Text & "%'"    '1027=personen
    End If
    If TLand.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1026 + Sprache) & " Like '%" & TLand.Text & "%'"        '1026=land
    End If
    SQL = SQL & " ORDER BY " & strDistinct & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    On Error Resume Next
    With adoRs
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        Msg = "errornumber=" & Err.Number & vbNewLine
        Msg = Msg & "errortext=" & Err.Description
        MsgBox Msg
        Exit Sub
    End If
    On Error GoTo 0
    Set Adodc1.Recordset = adoRs
    If Not Adodc1.Recordset.EOF Then
        strTemp = TLand.Text
        TLand.ComboItems.RemoveAll
        blnExitChange = True
        TLand.Text = strTemp
    End If
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset.Fields.Item(0)) And Adodc1.Recordset.Fields.Item(0) <> "" Then
            TLand.ComboItems.Add Adodc1.Recordset.Fields.Item(0)
        End If
        Adodc1.Recordset.MoveNext
        DoEvents
    Loop
    blnExitChange = False
End Sub

Private Sub ComboSituation()                                                                        'Gerbing 08.11.2012
    Dim strTemp As String
    Dim strDistinct As String
    
    strDistinct = LoadResString(1024 + Sprache)             '1024=situation
    SQL = "SELECT DISTINCT " & strDistinct & " From Fotos " & SQLJahresZahl
    If TLand.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1026 + Sprache) & " Like '%" & TLand.Text & "%'"        '1026=land
    End If
    If TOrt.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1025 + Sprache) & " Like '%" & TOrt.Text & "%'"         '1025=ort
    End If
    If TPersonen.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1027 + Sprache) & " Like '%" & TPersonen.Text & "%'"    '1027=personen
    End If
    If TSituation.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1024 + Sprache) & " Like '%" & TSituation.Text & "%'"   '1024=situation
    End If
    SQL = SQL & " ORDER BY " & strDistinct & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    On Error Resume Next
    With adoRs
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        Msg = "errornumber=" & Err.Number & vbNewLine
        Msg = Msg & "errortext=" & Err.Description
        MsgBox Msg
        Exit Sub
    End If
    On Error GoTo 0
    Set Adodc1.Recordset = adoRs
    If Not Adodc1.Recordset.EOF Then
        strTemp = TSituation.Text
        TSituation.ComboItems.RemoveAll
        blnExitChange = True
        TSituation.Text = strTemp
    End If
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset.Fields.Item(0)) And Adodc1.Recordset.Fields.Item(0) <> "" Then
            TSituation.ComboItems.Add Adodc1.Recordset.Fields.Item(0)
        End If
        Adodc1.Recordset.MoveNext
        DoEvents
    Loop
    blnExitChange = False
End Sub

Private Sub ComboPersonen()                                                                         'Gerbing 08.11.2012
    Dim strTemp As String
    Dim strDistinct As String
    
    strDistinct = LoadResString(1027 + Sprache)             '1027=personen
    SQL = "SELECT DISTINCT " & strDistinct & " From Fotos " & SQLJahresZahl
    If TLand.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1026 + Sprache) & " Like '%" & TLand.Text & "%'"        '1026=land
    End If
    If TOrt.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1025 + Sprache) & " Like '%" & TOrt.Text & "%'"         '1025=ort
    End If
    If TSituation.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1024 + Sprache) & " Like '%" & TSituation.Text & "%'"   '1024=situation
    End If
    If TPersonen.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1027 + Sprache) & " Like '%" & TPersonen.Text & "%'"    '1027=personen
    End If
    SQL = SQL & " ORDER BY " & strDistinct & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    On Error Resume Next
    With adoRs
        .ActiveConnection = DBado                                               'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If Err.Number <> 0 Then
        Msg = "errornumber=" & Err.Number & vbNewLine
        Msg = Msg & "errortext=" & Err.Description
        MsgBox Msg
        Exit Sub
    End If
    On Error GoTo 0
    Set Adodc1.Recordset = adoRs
    If Not Adodc1.Recordset.EOF Then
        strTemp = TPersonen.Text
        TPersonen.ComboItems.RemoveAll
        blnExitChange = True
        TPersonen.Text = strTemp
    End If
    Do Until Adodc1.Recordset.EOF
        If Not IsNull(Adodc1.Recordset.Fields.Item(0)) And Adodc1.Recordset.Fields.Item(0) <> "" Then
            TPersonen.ComboItems.Add Adodc1.Recordset.Fields.Item(0)
        End If
        Adodc1.Recordset.MoveNext
        DoEvents
    Loop
    blnExitChange = False
End Sub

Private Sub TOrt_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If TOrt.Text <> "*" Then
        If TSituation.Text = "*" Then Call ComboSituation
        If TLand.Text = "*" Then Call ComboLand
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub

Private Sub TOrt_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    If Not newSelectedItem Is Nothing Then
        TOrt.Text = newSelectedItem
        If TSituation.Text = "*" Then Call ComboSituation
        If TLand.Text = "*" Then Call ComboLand
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub

Private Sub TPersonen_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If TPersonen.Text <> "*" Then
        If TOrt.Text = "*" Then Call ComboOrt
        If TLand.Text = "*" Then Call ComboLand
        If TSituation.Text = "*" Then Call ComboSituation
    End If
End Sub

Private Sub TPersonen_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    If Not newSelectedItem Is Nothing Then
        TPersonen.Text = newSelectedItem
        If TSituation.Text = "*" Then Call ComboSituation
        If TOrt.Text = "*" Then Call ComboOrt
        If TLand.Text = "*" Then Call ComboLand
    End If
End Sub

Private Sub TSituation_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If TSituation.Text <> "*" Then
        If TOrt.Text = "*" Then Call ComboOrt
        If TLand.Text = "*" Then Call ComboLand
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub

Private Sub TSituation_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    If Not newSelectedItem Is Nothing Then
        TSituation.Text = newSelectedItem
        If TOrt.Text = "*" Then Call ComboOrt
        If TLand.Text = "*" Then Call ComboLand
        If TPersonen.Text = "*" Then Call ComboPersonen
    End If
End Sub
