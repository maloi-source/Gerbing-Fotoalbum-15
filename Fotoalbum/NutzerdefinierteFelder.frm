VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form ND 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Suche nach nutzerdefinierten Feldern"
   ClientHeight    =   5376
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   13764
   Icon            =   "NutzerdefinierteFelder.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5376
   ScaleWidth      =   13764
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   240
      Picture         =   "NutzerdefinierteFelder.frx":038A
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   38
      Top             =   160
      Width           =   360
   End
   Begin VB.CommandButton btnSucheGEODaten 
      Caption         =   "Suche Fotos mit &GEO-Daten"
      Height          =   492
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   37
      Top             =   120
      Width           =   5532
   End
   Begin CBLCtlsLibUCtl.ListBox ListNutzerdefinierteFelder 
      Height          =   252
      Left            =   5880
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   1932
      _cx             =   3408
      _cy             =   444
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048811
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
   Begin VB.ListBox ListDatenTyp 
      Height          =   240
      Left            =   5880
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   492
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   5532
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Keine Suche nach nutzerdefinierten Feldern"
      Height          =   492
      Left            =   8160
      TabIndex        =   18
      Top             =   4680
      Width           =   5532
   End
   Begin VB.Frame Frame2 
      Caption         =   "Wählen Sie Feldnamen und Vergleichsoperand und geben Sie Feldwerte ein"
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13572
      Begin VB.ComboBox cmbVG5 
         Height          =   288
         Left            =   6480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   24
         Top             =   3000
         Width           =   732
      End
      Begin VB.ComboBox cmbVG4 
         Height          =   288
         Left            =   6480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   2400
         Width           =   732
      End
      Begin VB.ComboBox cmbVG3 
         Height          =   288
         Left            =   6480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   1800
         Width           =   732
      End
      Begin VB.ComboBox cmbVG2 
         Height          =   288
         Left            =   6480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   21
         Top             =   1200
         Width           =   732
      End
      Begin VB.ComboBox cmbVG1 
         Height          =   288
         Left            =   6480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   20
         Top             =   600
         Width           =   732
      End
      Begin VB.Frame Frame6 
         Height          =   492
         Left            =   11040
         TabIndex        =   10
         Top             =   480
         Width           =   2412
         Begin VB.OptionButton SOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   12
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton Und1 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   11
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   528
         Left            =   11040
         TabIndex        =   7
         Top             =   1080
         Width           =   2412
         Begin VB.OptionButton OOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   9
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton Und2 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   8
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   492
         Left            =   11040
         TabIndex        =   4
         Top             =   1680
         Width           =   2412
         Begin VB.OptionButton LOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   6
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton Und3 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   5
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   528
         Left            =   11040
         TabIndex        =   1
         Top             =   2280
         Width           =   2412
         Begin VB.OptionButton SWFOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   3
            Top             =   150
            Width           =   1092
         End
         Begin VB.OptionButton Und4 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   2
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo1 
         Height          =   288
         Left            =   7320
         TabIndex        =   27
         Top             =   600
         Width           =   3612
         _cx             =   6371
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":31D7
         Text            =   "NutzerdefinierteFelder.frx":31F7
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo2 
         Height          =   288
         Left            =   7320
         TabIndex        =   33
         Top             =   1200
         Width           =   3612
         _cx             =   6371
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3217
         Text            =   "NutzerdefinierteFelder.frx":3237
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo3 
         Height          =   288
         Left            =   7320
         TabIndex        =   34
         Top             =   1800
         Width           =   3612
         _cx             =   6371
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3257
         Text            =   "NutzerdefinierteFelder.frx":3277
      End
      Begin CBLCtlsLibUCtl.ComboBox combo4 
         Height          =   288
         Left            =   7320
         TabIndex        =   35
         Top             =   2400
         Width           =   3612
         _cx             =   6371
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3297
         Text            =   "NutzerdefinierteFelder.frx":32B7
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo5 
         Height          =   288
         Left            =   7320
         TabIndex        =   36
         Top             =   3000
         Width           =   3612
         _cx             =   6371
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":32D7
         Text            =   "NutzerdefinierteFelder.frx":32F7
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld1 
         Height          =   288
         Left            =   1920
         TabIndex        =   26
         Top             =   600
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3317
         Text            =   "NutzerdefinierteFelder.frx":3337
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld2 
         Height          =   288
         Left            =   1920
         TabIndex        =   29
         Top             =   1200
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3357
         Text            =   "NutzerdefinierteFelder.frx":3377
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld3 
         Height          =   288
         Left            =   1920
         TabIndex        =   30
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
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3397
         Text            =   "NutzerdefinierteFelder.frx":33B7
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld4 
         Height          =   288
         Left            =   1920
         TabIndex        =   31
         Top             =   2400
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":33D7
         Text            =   "NutzerdefinierteFelder.frx":33F7
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld5 
         Height          =   288
         Left            =   1920
         TabIndex        =   32
         Top             =   3000
         Width           =   4452
         _cx             =   7853
         _cy             =   508
         AcceptNumbersOnly=   0   'False
         Appearance      =   3
         AutoHorizontalScrolling=   -1  'True
         BackColor       =   -2147483643
         BorderStyle     =   0
         CharacterConversion=   0
         DisabledEvents  =   5351
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
         CueBanner       =   "NutzerdefinierteFelder.frx":3417
         Text            =   "NutzerdefinierteFelder.frx":3437
      End
      Begin VB.Label lblFeldname1 
         Caption         =   "Feldname"
         Height          =   252
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1572
      End
      Begin VB.Label lblFeldname2 
         Caption         =   "Feldname"
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label lblFeldname3 
         Caption         =   "Feldname"
         Height          =   252
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1572
      End
      Begin VB.Label lblFeldname4 
         Caption         =   "Feldname"
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1572
      End
      Begin VB.Label lblFeldname5 
         Caption         =   "Feldname"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1572
      End
   End
End
Attribute VB_Name = "ND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Proversion Then
Option Explicit
'Gerbing 23.06.2011
'In der Entwicklungsumgebung eingestellt: ControlBox = False, weil bei Form_Unload auch ListNutzerdefinierteFelder.ListCount = 0
'gesetzt wird. Das führt zur Fehlermeldung
'"Sie können erst dann nutzerdefinierte Felder in die Suche einbeziehen,"
'"nachdem Sie nutzerdefinierte Felder angelegt haben."
'"Lesen Sie in der Hilfe, wie nutzerdefinierte Felder angelegt werden"
    Public AnzahlFelder As Long
    Public Plus1 As String
    Public Plus2 As String
    Public Plus3 As String
    Public Plus4 As String

Private Sub btnAbbrechen_Click()
    cmbFeld1.ComboItems.RemoveAll
    cmbFeld2.ComboItems.RemoveAll
    cmbFeld3.ComboItems.RemoveAll
    cmbFeld4.ComboItems.RemoveAll
    cmbFeld5.ComboItems.RemoveAll
    cmbVG1.Clear
    cmbVG2.Clear
    cmbVG3.Clear
    cmbVG4.Clear
    cmbVG5.Clear
    Combo1.ComboItems.RemoveAll
    Combo2.ComboItems.RemoveAll
    Combo3.ComboItems.RemoveAll
    combo4.ComboItems.RemoveAll
    Combo5.ComboItems.RemoveAll
    Combo1.Text = ""                                    'Gerbing 04.02.2007
    Combo2.Text = ""
    Combo3.Text = ""
    combo4.Text = ""
    Combo5.Text = ""
    AnzahlFelder = 0
    Query.CheckNutzerdefinierteFelder.Value = 0
    Query.CheckNutzerdefinierteFelder.Visible = False
    Query.lblNutzerdefinierteFelder.Visible = False     'Gerbing 25.06.2013
    gstrGEOStartPunkt = ""                              'Gerbing 05.09.2016
    gstrGEOEndPunkt = ""                                'Gerbing 05.09.2016
    Me.Hide
End Sub

Private Sub btnSucheGEODaten_Click()
    Dim SQL As String
    
    'Kontrollieren, ob es die Felder GPSLatitude und GPSLongitude in der Tabelle fotos gibt 'Gerbing 05.09.2016
    'wenn nicht, MsgBox zeigen und Abbrechen
    Dim rc As Integer
    
    rc = Form1.GPSFelderPrüfen                                                                  'Gerbing 02.10.2019
    If rc = 0 Then Exit Sub                                                                     'Gerbing 02.10.2019
    frmGPSRechteck.Show 1                                                                       'Gerbing 02.10.2019
End Sub

Private Sub cmbFeld1_KeyPress(KeyAscii As Integer)  'Gerbing 10.06.2005
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        'cmbFeld1.ListIndex = -1
        cmbFeld1.Text = ""
        cmbVG1.ListIndex = -1
        'Combo1.ListIndex = -1
        Combo1.Text = ""
    End If
End Sub

Private Sub cmbFeld1_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    Dim SQL As String
    Dim Msg As String
    
    If cmbFeld1.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
'    If Combo1.Text <> "" Then Exit Sub                  'Gerbing 29.12.2005
    Combo1.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    cmbVG1.Clear                                        'Gerbing 04.02.2007
    'Alle Werte nach Combo1 stellen, die zu diesem Feld in der Datenbank gefunden werden
    'aber nicht bei ...Fields(...).Type = dbBoolean, da gibts nur true oder False       'Gerbing 04.02.2007
    'abhängig von ...Fields(...).Type müssen die möglichen Vergleichsoperanden eingestellt werden
    'bei dbText gibt es nur = und <>
    'bei dbBoolean gibt es nur = und <>
    'bei Hyperlink gibt es nur = und <>
    Select Case DatenTyp(cmbFeld1)
        Case 1, 11 '1=dbBoolean '11=Boolean bei SQL-Server                              'Gerbing 23.11.2017
            ND.cmbVG1.AddItem "="
            ND.cmbVG1.AddItem "<>"
        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 23.11.2017
'            ND.cmbVG1.AddItem "="
'            ND.cmbVG1.AddItem "<>"

            ND.cmbVG1.AddItem "="
            ND.cmbVG1.AddItem ">"
            ND.cmbVG1.AddItem "<"
            ND.cmbVG1.AddItem ">="
            ND.cmbVG1.AddItem "<="
            ND.cmbVG1.AddItem "<>"
            ND.cmbVG1.AddItem "like"                                                    'Gerbing 08.05.2019

        Case 12   'Hyperlink
            ND.cmbVG1.AddItem "="
            ND.cmbVG1.AddItem "<>"

        Case Else 'Zahlen
            ND.cmbVG1.AddItem "="
            ND.cmbVG1.AddItem ">"
            ND.cmbVG1.AddItem "<"
            ND.cmbVG1.AddItem ">="
            ND.cmbVG1.AddItem "<="
            ND.cmbVG1.AddItem "<>"
    End Select
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld1.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld1.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld1.Text & "];"
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
    'wenn als Suchergebnis rstsql.EOF kommt, muss eine Warnung kommen, dass in diesem Feld keinerlei Werte
    'gespeichert sind                                                                   'Gerbing 04.02.2007
    If rstsql.EOF Then
'        msg = "Im Feld '" & cmbFeld1 & "' sind keine Werte gespeichert." & vbNewLine
'        msg = msg & "Sie können nach diesen Werten nicht suchen. Wählen Sie ein anderes Feld"
        Msg = LoadResString(2321 + Sprache) & cmbFeld1.Text & LoadResString(2322 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2323 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-------------------------------------
    If DatenTyp(cmbFeld1) = 1 Or DatenTyp(cmbFeld1) = 11 Then                   'Gerbing 23.11.2017
        If gblnSQLServerVersion = True Then
            Combo1.ComboItems.Add 0
            Combo1.ComboItems.Add 1
        Else
            'Combo1.AddItem "Wahr"
            'Combo1.AddItem "Falsch"
            Combo1.ComboItems.Add LoadResString(2324 + Sprache)
            Combo1.ComboItems.Add LoadResString(2325 + Sprache)
        End If
    Else
        Do Until rstsql.EOF
            If Not IsNull(rstsql.Fields(0)) Then
                Combo1.ComboItems.Add rstsql.Fields(0)
            End If
            rstsql.MoveNext
            DoEvents
        Loop
    End If
    If Combo1.ComboItems.Count <> 0 Then
        'Combo1.ListIndex = 0
    End If

End Sub

Private Sub cmbFeld2_KeyPress(KeyAscii As Integer)  'Gerbing 10.06.2005
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        'cmbFeld2.ListIndex = -1
        cmbFeld2.Text = ""
        cmbVG2.ListIndex = -1
        'Combo2.ListIndex = -1
        Combo2.Text = ""
    End If
End Sub

Private Sub cmbFeld2_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    Dim SQL As String
    Dim Msg As String
    
    If cmbFeld2.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
'    If Combo2.Text <> "" Then Exit Sub                  'Gerbing 29.12.2005
    Combo2.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    cmbVG2.Clear                                        'Gerbing 04.02.2007
    'Alle Werte nach Combo2 stellen, die zu diesem Feld in der Datenbank gefunden werden
    'aber nicht bei ...Fields(...).Type = dbBoolean, da gibts nur true oder False       'Gerbing 04.02.2007
    'abhängig von ...Fields(...).Type müssen die möglichen Vergleichsoperanden eingestellt werden
    'bei dbText gibt es nur = und <>
    'bei dbBoolean gibt es nur = und <>
    'bei Hyperlink gibt es nur = und <>
    Select Case DatenTyp(cmbFeld2)
        Case 1, 11 '1=dbBoolean '11=Boolean bei SQL-Server                              'Gerbing 23.11.2017
            ND.cmbVG2.AddItem "="
            ND.cmbVG2.AddItem "<>"
        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 23.11.2017
'            ND.cmbVG2.AddItem "="
'            ND.cmbVG2.AddItem "<>"
            ND.cmbVG2.AddItem "="
            ND.cmbVG2.AddItem ">"
            ND.cmbVG2.AddItem "<"
            ND.cmbVG2.AddItem ">="
            ND.cmbVG2.AddItem "<="
            ND.cmbVG2.AddItem "<>"
            ND.cmbVG2.AddItem "like"                                                    'Gerbing 08.05.2019
        Case 12   'Hyperlink
            ND.cmbVG2.AddItem "="
            ND.cmbVG2.AddItem "<>"
        Case Else 'Zahlen
            ND.cmbVG2.AddItem "="
            ND.cmbVG2.AddItem ">"
            ND.cmbVG2.AddItem "<"
            ND.cmbVG2.AddItem ">="
            ND.cmbVG2.AddItem "<="
            ND.cmbVG2.AddItem "<>"
    End Select
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld2.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld2.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld2.Text & "];"
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
    'wenn als Suchergebnis rstsql.EOF kommt, muss eine Warnung kommen, dass in diesem Feld keinerlei Werte
    'gespeichert sind                                                                   'Gerbing 04.02.2007
    If rstsql.EOF Then
'        msg = "Im Feld '" & cmbFeld2 & "' sind keine Werte gespeichert." & vbNewLine
'        msg = msg & "Sie können nach diesen Werten nicht suchen. Wählen Sie ein anderes Feld"
        Msg = LoadResString(2321 + Sprache) & cmbFeld2.Text & LoadResString(2322 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2323 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-------------------------------------
    If DatenTyp(cmbFeld2) = 1 Or DatenTyp(cmbFeld2) = 11 Then               'Gerbing 23.11.2017
        If gblnSQLServerVersion = True Then
            Combo2.ComboItems.Add 0
            Combo2.ComboItems.Add 1
        Else
            'Combo2.comboitems.add "Wahr"
            'Combo2.comboitems.add "Falsch"
            Combo2.ComboItems.Add LoadResString(2324 + Sprache)
            Combo2.ComboItems.Add LoadResString(2325 + Sprache)
        End If
    Else
        Do Until rstsql.EOF
            If Not IsNull(rstsql.Fields(0)) Then
                Combo2.ComboItems.Add rstsql.Fields(0)
            End If
            rstsql.MoveNext
            DoEvents
        Loop
    End If
    If Combo2.ComboItems.Count <> 0 Then
        'Combo2.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld3_KeyPress(KeyAscii As Integer)  'Gerbing 10.06.2005
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        'cmbFeld3.ListIndex = -1
        cmbFeld3.Text = ""
        cmbVG3.ListIndex = -1
        'Combo3.ListIndex = -1
        Combo3.Text = ""
    End If
End Sub

Private Sub cmbFeld3_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    Dim SQL As String
    Dim Msg As String
    
    If cmbFeld3.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
'    If Combo3.Text <> "" Then Exit Sub                  'Gerbing 29.12.2005
    Combo3.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    cmbVG3.Clear                                        'Gerbing 04.02.2007
    'Alle Werte nach Combo3 stellen, die zu diesem Feld in der Datenbank gefunden werden
    'aber nicht bei ...Fields(...).Type = dbBoolean, da gibts nur true oder False       'Gerbing 04.02.2007
    'abhängig von ...Fields(...).Type müssen die möglichen Vergleichsoperanden eingestellt werden
    'bei dbText gibt es nur = und <>
    'bei dbBoolean gibt es nur = und <>
    'bei Hyperlink gibt es nur = und <>
    Select Case DatenTyp(cmbFeld3.Text)
        Case 1, 11 '1=dbBoolean '11=Boolean bei SQL-Server                              'Gerbing 23.11.2017
            ND.cmbVG3.AddItem "="
            ND.cmbVG3.AddItem "<>"
        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 23.11.2017
'            ND.cmbVG3.AddItem "="
'            ND.cmbVG3.AddItem "<>"
            ND.cmbVG3.AddItem "="
            ND.cmbVG3.AddItem ">"
            ND.cmbVG3.AddItem "<"
            ND.cmbVG3.AddItem ">="
            ND.cmbVG3.AddItem "<="
            ND.cmbVG3.AddItem "<>"
            ND.cmbVG3.AddItem "like"                                                    'Gerbing 08.05.2019
        Case 12   'Hyperlink
            ND.cmbVG3.AddItem "="
            ND.cmbVG3.AddItem "<>"
        Case Else 'Zahlen
            ND.cmbVG3.AddItem "="
            ND.cmbVG3.AddItem ">"
            ND.cmbVG3.AddItem "<"
            ND.cmbVG3.AddItem ">="
            ND.cmbVG3.AddItem "<="
            ND.cmbVG3.AddItem "<>"
    End Select
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld3.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld3.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld3.Text & "];"
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
    'wenn als Suchergebnis rstsql.EOF kommt, muss eine Warnung kommen, dass in diesem Feld keinerlei Werte
    'gespeichert sind                                                                   'Gerbing 04.02.2007
    If rstsql.EOF Then
'        msg = "Im Feld '" & cmbFeld3 & "' sind keine Werte gespeichert." & vbNewLine
'        msg = msg & "Sie können nach diesen Werten nicht suchen. Wählen Sie ein anderes Feld"
        Msg = LoadResString(2321 + Sprache) & cmbFeld3.Text & LoadResString(2322 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2323 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-------------------------------------
    If DatenTyp(cmbFeld3) = 1 Or DatenTyp(cmbFeld3) = 11 Then                   'Gerbing 23.11.2017
        If gblnSQLServerVersion = True Then
            Combo3.ComboItems.Add 0
            Combo3.ComboItems.Add 1
        Else
            'Combo3.comboitems.add "Wahr"
            'Combo3.comboitems.add "Falsch"
            Combo3.ComboItems.Add LoadResString(2324 + Sprache)
            Combo3.ComboItems.Add LoadResString(2325 + Sprache)
        End If
    Else
        Do Until rstsql.EOF
            If Not IsNull(rstsql.Fields(0)) Then
                Combo3.ComboItems.Add rstsql.Fields(0)
            End If
            rstsql.MoveNext
            DoEvents
        Loop
    End If
    If Combo3.ComboItems.Count <> 0 Then
        'Combo3.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld4_KeyPress(KeyAscii As Integer)  'Gerbing 10.06.2005
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        'cmbFeld4.ListIndex = -1
        cmbFeld4.Text = ""
        cmbVG4.ListIndex = -1
        'combo4.ListIndex = -1
        combo4.Text = ""
    End If
End Sub

Private Sub cmbFeld4_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    Dim SQL As String
    Dim Msg As String
    
    If cmbFeld4.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
'    If Combo4.Text <> "" Then Exit Sub                  'Gerbing 29.12.2005
    combo4.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    cmbVG4.Clear                                       'Gerbing 04.02.2007
    'Alle Werte nach Combo4 stellen, die zu diesem Feld in der Datenbank gefunden werden
    'aber nicht bei ...Fields(...).Type = dbBoolean, da gibts nur true oder False       'Gerbing 04.02.2007
    'abhängig von ...Fields(...).Type müssen die möglichen Vergleichsoperanden eingestellt werden
    'bei dbText gibt es nur = und <>
    'bei dbBoolean gibt es nur = und <>
    'bei Hyperlink gibt es nur = und <>
    Select Case DatenTyp(cmbFeld4.Text)
        Case 1, 11 '1=dbBoolean '11=Boolean bei SQL-Server                              'Gerbing 23.11.2017
            ND.cmbVG4.AddItem "="
            ND.cmbVG4.AddItem "<>"
        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 23.11.2017
'            ND.cmbVG4.AddItem "="
'            ND.cmbVG4.AddItem "<>"
            ND.cmbVG4.AddItem "="
            ND.cmbVG4.AddItem ">"
            ND.cmbVG4.AddItem "<"
            ND.cmbVG4.AddItem ">="
            ND.cmbVG4.AddItem "<="
            ND.cmbVG4.AddItem "<>"
            ND.cmbVG4.AddItem "like"                                                    'Gerbing 08.05.2019
        Case 12   'Hyperlink
            ND.cmbVG4.AddItem "="
            ND.cmbVG4.AddItem "<>"
        Case Else 'Zahlen
            ND.cmbVG4.AddItem "="
            ND.cmbVG4.AddItem ">"
            ND.cmbVG4.AddItem "<"
            ND.cmbVG4.AddItem ">="
            ND.cmbVG4.AddItem "<="
            ND.cmbVG4.AddItem "<>"
    End Select
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld4.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld4.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld4.Text & "];"
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
    'wenn als Suchergebnis rstsql.EOF kommt, muss eine Warnung kommen, dass in diesem Feld keinerlei Werte
    'gespeichert sind                                                                   'Gerbing 04.02.2007
    If rstsql.EOF Then
'        msg = "Im Feld '" & cmbFeld4 & "' sind keine Werte gespeichert." & vbNewLine
'        msg = msg & "Sie können nach diesen Werten nicht suchen. Wählen Sie ein anderes Feld"
        Msg = LoadResString(2321 + Sprache) & cmbFeld4.Text & LoadResString(2322 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2323 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-------------------------------------
    If DatenTyp(cmbFeld4) = 1 Or DatenTyp(cmbFeld4.Text) = 11 Then              'Gerbing 23.11.2017
        If gblnSQLServerVersion = True Then
            combo4.ComboItems.Add 0
            combo4.ComboItems.Add 1
        Else
            'Combo4.comboitems.add "Wahr"
            'Combo4.comboitems.add "Falsch"
            combo4.ComboItems.Add LoadResString(2324 + Sprache)
            combo4.ComboItems.Add LoadResString(2325 + Sprache)
        End If
    Else
        Do Until rstsql.EOF
            If Not IsNull(rstsql.Fields(0)) Then
                combo4.ComboItems.Add rstsql.Fields(0)
            End If
            rstsql.MoveNext
            DoEvents
        Loop
    End If
    If combo4.ComboItems.Count <> 0 Then
        'combo4.ListIndex = 0
    End If
End Sub

Private Sub cmbFeld5_KeyPress(KeyAscii As Integer)  'Gerbing 10.06.2005
    If KeyAscii = 8 Or KeyAscii = 44 Then           '8=Return-Taste 44=Entf-Taste im numerischen Tastenfeld
        'cmbFeld5.ListIndex = -1
        cmbFeld5.Text = ""
        cmbVG5.ListIndex = -1
        'Combo5.ListIndex = -1
        Combo5.Text = ""
    End If
End Sub

Private Sub cmbFeld5_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    Dim SQL As String
    Dim Msg As String
    
    If cmbFeld5.Text = "" Then Exit Sub                      'Gerbing 10.06.2005
'    If Combo5.Text <> "" Then Exit Sub                  'Gerbing 29.12.2005
    Combo5.ComboItems.RemoveAll                                        'Gerbing 10.06.2005
    cmbVG5.Clear                                     'Gerbing 04.02.2007
    'Alle Werte nach Combo5 stellen, die zu diesem Feld in der Datenbank gefunden werden
    'aber nicht bei ...Fields(...).Type = dbBoolean, da gibts nur true oder False       'Gerbing 04.02.2007
    'abhängig von ...Fields(...).Type müssen die möglichen Vergleichsoperanden eingestellt werden
    'bei dbText gibt es nur = und <>
    'bei dbBoolean gibt es nur = und <>
    'bei Hyperlink gibt es nur = und <>
    Select Case DatenTyp(cmbFeld5.Text)
        Case 1, 11 '1=dbBoolean '11=Boolean bei SQL-Server                              'Gerbing 23.11.2017
            ND.cmbVG5.AddItem "="
            ND.cmbVG5.AddItem "<>"
        Case 10, 202    '10=Text 202 Text bei SQL-Server                                'Gerbing 23.11.2017
'            ND.cmbVG5.AddItem "="
'            ND.cmbVG5.AddItem "<>"
            ND.cmbVG5.AddItem "="
            ND.cmbVG5.AddItem ">"
            ND.cmbVG5.AddItem "<"
            ND.cmbVG5.AddItem ">="
            ND.cmbVG5.AddItem "<="
            ND.cmbVG5.AddItem "<>"
            ND.cmbVG5.AddItem "like"                                                    'Gerbing 08.05.2019
        Case 12   'Hyperlink
            ND.cmbVG5.AddItem "="
            ND.cmbVG5.AddItem "<>"
        Case Else 'Zahlen
            ND.cmbVG5.AddItem "="
            ND.cmbVG5.AddItem ">"
            ND.cmbVG5.AddItem "<"
            ND.cmbVG5.AddItem ">="
            ND.cmbVG5.AddItem "<="
            ND.cmbVG5.AddItem "<>"
    End Select
    SQL = "SELECT DISTINCT Fotos.[" & cmbFeld5.Text & "]"
    SQL = SQL & " From Fotos"
    SQL = SQL & " Where ((Not (Fotos.[" & cmbFeld5.Text & "]) Is Null))"
    SQL = SQL & " ORDER BY Fotos.[" & cmbFeld5.Text & "];"
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
    'wenn als Suchergebnis rstsql.EOF kommt, muss eine Warnung kommen, dass in diesem Feld keinerlei Werte
    'gespeichert sind                                                                   'Gerbing 04.02.2007
    If rstsql.EOF Then
'        msg = "Im Feld '" & cmbFeld5 & "' sind keine Werte gespeichert." & vbNewLine
'        msg = msg & "Sie können nach diesen Werten nicht suchen. Wählen Sie ein anderes Feld"
        Msg = LoadResString(2321 + Sprache) & cmbFeld5.Text & LoadResString(2322 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(2323 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    '-------------------------------------
    If DatenTyp(cmbFeld5) = 1 Or DatenTyp(cmbFeld5.Text) = 11 Then          'Gerbing 23.11.2017
        If gblnSQLServerVersion = True Then
            Combo5.ComboItems.Add 0
            Combo5.ComboItems.Add 1
        Else
            'Combo5.comboitems.add "Wahr"
            'Combo5.comboitems.add "Falsch"
            Combo5.ComboItems.Add LoadResString(2324 + Sprache)
            Combo5.ComboItems.Add LoadResString(2325 + Sprache)
        End If
    Else
        Do Until rstsql.EOF
            If Not IsNull(rstsql.Fields(0)) Then
                Combo5.ComboItems.Add rstsql.Fields(0)
            End If
            rstsql.MoveNext
            DoEvents
        Loop
    End If
    If Combo5.ComboItems.Count <> 0 Then
        'Combo5.ListIndex = 0
    End If
End Sub


Private Sub Form_Load()
    Dim rstErsterStart As ADODB.Recordset                                                   'Gerbing 05.09.2016
    Dim rstTemp As ADODB.Recordset                                                          'Gerbing 05.09.2016
    Dim SQL As String
    
    Call AnpassenNutzerWunsch(Me)                                                           'Gerbing 11.03.2017
    Me.Top = 0                                                                              'Gerbing 06.12.2006
    Me.Left = 0
    'If PublicLanguage <> 9 Then                                                            'Gerbing 20.11.2007
'    If PublicLanguage <> "9" Then                                                          'Gerbing 28.03.2010
'        If Query.chkFensterGrößeÄnderbar.Value = 1 Then                                    'Gerbing 06.12.2005
'            Me.Top = Form1.Top                                                             'Gerbing 06.12.2006
'            Me.Left = Form1.Left
'            Me.Width = Form1.Width
'        End If
'    End If
    Me.Caption = LoadResString(1089 + Sprache)     'Suche nach nutzerdefinierten Feldern  Gerbing 08.11.2005
    lblFeldname1.Caption = LoadResString(1033 + Sprache)  'Feldname
    lblFeldname2.Caption = LoadResString(1033 + Sprache)  'Feldname
    lblFeldname3.Caption = LoadResString(1033 + Sprache)  'Feldname
    lblFeldname4.Caption = LoadResString(1033 + Sprache)  'Feldname
    lblFeldname5.Caption = LoadResString(1033 + Sprache)  'Feldname
'    cmbFeld1.ToolTipText = LoadResString(2507 + Sprache) 'Irrtümlich gewählte Feldnamen können Sie durch Markieren dann Drücken der Return-Taste oder Entf-Taste auf dem numerischen Tastenfeld entfernen
'    cmbFeld2.ToolTipText = LoadResString(2507 + Sprache) 'Irrtümlich gewählte Feldnamen können Sie durch Markieren dann Drücken der Return-Taste oder Entf-Taste auf dem numerischen Tastenfeld entfernen
'    cmbFeld3.ToolTipText = LoadResString(2507 + Sprache) 'Irrtümlich gewählte Feldnamen können Sie durch Markieren dann Drücken der Return-Taste oder Entf-Taste auf dem numerischen Tastenfeld entfernen
'    cmbFeld4.ToolTipText = LoadResString(2507 + Sprache) 'Irrtümlich gewählte Feldnamen können Sie durch Markieren dann Drücken der Return-Taste oder Entf-Taste auf dem numerischen Tastenfeld entfernen
'    cmbFeld5.ToolTipText = LoadResString(2507 + Sprache) 'Irrtümlich gewählte Feldnamen können Sie durch Markieren dann Drücken der Return-Taste oder Entf-Taste auf dem numerischen Tastenfeld entfernen
    Frame2.Caption = LoadResString(3032 + Sprache)     'Wählen Sie Feldnamen und Vergleichsoperand und geben Sie Feldwerte ein
    Und1.Caption = LoadResString(3023 + Sprache) 'Und
    Und2.Caption = LoadResString(3023 + Sprache) 'Und
    Und3.Caption = LoadResString(3023 + Sprache) 'Und
    Und4.Caption = LoadResString(3023 + Sprache) 'Und
    SOder.Caption = LoadResString(3024 + Sprache)    'Oder
    OOder.Caption = LoadResString(3024 + Sprache)    'Oder
    LOder.Caption = LoadResString(3024 + Sprache)    'Oder
    SWFOder.Caption = LoadResString(3024 + Sprache)    'Oder
    btnOK.Caption = LoadResString(3001 + Sprache) '&OK
    btnAbbrechen.Caption = LoadResString(3033 + Sprache) '&Keine Suche nach nutzerdefinierten Feldern
    btnSucheGEODaten.Caption = LoadResString(3164 + Sprache)    'Suche Fotos mit GEO-Daten  'Gerbing 03.10.2016
    gstrGEOStartPunkt = ""                                                                  'Gerbing 05.09.2016
    gstrGEOEndPunkt = ""
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 14.2.2 gibt es in der Tabelle ErsterStart das Feld LetzterGEOPunkt und ZoomListIndex        'Gerbing 05.09.2016
    'die erzeugt das Programm selbst, wenn es die Professional Version ist
    If gblnProversion = True Then
        If gblnSchreibgeschützt = False Then
            On Error Resume Next
            SQL = "select LetzterGEOPunkt From ErsterStart;"
            Set rstErsterStart = New ADODB.Recordset
            With rstErsterStart
                .Source = SQL
                .ActiveConnection = DBado                                               'Gerbing 23.11.2017
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If Err.Number <> 0 Then
                'hier existiert das Feld LetzterGEOPunkt nicht
                If gblnSchreibgeschützt = False Then
                    If gblnSQLServerVersion = True Then
                        'SQL Server                                                     'Gerbing 23.11.2017
                        DBado.Execute _
                            "ALTER TABLE ErsterStart ADD LetzterGEOPunkt VARCHAR(255)"          'es heißt ADD und nicht ADD COLUMN
                        DBado.Execute _
                            "ALTER TABLE ErsterStart ADD ZoomListIndex VARCHAR(255)"
                    Else
                        'Access Version
                        'also wird Feld LetzterGEOPunkt und ZoomListIndex erzeugt
                        DBado.Execute _
                            "ALTER TABLE ErsterStart ADD COLUMN LetzterGEOPunkt TEXT"
                        DBado.Execute _
                            "ALTER TABLE ErsterStart ADD COLUMN ZoomListIndex TEXT"
                    End If
                End If
                rstErsterStart.Close
                On Error GoTo 0
            End If
        End If
   End If
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    cmbFeld1.SetFocus
    On Error GoTo 0
End Sub

Private Sub btnOK_Click()
    Dim Msg As String
    Dim pos As Long
    Dim pos1 As Long
    Dim rc As Long

    AnzahlFelder = 0

    'Query.CheckNutzerdefinierteFelder soll als Kennzeichen dienen
    'ob Suche nach nutzerdefinierten Feldern aktiv ist
    If gstrGEOStartPunkt <> "" Or cmbFeld1.Text <> "" Or cmbFeld2.Text <> "" Or cmbFeld3.Text <> "" Or cmbFeld4.Text <> "" Or cmbFeld5.Text <> "" Then  'Gerbing 05.09.2016
        Query.CheckNutzerdefinierteFelder.Value = 1
        If gblnProversion = True Then                               'Gerbing 10.06.2005
            Query.CheckNutzerdefinierteFelder.Visible = True
            Query.lblNutzerdefinierteFelder.Visible = True          'Gerbing 25.06.2013
            If gstrGEOStartPunkt <> "" Then                         'Gerbing 05.09.2016
                Me.Hide
                Exit Sub
            End If
        End If
    Else
        Query.CheckNutzerdefinierteFelder.Value = 0
        Query.CheckNutzerdefinierteFelder.Visible = False
        Query.lblNutzerdefinierteFelder.Visible = False             'Gerbing 25.06.2013
        Me.Hide
        Exit Sub                                                    'Gerbing 07.03.2005
    End If
    '-------------------------------------------------------------------
    'es dürfen kein Lücken entstehen, Mindestens 1 Feld muss belegt sein
    'AnzahlFelder muss gezählt werden
    AnzahlFelder = 1              'Anfangswert
    Do
        If cmbFeld2.Text = "" And cmbFeld3.Text = "" And cmbFeld4.Text = "" And cmbFeld5.Text = "" Then Exit Do
        AnzahlFelder = 2
        If cmbFeld2.Text <> "" Then
            If cmbFeld1.Text = "" Then
                'MsgBox "Sie müssen die Feldnamen lückenlos von oben nach unten ausfüllen"
                MsgBox LoadResString(2104 + Sprache)                  'Gerbing 08.11.2005
                Exit Sub
            End If
        End If
        If cmbFeld3.Text = "" And cmbFeld4.Text = "" And cmbFeld5.Text = "" Then Exit Do
        AnzahlFelder = 3
        If cmbFeld3.Text <> "" Then
            If cmbFeld1.Text = "" Or cmbFeld2.Text = "" Then
                'MsgBox "Sie müssen die Feldnamen lückenlos von oben nach unten ausfüllen"
                MsgBox LoadResString(2104 + Sprache)                  'Gerbing 08.11.2005
                Exit Sub
            End If
            If cmbFeld4.Text = "" And cmbFeld5.Text = "" Then Exit Do
        End If
        AnzahlFelder = 4
        If cmbFeld4.Text <> "" Then
            If cmbFeld1.Text = "" Or cmbFeld2.Text = "" Or cmbFeld3.Text = "" Then
                'MsgBox "Sie müssen die Feldnamen lückenlos von oben nach unten ausfüllen"
                MsgBox LoadResString(2104 + Sprache)                  'Gerbing 08.11.2005
                Exit Sub
            End If
            If cmbFeld5.Text = "" Then Exit Do
        End If
        AnzahlFelder = 5
        If cmbFeld5.Text <> "" Then
            If cmbFeld1.Text = "" Or cmbFeld2.Text = "" Or cmbFeld3.Text = "" Or cmbFeld4.Text = "" Then
                'MsgBox "Sie müssen die Feldnamen lückenlos von oben nach unten ausfüllen"
                MsgBox LoadResString(2104 + Sprache)                  'Gerbing 08.11.2005
                Exit Sub
            End If
        End If
        Exit Do
    Loop
    '-------------------------------------------------------------------------------------------------
    If cmbVG1.Text <> "" And cmbFeld1.Text = "" Then
        'MsgBox "Wenn Sie einen Vergleichsoperand auswählen, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2105 + Sprache)                  'Gerbing 08.11.2005
        cmbFeld1.SetFocus
        Exit Sub
    End If
    If cmbVG2.Text <> "" And cmbFeld2.Text = "" Then
        'MsgBox "Wenn Sie einen Vergleichsoperand auswählen, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2105 + Sprache)                  'Gerbing 08.11.2005
        cmbFeld2.SetFocus
        Exit Sub
    End If
    If cmbVG3.Text <> "" And cmbFeld3.Text = "" Then
        'MsgBox "Wenn Sie einen Vergleichsoperand auswählen, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2105 + Sprache)                  'Gerbing 08.11.2005
        cmbFeld3.SetFocus
        Exit Sub
    End If
    If cmbVG4.Text <> "" And cmbFeld4.Text = "" Then
        'MsgBox "Wenn Sie einen Vergleichsoperand auswählen, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2105 + Sprache)                  'Gerbing 08.11.2005
        cmbFeld4.SetFocus
        Exit Sub
    End If
    If cmbVG5.Text <> "" And cmbFeld5.Text = "" Then
        'MsgBox "Wenn Sie einen Vergleichsoperand auswählen, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2105 + Sprache)                  'Gerbing 08.11.2005
        cmbFeld5.SetFocus
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------
    If cmbFeld1.Text <> "" And cmbVG1.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Vergleichsoperand auswählen"
        MsgBox LoadResString(2106 + Sprache)                  'Gerbing 08.11.2005
        cmbVG1.SetFocus
        Exit Sub
    End If
    If cmbFeld2.Text <> "" And cmbVG2.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Vergleichsoperand auswählen"
        MsgBox LoadResString(2106 + Sprache)                  'Gerbing 08.11.2005
        cmbVG2.SetFocus
        Exit Sub
    End If
    If cmbFeld3.Text <> "" And cmbVG3.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Vergleichsoperand auswählen"
        MsgBox LoadResString(2106 + Sprache)                  'Gerbing 08.11.2005
        cmbVG3.SetFocus
        Exit Sub
    End If
    If cmbFeld4.Text <> "" And cmbVG4.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Vergleichsoperand auswählen"
        MsgBox LoadResString(2106 + Sprache)                  'Gerbing 08.11.2005
        cmbVG4.SetFocus
        Exit Sub
    End If
    If cmbFeld5.Text <> "" And cmbVG5.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Vergleichsoperand auswählen"
        MsgBox LoadResString(2106 + Sprache)                  'Gerbing 08.11.2005
        cmbVG5.SetFocus
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------
    If cmbFeld1.Text <> "" And Combo1.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Feldwert eingeben"
        MsgBox LoadResString(2107 + Sprache)
        Combo1.SetFocus
        Exit Sub
    End If
    If cmbFeld2.Text <> "" And Combo2.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Feldwert eingeben"
        MsgBox LoadResString(2107 + Sprache)
        Combo2.SetFocus
        Exit Sub
    End If
    If cmbFeld3.Text <> "" And Combo3.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Feldwert eingeben"
        MsgBox LoadResString(2107 + Sprache)
        Combo3.SetFocus
        Exit Sub
    End If
    If cmbFeld4.Text <> "" And combo4.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Feldwert eingeben"
        MsgBox LoadResString(2107 + Sprache)
        combo4.SetFocus
        Exit Sub
    End If
    If cmbFeld5.Text <> "" And Combo5.Text = "" Then
        'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch einen Feldwert eingeben"
        MsgBox LoadResString(2107 + Sprache)
        Combo5.SetFocus
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------
    If Combo1.Text <> "" And cmbFeld1.Text = "" Then
        'MsgBox "Wenn Sie einen Feldwert eingeben, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2108 + Sprache)
        cmbFeld1.SetFocus
        Exit Sub
    End If
    If Combo2.Text <> "" And cmbFeld2.Text = "" Then
        'MsgBox "Wenn Sie einen Feldwert eingeben, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2108 + Sprache)
        cmbFeld2.SetFocus
        Exit Sub
    End If
    If Combo3.Text <> "" And cmbFeld3.Text = "" Then
        'MsgBox "Wenn Sie einen Feldwert eingeben, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2108 + Sprache)
        cmbFeld3.SetFocus
        Exit Sub
    End If
    If combo4.Text <> "" And cmbFeld4.Text = "" Then
        'MsgBox "Wenn Sie einen Feldwert eingeben, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2108 + Sprache)
        cmbFeld4.SetFocus
        Exit Sub
    End If
    If Combo5.Text <> "" And cmbFeld5.Text = "" Then
        'MsgBox "Wenn Sie einen Feldwert eingeben, müssen Sie auch einen Feldnamen auswählen"
        MsgBox LoadResString(2108 + Sprache)
        cmbFeld5.SetFocus
        Exit Sub
    End If
    '-------------------------------------------------------------------------
    'Wenn ein Datum mit einer Uhrzeit gemischt wird, muss das abgelehnt werden
    If Sprache = 0 Then                                 'Gerbing 08.11.2005
        pos = InStr(1, Combo1.Text, ".", vbTextCompare)
        pos1 = InStr(1, Combo1.Text, ":", vbTextCompare)
    Else
        pos = InStr(1, Combo1.Text, ":", vbTextCompare)
        pos1 = InStr(1, Combo1.Text, "/", vbTextCompare)
    End If
    If pos <> 0 And pos1 <> 0 Then
        rc = Ablehnen(cmbFeld1.Text)                        'Gerbing 04.02.2007
        If rc = 1 Then Exit Sub
    End If
    If Sprache = 0 Then
        pos = InStr(1, Combo2.Text, ".", vbTextCompare)
        pos1 = InStr(1, Combo2.Text, ":", vbTextCompare)
    Else
        pos = InStr(1, Combo2.Text, ":", vbTextCompare)
        pos1 = InStr(1, Combo2.Text, "/", vbTextCompare)
    End If
    If pos <> 0 And pos1 <> 0 Then
        rc = Ablehnen(cmbFeld2.Text)                         'Gerbing 04.02.2007
        If rc = 1 Then Exit Sub
    End If
    If Sprache = 0 Then
        pos = InStr(1, Combo3.Text, ".", vbTextCompare)
        pos1 = InStr(1, Combo3.Text, ":", vbTextCompare)
    Else
        pos = InStr(1, Combo3.Text, ":", vbTextCompare)
        pos1 = InStr(1, Combo3.Text, "/", vbTextCompare)
    End If
    If pos <> 0 And pos1 <> 0 Then
        rc = Ablehnen(cmbFeld3.Text)                         'Gerbing 04.02.2007
        If rc = 1 Then Exit Sub
    End If
    If Sprache = 0 Then
        pos = InStr(1, combo4.Text, ".", vbTextCompare)
        pos1 = InStr(1, combo4.Text, ":", vbTextCompare)
    Else
        pos = InStr(1, combo4.Text, ":", vbTextCompare)
        pos1 = InStr(1, combo4.Text, "/", vbTextCompare)
    End If
    If pos <> 0 And pos1 <> 0 Then
        rc = Ablehnen(cmbFeld4.Text)                         'Gerbing 04.02.2007
        If rc = 1 Then Exit Sub
    End If
    If Sprache = 0 Then
        pos = InStr(1, Combo5.Text, ".", vbTextCompare)
        pos1 = InStr(1, Combo5.Text, ":", vbTextCompare)
    Else
        pos = InStr(1, Combo5.Text, ":", vbTextCompare)
        pos1 = InStr(1, Combo5.Text, "/", vbTextCompare)
    End If
    If pos <> 0 And pos1 <> 0 Then
        rc = Ablehnen(cmbFeld5.Text)                         'Gerbing 04.02.2007
        If rc = 1 Then Exit Sub
    End If
    '--------------------------------------------------------
    If cmbFeld1.Text <> "" Then Form1.F5Feld1 = cmbFeld1.Text
    If cmbFeld2.Text <> "" Then Form1.F5Feld2 = cmbFeld2.Text
    If cmbFeld3.Text <> "" Then Form1.F5Feld3 = cmbFeld3.Text
    If cmbFeld4.Text <> "" Then Form1.F5Feld4 = cmbFeld4.Text
    If cmbFeld5.Text <> "" Then Form1.F5Feld5 = cmbFeld5.Text
    
    Plus1 = " And "
    Plus2 = " And "
    Plus3 = " And "
    Plus4 = " And "
    If Und1.Value = False Then Plus1 = " Or "
    If Und2.Value = False Then Plus2 = " Or "
    If Und3.Value = False Then Plus3 = " Or "
    If Und4.Value = False Then Plus4 = " Or "

    Me.Hide
End Sub

Private Function Ablehnen(Feldname)
    Dim Msg As String
    'rc = 0 bei Hyperlink sind Punkte erlaubt
    'rc = 1 sonst sind Punkte nicht erlaubt
    
    Ablehnen = 0
    If DatenTyp(Feldname) = 12 Then '12=Hyperlink
        Exit Function
    End If
    'msg = "Das Programm hat die gleichzeitige Verwendung der Zeichen . und : im Feldwert entdeckt." & vbNewLine
    Msg = LoadResString(2109 + Sprache) & vbNewLine
    'msg = msg & "Darum wird vermutet, das Sie in einem Feld nach Datum und Uhrzeit gleichzeitig suchen." & vbNewLine
    Msg = Msg & LoadResString(2110 + Sprache) & vbNewLine
    'msg = msg & "Solche Felddefinitionen werden abgelehnt. Definieren Sie bitte getrennte Felder" & vbNewLine
    Msg = Msg & LoadResString(2111 + Sprache) & vbNewLine
    'msg = msg & "für Datum und Uhrzeit."
    Msg = Msg & LoadResString(2112 + Sprache)
    MsgBox Msg
    Ablehnen = 1
End Function

Private Function DatenTyp(Feldname) As Integer
    Dim n As Long
    Dim Msg As String
    
    'Gibt den Datentyp des Datenbankfeldes zurück                   'Gerbing 04.02.2007
    'die Datentypen stehen in ND.ListDatenTyp
    'Die Listbox ND.ListNutzerdefinierteFelder(unsortiert) speichert die nutzerdefinierten Felder in der
    'Reihenfolge der Columns von links nach rechts
    'Die Listbox ND.ListDatenTyp(unsortiert) speichert den Datentyp der nutzerdefinierten Felder in derselben
    'Reihenfolge wie die Listbox ND.ListNutzerdefinierteFelder(unsortiert)
    
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        If ND.ListNutzerdefinierteFelder.ListItems(n) = Feldname Then Exit For
    Next n
    DatenTyp = ND.ListDatenTyp.List(n)                          'Gerbing 04.02.2007
    Exit Function
End Function

#End If


Private Sub ListDatenTyp_Click()

End Sub
