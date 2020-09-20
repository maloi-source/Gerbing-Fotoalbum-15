VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form NeueDatensätzeGenerieren 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Neue Datensätze generieren"
   ClientHeight    =   10608
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   14928
   Icon            =   "NeueDatensätzeGenerieren.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10608
   ScaleWidth      =   14928
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   8400
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton optExif 
         Caption         =   "optExif"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optManuell 
         Caption         =   "optManuell"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optExtrahieren 
         Caption         =   "optExtrahieren"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   1935
      End
   End
   Begin EditCtlsLibUCtl.TextBox txtArbeitsfortschritt 
      Height          =   492
      Left            =   2280
      TabIndex        =   55
      Top             =   10080
      Width           =   9612
      _cx             =   16954
      _cy             =   868
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
      CueBanner       =   "NeueDatensätzeGenerieren.frx":038A
      Text            =   "NeueDatensätzeGenerieren.frx":03AA
   End
   Begin CBLCtlsLibUCtl.ListBox List1 
      Height          =   1092
      Left            =   120
      TabIndex        =   53
      Top             =   8880
      Width           =   11772
      _cx             =   20764
      _cy             =   1926
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048808
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
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
   Begin VB.Frame Frame3 
      Height          =   1092
      Left            =   120
      TabIndex        =   46
      Top             =   1200
      Width           =   11772
      Begin VB.OptionButton optProtokolldatei 
         Caption         =   "Eintragen in die Protokolldatei (pruef.log)"
         Height          =   372
         Left            =   5280
         TabIndex        =   48
         Top             =   240
         Width           =   6372
      End
      Begin VB.OptionButton optMsgbox 
         Caption         =   "Ich will bei jedem Fehler einen Fehlerhinweis erhalten"
         Height          =   372
         Left            =   5280
         TabIndex        =   47
         Top             =   600
         Width           =   6372
      End
      Begin VB.Label lblDatentypfehler 
         Caption         =   "Was soll bei Datentyp-Fehlern geschehen?"
         Height          =   492
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   4932
      End
   End
   Begin VB.Frame FrameExifIptc 
      Height          =   10572
      Left            =   12000
      TabIndex        =   41
      Top             =   0
      Width           =   2895
      Begin CBLCtlsLibUCtl.ListBox LstU 
         Height          =   9492
         Left            =   0
         TabIndex        =   61
         Top             =   1080
         Width           =   2652
         _cx             =   4678
         _cy             =   16743
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
      Begin VB.CheckBox chkIptcAnzeigen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IPTC-Felder"
         Height          =   372
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Setzen Sie hier ein Häkchen, wenn Sie die IPTC-Felder des markierten Satzes im Drag&Drop-Container sehen wollen"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkExifAnzeigen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EXIF-Felder"
         Height          =   372
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Setzen Sie hier ein Häkchen, wenn Sie die EXIF-Felder des markierten Satzes im Drag&Drop-Container sehen wollen"
         Top             =   240
         Width           =   2535
      End
      Begin EditCtlsLibUCtl.TextBox txtEXIFInfo 
         Height          =   9372
         Left            =   0
         TabIndex        =   71
         Top             =   1080
         Width           =   2652
         _cx             =   4678
         _cy             =   16531
         AcceptNumbersOnly=   0   'False
         AcceptTabKey    =   0   'False
         AllowDragDrop   =   -1  'True
         AlwaysShowSelection=   0   'False
         Appearance      =   1
         AutoScrolling   =   0
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":03CA
         Text            =   "NeueDatensätzeGenerieren.frx":03EA
      End
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   492
      Left            =   4680
      TabIndex        =   29
      Top             =   480
      Width           =   1692
   End
   Begin VB.CheckBox chkExif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIF/IPTC benutzen"
      Height          =   372
      Left            =   4680
      TabIndex        =   28
      ToolTipText     =   "Setzen Sie hier ein Häkchen, wenn Sie EXIF-Felder importieren wollen"
      Top             =   0
      Width           =   3372
   End
   Begin VB.CheckBox chkUnbeaufsichtigt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "unbeaufsichtigt"
      Height          =   372
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Setzen Sie hier ein Häkchen wenn die Datenbank automatisch erzeugt werden soll"
      Top             =   0
      Width           =   3012
   End
   Begin VB.Frame FrameNutzerDefiniert 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nutzerdefinierte Felder"
      Height          =   2292
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   11772
      Begin VB.ComboBox cmbEx5 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox cmbEx4 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cmbEx3 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmbEx2 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbEx1 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   2535
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo1 
         Height          =   288
         Left            =   4200
         TabIndex        =   56
         Top             =   480
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":040A
         Text            =   "NeueDatensätzeGenerieren.frx":042A
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo2 
         Height          =   288
         Left            =   4200
         TabIndex        =   57
         Top             =   840
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":044A
         Text            =   "NeueDatensätzeGenerieren.frx":046A
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo3 
         Height          =   288
         Left            =   4200
         TabIndex        =   58
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":048A
         Text            =   "NeueDatensätzeGenerieren.frx":04AA
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo4 
         Height          =   288
         Left            =   4200
         TabIndex        =   59
         Top             =   1560
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":04CA
         Text            =   "NeueDatensätzeGenerieren.frx":04EA
      End
      Begin CBLCtlsLibUCtl.ComboBox Combo5 
         Height          =   288
         Left            =   4200
         TabIndex        =   60
         Top             =   1920
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":050A
         Text            =   "NeueDatensätzeGenerieren.frx":052A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld1 
         Height          =   288
         Left            =   2040
         TabIndex        =   66
         Top             =   480
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":054A
         Text            =   "NeueDatensätzeGenerieren.frx":056A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld2 
         Height          =   288
         Left            =   2040
         TabIndex        =   67
         Top             =   840
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":058A
         Text            =   "NeueDatensätzeGenerieren.frx":05AA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld3 
         Height          =   288
         Left            =   2040
         TabIndex        =   68
         Top             =   1200
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":05CA
         Text            =   "NeueDatensätzeGenerieren.frx":05EA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld4 
         Height          =   288
         Left            =   2040
         TabIndex        =   69
         Top             =   1560
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":060A
         Text            =   "NeueDatensätzeGenerieren.frx":062A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld5 
         Height          =   288
         Left            =   2040
         TabIndex        =   70
         Top             =   1920
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":064A
         Text            =   "NeueDatensätzeGenerieren.frx":066A
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   11
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":068A
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   10
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":0ACC
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   9
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":0F0E
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   8
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":1350
         Stretch         =   -1  'True
         Top             =   960
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   132
         Index           =   7
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":1792
         Stretch         =   -1  'True
         Top             =   600
         Width           =   132
      End
      Begin VB.Label lblExifIptc2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "EXIF/IPTC-Feld"
         Height          =   252
         Left            =   9120
         TabIndex        =   33
         Top             =   240
         Width           =   2532
      End
      Begin VB.Label lblFeldinhalt 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldinhalt"
         Height          =   252
         Left            =   4200
         TabIndex        =   32
         Top             =   240
         Width           =   2532
      End
      Begin VB.Label lblFeldname5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1812
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4040
         TabIndex        =   19
         Top             =   1920
         Width           =   252
      End
      Begin VB.Label lblFeldname4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1812
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4040
         TabIndex        =   17
         Top             =   1560
         Width           =   252
      End
      Begin VB.Label lblFeldname3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4040
         TabIndex        =   15
         Top             =   1200
         Width           =   252
      End
      Begin VB.Label lblFeldname2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1812
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4040
         TabIndex        =   13
         Top             =   840
         Width           =   252
      End
      Begin VB.Label lblFeldname1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1812
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "="
         Height          =   252
         Left            =   4040
         TabIndex        =   11
         Top             =   480
         Width           =   252
      End
   End
   Begin VB.Frame FrameStandardWerte 
      BackColor       =   &H0080C0FF&
      Caption         =   "Standard-Felder"
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   11772
      Begin EditCtlsLibUCtl.TextBox txtKommentar 
         Height          =   300
         Left            =   2040
         TabIndex        =   54
         Top             =   3600
         Width           =   6612
         _cx             =   11663
         _cy             =   529
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
         CueBanner       =   "NeueDatensätzeGenerieren.frx":1BD4
         Text            =   "NeueDatensätzeGenerieren.frx":1BF4
      End
      Begin VB.ComboBox cmbSWFEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   2880
         Width           =   2535
      End
      Begin VB.ComboBox cmbKommentarEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   3600
         Width           =   2535
      End
      Begin VB.ComboBox cmbJahrEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   44
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cmbPersonenEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   2160
         Width           =   2535
      End
      Begin VB.ComboBox cmbLandEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox cmbOrtEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbSituationEx 
         Height          =   288
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton btnJahr 
         Caption         =   "&Jahr..."
         Height          =   372
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1452
      End
      Begin VB.Frame Frame2 
         Caption         =   "SWF"
         Height          =   1095
         Left            =   2040
         TabIndex        =   4
         Top             =   2520
         Width           =   6612
         Begin VB.OptionButton OptSV 
            BackColor       =   &H00E0E0E0&
            Caption         =   "SV (Schwarz/weiß-Video)"
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   600
            Width           =   3492
         End
         Begin VB.OptionButton OptFV 
            BackColor       =   &H00E0E0E0&
            Caption         =   "FV (Farb-Video)"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   2652
         End
         Begin VB.OptionButton OptF 
            BackColor       =   &H00E0E0E0&
            Caption         =   "F (Farb-Foto)"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   2652
         End
         Begin VB.OptionButton OptSW 
            BackColor       =   &H00E0E0E0&
            Caption         =   "SW (Schwarz/weiß-Foto)"
            Height          =   255
            Left            =   3000
            TabIndex        =   23
            Top             =   240
            Width           =   3492
         End
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbOrt 
         Height          =   288
         Left            =   2040
         TabIndex        =   63
         Top             =   1440
         Width           =   6612
         _cx             =   11663
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":1C14
         Text            =   "NeueDatensätzeGenerieren.frx":1C34
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbSituation 
         Height          =   288
         Left            =   2040
         TabIndex        =   62
         Top             =   1080
         Width           =   6612
         _cx             =   11663
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":1C54
         Text            =   "NeueDatensätzeGenerieren.frx":1C74
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbLand 
         Height          =   288
         Left            =   2040
         TabIndex        =   64
         Top             =   1800
         Width           =   6612
         _cx             =   11663
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":1C94
         Text            =   "NeueDatensätzeGenerieren.frx":1CB4
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbPersonen 
         Height          =   288
         Left            =   2040
         TabIndex        =   65
         Top             =   2160
         Width           =   6612
         _cx             =   11663
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
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "NeueDatensätzeGenerieren.frx":1CD4
         Text            =   "NeueDatensätzeGenerieren.frx":1CF4
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   6
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":1D14
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   5
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":2156
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   4
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":2598
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   3
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":29DA
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   2
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":2E1C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   1
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":325E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   132
      End
      Begin VB.Image Image1 
         Height          =   144
         Index           =   0
         Left            =   8880
         Picture         =   "NeueDatensätzeGenerieren.frx":36A0
         Stretch         =   -1  'True
         Top             =   840
         Width           =   132
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "standardmäßig wird das Jahr aus dem Dateiname extrahiert"
         Height          =   372
         Left            =   2040
         TabIndex        =   51
         Top             =   720
         Width           =   6612
      End
      Begin VB.Label lblExifIptc1 
         BackColor       =   &H0080C0FF&
         Caption         =   "EXIF/IPTC-Feld"
         Height          =   252
         Left            =   9120
         TabIndex        =   50
         Top             =   240
         Width           =   2532
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kommentar:"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   1692
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Personen:"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1692
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Land:"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1572
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ort:"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Situation:"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1812
      End
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   492
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1692
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "S&tart"
      Default         =   -1  'True
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1692
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   10200
      Width           =   2052
   End
   Begin VB.Label lblDateinamen 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Drag&&Drop-Container für Dateinamen vom Windows-Explorer"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   11772
   End
End
Attribute VB_Name = "NeueDatensätzeGenerieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'txtEXIFInfo ist jetzt ein Timosoft Control unicode fähig                           'Gerbing 13.11.2015



'* DEFINEs for field types - provided for reference only.
'   #DEFINE ADEMPTY               0
'   #DEFINE ADTINYINT            16
'   #DEFINE ADSMALLINT            2
'   #DEFINE ADINTEGER            3
'   #DEFINE ADBIGINT            20
'   #DEFINE ADUNSIGNEDTINYINT      17
'   #DEFINE ADUNSIGNEDSMALLINT      18
'   #DEFINE ADUNSIGNEDINT         19
'   #DEFINE ADUNSIGNEDBIGINT      21
'   #DEFINE ADSINGLE            4
'   #DEFINE ADDOUBLE            5
'   #DEFINE ADCURRENCY            6
'   #DEFINE ADDECIMAL            14
'   #DEFINE ADNUMERIC            131
'   #DEFINE ADBOOLEAN            11
'   #DEFINE ADERROR               10
'   #DEFINE ADUSERDEFINED         132
'   #DEFINE ADVARIANT            12
'   #DEFINE ADIDISPATCH            9
'   #DEFINE ADIUNKNOWN            13
'   #DEFINE ADGUID               72
'   #DEFINE ADDATE               7
'   #DEFINE ADDBDATE            133
'   #DEFINE ADDBTIME            134
'   #DEFINE ADDBTIMESTAMP         135
'   #DEFINE ADBSTR               8
'   #DEFINE ADCHAR               129
'   #DEFINE ADVARCHAR            200
'   #DEFINE ADLONGVARCHAR         201
'   #DEFINE ADWCHAR               130
'   #DEFINE ADVARWCHAR            202
'   #DEFINE ADLONGVARWCHAR         203
'   #DEFINE ADBINARY            128
'   #DEFINE ADVARBINARY            204
'   #DEFINE ADLONGVARBINARY         205
'   #DEFINE ADCHAPTER            136


Option Compare Text
    Dim DErwMsgBoxMussKommen As Boolean
    Dim NL As String
    Dim GefundeneJahresZahl As String
    Public GiltFürAlleFälle As Boolean
    Public AktuellerDateiname As String
    Public NutzerJahresZahl As String
    Public NichtNochmalWarnen As Boolean
    Public JahrFestLegenAbgebrochen As Boolean
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
    'Const aus win32api.txt
    Const LB_SETHORIZONTALEXTENT = &H194
    'Private iptc As New IPTCInfo.Reader
    Public blnOptGewählt As Boolean
    Public DummyJahr As String
    Dim strGefundenExifIptc As String
    Dim blnIPTCVorhanden As Boolean
    Dim MyAppPath As String
    Dim blnUserDefinedError As Boolean                                              'Gerbing 02.08.2016
    Dim blnDefaultFieldsNotEmpty As Boolean                                         'Gerbing 03.08.2016
    Dim blnUserdefinedFieldsNotEmpty As Boolean                                     'Gerbing 03.08.2016
    Dim blnErsterDurchlauf As Boolean
    Dim EXIFListInfo As String                                                      'Gerbing 12.10.2016
    Dim blnMinusGeprüft As Boolean                                                  'Gerbing 12.10.2016
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Dim rstsql As ADODB.Recordset                                                   'Gerbing 18.11.2019
    Dim blnMediaInfoInitialized As Boolean                                          'Gerbing 18.11.2019
    Dim Handle As Long                                                              'Gerbing 18.11.2019
    Private lngVideoDuration As Long                                                'Gerbing 15.11.2019
    Private glngStartMillisek As Long                                               'Gerbing 15.11.2019
    Private glngEndMillisek As Long                                                 'Gerbing 15.11.2019



Private Sub btnAbbrechen_Click()
    'dbs.Close                                                                      'Gerbing 29.05.2008
    'Close #DateiNummer
    oStream.Close
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

Private Sub btnJahr_Click()
    If List1.ListItems.Count = 0 Then
        'MsgBox "Sie haben noch keine Dateinamen in den Drag&Drop-Container gezogen"
        MsgBox LoadResString(1442 + Sprache)
        Exit Sub
    End If
    frmOptionJahr.Show 1
End Sub

Private Sub btnStart_Click()
    Dim i As Long
    Dim start As Long
    Dim PosBackSlash As Long
    Dim PosNextBackSlash As Long
    Dim Teilstring As String
    Dim SQL As String
    Dim DuplikatGefunden As Boolean
    Dim Pos As Long
    Dim strTemp As String
    Dim rstUserDef As ADODB.Recordset               'Gerbing 02.08.2016
    Dim rstDefFields As ADODB.Recordset             'Gerbing 03.08.2016
    Dim rstBH As ADODB.Recordset                    'Gerbing 10.07.2007
    Dim AviNummer As Long
    Dim DateinamenErweiterung As String
    Dim HinweisNötig As Boolean
    Dim Msg As String
    Dim antwort As Long
    Dim blnBreiteHoeheEingetragen As Boolean        'Gerbing 10.07.2007
    Dim MM As New MovieModule
    Dim Error As Long
    Dim lngBreite As Long
    Dim lngHöhe As Long
    Dim rstTemp As ADODB.Recordset
    Dim fs As New Scripting.FileSystemObject
    Dim f
    Dim rc As Boolean
    Dim strNDGFotoDatei As String                   'Gerbing 18.11.2019

    blnBreiteHoeheEingetragen = False               'Gerbing 10.07.2007
    Form1.FehlerGefunden = False
    If List1.ListItems.Count = 0 Then
        'MsgBox "Sie haben noch keine Dateinamen in den Drag&Drop-Container gezogen"
        MsgBox LoadResString(1442 + Sprache)
        Exit Sub
    End If
    
    'Gerbing 09.11.2006
    'btnStart wird nur akzeptiert, wenn vorher der Nutzer entschieden hat, was mit dem Jahr passieren soll
    'aber wenn chkExif nicht aktiviert ist, bleibt Extrahieren aus dem Dateiname Standard
    If chkExif.Value = 1 Then
        If NeueDatensätzeGenerieren.blnOptGewählt = False Then
            'MsgBox "Klicken Sie auf den Button 'Jahr...' und wählen Sie eine Einstellung"
            MsgBox LoadResString(1490 + Sprache)
            Exit Sub
        End If
    End If
    HinweisNötig = True
    '----------------------------------------------------------------------------------
    'Bei nutzerdefinierten Feldern auf korrekte Angaben prüfen      'Gerbing 14.02.2005
    If FrameNutzerDefiniert.Visible = True Then
        If chkExif.Value = 1 Then
            'hier ist chkExif.Value = 1
            'Hier wird mit EXIF/IPTC-Feldern gearbeitet
            On Error Resume Next                                            'Gerbing 02.08.2016
            If cmbEx1.Text <> "" And cmbFeld1.Text = "" Then
                'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
                MsgBox LoadResString(1519 + Sprache)
                cmbFeld1.SetFocus
                Exit Sub
            End If
            If cmbEx2.Text <> "" And cmbFeld2.Text = "" Then
                'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
                MsgBox LoadResString(1519 + Sprache)
                cmbFeld2.SetFocus
                Exit Sub
            End If
            If cmbEx3.Text <> "" And cmbFeld3.Text = "" Then
                'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
                MsgBox LoadResString(1519 + Sprache)
                cmbFeld3.SetFocus
                Exit Sub
            End If
            If cmbEx4.Text <> "" And cmbFeld4.Text = "" Then
                'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
                MsgBox LoadResString(1519 + Sprache)
                cmbFeld4.SetFocus
                Exit Sub
            End If
            If cmbEx5.Text <> "" And cmbFeld5.Text = "" Then
                'MsgBox "Wenn Sie ein EXIF/IPTC-Feld auswählen, müssen Sie auch einen Feldnamen auswählen"
                MsgBox LoadResString(1519 + Sprache)
                cmbFeld5.SetFocus
                Exit Sub
            End If
            If cmbFeld1.Text <> "" And (Combo1.Text = "" And cmbEx1.Text = "") Then
                'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch entweder einen Feldwert eingeben oder ein EXIF/IPTC-Feld auswählen"
                MsgBox LoadResString(1520 + Sprache)
                cmbEx1.SetFocus
                Exit Sub
            End If
            If cmbFeld2.Text <> "" And (Combo2.Text = "" And cmbEx2 = "") Then
                'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch entweder einen Feldwert eingeben oder ein EXIF/IPTC-Feld auswählen"
                MsgBox LoadResString(1520 + Sprache)
                cmbEx2.SetFocus
                Exit Sub
            End If
            If cmbFeld3.Text <> "" And (Combo3.Text = "" And cmbEx3 = "") Then
                'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch entweder einen Feldwert eingeben oder ein EXIF/IPTC-Feld auswählen"
                MsgBox LoadResString(1520 + Sprache)
                cmbEx3.SetFocus
                Exit Sub
            End If
            If cmbFeld4.Text <> "" And (Combo4.Text = "" And cmbEx4 = "") Then
                'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch entweder einen Feldwert eingeben oder ein EXIF/IPTC-Feld auswählen"
                MsgBox LoadResString(1520 + Sprache)
                cmbEx4.SetFocus
                Exit Sub
            End If
            If cmbFeld5.Text <> "" And (Combo5.Text = "" And cmbEx5 = "") Then
                'MsgBox "Wenn Sie einen Feldnamen auswählen, müssen Sie auch entweder einen Feldwert eingeben oder ein EXIF/IPTC-Feld auswählen"
                MsgBox LoadResString(1520 + Sprache)
                cmbEx5.SetFocus
                Exit Sub
            End If
        End If
        If blnUserDefinedError = False Then
            'nur wenn blnUserDefinedError = false                                   'Gerbing 02.08.2016
            'nur wenn die Tabelle UserDefined existiert                             'Gerbing 02.08.2016
            'und nur wenn die Tabelle UserDefined leer ist
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
        On Error GoTo 0
    End If
    '-------------------------------------------------------------------------------------------
    'nur wenn die Tabelle DefaultFields existiert                             'Gerbing 03.08.2016
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
    '---------------------------------End Gerbing 03.08.2016------------------------------------
    '-------------------------------------------------------------------------------------------
    'Alle Dateinamen, die in List1 stehen, werden zum Generieren eines neuen Satzes in der Datei
    'fotos.mdb benutzt
    'Dabei werden folgende Prüfungen gemacht
    '1. Es gibt 3 Optionen wie mit der Jahreszahl verfahren werden soll             'Gerbing 09.11.2006
    '   1.1 Im Pfadnamen muss eine 4-stellige Jahreszahl enthalten sein.
    '       Bei mehreren Vorkommen wird das erste genommen.
    '       Wenn keine Jahreszahl gefunden wird, wird das Dummy-Jahr benutzt
    '   1.2 Wenn keine Jahreszahl gefunden wird, soll ein Formular aufgehen und den Nutzer zur
    '       Festlegung einer Jahreszahl auffordern. Dabei soll der Nutzer entscheiden, ob
    '       diese Jahreszahl für alle Fälle gilt, wo kein Jahreszahl gefunden werden konnte,
    '       oder ob der Nutzer in jedem dieser Fälle wiederholt eine Jahreszahl angeben will.
    '   1.3 Die Jahreszahl wird aus einen EXIF-Feld entnommen.
    '       Wenn keine Jahreszahl gefunden wird, wird das Dummy-Jahr benutzt
    '2. Es dürfen keine Duplikate von Dateinamen in der Datei fotos.mdb entstehen.
    '   Die Duplikatkontrolle geschieht über eine SQL-Abfrage.
    '   2.1 Bei chkUnbeaufsichtigt.value = 1 kommt ein Eintrag in die Protokolldatei pruef.log
    '   2.2 Bei chkUnbeaufsichtigt.value = 0 Wenn ein Duplikat entdeckt wird, soll ein Formular aufgehen, wo
    '       der Nutzer entscheiden kann weitere Wanungen anzeigen/keine weiteren Warnungen anzeigen.
    '3  Einträge in die Protokolldatei pruef.log werden auch gemacht, bei Nichteinhaltung der 3-Einigkeit
    '   und bei den nutzerdefinierten Feldern, wenn in numerische Felder nichtnumerische Werte eingetragen
    '   werden sollen
    
    If List1.ListItems.Count <> 0 Then
        'ich will einen Recordset nicht mit sämtlichen fotos, sondern am besten einen leeren recordset
        'darum die Suche nach SWF='123' LoadResString(1029 + Sprache)
        'SQL = "Select * from Fotos Where SWF='123'"
        SQL = "Select * from Fotos Where " & LoadResString(1029 + Sprache) & "='123'"
        Set rstsql = New ADODB.Recordset
        With rstsql
            .Source = SQL
            .ActiveConnection = Form1.DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        
        Form1.SpaltenbreiteWiederherstellen
        For i = 0 To List1.ListItems.Count - 1
            blnIPTCVorhanden = False                            'Gerbing 04.02.2008
            '-------------------------------
            '1. 4-stellige Jahreszahl finden
            start = 1
            Do
                EXIFListInfo = ""                               'Gerbing 12.10.2016
                blnMinusGeprüft = False                         'Gerbing 12.10.2016
                AktuellerDateiname = List1.ListItems(i)
                If left(AktuellerDateiname, 1) = "+" Then
                    AktuellerDateiname = Replace(AktuellerDateiname, "+:", MyAppPath)
                    strNDGFotoDatei = AktuellerDateiname            'Gerbing 18.11.2019
                End If
                txtArbeitsfortschritt.Text = AktuellerDateiname                      'Gerbing 11.04.2005
                PosBackSlash = InStr(start, List1.ListItems(i), "\")
                If PosBackSlash = 0 Then
                    If optManuell.Value = True Then
                        Call FrageNutzerJahreszahl
                    End If
                    Exit Do
                End If
                PosNextBackSlash = InStr(PosBackSlash + 1, List1.ListItems(i), "\")
                If PosNextBackSlash = 0 Then
                    If optManuell.Value = True Then
                        Call FrageNutzerJahreszahl
                    End If
                    Exit Do
                Else
                    'prüfe den Teilstring ob es eine 4-stellige Jahreszahl ist
                    Teilstring = Mid(List1.ListItems(i), PosBackSlash + 1, PosNextBackSlash - PosBackSlash - 1)
                    If Not IsNumeric(Teilstring) Or Len(Teilstring) <> 4 Then
                        'es ist keine 4-stellige Jahreszahl
                        GefundeneJahresZahl = ""
                    Else
                        'es ist eine 4-stellige Jahreszahl
                        GefundeneJahresZahl = Teilstring
                        Exit Do
                    End If
                End If
                start = PosBackSlash + 1
            Loop
            If chkExif.Value = 0 Then
                If GefundeneJahresZahl = "" Then
                    If optManuell = False Then
                        Call FrageNutzerJahreszahl
                    End If
                End If
            End If
            If JahrFestLegenAbgebrochen = True Then
                'MsgBox "Jahr festlegen wurde abgebrochen. Neue Datensätze generieren wird beendet"
                MsgBox LoadResString(1491 + Sprache)
                'dbs.Close                                                                   'Gerbing 29.05.2008
                Unload Me
                Screen.MousePointer = vbDefault              'Gerbing 17.03.2005
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass                'Gerbing 17.03.2005
            '-------------------------------------------------------------------
            'Beim sql server  wie auch Access-version müssen einfache Hochkommas genommen werden    'Gerbing 22.01.2018
            'und ' im Dateiname muss durch 2 Hochkommas ersetzt werden                             'Gerbing 22.01.2018
            '2. Duplikatprüfung
            SQL = "SELECT * From Fotos"
            strTemp = Replace(AktuellerDateiname, MyAppPath, "+:")                      'Gerbing 11.04.2005
            strTemp = Replace(strTemp, "'", "''")                                       'Gerbing 23.01.2018
            'SQL = SQL & " WHERE Dateiname='"  & strTemp & "'" & ";"
            SQL = SQL & " WHERE " & LoadResString(1028 + Sprache) & "='" & strTemp & "'" & ";"
            'Die Selektions-Begriffe stehen im String SQL
            On Error Resume Next                                                            'Gerbing 22.07.2017
            Set rstTemp = New ADODB.Recordset
            With rstTemp
                .Source = SQL
                .ActiveConnection = Form1.DBsql
                .CursorType = adOpenForwardOnly
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            DuplikatGefunden = False
            If Not rstTemp.EOF Then
                DuplikatGefunden = True
                If chkUnbeaufsichtigt.Value = 0 Then            'Gerbing 09.11.2006
                    If NichtNochmalWarnen = False Then
                        Screen.MousePointer = vbDefault          'Gerbing 17.03.2005
                        DuplikatWarnung.Show 1
                        Screen.MousePointer = vbHourglass       'Gerbing 17.03.2005
                    End If
                End If
            End If
            rstTemp.Close
            On Error GoTo 0                                                                 'Gerbing 22.07.2017
            '-------------------------------------------------------------------
            '2A. Prüfung ob vor der Dateinamen-Erweiterung ein Punkt steht                  'Gerbing 09.09.2014
                rc = PrüfePunktVorDateinamenErweiterung(AktuellerDateiname)
                If rc <> 0 Then GoTo Nexti
            '-------------------------------------------------------------------
            '3. Neuen Satz in der Datenbank erzeugen
            'zuerst das Jahr
            If GefundeneJahresZahl = "" And NutzerJahresZahl = "" Then
                'es wurde keine Jahreszahl in Dateiname gefunden
                If optExtrahieren.Value = True Then
                    NutzerJahresZahl = DummyJahr
                End If
                'bei optManuell muss ein Wert in NutzerJahresZahl stehen
            End If
            If optExif.Value = True Then
                GefundeneJahresZahl = ""
                Call EXIFJahrEintragen
                If GefundeneJahresZahl = "" Then
                    NutzerJahresZahl = DummyJahr
                End If
            End If
            If DuplikatGefunden = True And chkUnbeaufsichtigt.Value = 1 Then
            'Bei chkUnbeaufsichtigt.value = 1 kommt ein Eintrag in die Protokolldatei pruef.log
                Call SchreibePruefLogDuplikat
                GoTo Nexti
            End If
            If DuplikatGefunden = False Then
                On Error Resume Next
                rstsql.AddNew       '2x Addnew ohne Update würde einen Fehler erzeugen   'Gerbing 29.11.2011
                On Error GoTo 0
                '----------------------------------------------------------------------
                'wenn ich hier Null eintrage, wird der Inhalt des Feldes leer.      'Gerbing 24.11.2019
                'wenn ich hier nichts eintrage wird der Standardwert '0' eingetragen    'Gerbing 24.11.2019
                rstsql.Fields(LoadResString(1106 + Sprache)) = Null 'Breite         'Gerbing 24.11.2019
                rstsql.Fields(LoadResString(1107 + Sprache)) = Null 'Hoehe          'Gerbing 24.11.2019
                rstsql.Fields("VideoDuration") = Null                               'Gerbing 24.11.2019
                rstsql.Fields("GPSLatitude") = Null                                 'Gerbing 24.11.2019
                rstsql.Fields("GPSLongitude") = Null                                'Gerbing 24.11.2019

                'Die Jahreszahl steht im Feld GefundeneJahresZahl oder in NutzerJahresZahl
                If GefundeneJahresZahl = "" Then
                    'rstsql!Jahr = NutzerJahresZahl
                    rstsql.Fields(LoadResString(1023 + Sprache)) = NutzerJahresZahl
                Else
                    'rstsql!Jahr = GefundeneJahresZahl
                    rstsql.Fields(LoadResString(1023 + Sprache)) = GefundeneJahresZahl
                End If
                'dann Merker
                'rstsql.Fields("Merker") = False
                rstsql.Fields(LoadResString(2524 + Sprache)) = False
                '------------------------------------------------------------------
                'dann die Standardfelder Situation Ort Land Personen
                If StrComp(Right(AktuellerDateiname, 3), "JPG", vbTextCompare) = 0 Then     'Gerbing 26.10.2013
                    'Nur Dateien *.jpg werden auf IPTC-Felder untersucht
                    'IPTC-Felder nur einmal lesen                                           'Gerbing 26.10.2013
                    IPTCItemsDelimiter = ";"
                    rc = LeseIPTC(AktuellerDateiname, LstU, False)               'ohne Ausgabe in LstU
                End If
                If cmbSituationEx.Text = "" Or cmbSituation.Text <> "" Or chkExif.Value = 0 Then 'Gerbing 26.10.2013 10.11.2016
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1024 + Sprache), cmbSituation.Text)
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbSituationEx.Text)
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1024 + Sprache), strGefundenExifIptc)
                    'rstsql.Fields(LoadResString(1024 + Sprache)) = strGefundenExifIptc
                End If
                If cmbOrtEx.Text = "" Or cmbOrt.Text <> "" Or chkExif.Value = 0 Then 'Gerbing 26.10.2013 10.11.2016
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1025 + Sprache), cmbOrt.Text)
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbOrtEx.Text)
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1025 + Sprache), strGefundenExifIptc)
                End If
                If cmbLandEx.Text = "" Or cmbLand.Text <> "" Or chkExif.Value = 0 Then 'Gerbing 26.10.2013 10.11.2016
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1026 + Sprache), cmbLand.Text)
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbLandEx.Text)
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1026 + Sprache), strGefundenExifIptc)
                End If
                If cmbPersonenEx.Text = "" Or cmbPersonen.Text <> "" Or chkExif.Value = 0 Then 'Gerbing 26.10.2013 10.11.2016
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1027 + Sprache), cmbPersonen.Text)
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbPersonenEx.Text)
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1027 + Sprache), strGefundenExifIptc)
                End If
                '------------------------------------------------------------------
                'Standardfeld Dateiname
                strTemp = Replace(AktuellerDateiname, MyAppPath, "+:")     'Gerbing 11.04.2005
                If left(strTemp, 2) <> "+:" Then
                    If chkUnbeaufsichtigt.Value = 1 Or optProtokolldatei.Value = True Then              'Gerbing 29.12.2011
                        Call SchreibePruefLog3Einig
                        GoTo Nexti
                    Else
                        'Verstöße gegen die 3-Einigkeit liegen vor, wenn nicht +: am Anfang von strTemp
                        'gefunden wird. Verstöße werden nicht in die Datenbank aufgenommen. Es kommt ein Hinweis.
                        'Der Hinweis kommt nur einmal und wird dann unterdrückt.
                        If HinweisNötig = True Then
                            'msg = "Seit Version 12.0.0.0 verlangt das Programm die 3-Einigkeit, d.h. dass alle Fotos oder Videos oder andere Dateien" & vbNewLine
                            Msg = LoadResString(2160 + Sprache) & vbNewLine
                            'msg = msg & "unterhalb von AppPath stehen. AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
                            Msg = Msg & LoadResString(2161 + Sprache) & vbNewLine
    '                        Msg = Msg & AktuellerDateiName & " hält diese Forderung nicht ein." & vbNewLine
    '                        Msg = Msg & "Dateinamen, die diese Forderung nicht einhalten, werden nicht in die Datenbank aufgenommen." & vbNewLine
    '                        Msg = Msg & "Dieser Hinweis wird nicht wiederholt."
                            Msg = Msg & AktuellerDateiname & LoadResString(1409 + Sprache) & vbNewLine
                            Msg = Msg & LoadResString(1410 + Sprache) & vbNewLine
                            Msg = Msg & LoadResString(2094 + Sprache)
                            MsgBox Msg
                            HinweisNötig = False
                        End If
                    End If
                    GoTo Nexti
                End If
                'rstsql!Dateiname = strTemp                                        'Gerbing 11.04.2005
                Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1028 + Sprache), strTemp)
                '----------------------------------------------------
                '4 mögliche Bezeichnungen Foto oder Video Farbe oder Schwarz/weiß
                If cmbSWFEx.Text = "" Then                                                  'Gerbing 04.01.2009
                    If OptF = True Then
                        'rstsql!SWF = "F"
                        rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1437 + Sprache) '"F"
                    End If
                    If OptSW = True Then
                        rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1438 + Sprache) '"SW"
                    End If
                    If OptFV = True Then
                        rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1439 + Sprache) '"FV"
                    End If
                    If OptSV = True Then
                        rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1440 + Sprache) '"SV"
                    End If
                Else                                                                        'Gerbing 04.01.2009
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbSWFEx.Text)
                    If strGefundenExifIptc = "" Then                                        'Gerbing 12.12.2016
                        'wenn zwar eine IPTC-Quelle angegeben ist, aber dort nichts steht, nehme ich eine von 4 Standardangaben
                        If OptF = True Then
                            'rstsql!SWF = "F"
                            rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1437 + Sprache) '"F"
                        End If
                        If OptSW = True Then
                            rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1438 + Sprache) '"SW"
                        End If
                        If OptFV = True Then
                            rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1439 + Sprache) '"FV"
                        End If
                        If OptSV = True Then
                            rstsql.Fields(LoadResString(1029 + Sprache)) = LoadResString(1440 + Sprache) '"SV"
                        End If
                    End If
                    On Error Resume Next                                                    'Gerbing 16.06.2017
                    strGefundenExifIptc = rstsql.Fields(LoadResString(1029 + Sprache))      'Gerbing 26.02.2017
                    '20.04.2012 keine Fehler melden wenn Fehler bei Feld SW (BWC)
                    On Error GoTo 0                                                         'Gerbing 16.06.2017
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1029 + Sprache), strGefundenExifIptc)
                End If
                '----------------------------------------------------------------------------------
                'Standardfeld Kommentar
                If cmbKommentarEx.Text = "" Or txtKommentar.Text <> "" Or chkExif.Value = 0 Then 'Gerbing 26.10.2013 10.11.2016
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1030 + Sprache), txtKommentar)
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbKommentarEx.Text)
                    Call FeldEintragenMitLängenprüfung(rstsql, LoadResString(1030 + Sprache), strGefundenExifIptc)
                End If
                '---------------------------------------------------------------------------------
                'DateinameKurz und DDatum eintragen                            'Gerbing 10.10.2004
                start = 1
                Do
                    Pos = InStr(start, AktuellerDateiname, "\")
                    If Pos = 0 Then Exit Do
                    start = Pos + 1
                Loop
                'rstsql("DateinameKurz") = Right(AktuellerDateiName, Len(AktuellerDateiName) - start + 1)
                rstsql(LoadResString(1031 + Sprache)) = Right(AktuellerDateiname, Len(AktuellerDateiname) - start + 1)
                If left(AktuellerDateiname, 2) = "+:" Then                                  'Gerbing 11.04.2005
                    AktuellerDateiname = Replace(AktuellerDateiname, "+:", MyAppPath)        'Gerbing 11.04.2005
                End If
                On Error Resume Next                                                'Gerbing 13.06.2011
                ERR = 0
                'strTemp = FileDateTime(AktuellerDateiName)
                Set f = fs.GetFile(AktuellerDateiname)
                strTemp = f.DateLastModified
                If ERR.Number <> 0 Or strTemp = "00:00:00" Then
                    If optProtokolldatei.Value = True Then
                        Call SchreibePruefLogVerbotenerDateiname                    'Gerbing 09.09.2014
                    Else
    '                    msg = "Fehler im Dateiname " & AktuellerDateiName & vbNewLine
    '                    msg = msg & "error number=" & Err.Number & vbNewLine
    '                    msg = msg & "error text=" & Err.Description & vbNewLine
    '                    msg = msg & "Möglicherweise verbotene Zeichen im Dateiname" & vbnewline
    '                    msg = msg & "oder ungültiges Datei-Datum"
                        Msg = LoadResString(2455 + Sprache) & AktuellerDateiname & vbNewLine
                        Msg = Msg & "error number=" & ERR.Number & vbNewLine
                        Msg = Msg & "error text=" & ERR.Description & vbNewLine
                        Msg = Msg & LoadResString(2456 + Sprache) & vbNewLine
                        Msg = Msg & LoadResString(2464 + Sprache) & vbNewLine
            '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                        Msg = Msg & LoadResString(1507 + Sprache)
                        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                        If antwort = vbYes Then
                            optProtokolldatei.Value = True
                            Call SchreibePruefLogVerbotenerDateiname                    'Gerbing 09.09.2014
                        End If
                    End If
                    On Error GoTo 0
                    Screen.MousePointer = vbDefault
                    GoTo Nexti
                End If
                On Error GoTo 0
                Pos = InStr(1, strTemp, " ")
                If Pos <> 0 Then                                                    'Gerbing 04.11.2010
                    strTemp = left(strTemp, Pos - 1)
                End If                                                              'Gerbing 04.11.2010
                'rstsql("DDatum") = strTemp
                rstsql(LoadResString(1032 + Sprache)) = strTemp
                'BreitePixel und HoehePixel eintragen
                ERR = 0
                Form1.pintBreite = 0
                DateinamenErweiterung = Right(AktuellerDateiname, 3)
                DateinamenErweiterung = UCase(DateinamenErweiterung)                'Gerbing 18.11.2019
                '-----------------------------------------------------------------------
                'für Bilddateien oder Video Breite und Höhe ermitteln
                'bei Videos mciSendString siehe MovieModule.cls benutzen            'Gerbing 26.10.2011
                'bei Fotos Call LoadPicBox
                Form1.pintHoehe = 0
                Form1.lngVideoDuration = 0
                Form1.blnMediaPlayerStopped = False
                Form1.blnMediaPlayerError = False
                '--------------------------------------------------------
                Select Case DateinamenErweiterung
                    Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"  'Gerbing 10.12.2017
                        Form1.WMP.settings.autoStart = False
                        Form1.WMP.Width = 1
                        Form1.WMP.URL = strNDGFotoDatei                             'Gerbing 18.11.2019
                        Form1.WMP.Visible = True     'erst nach Form1.WMP.URL = ...27.11.2016                                                             'Gerbing 01.09.2008
                        On Error Resume Next
                        ERR = 0
                        Form1.WMP.Controls.play
                        'jetzt muss ich warten bis 'player .playState=1(stopped) kommt
                        'bei Fehlern und wenn ich sage 'ja' bei 'soll der player versuchen den Inhalt wiederzugeben' gibt es keinen Loop
                        'bei Fehlern und wenn ich sage 'nein' bei 'soll der player versuchen den Inhalt wiederzugeben' gibt es einen Loop
                        'nach einer Sekunde beende ich den Loop
                        glngStartMillisek = timeGetTime
                        Do
                            glngEndMillisek = timeGetTime
                            If glngEndMillisek - glngStartMillisek > 1000 Then Exit Do
                            If Form1.blnMediaPlayerStopped = True Then Exit Do
                            If Form1.blnMediaPlayerError = True Then Exit Do
                            DoEvents
                        Loop
                        'wenn mp4 oder mov file, hier einfügen Suche mit Mediainfo.DLL               'Gerbing 18.11.2019
                        'If DateinamenErweiterung = "MP4" Or DateinamenErweiterung = "MOV" Then
                            If blnMediaInfoInitialized = False Then
                                If Not file_exist(AppPath + "\MediaInfo.dll") Then
                                    Msg = LoadResString(1558 + Sprache) & vbCr   'Sorry, the MediaInfo.dll(i386 version) not found in the current path!
                                    Msg = Msg & LoadResString(1559 + Sprache) & vbCr 'Put the {MediaInfo.dll} into current path before runnig this Application!
                                    Msg = Msg & LoadResString(1560 + Sprache) & vbNewLine & vbNewLine 'We need Mediainfo.dll for checking GPS info in mp4 videos
                                    
                                    Msg = Msg & LoadResString(1561 + Sprache)   'Do you want to continue without MediaInfo.dll?
                                    antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo)
                                    If antwort = vbNo Then
                                        End                                                                     'Gerbing 18.11.2019
                                    Else
                                        GoTo WeiterOhneMediaInfoDLL                                             'Gerbing 05.06.2019
                                    End If
                                End If
                                On Error Resume Next
                                Handle = MediaInfo_New()    'hier bekomme ich Laufzeitfehler '48' Datei nicht gefunden mediainfo.dll wenn ich das Programm aus der IDE heraus
                                                            'zum zweitenmal starte
                                If ERR.Number <> 0 Then
                                    MsgBox "invalid DLL-Version(must be i386 version)"
                                    End                                                                     'Gerbing 18.11.2019
                                End If
                            End If
                            Call GetMediaInfo
                            blnMediaInfoInitialized = True
                        'End If
WeiterOhneMediaInfoDLL:
                    Case Else
                        'Call LoadPicBox(AktuellerDateiName, Form1.PrüfPicture) 'Gerbing 01.09.2007
                        Call LoadPicBox(AktuellerDateiname) 'Gerbing 01.09.2007
                        Form1.pintBreite = gsngPicWidth
                        Form1.pintHoehe = gsngPicHeight
                End Select
                On Error Resume Next                                        'Gerbing 24.10.2007
                If Form1.lngVideoDuration = 0 Then                           'Gerbing 24.10.2007
                    rstsql.Fields("VideoDuration") = Null
                Else
                    rstsql.Fields("VideoDuration") = Form1.lngVideoDuration
                End If
                On Error GoTo 0                                             'Gerbing 24.10.2007
                If Form1.pintBreite <> 0 Then
'                    rstsql.Fields("BreitePixel") = intBreite
'                    rstsql.Fields("HoehePixel") = intHoehe
                    rstsql.Fields(LoadResString(1106 + Sprache)) = Form1.pintBreite
                    rstsql.Fields(LoadResString(1107 + Sprache)) = Form1.pintHoehe
                    blnBreiteHoeheEingetragen = True                        'Gerbing 10.07.2007
                End If
                'AudioFileExists
                rstsql.Fields(LoadResString(2537 + Sprache)) = False
                '--------------------------------------------------------
                '====================================================================================
                'Nutzerdefinierte Felder eintragen
                'Eintrag in die Protokolldatei (Pruef.log) bei den nutzerdefinierten Feldern, wenn in
                'numerische Felder nichtnumerische Werte eingetragen werden sollen oder andere
                'Typ-Unverträglichkeiten wegen Datum oder ja/nein Feld
                
                'chkExif braucht nicht eingeschaltet zu sein um die Felder GPSLatitude, GPSLongitude, ExifDateTimeOriginal zu füllen
                'Hauptsache ist, dass in Tabelle UserDefined eine Zuordnung getroffen wurde
                
                If cmbEx1.Text = "" Then
                    If cmbFeld1.Text <> "" Then
                        On Error Resume Next
                        rstsql.Fields(cmbFeld1.Text) = Combo1
                        Call TypeError(cmbFeld1.Text, Combo1)
                        On Error GoTo 0
                    End If
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbEx1.Text)
                    If strGefundenExifIptc <> "" Then                                                   'Gerbing 21.10.2016
                        On Error Resume Next
                        If cmbFeld1.Text = "GPSLatitude" Or cmbFeld1.Text = "GPSLongitude" Then         'Gerbing 12.10.2016
                            rstsql.Fields(cmbFeld1.Text) = CDbl(strGefundenExifIptc)
                        Else
                            rstsql.Fields(cmbFeld1.Text) = strGefundenExifIptc
                        End If
                        If cmbFeld1.Text = "ExifDateTimeOriginal" Then                                  'Gerbing 25.11.2019
                            rstsql.Fields(cmbFeld1.Text) = strGefundenExifIptc                          'Gerbing 25.11.2019
                        End If                                                                          'Gerbing 25.11.2019
                        Call TypeError(cmbFeld1.Text, strGefundenExifIptc)
                        On Error GoTo 0
                    End If
                    If gstrLatXMP <> "" Then                                                            'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        'Fehler 3265, wenn Feld GPSLatitude nicht vorhanden ist, aber befüllt werden soll
                        rstsql.Fields("GPSLatitude") = gstrLat
                    End If
                    If gstrLongXMP <> "" Then                                                           'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLongitude") = gstrLong
                    End If
                End If
                If cmbEx2.Text = "" Then
                    If cmbFeld2.Text <> "" Then
                        On Error Resume Next
                        rstsql.Fields(cmbFeld2.Text) = Combo2
                        Call TypeError(cmbFeld2.Text, Combo2)
                        On Error GoTo 0
                    End If
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbEx2.Text)
                    If strGefundenExifIptc <> "" Then                                                   'Gerbing 21.10.2016
                        On Error Resume Next
                        If cmbFeld2.Text = "GPSLatitude" Or cmbFeld2.Text = "GPSLongitude" Then         'Gerbing 12.10.2016
                            rstsql.Fields(cmbFeld2.Text) = CDbl(strGefundenExifIptc)
                        Else
                            rstsql.Fields(cmbFeld2.Text) = strGefundenExifIptc
                        End If
                        If cmbFeld2.Text = "ExifDateTimeOriginal" Then                                  'Gerbing 25.11.2019
                            rstsql.Fields(cmbFeld2.Text) = strGefundenExifIptc                          'Gerbing 25.11.2019
                        End If                                                                          'Gerbing 25.11.2019
                        On Error GoTo 0
                    End If
                    If gstrLatXMP <> "" Then                                                            'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLatitude") = gstrLat
                    End If
                    If gstrLongXMP <> "" Then                                                           'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLongitude") = gstrLong
                    End If
                End If
                If cmbEx3.Text = "" Then
                    If cmbFeld3.Text <> "" Then
                        On Error Resume Next
                        rstsql.Fields(cmbFeld3.Text) = Combo3
                        Call TypeError(cmbFeld3.Text, Combo3)
                        On Error GoTo 0
                    End If
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbEx3.Text)
                    If strGefundenExifIptc <> "" Then                                                   'Gerbing 21.10.2016
                        On Error Resume Next
                        If cmbFeld3.Text = "GPSLatitude" Or cmbFeld3.Text = "GPSLongitude" Then         'Gerbing 12.10.2016
                            rstsql.Fields(cmbFeld3.Text) = CDbl(strGefundenExifIptc)
                        Else
                            rstsql.Fields(cmbFeld3.Text) = strGefundenExifIptc
                        End If
                        If cmbFeld3.Text = "ExifDateTimeOriginal" Then                                  'Gerbing 25.11.2019
                            rstsql.Fields(cmbFeld3.Text) = strGefundenExifIptc                          'Gerbing 25.11.2019
                        End If                                                                          'Gerbing 25.11.2019
                        On Error GoTo 0
                    End If
                    If gstrLatXMP <> "" Then                                                            'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLatitude") = gstrLat
                    End If
                    If gstrLongXMP <> "" Then                                                           'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLongitude") = gstrLong
                    End If
                End If
                If cmbEx4.Text = "" Then
                    If cmbFeld4.Text <> "" Then
                        On Error Resume Next
                        rstsql.Fields(cmbFeld4.Text) = Combo4
                        Call TypeError(cmbFeld4.Text, Combo4)
                        On Error GoTo 0
                    End If
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbEx4.Text)
                    If strGefundenExifIptc <> "" Then                                                   'Gerbing 21.10.2016
                        On Error Resume Next
                        If cmbFeld4.Text = "GPSLatitude" Or cmbFeld4.Text = "GPSLongitude" Then         'Gerbing 12.10.2016
                            rstsql.Fields(cmbFeld4.Text) = CDbl(strGefundenExifIptc)
                        Else
                            rstsql.Fields(cmbFeld4.Text) = strGefundenExifIptc
                        End If
                        If cmbFeld4.Text = "ExifDateTimeOriginal" Then                                  'Gerbing 25.11.2019
                            rstsql.Fields(cmbFeld4.Text) = strGefundenExifIptc                          'Gerbing 25.11.2019
                        End If                                                                          'Gerbing 25.11.2019
                        On Error GoTo 0
                    End If
                    If gstrLatXMP <> "" Then                                                            'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLatitude") = gstrLat
                    End If
                    If gstrLongXMP <> "" Then                                                           'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLongitude") = gstrLong
                    End If
                End If
                If cmbEx5.Text = "" Then
                    If cmbFeld5.Text <> "" Then
                        On Error Resume Next
                        rstsql.Fields(cmbFeld5.Text) = Combo5
                        Call TypeError(cmbFeld5.Text, Combo5)
                        On Error GoTo 0
                    End If
                Else
                    strGefundenExifIptc = ""
                    Call FindeExifIptc(cmbEx5.Text)
                    If strGefundenExifIptc <> "" Then                                                   'Gerbing 21.10.2016
                        On Error Resume Next
                        If cmbFeld5.Text = "GPSLatitude" Or cmbFeld5.Text = "GPSLongitude" Then         'Gerbing 12.10.2016
                            rstsql.Fields(cmbFeld5.Text) = CDbl(strGefundenExifIptc)
                        Else
                            rstsql.Fields(cmbFeld5.Text) = strGefundenExifIptc
                        End If
                        If cmbFeld5.Text = "ExifDateTimeOriginal" Then                                  'Gerbing 25.11.2019
                            rstsql.Fields(cmbFeld5.Text) = strGefundenExifIptc                          'Gerbing 25.11.2019
                        End If                                                                          'Gerbing 25.11.2019
                        On Error GoTo 0
                    End If
                    If gstrLatXMP <> "" Then                                                            'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLatitude") = gstrLat
                    End If
                    If gstrLongXMP <> "" Then                                                           'Gerbing 08.04.2019
                        Call Form1.GEOKoordinatenUmrechnenXMP
                        rstsql.Fields("GPSLongitude") = gstrLong
                    End If
                End If
                '-------------------------------------------------------
                If blnIPTCVorhanden = True Then                                 'Gerbing 04.02.2008
                    On Error Resume Next    'falls jemand mit der exe ab 13.3.5 mit einer fotos.mdb arbeitet, wo das Feld IPTCPresent nicht drin ist weiss der Kuckuck warum nicht
                    rstsql.Fields("IPTCPresent") = True
                    On Error GoTo 0
                Else
                    On Error Resume Next    'falls jemand mit der exe ab 13.3.5 mit einer fotos.mdb arbeitet, wo das Feld IPTCPresent nicht drin ist weiss der Kuckuck warum nicht
                    rstsql.Fields("IPTCPresent") = False
                    On Error GoTo 0
                End If
                On Error GoTo rstsqlUpdate
                rstsql.Update
            End If
Nexti:
        DoEvents
        btnStart.Enabled = False
        Next i
        
        If blnBreiteHoeheEingetragen = True Then    'Gerbing 10.07.2007
            'jetzt wird Tabelle ErsterStart, Feld DatumBreiteHoehe mit dem Datum von heute aktualisiert
            SQL = "select * From ErsterStart;"
            'SQL = "SELECT * From " & LoadResString(2527 + Sprache) & ";"
            Set rstBH = New ADODB.Recordset
            With rstBH
                .Source = SQL
                .ActiveConnection = Form1.DBsql
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            If rstBH.EOF Then
                If gblnSchreibgeschützt = False Then
                    rstBH.AddNew
                End If
            Else
                If gblnSchreibgeschützt = False Then
                    'rstBH.Edit
                End If
            End If
            If gblnSchreibgeschützt = False Then
                rstBH.Fields("DatumBreiteHoehe") = Date         '2528=DatumBreiteHoehe
                rstBH.Fields("ErsterStart") = 0                 '2527=ErsterStart
                rstBH.Update
            End If
            rstBH.Close
        End If
        'Msg = "Neue Datensätze generieren ist beendet." & NL
        Msg = LoadResString(1441 + Sprache) & NL
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        'Close #DateiNummer
        oStream.Close
        Screen.MousePointer = vbDefault              'Gerbing 17.03.2005
        '-----------------------------------------------
        'Gerbing 09.11.2006
        'Wenn pruef.log nicht leer ist, muss eine Msgbox kommen, wo der Nutzer entscheiden kann, ob er sofort
        'den Inhalt von pruef.log sehen will
        If Form1.FehlerGefunden = True Then
            'Msg = Msg & "Beim Generieren neuer Datensätze sind Fehler aufgetreten." & NL
            'Msg = Msg & "Wollen Sie die Fehlerprotokolldatei (pruef.log) öffnen?"
            Msg = Msg & LoadResString(1501 + Sprache) & NL
            Msg = Msg & LoadResString(1502 + Sprache) & NL
            Msg = Msg & PruefLogFile                                                    'Gerbing 02.08.2016
        End If
        If Form1.FehlerGefunden = False Then
            MsgBox Msg
        Else
            antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            If antwort = vbYes Then
                frmLogLesenGenerierung.Show 1
            End If
        End If
        'dbs.Close                                                                      'Gerbing 29.05.2008
        Do Until NachPrüfen3Löschen.KollZusätzlicheDateien.Count = 0                    'Gerbing 26.10.2013
            NachPrüfen3Löschen.KollZusätzlicheDateien.Remove 1
        Loop
        Do Until NachPrüfen3Aufnehmen.KollZusätzlicheDateien.Count = 0                  'Gerbing 26.10.2013
            NachPrüfen3Aufnehmen.KollZusätzlicheDateien.Remove 1
        Loop
        Unload Me
    End If
    Exit Sub

rstsqlUpdate:                                                                           'Gerbing 06.09.2013
    Msg = "errornumber=" & ERR.Number & vbNewLine
    Msg = Msg & "errortext=" & ERR.Description & vbNewLine
    Msg = Msg & "Kontrollieren Sie ob sich Datenbank-Felder mit MS Access ändern lassen"
    MsgBox Msg
    'End
    Resume Next
End Sub

Private Sub chkExif_Click()
    Dim i As Long
    
    FrameStandardWerte.ToolTipText = LoadResString(2564 + Sprache)      'Mit Rechtsklick können Sie die Feld-Zuordnung zurücksetzen 'Gerbing 12.10.2019
    FrameNutzerDefiniert.ToolTipText = LoadResString(2564 + Sprache)    'Mit Rechtsklick können Sie die Feld-Zuordnung zurücksetzen 'Gerbing 12.10.2019
    If chkExif.Value = 1 Then
        'chkExif Häkchen ist gesetzt
        For i = 0 To 11                                                                     'Gerbing 23.06.2011
            Image1(i).Visible = True
        Next i
        FrameExifIptc.Visible = True
        lblExifIptc1.Visible = True
        lblExifIptc2.Visible = True
        cmbJahrEx.Visible = True
        cmbSituationEx.Visible = True
        cmbOrtEx.Visible = True
        cmbLandEx.Visible = True
        cmbPersonenEx.Visible = True
        cmbSWFEx.Visible = True                                                             'Gerbing 04.01.2009
        cmbKommentarEx.Visible = True
        cmbEx1.Visible = True
        cmbEx2.Visible = True
        cmbEx3.Visible = True
        cmbEx4.Visible = True
        cmbEx5.Visible = True
        cmbJahrEx.AddItem "EXIF-DateTime"
        cmbJahrEx.AddItem "IPTC-Date created"                           'Gerbing 23.01.2008
        cmbJahrEx.Text = cmbJahrEx.List(0)
        cmbSituation.Enabled = True
        cmbOrt.Enabled = True
        cmbLand.Enabled = True
        txtKommentar.Enabled = True
    Else
        'chkExif Häkchen ist nicht gesetzt
        For i = 0 To 11                                                                     'Gerbing 23.06.2011
            Image1(i).Visible = False
        Next i
        chkUnbeaufsichtigt.Value = 0
        FrameExifIptc.Visible = False
        lblExifIptc1.Visible = False
        lblExifIptc2.Visible = False
        cmbJahrEx.Visible = False
        cmbSituationEx.Visible = False
        cmbOrtEx.Visible = False
        cmbLandEx.Visible = False
        cmbPersonenEx.Visible = False
        cmbSWFEx.Visible = False                                                            'Gerbing 04.01.2009
        cmbKommentarEx.Visible = False
        cmbEx1.Visible = False
        cmbEx2.Visible = False
        cmbEx3.Visible = False
        cmbEx4.Visible = False
        cmbEx5.Visible = False
        If optExif.Value = True Then
            'MsgBox "Sie müssen die Einstellung für das Jahr korrigieren"
            MsgBox LoadResString(1494 + Sprache)
            Call btnJahr_Click
        End If
    End If
End Sub

Private Sub chkExifAnzeigen_Click()                                     'Gerbing 09.11.2006
    Dim tempFilename As String
    Dim n As Long
    Dim blnGefunden As Boolean

    chkIptcAnzeigen.Value = 0
    If chkExifAnzeigen.Value = 0 Then
        LstU.Visible = False
        txtEXIFInfo.Visible = False                                     'Gerbing 07.05.2007
    Else
        LstU.Visible = False
        txtEXIFInfo.Visible = True
        blnGefunden = False
        For n = 0 To List1.ListItems.Count - 1
            If List1.ListItems(n).Selected = True Then
                blnGefunden = True
                tempFilename = Replace(List1.ListItems(n).Text, "+:\", MyAppPath & "\")   'Gerbing 07.05.2007
                Form1.EXF.ImageFile = tempFilename 'set the image file property
                '
                'EXF.ListInfo ist ein String mit vbCrLf
                '
                txtEXIFInfo.Text = Form1.EXF.ListInfo 'list all tags into the text box
            End If
        Next n
        If blnGefunden = False Then
            txtEXIFInfo.Text = LoadResString(1483 + Sprache) 'keine Datei markiert
        End If
    End If
End Sub

Private Sub chkIptcAnzeigen_Click()                                     'Gerbing 09.11.2006
    Dim tempFilename As String
    Dim strTemp As String
    Dim start As Long
    Dim Länge As Long
    Const Standardlänge As Long = 70
    Dim n As Long
    Dim blnGefunden As Boolean
    Dim rc As Boolean
    
    chkExifAnzeigen.Value = 0
    If chkIptcAnzeigen.Value = 0 Then
        LstU.Visible = False
        txtEXIFInfo.Visible = False
    Else
        LstU.Visible = True
        txtEXIFInfo.Visible = False                                     'Gerbing 07.05.2007
        LstU.ListItems.RemoveAll
        blnGefunden = False
        For n = 0 To List1.ListItems.Count - 1
            If List1.ListItems(n).Selected = True Then
                blnGefunden = True
                IPTCItemsDelimiter = ";"
                rc = LeseIPTC(Replace(List1.ListItems(n).Text, "+:\", MyAppPath & "\"), LstU, True)  'mit Ausgabe in LstU
            End If
        Next n
        If blnGefunden = False Then
            LstU.ListItems.Add LoadResString(1483 + Sprache)  'keine Datei markiert
        End If
    End If
End Sub

Private Sub chkUnbeaufsichtigt_Click()
    Dim SQL As String
    
    If chkUnbeaufsichtigt.Value = 1 Then
        optProtokolldatei.Value = True
        optMsgbox.Enabled = False
        optProtokolldatei.Enabled = False
        If optManuell.Value = True Then
            'MsgBox "Sie müssen die Einstellung für das Jahr korrigieren"
            MsgBox LoadResString(1494 + Sprache)
            Call btnJahr_Click
        End If
    Else
        optMsgbox.Value = True
        optMsgbox.Enabled = True
        optProtokolldatei.Enabled = True
    End If
End Sub

Private Sub cmbEx1_Click()
     Combo1.ComboItems.RemoveAll
    Combo1.Enabled = False
End Sub

Private Sub cmbEx2_Click()
    Combo2.ComboItems.RemoveAll
    Combo2.Enabled = False
End Sub

Private Sub cmbEx3_Click()
    Combo3.ComboItems.RemoveAll
    Combo3.Enabled = False
End Sub

Private Sub cmbEx4_Click()
    Combo4.ComboItems.RemoveAll
    Combo4.Enabled = False
End Sub

Private Sub cmbEx5_Click()
    Combo5.ComboItems.RemoveAll
    Combo5.Enabled = False
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
    ERR.Number = 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then                                 'Gerbing 02.08.2016
        blnUserDefinedError = True
        'MsgBox "Feldname error"
        MsgBox LoadResString(1033 + Sprache) & " error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0                                         'Gerbing 02.08.2016
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld1.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            If Form1.rstsql.Fields(0).Type = 11 Then                'Siehe DEFINEs for field types
                Combo1.ComboItems.Add 0
            Else
                Combo1.ComboItems.Add Form1.rstsql.Fields(0)
            End If
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
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
    ERR.Number = 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then                                 'Gerbing 02.08.2016
        blnUserDefinedError = True
        'MsgBox "Feldname error"
        MsgBox LoadResString(1033 + Sprache) & " error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0                                         'Gerbing 02.08.2016
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld2.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            If Form1.rstsql.Fields(0).Type = 11 Then                'Siehe DEFINEs for field types
                Combo2.ComboItems.Add 0
            Else
                Combo2.ComboItems.Add Form1.rstsql.Fields(0)
            End If
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
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
    ERR.Number = 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then                                 'Gerbing 02.08.2016
        blnUserDefinedError = True
        'MsgBox "Feldname error"
        MsgBox LoadResString(1033 + Sprache) & " error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0                                         'Gerbing 02.08.2016
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld3.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            If Form1.rstsql.Fields(0).Type = 11 Then                'Siehe DEFINEs for field types
                Combo3.ComboItems.Add 0
            Else
                Combo3.ComboItems.Add Form1.rstsql.Fields(0)
            End If
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
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
    ERR.Number = 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then                                 'Gerbing 02.08.2016
        blnUserDefinedError = True
        'MsgBox "Feldname error"
        MsgBox LoadResString(1033 + Sprache) & " error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0                                         'Gerbing 02.08.2016
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld4.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            If Form1.rstsql.Fields(0).Type = 11 Then                'Siehe DEFINEs for field types
                Combo4.ComboItems.Add 0
            Else
                Combo4.ComboItems.Add Form1.rstsql.Fields(0)
            End If
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
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
    ERR.Number = 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then                                 'Gerbing 02.08.2016
        blnUserDefinedError = True
        'MsgBox "Feldname error"
        MsgBox LoadResString(1033 + Sprache) & " error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0                                         'Gerbing 02.08.2016
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields(cmbFeld5.Text)) Then
        If Not IsNull(Form1.rstsql.Fields(0)) Then
            If Form1.rstsql.Fields(0).Type = 11 Then                'Siehe DEFINEs for field types
                Combo5.ComboItems.Add 0
            Else
                Combo5.ComboItems.Add Form1.rstsql.Fields(0)
            End If
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    If Combo5.ComboItems.Count <> 0 Then
        'Combo5.ListIndex = 0
    End If
End Sub

Private Sub cmbKommentarEx_Click()
    txtKommentar.Text = ""
    txtKommentar.Enabled = False
End Sub

Private Sub cmbLandEx_Click()
    cmbLand.ComboItems.RemoveAll
    cmbLand.Enabled = False
End Sub

Private Sub cmbOrtEx_Click()
    cmbOrt.ComboItems.RemoveAll
    cmbOrt.Enabled = False
End Sub

Private Sub cmbPersonenEx_Click()
    cmbPersonen.ComboItems.RemoveAll
    cmbPersonen.Enabled = False
End Sub

Private Sub cmbSituationEx_Click()
    cmbSituation.ComboItems.RemoveAll                                                                      'Gerbing 04.01.2009
    cmbSituation.Enabled = False                                                            'Gerbing 04.01.2009
End Sub

Private Sub cmbSWFEx_Click()
    OptF.Enabled = False                                                                    'Gerbing 04.01.2009
    OptFV.Enabled = False
    OptSV.Enabled = False
    OptSW.Enabled = False
End Sub

Private Sub Form_Load()
    Dim Feldname As String
    Dim Gefunden As Boolean
    Dim n As Long
    Dim SQL As String
    Dim Faktor As Integer
    Dim rc As Long
    Dim Pixel As Integer
    Dim Msg As String
    Dim rstUserDef As ADODB.Recordset               'Gerbing 02.08.2016
    Dim rstDefFields As ADODB.Recordset             'Gerbing 03.08.2016
    
    blnErsterDurchlauf = True                       'Gerbing 03.08.2016
    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    Call AnpassenNutzerWunsch(Me)                                                       'Gerbing 11.03.2017
    For n = 0 To 11                                                                     'Gerbing 23.06.2011
        Image1(n).Visible = False
    Next n
    Me.Caption = LoadResString(1350 + Sprache)  'Neue Datensätze generieren
    Label1.Caption = LoadResString(1351 + Sprache)     'standardmäßig wird das Jahr aus dem Dateiname extrahiert
    btnStart.Caption = LoadResString(3101 + Sprache)        'S&tart
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)        '&Abbrechen
    FrameStandardWerte.Caption = LoadResString(1352 + Sprache)      'Standard-Felder
    'Label2.Caption = LoadResString(1353 + Sprache)      'Das Jahr wird aus dem Dateinamen extrahiert
    Label3.Caption = LoadResString(1024 + Sprache) & ":"      'Situation:
    Label4.Caption = LoadResString(1025 + Sprache) & ":"       'Ort:
    Label5.Caption = LoadResString(1026 + Sprache) & ":"       'Land:
    Label6.Caption = LoadResString(1027 + Sprache) & ":"       'Personen:
    Frame2.Caption = LoadResString(1029 + Sprache) & ":"       'SWF
    OptF.Caption = LoadResString(1354 + Sprache)        'F (Farb-Foto)
    OptSW.Caption = LoadResString(1355 + Sprache)       'SW (Schwarz/weiß-Foto)
    OptFV.Caption = LoadResString(1356 + Sprache)       'FV (Farb-Video)
    OptSV.Caption = LoadResString(1357 + Sprache)       'SV (Schwarz/weiß-Video)
    Label7.Caption = LoadResString(1030 + Sprache) & ":"      'Kommentar:
    FrameNutzerDefiniert.Caption = LoadResString(1358 + Sprache) 'Nutzerdefinierte Felder
    If gblnSQLServerVersion = True Then
        FrameNutzerDefiniert.Visible = True
    Else
    'die Abfrage ob Professional Version entfällt bei SQL Server version
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
    lblDateinamen.Caption = LoadResString(1359 + Sprache)       'Drag&&Drop-Container für Dateinamen vom Windows-Explorer
    Label17.Caption = LoadResString(1014 + Sprache) & ":"       'Arbeitsfortschritt:
    '09.11.2006
    chkUnbeaufsichtigt.Caption = LoadResString(1472 + Sprache) 'unbeaufsichtigt
    chkUnbeaufsichtigt.ToolTipText = LoadResString(1473 + Sprache) 'Setzen Sie hier ein Häkchen, wenn die Datenbank automatisch erzeugt werden soll
    chkExif.Caption = LoadResString(1475 + Sprache) 'EXIF/IPTC benutzen
    chkExif.ToolTipText = LoadResString(1474 + Sprache) 'Setzen Sie hier ein Häkchen, wenn Sie EXIF-Felder importieren wollen
    btnHilfe.Caption = LoadResString(3014 + Sprache)    '&Hilfe
    btnJahr.Caption = LoadResString(1478 + Sprache)  '&Jahr...
    lblExifIptc1.Caption = LoadResString(1479 + Sprache)    'EXIF/IPTC-Feld
    lblExifIptc2.Caption = LoadResString(1479 + Sprache)    'EXIF/IPTC-Feld
    lblFeldinhalt.Caption = LoadResString(1480 + Sprache)   'Feldinhalt
    chkExifAnzeigen.Caption = LoadResString(1116 + Sprache) 'EXIF-Felder
    chkExifAnzeigen.ToolTipText = LoadResString(1481 + Sprache) 'Setzen Sie hier ein Häkchen, wenn Sie die EXIF-Felder des markiertenSatzes im Drag&Drop-Container sehen wollen
    chkIptcAnzeigen.Caption = LoadResString(1117 + Sprache) 'IPTC-Felder
    chkIptcAnzeigen.ToolTipText = LoadResString(1482 + Sprache) 'Setzen Sie hier ein Häkchen, wenn Sie die IPTC-Felder des markiertenSatzes im Drag&Drop-Container sehen wollen
    'Horizontalen Scrollbar an Listbox LstU anbringen        'Gerbing 09.11.2006
    Faktor = 4
    Pixel = LstU.Width / Screen.TwipsPerPixelX * Faktor
    rc = SendMessage(LstU.hWnd, LB_SETHORIZONTALEXTENT, Pixel, 0)
    FrameExifIptc.Visible = False
    lblExifIptc1.Visible = False
    lblExifIptc2.Visible = False
    cmbJahrEx.Visible = False
    cmbSituationEx.Visible = False
    cmbOrtEx.Visible = False
    cmbLandEx.Visible = False
    cmbPersonenEx.Visible = False
    cmbSWFEx.Visible = False                                                                'Gerbing 04.01.2009
    cmbKommentarEx.Visible = False
    cmbEx1.Visible = False
    cmbEx2.Visible = False
    cmbEx3.Visible = False
    cmbEx4.Visible = False
    cmbEx5.Visible = False
    'Öffne die Datei pruef.log
    On Error Resume Next
    Form1.DateiNummer = FreeFile  ' neue Datei-Nr.
    ERR = 0
    Set oStream = PruefFso.CreateTextFile(PruefLogFile, True, True)
    If ERR <> 0 And ERR <> 55 Then                                                          'Gerbing 23.06.2011
        'Msg = "Die Datei " & PruefLogFile & " kann nicht geöffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & PruefLogFile & " " & LoadResString(1372 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number
        MsgBox Msg, vbCritical
        End
    End If
    On Error GoTo 0
    optMsgbox.Value = True
    optMsgbox.Enabled = True
    optProtokolldatei.Enabled = True
    lblDatentypfehler.Caption = LoadResString(1487 + Sprache) 'Was soll bei Fehlern geschehen?
    optProtokolldatei.Caption = LoadResString(1503 + Sprache)   'Eintragen in die Protokolldatei (pruef.log)
    optMsgbox.Caption = LoadResString(1504 + Sprache)   'Ich will bei jedem Fehler einen Fehlerhinweis erhalten
    
    NL = Chr(10) & Chr(13)
    DErwMsgBoxMussKommen = True
    GiltFürAlleFälle = False                'Anfangswert
    NichtNochmalWarnen = False              'Anfangswert
    JahrFestLegenAbgebrochen = False        'Anfangswert
    'Set wrk = CreateWorkspace("", "Admin", "")
    If gblnSchreibgeschützt = True Then
        MsgBox LoadResString(1371 + Sprache)                                            '"Bei einer schreibgeschützten Datenbank können keine neuen Sätze generiert werden."
        Unload Me
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------
    'jetzt wird untersucht, ob es nutzerdefinierte Felder gibt
    'die haben dann andere Feldnamen als die Namen der Standardfelder
    On Error Resume Next
    Form1.rstsql.Close
    On Error GoTo 0
    With Form1.rstsql
        .Source = "select * from Fotos"                 'select * from fotos
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    For n = 0 To Form1.rstsql.Fields.Count - 1
        Feldname = Form1.rstsql.Fields(n).Name
        Feldname = LCase(Feldname)
        Select Case Feldname
            'AudioFileExists gehört nicht zu den nutzerdefinierten Feldern              'Gerbing 14.05.2006
            'IPTCPresent gehört nicht zu den nutzerdefinierten Feldern                  'Gerbing 04.02.2008
            'Case "merker", "jahr", "situation", "ort", "land", "personen", "dateiname", "swf", "kommentar", "dateinamekurz", "ddatum", "breitepixel", "hoehepixel", "audiofileexists"
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
    'für alle Felder wird eine Combobox mit den schon vorhandenen Werten angeboten  Gerbing 10.06.2005
    Call FülleComboBoxen
    Call AddExifFelder(cmbSituationEx)
    Call AddExifFelder(cmbOrtEx)
    Call AddExifFelder(cmbLandEx)
    Call AddExifFelder(cmbPersonenEx)
    Call AddExifFelder(cmbSWFEx)                                                        'Gerbing 04.01.2009
    Call AddExifFelder(cmbKommentarEx)
    Call AddExifFelder(cmbEx1)
    Call AddExifFelder(cmbEx2)
    Call AddExifFelder(cmbEx3)
    Call AddExifFelder(cmbEx4)
    Call AddExifFelder(cmbEx5)
    '-----------------------------------------------------------------------------------------
    'Tabelle UserDefined auswerten                                                      'Gerbing 02.08.2016
    If cmbFeld1.ComboItems.Count = 0 Then FrameNutzerDefiniert.Visible = False
    If cmbFeld1.ComboItems.Count <> 0 Then                                              'Gerbing 02.08.2016
        'wenn in der Tabelle UserDefined etwas steht,                                   'Gerbing 02.08.2016
        'dann muss ich die Feldnamen und 'Quell EXIF/IPTC-Felder' aus der Tabelle UserDefined auffüllen
        'und den Schalter blnDefaultFieldsNotEmpty einschalten
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
            On Error GoTo 0
            If Not rstUserDef.EOF Then
                'hier ist die Tabelle UserDefined nicht leer
                blnUserdefinedFieldsNotEmpty = True
                If rstUserDef.Fields("FieldName1") <> "" Then
                    cmbFeld1.Text = rstUserDef.Fields("FieldName1")
                    cmbEx1.Text = rstUserDef.Fields("SourceField1")
                End If
                If rstUserDef.Fields("FieldName2") <> "" Then
                    cmbFeld2.Text = rstUserDef.Fields("FieldName2")
                    cmbEx2.Text = rstUserDef.Fields("SourceField2")
                End If
                If rstUserDef.Fields("FieldName3") <> "" Then
                    cmbFeld3.Text = rstUserDef.Fields("FieldName3")
                    cmbEx3.Text = rstUserDef.Fields("SourceField3")
                End If
                If rstUserDef.Fields("FieldName4") <> "" Then
                    cmbFeld4.Text = rstUserDef.Fields("FieldName4")
                    cmbEx4.Text = rstUserDef.Fields("SourceField4")
                End If
                If rstUserDef.Fields("FieldName5") <> "" Then
                    cmbFeld5.Text = rstUserDef.Fields("FieldName5")
                    cmbEx5.Text = rstUserDef.Fields("SourceField5")
                End If
            End If
        End If
        rstUserDef.Close
    End If
    On Error GoTo 0
    'Tabelle UserDefined auswerten End                                                      'Gerbing 02.08.2016
    '----------------------------------------------------------------------------------------------------------
    'Tabelle DefaultFields auswerten                                                        'Gerbing 03.08.2016
    'wenn in der Tabelle DefaultFields etwas steht,                                         'Gerbing 03.08.2016
    'dann muss ich den Schalter blnDefaultFieldsNotEmpty einschalten
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
        'hier existiert die Tabelle UserDefined
        If Not rstDefFields.EOF Then
            'hier ist die Tabelle UserDefined nicht leer
            blnDefaultFieldsNotEmpty = True
            cmbSituationEx.Text = rstDefFields.Fields("SituationSource")
            cmbOrtEx.Text = rstDefFields.Fields("LocationSource")
            cmbLandEx.Text = rstDefFields.Fields("CountrySource")
            cmbPersonenEx.Text = rstDefFields.Fields("PeopleSource")
            cmbSWFEx.Text = rstDefFields.Fields("BWCSource")
            cmbKommentarEx = rstDefFields.Fields("CommentSource")
        Else
            'für die Standarddatenbankfelder gibt es einige IPTC-Felder, die ich voreinstelle (Eigenbedarf)
            'cmbJahrEx.ListIndex = 9                             '9=Date created
            cmbSituationEx.ListIndex = 11 + 49 + 52 + 5                   '11=Headline
            cmbOrtEx.ListIndex = 5 + 49 + 52 + 5                         '5=City
            cmbLandEx.ListIndex = 7 + 49 + 52 + 5                        '7=Country
            cmbKommentarEx.ListIndex = 2 + 49 + 52 + 5                    '2=Caption
            cmbSituation.Enabled = True                         'Gerbing 17.03.2017
            cmbOrt.Enabled = True                               'Gerbing 17.03.2017
            cmbLand.Enabled = True                              'Gerbing 17.03.2017
        End If
    End If
    'Tabelle UserDefined auswerten End                                                      'Gerbing 03.08.2016
    '----------------------------------------------------------------------------------------------------------
    Me.top = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnMediaInfoInitialized = True Then
        Call MediaInfo_Close(Handle)                                                    'Gerbing 18.11.2019
        Call MediaInfo_Delete(Handle)
    End If
End Sub

Private Sub FrameNutzerDefiniert_MouseDown(Button As Integer, shift As Integer, x As Single, Y As Single)
    Dim antwort As Long
    Dim rstUserDef As ADODB.Recordset                                               'Gerbing 02.08.2016
    Dim SQL As String
    
    If Button = vbRightButton Then                                                  'Gerbing 02.08.2016
        If blnUserdefinedFieldsNotEmpty = True Then
            'antwort = MsgBox("Wollen Sie abbrechen und die Feld-Zuordnung zurücksetzen?", vbYesNo)
            antwort = MsgBox(LoadResString(2469 + Sprache), vbYesNo)
            If antwort = vbYes Then
                On Error Resume Next
                If gblnSQLServerVersion = True Then
                    'beim SQL Server muss es heißen 'Delete from table
                    SQL = "DELETE From UserDefined"
                Else
                    SQL = "DELETE * FROM UserDefined"
                End If
                Set rstUserDef = New ADODB.Recordset
                With rstUserDef
                    .Source = SQL
                    .ActiveConnection = Form1.DBsql
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .CursorLocation = adUseClient
                    .Open
                End With
                rstUserDef.Close
                'Wiederherstellen Standard-Werte
                cmbFeld1.Text = ""
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
        End If
    End If
End Sub

Private Sub FrameStandardWerte_MouseDown(Button As Integer, shift As Integer, x As Single, Y As Single)
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
'                cmbJahrEx.ListIndex = 9                             '9=Date created
'                cmbSituationEx.ListIndex = 11 + 49 + 52 + 5                   '11=Headline
'                cmbOrtEx.ListIndex = 5 + 49 + 52 + 5                         '5=City
'                cmbLandEx.ListIndex = 7 + 49 + 52 + 5                        '7=Country
'                cmbKommentarEx.ListIndex = 2 + 49 + 52 + 5                    '2=Caption
'                cmbPersonenEx.Text = ""
'                cmbSWFEx.Text = ""
                               
                cmbJahrEx.ListIndex = -1                            'Gerbing 27.10.2016
                cmbSituationEx.ListIndex = -1
                cmbOrtEx.ListIndex = -1
                cmbLandEx.ListIndex = -1
                cmbKommentarEx.ListIndex = -1
                cmbSWFEx.ListIndex = -1
            End If
        'End If
    End If
End Sub

Private Sub List1_AbortedDrag()
    List1.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, List1.ListItems(0)
End Sub

Private Sub List1_Click(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim n As Long
    Dim tempFilename As String
    Dim rc As Boolean
    
    LstU.ListItems.RemoveAll                                                        'Gerbing 21.03.2016
    txtEXIFInfo.Text = ""                                                           'Gerbing 21.03.2016
    If chkExifAnzeigen.Value <> 0 Then                                              'Gerbing 21.03.2016
        For n = 0 To List1.ListItems.Count - 1
            If List1.ListItems(n).Selected = True Then
                tempFilename = Replace(List1.ListItems(n).Text, "+:\", MyAppPath & "\")
                Form1.EXF.ImageFile = tempFilename 'set the image file property
                '
                'EXF.ListInfo ist ein String mit vbCrLf
                '
                txtEXIFInfo.Text = Form1.EXF.ListInfo 'list all tags into the text box
                If gstrLatXMP <> "" Then                                                                'Gerbing 02.10.2019
                    txtEXIFInfo.Text = txtEXIFInfo.Text & "GPSLatitude:" & gstrLatXMP & vbNewLine
                End If
                If gstrLongXMP <> "" Then                                                               'Gerbing 02.10.2019
                    txtEXIFInfo.Text = txtEXIFInfo.Text & "GPSLongitude:" & gstrLongXMP & vbNewLine
                End If
            End If
        Next n
    End If
    If chkIptcAnzeigen.Value <> 0 Then                                              'Gerbing 21.03.2016
        For n = 0 To List1.ListItems.Count - 1
            If List1.ListItems(n).Selected = True Then
                IPTCItemsDelimiter = ";"
                rc = LeseIPTC(Replace(List1.ListItems(n).Text, "+:\", MyAppPath & "\"), LstU, True)  'mit Ausgabe in LstU
            End If
        Next n
    End If
End Sub

Private Sub List1_OLEDragDrop(ByVal Data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim files() As String
    Dim n, i As Integer
    Dim KeineÜbereinstimmung As Boolean
    Dim DateinamenErweiterung As String
    Dim Msg As String
    Dim blnErweiterungGefunden As Boolean
    Dim insertAfter As Variant

    'List1 hat die Eigenschaft Sorted=True
        Screen.MousePointer = vbArrowHourglass
        files = Data.GetData(vbCFFiles)
        For n = LBound(files) To UBound(files)
        KeineÜbereinstimmung = True
        'Sämtliche Einträge in List1 prüfen, ob sie übereinstimmen mit Data.Files(n)
        If List1.ListItems.Count <> 0 Then
            KeineÜbereinstimmung = True
            'List1.ListIndex = 0
            For i = 0 To List1.ListItems.Count - 1
                'If List1.ListItems(i) = Data.files(n) Then
                If StrComp(List1.ListItems(i), files(n), vbTextCompare) = 0 Then
                    KeineÜbereinstimmung = False
                    Exit For
                End If
            Next i
        End If
        'Nur Einträge mit keineÜbereinstimmung übernehmen, damit werden Duplikate verhindert
        'Duplikate würden entstehen, wenn der Nutzer zweimal Drag&Drop mit denselben Dateinamen macht
        If KeineÜbereinstimmung = True Then
            blnErweiterungGefunden = False
            DateinamenErweiterung = Right(files(n), 3)
            DateinamenErweiterung = UCase(DateinamenErweiterung)
            Select Case DateinamenErweiterung
                '3-stellige
                'Gerbing 11.12.2005 und 09.01.2008 10.12.2017
                Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF", "AVI", "MPG", "PEG", "MOV", "MKV", "FLV", _
                    "MPE", "ASF", "ASX", "WMV", "HTM", "PDF", "XLS", _
                    "ANI", "B3D", "CAM", "CLP", "CPT", "CRW", "CR2", "DCM", "ACR", "IMA", "DCX", "DDS", _
                    "DXF", "DWG", "ECW", "EMF", "EPS", "FPX", "FSH", "ICL", _
                    "ICS", "IFF", "LBM", "IMG", "JP2", "JPC", "J2K", "JPM", "KDC", "LWF", _
                    "MNG", "JNG", "SID", "DNG", "EEF", "NEF", "MRW", "ORF", "RAF", _
                    "DCR", "SRF", "PEF", "X3F", "NLM", "NOL", "NGG", "PBM", "PCD", "PCX", "PGM", "PIC", _
                    "PNG", "PPM", "PSD", "PSP", "RAS", "SUN", "RAW", "RLE", "SFF", "SFW", "SGI", "RGB", _
                    "SWF", "TGA", "TIF", "TTF", "WAD", "WAL", "XBM", "XPM", _
                    "3FR", "ARW", "CS1", "CS4", "DCS", "ERF", "MEF", "SR2"
                    'List1.AddItem Data.files(n)
                    Set insertAfter = List1.ListItems.Add(files(n), , 1)
                    blnErweiterungGefunden = True
                Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"                'Gerbing 10.12.2017
                    'List1.AddItem Data.files(n)
                    Set insertAfter = List1.ListItems.Add(files(n), , 1)
                    blnErweiterungGefunden = True
                Case "HTM", "PDF", "XLS", "DOC"                                                                 'Gerbing 18.01.2014
                    'List1.AddItem Data.files(n)
                    Set insertAfter = List1.ListItems.Add(files(n), , 1)
                    blnErweiterungGefunden = True
            End Select
            If blnErweiterungGefunden = False Then
                'DateinamenErweiterung = Right(Data.files(n), 4)
                DateinamenErweiterung = Right(files(n), 4)
                DateinamenErweiterung = UCase(DateinamenErweiterung)
                Select Case DateinamenErweiterung
                    '4-stellige
                    'Gerbing 09.01.2008
                    Case "WBMP", "TIFF", "PICT", "QTIF", "JPEG", "FITS", "HPGL", "IW44", "DJVU", "CS16", "DOCX" 'Gerbing 19.01.2014
                        'List1.AddItem Data.files(n)
                        Set insertAfter = List1.ListItems.Add(files(n), , 1)
                        blnErweiterungGefunden = True
                End Select
            End If
            If blnErweiterungGefunden = False Then
                'DateinamenErweiterung = Right(Data.files(n), 5)
                DateinamenErweiterung = Right(files(n), 5)
                DateinamenErweiterung = UCase(DateinamenErweiterung)
                Select Case DateinamenErweiterung
                    '5-stellige
                    Case "MRSID"
                        'List1.AddItem Data.files(n)
                        Set insertAfter = List1.ListItems.Add(files(n), , 1)
                        blnErweiterungGefunden = True
                End Select
            End If
            If blnErweiterungGefunden = False Then
                'DateinamenErweiterung = Right(Data.files(n), 2)
                DateinamenErweiterung = Right(files(n), 2)
                DateinamenErweiterung = UCase(DateinamenErweiterung)
                Select Case DateinamenErweiterung
                    '2-stellige
                    Case "G3"
                        'List1.AddItem Data.files(n)
                        Set insertAfter = List1.ListItems.Add(files(n), , 1)
                        blnErweiterungGefunden = True
                End Select
            End If
            If blnErweiterungGefunden = False Then
                If DErwMsgBoxMussKommen = True Then
                    'Msg = "Nur Dateinamen-Erweiterungen "
                    '"3FR", "ARW", "CS1", "CS4", "DCS", "ERF", "MEF", "SR2", "CS16"      'Gerbing 09.01.2008
                    Msg = LoadResString(2304 + Sprache)
                    Msg = Msg & "'BMP', 'CUR', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'WMF', " & vbNewLine
                    Msg = Msg & "'AVI', 'MPG', 'MPEG', 'MOV', 'MPE', 'ASF', 'ASX', 'WMV' " & vbNewLine
                    Msg = Msg & "'HTM', 'PDF', 'XLS' " & vbNewLine

                    Msg = Msg & "'3FR', 'ACR', 'ANI', 'ARW', 'B3D', "
                    Msg = Msg & "'CAM', 'CLP', 'CPT', 'CRW', 'CR2', 'CS1', 'CS4', 'CS16', 'DCM', 'DCS', 'DCR', "
                    Msg = Msg & "'DCX', 'DDS', 'DJVU', 'DNG', "
                    Msg = Msg & "'DXF', 'DWG', 'ECW', 'EEF', 'EMF', 'EPS', 'ERF', 'FITS', 'FPX', 'FSH', 'G3', 'HPGL', 'ICL',  "
                    Msg = Msg & "'ICS', 'IFF', 'IMA', 'IMG', 'IW44', 'J2K', 'JP2', 'JNG', 'JPC', 'JPEG', "
                    Msg = Msg & "'JPM', 'KDC', 'LBM', 'LWF', "
                    Msg = Msg & "'MEF', 'MNG', 'MRW', 'MRSID' 'NEF', 'NGG', 'NLM', 'NOL', 'ORF', "
                    Msg = Msg & "'PBM', 'PCD', PCX', 'PEF', "
                    Msg = Msg & "'PGM', 'PIC', 'PICT', "
                    Msg = Msg & "'PNG', 'PPM', 'PSD', 'PSP', 'QTIF', 'RAF', 'RAS', 'RAW', 'RGB', 'RLE', 'SFF', 'SFW', "
                    Msg = Msg & "'SGI', 'SRF', 'SID', 'SUN', "
                    Msg = Msg & "'SWF', 'TGA', 'TIF', 'TIFF', 'TTF', 'WAD', 'WAL', 'WBMP', 'X3F', 'XBM', 'XPM', "
                    'Msg = Msg & "sind erlaubt." & NL
                    Msg = Msg & LoadResString(2092 + Sprache) & NL
                    'Msg = Msg & "Nur die gültigen Dateinamen-Erweiterungen werden übernommen." & NL
                    Msg = Msg & LoadResString(2306 + Sprache) & NL
                    'Msg = Msg & "Dieser Hinweis wird nicht wiederholt."
                    Msg = Msg & LoadResString(2094 + Sprache)
                    MsgBox Msg
                    DErwMsgBoxMussKommen = False
                End If
            End If
        End If
        If n > 32766 Then                                                               'Gerbing 26.01.2009
            'MsgBox "Das Programm kann in einem Durchlauf maximal 32767 Dateien aufnehmen. Wiederholen Sie die Programmfunktion für die noch nicht aufgenommenen Dateien."
            MsgBox LoadResString(2333 + Sprache)
            Exit For
        End If
    Next n
    Screen.MousePointer = vbDefault
End Sub

Private Sub List1_OLEDragOver(Data As DataObject, effect As Long, Button As Integer, shift As Integer, x As Single, Y As Single, State As Integer)
    effect = vbDropEffectCopy       'zeigt das Plus als Kopiersymbol
End Sub

Private Sub FrageNutzerJahreszahl()
    'Wenn der Nutzer bereits eine Jahreszahl festgelegt hat, die für alle Fälle gelten soll
    'steht diese im Feld NutzerJahresZahl, die Prozedur wird verlassen
    
    'Wenn der Nutzer noch keine Jahreszahl festgelegt hat, oder jedesmal eine neue Jahreszahl festlegen wollte
    'Öffnet sich das Formular 'JahrFestlegen'
    
    GefundeneJahresZahl = ""
    If GiltFürAlleFälle = True Then Exit Sub
    JahrFestlegen.Show 1
End Sub

Private Sub AddExifFelder(Control)
    Control.AddItem "EXIF-ImageDescription"                     'Gerbing 07.05.2007
    Control.AddItem "EXIF-Make"
    Control.AddItem "EXIF-Model"
    Control.AddItem "EXIF-Orientation"
    Control.AddItem "EXIF-XResolution"
    Control.AddItem "EXIF-YResolution"
    Control.AddItem "EXIF-ResolutionUnit"
    Control.AddItem "EXIF-Software"                             'Gerbing 07.05.2007
    Control.AddItem "EXIF-DateTime"
    Control.AddItem "EXIF-Artist"
    Control.AddItem "EXIF-WhitePoint"
    Control.AddItem "EXIF-PrimaryChromaticities"
    Control.AddItem "EXIF-YCbCrCoefficients"
    Control.AddItem "EXIF-YCbCrPositioning"
    Control.AddItem "EXIF-ReferenceBlackWhite"
    Control.AddItem "EXIF-Copyright"                            'Gerbing 07.05.2007
    Control.AddItem "EXIF-EXIFOffset"
    Control.AddItem "EXIF-Compression"
    Control.AddItem "EXIF-JPEGInterchangeFormatOffset"
    Control.AddItem "EXIF-JPEGInterchangeFormatLength"
    Control.AddItem "EXIF-ExposureTime"
    Control.AddItem "EXIF-FNumber"
    Control.AddItem "EXIF-ExposureProgram"                      'Gerbing 07.05.2007
    Control.AddItem "EXIF-ISO"                                  'Gerbing 07.05.2007
    Control.AddItem "EXIF-EXIFVersion"
    Control.AddItem "EXIF-DateTimeOriginal"
    Control.AddItem "EXIF-DateTimeDigitized"
    Control.AddItem "EXIF-ComponentsConfiguration"
    Control.AddItem "EXIF-CompressedBitsPerPixel"
    Control.AddItem "EXIF-ShutterSpeedValue"
    Control.AddItem "EXIF-ApertureValue"
    Control.AddItem "EXIF-BrightnessValue"                      'Gerbing 07.05.2007
    Control.AddItem "EXIF-ExposureBiasValue"
    Control.AddItem "EXIF-MaxApertureValue"
    Control.AddItem "EXIF-SubjectDistance"                      'Gerbing 07.05.2007
    Control.AddItem "EXIF-MeteringMode"
    Control.AddItem "EXIF-LightSource"                          'Gerbing 07.05.2007
    Control.AddItem "EXIF-Flash"
    Control.AddItem "EXIF-FocalLength"
    Control.AddItem "EXIF-MakerNote"
    Control.AddItem "EXIF-UserComment"                          'Gerbing 07.05.2007
    Control.AddItem "EXIF-SubsecTime"
    Control.AddItem "EXIF-SubsecTimeOriginal"
    Control.AddItem "EXIF-SubsecTimeDigitized"
    Control.AddItem "EXIF-FlashPixVersion"
    Control.AddItem "EXIF-ColorSpace"
    Control.AddItem "EXIF-EXIFImageWidth"
    Control.AddItem "EXIF-EXIFImageHeight"
    Control.AddItem "EXIF-RelatedSoundFile"                     'Gerbing 07.05.2007
    Control.AddItem "EXIF-InteroperatibilityOffset"
    Control.AddItem "EXIF-ExposureIndex"                        'Gerbing 07.05.2007
    Control.AddItem "EXIF-FocalPlaneXResolution"
    Control.AddItem "EXIF-FocalPlaneYResolution"
    Control.AddItem "EXIF-FocalPlaneResolutionUnit"
    Control.AddItem "EXIF-ExposureIndex"
    Control.AddItem "EXIF-SensingMethod"
    Control.AddItem "EXIF-FileSource"
    Control.AddItem "EXIF-SceneType"                            'Gerbing 07.05.2007
    Control.AddItem "EXIF-CFAPattern"
    Control.AddItem "EXIF-CostumRendered"
    Control.AddItem "EXIF-ExposureMode"
    Control.AddItem "EXIF-WhiteBalance"
    Control.AddItem "EXIF-DigitalZoomRatio"
    Control.AddItem "EXIF-SceneCaptureType"
    Control.AddItem "EXIF-GainControl"
    Control.AddItem "EXIF-Contrast"
    Control.AddItem "EXIF-Saturation"
    Control.AddItem "EXIF-Sharpness"
    Control.AddItem "EXIF-SubjectDistanceRange"
'-----------------------------------------------------------------------------------
'   ab GERBING Fotoalbum 14
    Control.AddItem "EXIF- "
    Control.AddItem "EXIF-GPSLatitudeRef"
    Control.AddItem "EXIF-GPSLatitude"
    Control.AddItem "EXIF-GPSLongitudeRef"
    Control.AddItem "EXIF-GPSLongitude"
    Control.AddItem "EXIF-GPSAltitudeRef"
    Control.AddItem "EXIF-GPSAltitude"
    Control.AddItem "EXIF-GPSTimeStamp"
    Control.AddItem "EXIF-GPSSatellites"
    Control.AddItem "EXIF-GPSStatus"
    Control.AddItem "EXIF-GPSMeasureMode"
    Control.AddItem "EXIF-GPSDOP"
    Control.AddItem "EXIF-GPSSpeedRef"
    Control.AddItem "EXIF-GPSSpeed"
    Control.AddItem "EXIF-GPSTrackRef"
    Control.AddItem "EXIF-GPSTrack"
    Control.AddItem "EXIF-GPSImgDirectionRef"
    Control.AddItem "EXIF-GPSImgDirection"
    Control.AddItem "EXIF-GPSMapDatum"
    Control.AddItem "EXIF-GPSDestLatitudeRef"
    Control.AddItem "EXIF-GPSDestLatitude"
    Control.AddItem "EXIF-GPSDestLongitudeRef"
    Control.AddItem "EXIF-GPSDestLongitude"
    Control.AddItem "EXIF-GPSDestBearingRef"
    Control.AddItem "EXIF-GPSDestBearing"
    Control.AddItem "EXIF-GPSDestDistanceRef"
    Control.AddItem "EXIF-GPSDestDistance"
    Control.AddItem "EXIF-GPSProcessingMethod"
    Control.AddItem "EXIF-GPSAreaInformation"
    Control.AddItem "EXIF-GPSDateStamp"
    Control.AddItem "EXIF-GPSDifferential"
    Control.AddItem "EXIF-GPSHPositioningError"
'-----------------------------------------------------------------------------------
    '   ab GERBING Fotoalbum 14.0.1                          'Gerbing 02.010.2013
    Control.AddItem "EXIF-XPTitle"
    Control.AddItem "EXIF-XPSubject"
    Control.AddItem "EXIF-XPKeywords"
    Control.AddItem "EXIF-XPComment"
    Control.AddItem "EXIF-XPAuthor"
'-----------------------------------------------------------------------------------
    Control.AddItem "IPTC-Object Name"
    Control.AddItem "IPTC-Byline"
    Control.AddItem "IPTC-Byline title"
    Control.AddItem "IPTC-Caption"
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
    Control.AddItem "IPTC-Credits"                              'Gerbing 10.01.2008
    Control.AddItem "IPTC-Originating Program"                  'Gerbing 10.01.2008
    '-----------------------------------------
    'jetzt die restlichen 8 IPTC-Felder                         'Gerbing 04.03.2013
    '-----------------------------------------
    Control.AddItem "IPTC-Release date"
    Control.AddItem "IPTC-Release time"
    Control.AddItem "IPTC-Object cycle"
    Control.AddItem "IPTC-Location code"
    Control.AddItem "IPTC-Sublocation"
    Control.AddItem "IPTC-Program version"
    Control.AddItem "IPTC-Edit status"
    Control.AddItem "IPTC-JobID"
    '-----------------------------------------------
End Sub

Private Sub SchreibePruefLog3Einig()
    Dim Msg As String
    
    Form1.FehlerGefunden = True
    'Msg = "Verstoß gegen die 3-Einigkeit-" & AktuellerDateiName
    Msg = LoadResString(1499 + Sprache) & AktuellerDateiname
    oStream.WriteLine Msg
End Sub

Private Sub SchreibePruefLogDuplikat()
    Dim Msg As String
    
    Form1.FehlerGefunden = True
    'Msg = "Abgelehnt-Mit diesem Dateiname würde ein Duplikat entstehen-" & AktuellerDateiName
    Msg = LoadResString(1500 + Sprache) & AktuellerDateiname
    oStream.WriteLine Msg
End Sub

Private Sub SchreibePruefLogKeinPunkt()                                     'Gerbing 09.09.2014
    Dim Msg As String
    
    Form1.FehlerGefunden = True
'    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
    Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine & "----------------------------"
    oStream.WriteLine Msg
End Sub

Private Sub SchreibePruefLogHochkommaErsetzt()                              'Gerbing 03.10.2017
    Dim Msg As String
    
    Form1.FehlerGefunden = True
'    Msg = "Hochkomma(') wurde durch - ersetzt " & AktuellerDateiName & vbNewLine
    Msg = LoadResString(2470 + Sprache) & AktuellerDateiname & vbNewLine & "----------------------------"
    oStream.WriteLine Msg
End Sub

Private Sub SchreibePruefLogVerbotenerDateiname()                           'Gerbing 09.09.2014
    Dim Msg As String
    
    Form1.FehlerGefunden = True
'    Msg = "Fehler im Dateiname " & AktuellerDateiName & vbNewLine
'    Msg = Msg & "error number=" & ERR.Number & vbNewLine
'    Msg = Msg & "error text=" & ERR.Description & vbNewLine
'    Msg = Msg & "Möglicherweise verbotene Zeichen im Dateiname" & vbNewLine
'    Msg = Msg & "oder ungültiges Datei-Datum"
    Msg = LoadResString(2455 + Sprache) & AktuellerDateiname & vbNewLine
    Msg = Msg & LoadResString(2456 + Sprache) & vbNewLine
    Msg = Msg & LoadResString(2464 + Sprache) & vbNewLine & "----------------------------"
    'Print #Form1.DateiNummer, Msg
    oStream.WriteLine Msg
End Sub

Private Sub SchreibePruefLogTypFehler(Feldname, Feldwert)
    Dim Msg As String
    
    Form1.FehlerGefunden = True
    'msg = "Datentypfehler beim Feldname " & FeldName & " Feldwert " & Feldwert & " " & AktuellerDateiName
    Msg = LoadResString(1505 + Sprache) & Feldname & LoadResString(1506 + Sprache) & "'" & Feldwert & "'" & " " & AktuellerDateiname    'Gerbing 04.01.2009
    oStream.WriteLine Msg
End Sub

Private Sub SchreibeSWFFehler(Feldname, Feldwert)                                           'Gerbing 04.01.2009
    Dim Msg As String
    
    Form1.FehlerGefunden = True
    'msg = "Falscher Inhalt beim Feldname " & FeldName & " Feldwert " & Feldwert & " " & AktuellerDateiName
    Msg = LoadResString(1555 + Sprache) & Feldname & LoadResString(1506 + Sprache) & Feldwert & " " & AktuellerDateiname
    oStream.WriteLine Msg
End Sub

Private Sub EXIFJahrEintragen()
    Dim ExifFeld As String
    Dim Pos As Long
    Dim pos1 As Long
    Dim rc As Boolean

    If left(cmbJahrEx.Text, 4) = "EXIF" Then
        ExifFeld = Mid(cmbJahrEx.Text, 6, Len(cmbJahrEx.Text) - 5)
        If EXIFListInfo = "" Then                               'Gerbing 12.10.2016
            Form1.EXF.ImageFile = AktuellerDateiname            'Gerbing 07.05.2007
            EXIFListInfo = Form1.EXF.ListInfo
        End If
        'Wenn DateTime nicht gefunden wird, dann wird auch zB DateTimeOriginal akzeptiert
        Pos = InStr(1, EXIFListInfo, ExifFeld, vbTextCompare)
        If Pos <> 0 Then
            pos1 = InStr(Pos, EXIFListInfo, ":", vbTextCompare)
            If pos1 <> 0 Then
                GefundeneJahresZahl = Mid(EXIFListInfo, pos1 + 2, 4)
            End If
        End If
    Else                                                        'Gerbing 23.01.2008
        IPTCItemsDelimiter = ";"
        rc = LeseIPTC(AktuellerDateiname, LstU, False)               'ohne Ausgabe in LstU
        GefundeneJahresZahl = left(iptc.DateCreated, 4)
    End If
End Sub

Private Sub FindeExifIptc(ExFeldText)
    Dim ExifIptcFeld As String
    Dim Pos As Long
    Dim posA As Long
    Dim pos1 As Long
    Dim pos1A As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim rc As Boolean

    ExifIptcFeld = Mid(ExFeldText, 6, Len(ExFeldText) - 5)
    'ExifIptcFeld = ExifIptcFeld & ":"
    If left(ExFeldText, 4) = "EXIF" Then
        'hier ist es ein EXIF-Feld
        ExifIptcFeld = ExifIptcFeld & ":"                       'Gerbing 24.07.2013
        If EXIFListInfo = "" Then                               'Gerbing 12.10.2016
            Form1.EXF.ImageFile = AktuellerDateiname            'Gerbing 07.05.2007
            EXIFListInfo = Form1.EXF.ListInfo
        End If
        'Gerbing                                                'Gerbing 12.10.2016
        'Wenn in EXIFListInfo GPSLatitude: und GPSLongitude: auftauchen, dann habe ich hier die Gelegenheit ein zusätzliche Minus
        'davorzuschreiben. Minus wird geschrieben, wenn 'GPSLatitudeRef: S' gefunden wird
        'oder wenn 'GPSLongitude: W' gefunden wird
        If blnMinusGeprüft = False Then
            Pos = InStr(1, EXIFListInfo, "GPSLatitude:", vbTextCompare)
            If Pos <> 0 Then
                posA = InStr(Pos, EXIFListInfo, ":", vbTextCompare)
                If Pos <> 0 Then
                    pos2 = InStr(1, EXIFListInfo, "GPSLatitudeRef: S", vbTextCompare)
                    If pos2 <> 0 Then
                        pos3 = InStr(Pos, EXIFListInfo, vbNewLine)
                        EXIFListInfo = left(EXIFListInfo, Pos) & "GPSLatitude: -" & Mid(EXIFListInfo, posA + 2, Len(EXIFListInfo) - posA + 2)
                    End If
                End If
            End If
            pos1 = InStr(1, EXIFListInfo, "GPSLongitude:", vbTextCompare)
            If pos1 <> 0 Then
                pos1A = InStr(pos1, EXIFListInfo, ":", vbTextCompare)
                If pos1 <> 0 Then
                    pos2 = InStr(1, EXIFListInfo, "GPSLongitudeRef: W", vbTextCompare)
                    If pos2 <> 0 Then
                        pos3 = InStr(pos1, EXIFListInfo, vbNewLine)
                        EXIFListInfo = left(EXIFListInfo, pos1 - 1) & "GPSLongitude: -" & Mid(EXIFListInfo, pos1A + 2, Len(EXIFListInfo) - pos1A + 2)
                    End If
                End If
            End If
            blnMinusGeprüft = True
        End If
        '-------------------------------------------------------------------------------------------------------------------------------
        Pos = InStr(1, EXIFListInfo, ExifIptcFeld, vbTextCompare)
        If Pos <> 0 Then
            pos1 = InStr(Pos, EXIFListInfo, ":", vbTextCompare)
            If pos1 <> 0 Then
                pos2 = InStr(Pos, EXIFListInfo, vbCrLf, vbTextCompare)
                If pos2 <> 0 Then
                    strGefundenExifIptc = Mid(EXIFListInfo, pos1 + 2, pos2 - pos1 - 2)
                End If
            End If
        End If
    Else
        'hier ist es ein IPTC-Feld
        Call IptcStringHolen(ExifIptcFeld)
        If strGefundenExifIptc <> "" Then                       'Gerbing 04.02.2008
            If rc = True Then
                blnIPTCVorhanden = True
            End If
        End If
    End If
End Sub

Private Sub IptcStringHolen(ExifIptcFeld As String)
    If ExifIptcFeld = "Object name" Then strGefundenExifIptc = iptc.ObjectName
    If ExifIptcFeld = "Byline" Then strGefundenExifIptc = iptc.Byline
    If ExifIptcFeld = "Byline title" Then strGefundenExifIptc = iptc.BylineTitle
    If ExifIptcFeld = "Caption" Then strGefundenExifIptc = iptc.Caption
    If ExifIptcFeld = "Caption writer" Then strGefundenExifIptc = iptc.CaptionWriter
    If ExifIptcFeld = "Copyright notice" Then strGefundenExifIptc = iptc.Copyright
    If ExifIptcFeld = "Special instructions" Then strGefundenExifIptc = iptc.SpecialInstructions
    If ExifIptcFeld = "Urgency" Then strGefundenExifIptc = iptc.Urgency
    If ExifIptcFeld = "Date created" Then strGefundenExifIptc = iptc.DateCreated
    If ExifIptcFeld = "Time created" Then strGefundenExifIptc = iptc.TimeCreated
    If ExifIptcFeld = "City" Then strGefundenExifIptc = iptc.City
    If ExifIptcFeld = "Province/State" Then strGefundenExifIptc = iptc.ProvinceState
    If ExifIptcFeld = "Country" Then strGefundenExifIptc = iptc.Country
    If ExifIptcFeld = "Credits" Then strGefundenExifIptc = iptc.Credits                            'Gerbing 04.01.2009
    If ExifIptcFeld = "Source" Then strGefundenExifIptc = iptc.Source
    If ExifIptcFeld = "Headline" Then strGefundenExifIptc = iptc.Headline
    If ExifIptcFeld = "Transmission reference" Then strGefundenExifIptc = iptc.OriginalTransmissionReference
    If ExifIptcFeld = "Category" Then strGefundenExifIptc = iptc.Category
    If ExifIptcFeld = "Supplemental categories" Then strGefundenExifIptc = iptc.SupplementalCategories
    If ExifIptcFeld = "Keywords" Then strGefundenExifIptc = iptc.Keywords
    If ExifIptcFeld = "Originating Program" Then strGefundenExifIptc = iptc.OriginatingProgram  'Gerbing 10.01.2008
    '--------------------------------------------------------------------------------------------------------------
    'jetzt die anderen 8 Felder
    '--------------------------------------------------------------------------------------------------------------
    If ExifIptcFeld = "Release date" Then strGefundenExifIptc = iptc.ReleaseDate
    If ExifIptcFeld = "Release time" Then strGefundenExifIptc = iptc.ReleaseTime
    If ExifIptcFeld = "Object cycle" Then strGefundenExifIptc = iptc.Objectcycle
    If ExifIptcFeld = "Location code" Then strGefundenExifIptc = iptc.LocationCode
    If ExifIptcFeld = "Sublocation" Then strGefundenExifIptc = iptc.SubLocation
    If ExifIptcFeld = "Program version" Then strGefundenExifIptc = iptc.ProgramVersion
    If ExifIptcFeld = "Edit status" Then strGefundenExifIptc = iptc.EditStatus
    If ExifIptcFeld = "JobID" Then strGefundenExifIptc = iptc.JobId
End Sub

Private Sub TypeError(Feldname, Feldwert)
    Dim Msg As String
    Dim antwort As Long
    
    If ERR.Number <> 0 Then
        If optProtokolldatei.Value = True Then
            Call SchreibePruefLogTypFehler(Feldname, Feldwert)
        Else
'            msg = "Datentypfehler beim Feldname " & cmbFeld1.Text & " Feldwert " & Combo1 & " " & AktuellerDateiName & NL
'            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
            Msg = LoadResString(1505 + Sprache) & Feldname & LoadResString(1506 + Sprache) & Feldwert & " " & AktuellerDateiname & NL
            Msg = Msg & LoadResString(1507 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
            If antwort = vbYes Then
                optProtokolldatei.Value = True
                Call SchreibePruefLogTypFehler(Feldname, Feldwert)                  'Gerbing 05.08.2016
            End If
        End If
    End If
End Sub

Private Sub FeldEintragenMitLängenprüfung(rst, Zielfeld, Quellwert)
    Dim Msg As String
    Dim antwort As Long
    
    On Error Resume Next
    rst.Fields(Zielfeld) = Quellwert
    If ERR.Number <> 0 Then
        If optProtokolldatei.Value = True Then
            Call SchreibePruefLogTypFehler(Zielfeld, Quellwert)
        Else
'            msg = "Feldlängenfehler beim Feldname " & Zielfeld & " Feldwert " & Quellwert & " " & AktuellerDateiName & NL
'            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
            Msg = LoadResString(1521 + Sprache) & Zielfeld & LoadResString(1506 + Sprache) & Quellwert & NL & AktuellerDateiname & NL
            Msg = Msg & LoadResString(1507 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
            If antwort = vbYes Then
                optProtokolldatei.Value = True
                Call SchreibePruefLogTypFehler(Zielfeld, Quellwert)                         'Gerbing 05.08.2016
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub SWFFehler(rst, Zielfeld, Quellwert)                                             'Gerbing 04.01.2009
    Dim Msg As String
    Dim antwort As Long
    
    On Error Resume Next
    rst.Fields(Zielfeld) = Quellwert
    If ERR.Number <> 0 Then
        If optProtokolldatei.Value = True Then
            Call SchreibeSWFFehler(Zielfeld, Quellwert)
        Else
'            msg = "Falscher Inhalt beim Feldname " & Zielfeld & " Feldwert " & Quellwert & " " & AktuellerDateiName & NL
'            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
            Msg = LoadResString(1555 + Sprache) & Zielfeld & LoadResString(1506 + Sprache) & Quellwert & NL & AktuellerDateiname & NL
            Msg = Msg & LoadResString(1507 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
            If antwort = vbYes Then
                optProtokolldatei.Value = True
                Call SchreibeSWFFehler(Zielfeld, Quellwert)                                 'Gerbing 05.08.2016
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub List1_KeyDown(keyCode As Integer, ByVal shift As Integer)
    If keyCode = KeyCodeConstants.vbKeyEscape Then
        If Not (List1.DraggedItems Is Nothing) Then List1.EndDrag True
    End If
End Sub

Private Sub FülleComboBoxen()
    Dim SQL As String

    'für alle Felder wird eine Combobox mit den schon vorhandenen Werten angeboten  Gerbing 10.06.2005
    'SQL = "SELECT DISTINCT Fotos.Situation From Fotos WHERE ((Not (Fotos.Situation)='')) ORDER BY Situation;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1024 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1024 + Sprache) & ")='')) ORDER BY " & LoadResString(1024 + Sprache) & ";"
    On Error Resume Next
    Form1.rstsql.Close
    On Error GoTo 0
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    Do Until Form1.rstsql.EOF
        'If Not IsNull(Form1.rstsql.Fields("Situation")) Then
        If Not IsNull(Form1.rstsql.Fields(LoadResString(1024 + Sprache))) Then
            cmbSituation.ComboItems.Add Form1.rstsql.Fields(LoadResString(1024 + Sprache))
            'rc = MessageBox(Form1.rstsql.Fields(LoadResString(1024 + Sprache)), "Caption", vbOKOnly)
        End If
        Form1.rstsql.Movenext
        DoEvents
    Loop
    Form1.rstsql.Close                                                                       'Gerbing 29.05.2008
    'SQL = "SELECT DISTINCT Fotos.Ort From Fotos WHERE ((Not (Fotos.Ort)='')) ORDER BY Ort;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1025 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1025 + Sprache) & ")='')) ORDER BY " & LoadResString(1025 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
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
    Form1.rstsql.Close                                                                       'Gerbing 29.05.2008
    'SQL = "SELECT DISTINCT Fotos.Land From Fotos WHERE ((Not (Fotos.Land)='')) ORDER BY Land;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1026 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1026 + Sprache) & ")='')) ORDER BY " & LoadResString(1026 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
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
    Form1.rstsql.Close                                                                          'Gerbing 29.05.2008
    'SQL = "SELECT DISTINCT Fotos.Personen From Fotos WHERE ((Not (Fotos.Personen)='')) ORDER BY Personen;"
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1027 + Sprache) & " From Fotos WHERE ((Not (Fotos." & LoadResString(1027 + Sprache) & ")='')) ORDER BY " & LoadResString(1027 + Sprache) & ";"
    With Form1.rstsql
        .Source = SQL
        .ActiveConnection = Form1.DBsql
        '.CursorType = adOpenStatic
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
    Form1.rstsql.Close                                                                          'Gerbing 29.05.2008
End Sub

Private Function PrüfePunktVorDateinamenErweiterung(AktuellerDateiname)                         'Gerbing 09.09.2014
    Dim strPunkt As String
    Dim DateinamenErweiterung As String
    Dim blnErweiterungGefunden As Boolean
    Dim Msg As String
    Dim antwort As Long
    
    DateinamenErweiterung = Right(AktuellerDateiname, 3)
    DateinamenErweiterung = UCase(DateinamenErweiterung)
    PrüfePunktVorDateinamenErweiterung = 1                              'rc=1 Voreinstellung falls Fehler
    Select Case DateinamenErweiterung
        '3-stellige
        'Gerbing 11.12.2005 und 09.01.2008 10.12.2017
        Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF", "AVI", "MPG", "PEG", "MOV", "MKV", "FLV", _
            "MPE", "ASF", "ASX", "WMV", "HTM", "PDF", "XLS", _
            "ANI", "B3D", "CAM", "CLP", "CPT", "CRW", "CR2", "DCM", "ACR", "IMA", "DCX", "DDS", _
            "DXF", "DWG", "ECW", "EMF", "EPS", "FPX", "FSH", "ICL", _
            "ICS", "IFF", "LBM", "IMG", "JP2", "JPC", "J2K", "JPM", "KDC", "LWF", _
            "MNG", "JNG", "SID", "DNG", "EEF", "NEF", "MRW", "ORF", "RAF", _
            "DCR", "SRF", "PEF", "X3F", "NLM", "NOL", "NGG", "PBM", "PCD", "PCX", "PGM", "PIC", _
            "PNG", "PPM", "PSD", "PSP", "RAS", "SUN", "RAW", "RLE", "SFF", "SFW", "SGI", "RGB", _
            "SWF", "TGA", "TIF", "TTF", "WAD", "WAL", "XBM", "XPM", _
            "3FR", "ARW", "CS1", "CS4", "DCS", "ERF", "MEF", "SR2"
            strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 3, 1)                                              'Gerbing 09.09.2014
            If strPunkt = "." Then
                blnErweiterungGefunden = True
            Else
                If optProtokolldatei.Value = True Then
                    Call SchreibePruefLogKeinPunkt
                Else
                '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & vbNewLine & AktuellerDateiName & vbNewLine
                    Msg = LoadResString(2468 + Sprache) & vbNewLine & AktuellerDateiname & vbNewLine
        '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                    Msg = Msg & LoadResString(1507 + Sprache)
                    'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                    antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                    If antwort = vbYes Then
                        optProtokolldatei.Value = True
                        Call SchreibePruefLogKeinPunkt
                    End If
                End If
                Exit Function
            End If
        Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"                                    'Gerbing 10.12.2017
            strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 3, 1)                                              'Gerbing 09.09.2014
            If strPunkt = "." Then
                blnErweiterungGefunden = True
            Else
                If optProtokolldatei.Value = True Then
                    Call SchreibePruefLogKeinPunkt
                Else
                '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
                    Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine
        '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                    Msg = Msg & LoadResString(1507 + Sprache)
                    'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                    antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                    If antwort = vbYes Then
                        optProtokolldatei.Value = True
                        Call SchreibePruefLogKeinPunkt
                    End If
                End If
                Exit Function
            End If
            blnErweiterungGefunden = True
        Case "HTM", "PDF", "XLS", "DOC"                                                                 'Gerbing 18.01.2014
            strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 3, 1)                                              'Gerbing 09.09.2014
            If strPunkt = "." Then
                blnErweiterungGefunden = True
            Else
                If optProtokolldatei.Value = True Then
                    Call SchreibePruefLogKeinPunkt
                Else
                '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
                    Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine
        '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                    Msg = Msg & LoadResString(1507 + Sprache)
                    'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                    antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                    If antwort = vbYes Then
                        optProtokolldatei.Value = True
                        Call SchreibePruefLogKeinPunkt
                    End If
                End If
                Exit Function
            End If
            blnErweiterungGefunden = True
    End Select
    If blnErweiterungGefunden = False Then
    DateinamenErweiterung = Right(AktuellerDateiname, 4)
    DateinamenErweiterung = UCase(DateinamenErweiterung)
        Select Case DateinamenErweiterung
            '4-stellige
            'Gerbing 09.01.2008
            Case "WBMP", "TIFF", "PICT", "QTIF", "JPEG", "FITS", "HPGL", "IW44", "DJVU", "CS16", "DOCX" 'Gerbing 19.01.2014
                strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 4, 1)                                              'Gerbing 09.09.2014
                If strPunkt = "." Then
                    blnErweiterungGefunden = True
                Else
                    If optProtokolldatei.Value = True Then
                        Call SchreibePruefLogKeinPunkt
                    Else
                    '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
                        Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine
            '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                        Msg = Msg & LoadResString(1507 + Sprache)
                        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                        If antwort = vbYes Then
                            optProtokolldatei.Value = True
                            Call SchreibePruefLogKeinPunkt
                        End If
                    End If
                    Exit Function
                End If
                blnErweiterungGefunden = True
        End Select
    End If
    If blnErweiterungGefunden = False Then
        DateinamenErweiterung = Right(AktuellerDateiname, 5)
        DateinamenErweiterung = UCase(DateinamenErweiterung)
        Select Case DateinamenErweiterung
            '5-stellige
            Case "MRSID"
                strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 5, 1)                                              'Gerbing 09.09.2014
                If strPunkt = "." Then
                    blnErweiterungGefunden = True
                Else
                    If optProtokolldatei.Value = True Then
                        Call SchreibePruefLogKeinPunkt
                    Else
                    '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
                        Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine
            '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                        Msg = Msg & LoadResString(1507 + Sprache)
                        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                        If antwort = vbYes Then
                            optProtokolldatei.Value = True
                            Call SchreibePruefLogKeinPunkt
                        End If
                    End If
                    Exit Function
                End If
                blnErweiterungGefunden = True
        End Select
    End If
    If blnErweiterungGefunden = False Then
        DateinamenErweiterung = Right(AktuellerDateiname, 2)
        DateinamenErweiterung = UCase(DateinamenErweiterung)
        Select Case DateinamenErweiterung
            '2-stellige
            Case "G3"
                strPunkt = Mid(AktuellerDateiname, Len(AktuellerDateiname) - 5, 1)                                              'Gerbing 09.09.2014
                If strPunkt = "." Then
                    blnErweiterungGefunden = True
                Else
                    If optProtokolldatei.Value = True Then
                        Call SchreibePruefLogKeinPunkt
                    Else
                    '    Msg = "Kein Punkt vor der Dateinamen-Erweiterung " & AktuellerDateiName & vbNewLine
                        Msg = LoadResString(2468 + Sprache) & AktuellerDateiname & vbNewLine
            '            msg = msg & "Sollen weitere Fehlermeldungen in die Protokolldatei (pruef.log) geschrieben werden?"
                        Msg = Msg & LoadResString(1507 + Sprache)
                        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
                        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
                        If antwort = vbYes Then
                            optProtokolldatei.Value = True
                            Call SchreibePruefLogKeinPunkt
                        End If
                    End If
                    Exit Function
                End If
                blnErweiterungGefunden = True
        End Select
    End If
    PrüfePunktVorDateinamenErweiterung = 0                                              'rc=0 kein Fehler
End Function

Private Sub GetMediaInfo()                                                      'Gerbing 18.11.2019
Dim display As String
    Dim GPS As String
    Dim strLat As String
    Dim strLong As String
    Dim InfoMediadll As String
    Dim strDateTime As String
    Dim Pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim rc As Long
    Const SW_SHOWNORMAL = 1
    
    rc = MediaInfo_Open(Handle, StrPtr(AktuellerDateiname))
    display = InfoMediadll
    Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr(""))
    display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
    Call MediaInfo_Close(Handle)
    
    'Die Felder GPSLatitude und GPSLongitude auffüllen
    'xyz Suchen                                                                'Gerbing 18.11.2019
    Pos = InStr(1, display, "xyz")
    If Pos <> 0 Then
        pos1 = InStr(Pos, display, ":")
        pos2 = InStr(pos1, display, "/")
        pos3 = InStr(pos1 + 3, display, "+")
        If pos3 = 0 Then
            pos3 = InStr(pos1 + 3, display, "-")
        End If
        GPS = Mid(display, pos1 + 2, pos2 - pos1 - 2)
        'MsgBox GFfPS
        'zB GPS = "+50.8314+12.8311"
        strLat = Mid(GPS, 1, pos3 - pos1 - 2)
        strLong = Mid(GPS, Len(strLat) + 1, pos2 - pos3)
        rstsql.Fields("GPSLatitude") = strLat
        rstsql.Fields("GPSLongitude") = strLong
    End If
    '---------------------------------------------------------------------------------------------------------------------
    Pos = InStr(1, display, "Encoded Date", vbTextCompare)                      'das erste "Encoded Date" ist das richtige
    If Pos <> 0 Then
        pos1 = InStr(Pos, display, ": UTC")
        If Pos <> 0 Then
            strDateTime = Mid(display, pos1 + 6, 19)
            strDateTime = Replace(strDateTime, "-", ":")
            rstsql.Fields("ExifDateTimeOriginal") = strDateTime
        End If
    End If
    display = Empty
    Exit Sub
End Sub
