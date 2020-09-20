VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form F5MehrereZeilen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Angaben zum aktuellen Bild"
   ClientHeight    =   10524
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   13932
   Icon            =   "MehrereZeilen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10524
   ScaleWidth      =   13932
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkIptc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "IPTC-Felder"
      Height          =   372
      Left            =   11040
      TabIndex        =   22
      Top             =   360
      Width           =   2772
   End
   Begin VB.CheckBox chkExif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIF-Felder"
      Height          =   372
      Left            =   11040
      TabIndex        =   21
      Top             =   0
      Width           =   2772
   End
   Begin VB.TextBox txtHoehePixel 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox txtBreitePixel 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox txtDDatum 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox txtSWF 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtJahr 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   600
      Width           =   732
   End
   Begin VB.Frame FrameNutzer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nutzerdefinierte Felder"
      Height          =   3132
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   10812
      Begin EditCtlsLibUCtl.TextBox txtFeldname5 
         Height          =   372
         Left            =   240
         TabIndex        =   38
         Top             =   2520
         Width           =   2892
         _cx             =   5101
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
         CueBanner       =   "MehrereZeilen.frx":038A
         Text            =   "MehrereZeilen.frx":03AA
      End
      Begin EditCtlsLibUCtl.TextBox txtFeldname4 
         Height          =   372
         Left            =   240
         TabIndex        =   37
         Top             =   2040
         Width           =   2892
         _cx             =   5101
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
         CueBanner       =   "MehrereZeilen.frx":03CA
         Text            =   "MehrereZeilen.frx":03EA
      End
      Begin EditCtlsLibUCtl.TextBox txtFeldname3 
         Height          =   372
         Left            =   240
         TabIndex        =   36
         Top             =   1560
         Width           =   2892
         _cx             =   5101
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
         CueBanner       =   "MehrereZeilen.frx":040A
         Text            =   "MehrereZeilen.frx":042A
      End
      Begin EditCtlsLibUCtl.TextBox txtFeldname2 
         Height          =   372
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   2892
         _cx             =   5101
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
         CueBanner       =   "MehrereZeilen.frx":044A
         Text            =   "MehrereZeilen.frx":046A
      End
      Begin EditCtlsLibUCtl.TextBox txtFeldname1 
         Height          =   372
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   2892
         _cx             =   5101
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
         CueBanner       =   "MehrereZeilen.frx":048A
         Text            =   "MehrereZeilen.frx":04AA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld5 
         Height          =   288
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   2892
         _cx             =   5101
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
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":04CA
         Text            =   "MehrereZeilen.frx":04EA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld4 
         Height          =   288
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Width           =   2892
         _cx             =   5101
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
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":050A
         Text            =   "MehrereZeilen.frx":052A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld3 
         Height          =   288
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   2892
         _cx             =   5101
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
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":054A
         Text            =   "MehrereZeilen.frx":056A
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld2 
         Height          =   288
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   2892
         _cx             =   5101
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
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":058A
         Text            =   "MehrereZeilen.frx":05AA
      End
      Begin CBLCtlsLibUCtl.ComboBox cmbFeld1 
         Height          =   288
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   2892
         _cx             =   5101
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
         Style           =   1
         SupportOLEDragImages=   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":05CA
         Text            =   "MehrereZeilen.frx":05EA
      End
      Begin EditCtlsLibUCtl.TextBox txtFeld1 
         Height          =   372
         Left            =   3360
         TabIndex        =   39
         Top             =   600
         Width           =   7332
         _cx             =   12933
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
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   1
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":060A
         Text            =   "MehrereZeilen.frx":062A
      End
      Begin EditCtlsLibUCtl.TextBox txtFeld2 
         Height          =   372
         Left            =   3360
         TabIndex        =   40
         Top             =   1080
         Width           =   7332
         _cx             =   12933
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
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   1
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":064E
         Text            =   "MehrereZeilen.frx":066E
      End
      Begin EditCtlsLibUCtl.TextBox txtFeld3 
         Height          =   372
         Left            =   3360
         TabIndex        =   41
         Top             =   1560
         Width           =   7332
         _cx             =   12933
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
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   1
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":068E
         Text            =   "MehrereZeilen.frx":06AE
      End
      Begin EditCtlsLibUCtl.TextBox txtFeld4 
         Height          =   372
         Left            =   3360
         TabIndex        =   42
         Top             =   2040
         Width           =   7332
         _cx             =   12933
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
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   1
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":06CE
         Text            =   "MehrereZeilen.frx":06EE
      End
      Begin EditCtlsLibUCtl.TextBox txtFeld5 
         Height          =   372
         Left            =   3360
         TabIndex        =   43
         Top             =   2520
         Width           =   7332
         _cx             =   12933
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
         MultiLine       =   -1  'True
         OLEDragImageStyle=   0
         PasswordChar    =   0
         ProcessContextMenuKeys=   -1  'True
         ReadOnly        =   -1  'True
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   1
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "MehrereZeilen.frx":070E
         Text            =   "MehrereZeilen.frx":072E
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldinhalt"
         Height          =   372
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1812
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Feldname"
         Height          =   372
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1692
      End
   End
   Begin VB.CommandButton btnNutzerdef 
      Caption         =   "Nutzerdefinierte Felder einstellen"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "Wählen Sie maximal 5 feldnamen aus, deren Inhalt Sie sehen wollen"
      Top             =   6840
      Width           =   8772
   End
   Begin EditCtlsLibUCtl.TextBox txtSituation 
      Height          =   612
      Left            =   2160
      TabIndex        =   23
      Top             =   1080
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":074E
      Text            =   "MehrereZeilen.frx":076E
   End
   Begin EditCtlsLibUCtl.TextBox txtLand 
      Height          =   612
      Left            =   2160
      TabIndex        =   25
      Top             =   2520
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":078E
      Text            =   "MehrereZeilen.frx":07AE
   End
   Begin EditCtlsLibUCtl.TextBox txtPersonen 
      Height          =   612
      Left            =   2160
      TabIndex        =   26
      Top             =   3240
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":07CE
      Text            =   "MehrereZeilen.frx":07EE
   End
   Begin EditCtlsLibUCtl.TextBox txtDateiname 
      Height          =   612
      Left            =   2160
      TabIndex        =   27
      Top             =   3960
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":080E
      Text            =   "MehrereZeilen.frx":082E
   End
   Begin EditCtlsLibUCtl.TextBox txtDateinameKurz 
      Height          =   612
      Left            =   2160
      TabIndex        =   28
      Top             =   6000
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":084E
      Text            =   "MehrereZeilen.frx":086E
   End
   Begin EditCtlsLibUCtl.TextBox txtKommentar 
      Height          =   1212
      Left            =   2160
      TabIndex        =   44
      Top             =   4680
      Width           =   8772
      _cx             =   15473
      _cy             =   2138
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
      ReadOnly        =   -1  'True
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
      CueBanner       =   "MehrereZeilen.frx":088E
      Text            =   "MehrereZeilen.frx":08AE
   End
   Begin EditCtlsLibUCtl.TextBox txtOrt 
      Height          =   612
      Left            =   2160
      TabIndex        =   24
      Top             =   1800
      Width           =   8772
      _cx             =   15473
      _cy             =   1080
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
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   2
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "MehrereZeilen.frx":08CE
      Text            =   "MehrereZeilen.frx":08EE
   End
   Begin EditCtlsLibUCtl.TextBox txtEXIFInfo 
      Height          =   4452
      Left            =   11040
      TabIndex        =   45
      Top             =   840
      Width           =   2892
      _cx             =   5101
      _cy             =   7853
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
      CueBanner       =   "MehrereZeilen.frx":090E
      Text            =   "MehrereZeilen.frx":092E
   End
   Begin EditCtlsLibUCtl.TextBox txtIPTCInfo 
      Height          =   4452
      Left            =   11040
      TabIndex        =   46
      Top             =   840
      Width           =   2892
      _cx             =   5101
      _cy             =   7853
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
      CueBanner       =   "MehrereZeilen.frx":094E
      Text            =   "MehrereZeilen.frx":096E
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "HoehePixel:"
      Height          =   372
      Left            =   9360
      TabIndex        =   19
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BreitePixel:"
      Height          =   372
      Left            =   7800
      TabIndex        =   17
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DDatum:"
      Height          =   372
      Left            =   4920
      TabIndex        =   9
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DateinameKurz:"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   1932
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kommentar:"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   1692
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SWF:"
      Height          =   372
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dateiname:"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1692
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Personen:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1452
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Land:"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   972
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ort:"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   972
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Situation"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Jahr:"
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "F5MehrereZeilen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'txtEXIFInfo ist jetzt ein Timosoft Control unicode fähig                           'Gerbing 13.11.2015
Option Explicit
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
    'Const aus win32api.txt
    Const LB_SETHORIZONTALEXTENT = &H194
    
Private Sub btnNutzerdef_Click()
    Dim n As Long
    
    cmbFeld1.ComboItems.RemoveAll
    cmbFeld2.ComboItems.RemoveAll
    cmbFeld3.ComboItems.RemoveAll
    cmbFeld4.ComboItems.RemoveAll
    cmbFeld5.ComboItems.RemoveAll
    cmbFeld1.Visible = True
    cmbFeld2.Visible = True
    cmbFeld3.Visible = True
    cmbFeld4.Visible = True
    cmbFeld5.Visible = True
    txtFeldname1.Visible = False
    txtFeldname2.Visible = False
    txtFeldname3.Visible = False
    txtFeldname4.Visible = False
    txtFeldname5.Visible = False
    txtFeld1 = ""
    txtFeld2 = ""
    txtFeld3 = ""
    txtFeld4 = ""
    txtFeld5 = ""
    txtFeldname1 = ""
    txtFeldname2 = ""
    txtFeldname3 = ""
    txtFeldname4 = ""
    txtFeldname5 = ""

    'ND.ListNutzerdefinierteFelder.ListIndex = 0
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        cmbFeld1.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
    Next n
    'ND.ListNutzerdefinierteFelder.ListIndex = 0
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        cmbFeld2.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
    Next n
    'ND.ListNutzerdefinierteFelder.ListIndex = 0
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        cmbFeld3.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
    Next n
    'ND.ListNutzerdefinierteFelder.ListIndex = 0
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        cmbFeld4.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
    Next n
    'ND.ListNutzerdefinierteFelder.ListIndex = 0
    For n = 0 To ND.ListNutzerdefinierteFelder.ListItems.Count - 1
        cmbFeld5.ComboItems.Add ND.ListNutzerdefinierteFelder.ListItems(n)
    Next n
End Sub

Private Sub chkExif_Click()                         'Gerbing 09.11.2006
    Dim tempDateiname As String
    Dim strImageDescription As String
    Dim strXPTitle As String

    'If StrComp(Right(txtDateiname, 3), "JPG", vbTextCompare) <> 0 Then Exit Sub             'Gerbing 27.09.2013 'Gerbing 12.11.2015
    
    If chkExif.Value = 0 Then
        txtIPTCInfo.Visible = False                                                         'Gerbing 11.03.2016
        txtEXIFInfo.Visible = False                                                         'Gerbing 07.05.2007
    Else
        chkIptc.Value = 0
        txtIPTCInfo.Visible = False                                                         'Gerbing 11.03.2016
        txtEXIFInfo.Visible = True
        tempDateiname = Replace(txtDateiname, "+:\", gstrFotosMdbLocation & "\")            'Gerbing 07.11.2011
        Form1.EXF.ImageFile = tempDateiname 'set the image file property, read metainfo, parse metainfo
        '
        'EXF.ListInfo ist ein String mit vbCrLf
        '
        txtEXIFInfo.Text = Form1.EXF.ListInfo 'list all tags into the text box
    End If
End Sub

Public Sub chkIptc_Click()                         'Gerbing 09.11.2006
    Dim tempDateiname As String
    Dim rc As Boolean

    txtEXIFInfo.Text = ""                                                                   'Gerbing 12.11.2015
    'If StrComp(Right(txtDateiname, 3), "JPG", vbTextCompare) <> 0 Then Exit Sub             'Gerbing 27.09.2013 'Gerbing 12.11.2015
    
    If chkIptc.Value = 0 Then
        gblnIPTCAusgewählt = False                                                          'Gerbing 29.03.2012
        txtIPTCInfo.Visible = False                                                         'Gerbing 11.03.2016
        If StrComp(Right(txtDateiname, 3), "JPG", vbTextCompare) <> 0 Then Exit Sub         'Gerbing 12.11.2015
    Else
        gblnIPTCAusgewählt = True                                                          'Gerbing 29.03.2012
        chkExif.Value = 0
        txtIPTCInfo.Visible = True                                                          'Gerbing 11.03.2016
        If StrComp(Right(txtDateiname, 3), "JPG", vbTextCompare) <> 0 Then Exit Sub         'Gerbing 12.11.2015
        tempDateiname = Replace(txtDateiname, "+:\", gstrFotosMdbLocation & "\")            'Gerbing 07.11.2011
        'IPTC.ItemsDelimiter = vbCrLf
        IPTCItemsDelimiter = ";"
        rc = LeseIPTC(tempDateiname, txtIPTCInfo, True)  'mit Ausgabe in txtIPTCInfo        'Gerbing 11.03.2016
    End If
End Sub

Private Sub cmbFeld1_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    On Error Resume Next                                'Gerbing 22.11.2017
    Form1.F5Feld1 = cmbFeld1.Text
    If Not IsNull(frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld1.Text)) Then
        txtFeld1 = frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld1.Text)
    End If
End Sub

Private Sub cmbFeld2_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    On Error Resume Next                                'Gerbing 22.11.2017
    Form1.F5Feld2 = cmbFeld2.Text
    If Not IsNull(frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld2.Text)) Then
        txtFeld2 = frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld2.Text)
    End If
End Sub

Private Sub cmbFeld3_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    On Error Resume Next                                'Gerbing 22.11.2017
    Form1.F5Feld3 = cmbFeld3.Text
    If Not IsNull(frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld3.Text)) Then
        txtFeld3 = frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld3.Text)
    End If
End Sub

Private Sub cmbFeld4_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    On Error Resume Next                                'Gerbing 22.11.2017
    Form1.F5Feld4 = cmbFeld4.Text
    If Not IsNull(frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld4.Text)) Then
        txtFeld4 = frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld4.Text)
    End If
End Sub

Private Sub cmbFeld5_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
    On Error Resume Next                                'Gerbing 22.11.2017
    Form1.F5Feld5 = cmbFeld5.Text
    If Not IsNull(frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld5.Text)) Then
        txtFeld5 = frmGridAndThumb.Adodc1.Recordset.Fields(cmbFeld5.Text)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4
        Call MeUnload                                                               'Gerbing 23.08.2014
        'Tastatur-Eingabe weiterreichen
        '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
         Call Form1.Form_KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub Form_Load()
    Dim Faktor As Integer
    Dim rc As Long
    Dim Pixel As Integer
    Dim tempFRODN As String
    Dim mark As Variant
    Dim Msg As String

    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
        
    Me.Top = 0                                                      'Gerbing 06.12.2006
    Me.Left = 0
    'If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
        Me.Top = Form1.Top                                          'Gerbing 06.12.2006
        Me.Left = Form1.Left
        Me.width = Form1.width
        If screenWidth >= 1920 Then                                 'Gerbing 25.11.2013
            If Me.width - txtIPTCInfo.Left - 200 > 0 Then
                txtIPTCInfo.width = Me.width - txtIPTCInfo.Left - 200
            End If
            If Me.width - txtEXIFInfo.Left - 200 > 0 Then
                txtEXIFInfo.width = Me.width - txtEXIFInfo.Left - 200
            End If
        End If
    'End If

    Me.Caption = LoadResString(1022 + Sprache)    '"Angaben zum aktuellen Bild"   'Gerbing 08.11.2005
    Label1.Caption = LoadResString(1023 + Sprache)    'Jahr:
    Label2.Caption = LoadResString(1024 + Sprache)    'Situation:
    Label3.Caption = LoadResString(1025 + Sprache)    'Ort:
    Label4.Caption = LoadResString(1026 + Sprache)    'Land:
    Label5.Caption = LoadResString(1027 + Sprache)    'Personen:
    Label6.Caption = LoadResString(1028 + Sprache)    'Dateiname:
    Label7.Caption = LoadResString(1029 + Sprache)    'SWF:
    Label8.Caption = LoadResString(1030 + Sprache)    'Kommentar:
    Label9.Caption = LoadResString(1031 + Sprache)    'DateinameKurz:
    Label10.Caption = LoadResString(1032 + Sprache)    'DDatum:
    Label11.Caption = LoadResString(1033 + Sprache)    'Feldname
    Label12.Caption = LoadResString(1034 + Sprache)    'Feldinhalt
    Label13.Caption = LoadResString(1106 + Sprache)    'BreitePixel
    Label14.Caption = LoadResString(1107 + Sprache)    'HoehePixel
    btnNutzerdef.Caption = LoadResString(3016 + Sprache)   'Nutzerdefinierte Felder einstellen
    FrameNutzer.Caption = LoadResString(3070 + Sprache)  'Nutzerdefinierte Felder
    btnNutzerdef.tooltipText = LoadResString(3071 + Sprache) '"Wählen Sie maximal 5 Feldnamen aus, deren Inhalt Sie sehen wollen"
    txtFeldname1.Text = LoadResString(1111 + Sprache) '"noch nicht eingestellt"
    txtFeldname2.Text = LoadResString(1111 + Sprache)
    txtFeldname3.Text = LoadResString(1111 + Sprache)
    txtFeldname4.Text = LoadResString(1111 + Sprache)
    txtFeldname5.Text = LoadResString(1111 + Sprache)
    chkExif.Caption = LoadResString(1116 + Sprache) 'EXIF-Felder    'Gerbing 09.11.2006
    chkIptc.Caption = LoadResString(1117 + Sprache) 'IPTC-Felder    'Gerbing 09.11.2006
    txtIPTCInfo.Visible = False                                     'Gerbing 09.11.2006
    txtEXIFInfo.Visible = False
        
    If ND.ListNutzerdefinierteFelder.ListItems.Count <> 0 Then
        If Form1.F5Feld1 <> "" Then
            txtFeldname1.Text = Form1.F5Feld1
        End If
        If Form1.F5Feld2 <> "" Then
            txtFeldname2.Text = Form1.F5Feld2
        End If
        If Form1.F5Feld3 <> "" Then
            txtFeldname3.Text = Form1.F5Feld3
        End If
        If Form1.F5Feld4 <> "" Then
            txtFeldname4.Text = Form1.F5Feld4
        End If
        If Form1.F5Feld5 <> "" Then
            txtFeldname5.Text = Form1.F5Feld5
        End If
    End If
    
    mark = 1                                                                                                        'Gerbing 27.10.2012
    tempFRODN = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")                                               'Gerbing 27.10.2012
    tempFRODN = Replace(tempFRODN, "'", "''")
    'Bei Dateinamen mit Hochkomma bringt ..Find Laufzeitfehler -> ersetzen durch 2 Hochkommas                            'Gerbing 23.01.2018
    'frmGridAndThumb.rsDataGrid.Find LoadResString(1028 + Sprache) & " = '" & tempFRODN & "'"                            'Gerbing 27.10.2012
    frmGridAndThumb.rsDataGrid.Find LoadResString(1028 + Sprache) & " = '" & tempFRODN & "'", 0, adSearchForward, mark   'Gerbing 27.10.2012
    
    If frmGridAndThumb.rsDataGrid.EOF Then
        Msg = tempFRODN & vbNewLine
        'Msg = Msg & "frmGridAndThumb.rsDataGrid.Find"
        Msg = Msg & LoadResString(2289 + Sprache) & vbNewLine   '"Es wurde Refresh ausgeführt."                             'Gerbing 12.09.2020
        Msg = Msg & LoadResString(2290 + Sprache)               '"Dieses Bild entspricht nicht mehr den Such-Kriterien."    'Gerbing 12.09.2020
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub
    End If
    
    txtDateiname = frmGridAndThumb.rsDataGrid(LoadResString(1028 + Sprache))            'Gerbing 08.11.2005
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1031 + Sprache))) Then
        txtDateinameKurz = frmGridAndThumb.rsDataGrid(LoadResString(1031 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1032 + Sprache))) Then
        txtDDatum = frmGridAndThumb.rsDataGrid(LoadResString(1032 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1023 + Sprache))) Then
        txtJahr = frmGridAndThumb.rsDataGrid(LoadResString(1023 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1026 + Sprache))) Then
        txtLand = frmGridAndThumb.rsDataGrid(LoadResString(1026 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1025 + Sprache))) Then
        txtOrt = frmGridAndThumb.rsDataGrid(LoadResString(1025 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1027 + Sprache))) Then
        txtPersonen = frmGridAndThumb.rsDataGrid(LoadResString(1027 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1024 + Sprache))) Then
        txtSituation = frmGridAndThumb.rsDataGrid(LoadResString(1024 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1029 + Sprache))) Then
        txtSWF = frmGridAndThumb.rsDataGrid(LoadResString(1029 + Sprache))
    End If
    
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1030 + Sprache))) Then
        txtKommentar.Text = frmGridAndThumb.rsDataGrid(LoadResString(1030 + Sprache))     'Gerbing 04.03.2013
    Else
        txtKommentar.Text = ""                                                       'Gerbing 04.03.2013
    End If
    
    'es fehlen 2 Felder Gerbing 08.11.2005
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1106 + Sprache))) Then
        txtBreitePixel = frmGridAndThumb.rsDataGrid(LoadResString(1106 + Sprache))
    End If
    If Not IsNull(frmGridAndThumb.rsDataGrid(LoadResString(1107 + Sprache))) Then
        txtHoehePixel = frmGridAndThumb.rsDataGrid(LoadResString(1107 + Sprache))
    End If
    
    'nutzerdefinierte Felder
    If Form1.F5Feld1 <> "" Then
        If Not IsNull(frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld1)) Then
            F5MehrereZeilen.txtFeld1 = frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld1)
        End If
    End If
    If Form1.F5Feld2 <> "" Then
        If Not IsNull(frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld2)) Then
            F5MehrereZeilen.txtFeld2 = frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld2)
        End If
    End If
    If Form1.F5Feld3 <> "" Then
        If Not IsNull(frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld3)) Then
            F5MehrereZeilen.txtFeld3 = frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld3)
        End If
    End If
    If Form1.F5Feld4 <> "" Then
        If Not IsNull(frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld4)) Then
            F5MehrereZeilen.txtFeld4 = frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld4)
        End If
    End If
    If Form1.F5Feld5 <> "" Then
        If Not IsNull(frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld5)) Then
            F5MehrereZeilen.txtFeld5 = frmGridAndThumb.rsDataGrid.Fields(Form1.F5Feld5)
        End If
    End If

    If gblnIPTCAusgewählt = True Then                               'Gerbing 29.03.2012
        chkIptc.Value = 1
    Else
        chkIptc.Value = 0
    End If

    #If Proversion Then
        'wenn es keine nutzerdefinierten Felder gibt, Button und Frame unsichtbar machen
        If ND.ListNutzerdefinierteFelder.ListItems.Count = 0 Then
                Me.height = 6400                        'Gerbing 23.06.2011
                txtIPTCInfo.height = 5200               'Gerbing 23.06.2011
                txtEXIFInfo.height = 5200               'Gerbing 23.06.2011
                btnNutzerdef.Visible = False
                FrameNutzer.Visible = False
                Exit Sub
        Else
                Me.height = 11004                       'Gerbing 23.06.2011
                txtIPTCInfo.height = 9800             'Gerbing 23.06.2011
                txtEXIFInfo.height = 9800              'Gerbing 23.06.2011
                btnNutzerdef.Visible = True
                FrameNutzer.Visible = True
        End If
    #Else
        Me.height = 6400                        'Gerbing 23.06.2011
        txtIPTCInfo.height = 5200               'Gerbing 23.06.2011
        txtEXIFInfo.height = 5200               'Gerbing 23.06.2011
        btnNutzerdef.Visible = False
        FrameNutzer.Visible = False
        Exit Sub
    #End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)   'Gerbing 23.08.2014
    Dim Msg As String
    
    If Button = vbRightButton Then
        'MsgBox "erlaubt sind die Tasten F1 F2 F3 F4"
        Msg = LoadResString(2467 + Sprache)
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MeUnload                                                               'Gerbing 23.08.2014
End Sub

Private Sub MeUnload()                                                          'Gerbing 23.08.2014
    Dim Msg As String
    
    Unload Me
    If Form1.blnComeFromError = True Then                                       'Gerbing 15.02.2014
        Form1.Picture1 = LoadPicture()
        MakeGradient Form1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        MakeGradient Form1.Picture1, vbBlue, vbGreen, GRADIENT_FILL_RECT_V
        'MsgBox gstrFRODN & " Bild kann nicht geladen werden"
        Msg = gstrFRODN & " " & LoadResString(2056 + Sprache)
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        Exit Sub                                        'Gerbing 15.02.2014
    End If
    If gblnComefromVideo = True Then                                            'Gerbing 16.06.2012
        On Error Resume Next
        frmVideo.Show
        On Error GoTo 0
    End If
    'Set IPTC = Nothing
End Sub

