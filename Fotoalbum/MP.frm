VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form MP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Weitere Filter"
   ClientHeight    =   10308
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9720
   Icon            =   "MP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10308
   ScaleWidth      =   9720
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame FrameVideoFilter 
      Caption         =   "Video-Filter"
      Height          =   2412
      Left            =   120
      TabIndex        =   39
      Top             =   5160
      Width           =   9492
      Begin VB.ComboBox cmbVG3 
         Height          =   288
         Left            =   1560
         Style           =   2  'Dropdown-Liste
         TabIndex        =   54
         Top             =   1680
         Width           =   732
      End
      Begin VB.ComboBox cmbVG2 
         Height          =   288
         Left            =   1560
         Style           =   2  'Dropdown-Liste
         TabIndex        =   53
         Top             =   1080
         Width           =   732
      End
      Begin VB.Frame Frame10 
         Height          =   492
         Left            =   3360
         TabIndex        =   50
         Top             =   960
         Width           =   2652
         Begin VB.OptionButton Und6 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   52
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Oder6 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   51
            Top             =   150
            Width           =   1332
         End
      End
      Begin VB.TextBox txtDauer 
         Height          =   372
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   49
         Top             =   1680
         Width           =   732
      End
      Begin VB.TextBox txtHöhe 
         Height          =   372
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   48
         Top             =   1080
         Width           =   732
      End
      Begin VB.Frame Frame9 
         Height          =   492
         Left            =   3360
         TabIndex        =   43
         Top             =   360
         Width           =   2652
         Begin VB.OptionButton Oder5 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   45
            Top             =   150
            Width           =   1332
         End
         Begin VB.OptionButton Und5 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   44
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtBreite 
         Height          =   372
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   42
         Top             =   480
         Width           =   732
      End
      Begin VB.ComboBox cmbVG1 
         Height          =   288
         Left            =   1560
         Style           =   2  'Dropdown-Liste
         TabIndex        =   41
         Top             =   480
         Width           =   732
      End
      Begin VB.Label lblDauer 
         Caption         =   "Dauer:"
         Height          =   372
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label lblHöhe 
         Caption         =   "Höhe:"
         Height          =   372
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label lblBreite 
         Caption         =   "Breite:"
         Height          =   372
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   1092
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Datums-Filter"
      Height          =   2052
      Left            =   120
      TabIndex        =   23
      Top             =   7680
      Width           =   9492
      Begin VB.TextBox txtDatumVon 
         Height          =   372
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "beliebig"
         Top             =   720
         Width           =   2052
      End
      Begin VB.TextBox txtDatumBis 
         Height          =   372
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "beliebig"
         Top             =   1440
         Width           =   2052
      End
      Begin VB.CommandButton btnDatumVon 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   372
         Left            =   8640
         TabIndex        =   28
         Top             =   720
         Width           =   372
      End
      Begin VB.CommandButton btnDatumBis 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   372
         Left            =   8640
         TabIndex        =   27
         Top             =   1440
         Width           =   372
      End
      Begin VB.Frame Frame8 
         Height          =   1212
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   5172
         Begin VB.OptionButton optDatumEinbeziehen 
            Caption         =   "in den Vergleich einbeziehen"
            Height          =   372
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   4932
         End
         Begin VB.OptionButton optDatumNichtEinbeziehen 
            Caption         =   "nicht einbeziehen(Standard)"
            Height          =   372
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Value           =   -1  'True
            Width           =   4932
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Geändert am:"
         Height          =   252
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Label5 
         Caption         =   "von:"
         Height          =   252
         Left            =   5400
         TabIndex        =   32
         Top             =   720
         Width           =   612
      End
      Begin VB.Label Label6 
         Caption         =   "bis:"
         Height          =   252
         Left            =   5400
         TabIndex        =   31
         Top             =   1440
         Width           =   612
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sortier-Filter"
      Height          =   1452
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   9492
      Begin VB.OptionButton OptNurNachDateiname 
         Caption         =   "DateinameKurz"
         Height          =   372
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Alphabetisch aufsteigend sortiert nach DateinameKurz"
         Top             =   840
         Width           =   2892
      End
      Begin VB.OptionButton OptNachKomplettemDateiename 
         Caption         =   "Dateiname"
         Height          =   372
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Alphabetisch aufsteigend sortiert nach dem kompletten Dateiname(Standard)"
         Top             =   360
         Value           =   -1  'True
         Width           =   2892
      End
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Alle Filter aus"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   9840
      Width           =   2892
   End
   Begin VB.Frame Frame2 
      Caption         =   "Personen-Filter"
      Height          =   3372
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9492
      Begin VB.Frame Frame5 
         Height          =   492
         Left            =   6720
         TabIndex        =   11
         Top             =   2040
         Width           =   2652
         Begin VB.OptionButton Und4 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton SWFOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   12
            Top             =   150
            Width           =   1332
         End
      End
      Begin VB.Frame Frame4 
         Height          =   492
         Left            =   6720
         TabIndex        =   8
         Top             =   1440
         Width           =   2652
         Begin VB.OptionButton Und3 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   10
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton LOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   9
            Top             =   150
            Width           =   1332
         End
      End
      Begin VB.Frame Frame3 
         Height          =   492
         Left            =   6720
         TabIndex        =   5
         Top             =   840
         Width           =   2652
         Begin VB.OptionButton Und2 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   6
            Top             =   150
            Width           =   1332
         End
      End
      Begin VB.Frame Frame6 
         Height          =   492
         Left            =   6720
         TabIndex        =   2
         Top             =   240
         Width           =   2652
         Begin VB.OptionButton Und1 
            Caption         =   "Und"
            Height          =   312
            Left            =   120
            TabIndex        =   4
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton SOder 
            Caption         =   "Oder"
            Height          =   312
            Left            =   1200
            TabIndex        =   3
            Top             =   150
            Width           =   1332
         End
      End
      Begin CBLCtlsLibUCtl.ComboBox TPerson1 
         Height          =   288
         Left            =   1680
         TabIndex        =   34
         Top             =   360
         Width           =   4932
         _cx             =   8700
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
         CueBanner       =   "MP.frx":038A
         Text            =   "MP.frx":03AA
      End
      Begin CBLCtlsLibUCtl.ComboBox TPerson2 
         Height          =   288
         Left            =   1680
         TabIndex        =   35
         Top             =   960
         Width           =   4932
         _cx             =   8700
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
         CueBanner       =   "MP.frx":03CA
         Text            =   "MP.frx":03EA
      End
      Begin CBLCtlsLibUCtl.ComboBox TPerson3 
         Height          =   288
         Left            =   1680
         TabIndex        =   36
         Top             =   1560
         Width           =   4932
         _cx             =   8700
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
         CueBanner       =   "MP.frx":040A
         Text            =   "MP.frx":042A
      End
      Begin CBLCtlsLibUCtl.ComboBox TPerson4 
         Height          =   288
         Left            =   1680
         TabIndex        =   37
         Top             =   2160
         Width           =   4932
         _cx             =   8700
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
         CueBanner       =   "MP.frx":044A
         Text            =   "MP.frx":046A
      End
      Begin CBLCtlsLibUCtl.ComboBox TPerson5 
         Height          =   288
         Left            =   1680
         TabIndex        =   38
         Top             =   2760
         Width           =   4932
         _cx             =   8700
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
         CueBanner       =   "MP.frx":048A
         Text            =   "MP.frx":04AA
      End
      Begin VB.Label lbl5Person 
         Caption         =   "5. Person:"
         Height          =   372
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1452
      End
      Begin VB.Label lbl4Person 
         Caption         =   "4. Person:"
         Height          =   372
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1452
      End
      Begin VB.Label lbl3Person 
         Caption         =   "3. Person:"
         Height          =   372
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label lbl2Person 
         Caption         =   "2. Person"
         Height          =   372
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label lbl1Person 
         Caption         =   "1. Person:"
         Height          =   372
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   9840
      Width           =   2892
   End
End
Attribute VB_Name = "MP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Gerbing 23.06.2011
'In der Entwicklungsumgebung eingestellt: ControlBox = False, weil nach Form_Unload trotzdem noch angezeigt wird
'weitere Filter sind aktiv

    Public AnzahlPersonen As Long
    Public Plus1 As String
    Public Plus2 As String
    Public Plus3 As String
    Public Plus4 As String
    Dim adoRs As ADODB.Recordset                        'Gerbing 04.01.2006
       
Private Sub btnAbbrechen_Click()
    TPerson1.Text = ""
    TPerson2.Text = ""
    TPerson3.Text = ""
    TPerson4.Text = ""
    TPerson5.Text = ""
    AnzahlPersonen = 0                                  'Gerbing 15.11.2004
    txtBreite.Text = ""                                 'Gerbing 12.06.2016
    txtHöhe.Text = ""                                   'Gerbing 12.06.2016
    txtDauer.Text = ""                                  'Gerbing 12.06.2016
    cmbVG1.ListIndex = 0                                'Gerbing 12.06.2016
    cmbVG2.ListIndex = 0                                'Gerbing 12.06.2016
    cmbVG3.ListIndex = 0                                'Gerbing 12.06.2016
    optDatumEinbeziehen.Value = False
    OptNachKomplettemDateiename.Value = True
    OptNurNachDateiname.Value = False
    optDatumNichtEinbeziehen.Value = True
    Query.CheckWeitereFilterAktiv.Value = 0
    Query.CheckWeitereFilterAktiv.Visible = False
    Query.lblWeitereFilterAktiv.Visible = False         'Gerbing 25.06.2013
    Query.TPersonen.Text = "*"                        'Gerbing 07.11.2004
    Query.TPersonen.Enabled = True
    Me.Hide
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
        MsgBox LoadResString(2018 + Sprache), , LoadResString(2019 + Sprache)
        Exit Sub
    End If
    txtDatumBis = DatumBis

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
        MsgBox LoadResString(2018 + Sprache), , LoadResString(2019 + Sprache)
        Exit Sub
    End If
    txtDatumVon = DatumVon

End Sub

Private Sub btnOK_Click()
    Dim msg As String

    'Query.CheckWeitereFilterAktiv soll als Kennzeichen dienen ob weitere Filter aktiv sind
    If TPerson1.Text <> "" Or TPerson2.Text <> "" Or TPerson3.Text <> "" Or TPerson4.Text <> "" Or TPerson5.Text <> "" _
            Or txtBreite.Text <> "" Or txtHöhe.Text <> "" Or txtDauer.Text <> "" _
            Or optDatumEinbeziehen.Value = True Or OptNurNachDateiname.Value = True Then
        Query.CheckWeitereFilterAktiv.Value = 1
        Query.CheckWeitereFilterAktiv.Visible = True
        Query.lblWeitereFilterAktiv.Visible = True                  'Gerbing 25.06.2013
    Else
        Query.CheckWeitereFilterAktiv.Value = 0
        Query.CheckWeitereFilterAktiv.Visible = False
        Query.lblWeitereFilterAktiv.Visible = False               'Gerbing 25.06.2013
    End If
    
    Query.TPersonen.Text = "*" 'Beliebig
    AnzahlPersonen = 0
    
    If TPerson1.Text = "" And TPerson2.Text = "" And TPerson3.Text = "" And TPerson4.Text = "" And TPerson5.Text = "" Then
        Me.Hide
        Exit Sub
    Else
        'es dürfen kein Lücken entstehen, Mindestens 2 Felder müssen belegt sein
        'AnzahlPersonen muss gezählt werden
        AnzahlPersonen = 2              'Anfangswert
        Do
            If TPerson1.Text = "" Or TPerson2.Text = "" Then
                'MsgBox "Mindestens die 1. Person und die 2. Person müssen ausgefüllt sein"
                MsgBox LoadResString(2102 + Sprache)               'Gerbing 08.11.2005
                Exit Sub
            End If
            If TPerson3.Text = "" And TPerson4.Text = "" And TPerson5.Text = "" Then Exit Do
            AnzahlPersonen = 3
            If TPerson3.Text <> "" Then
                If TPerson1.Text = "" Or TPerson2.Text = "" Then
                    'MsgBox "Sie müssen die Personen lückenlos ausfüllen"
                    MsgBox LoadResString(2103 + Sprache)          'Gerbing 08.11.2005
                    Exit Sub
                End If
                If TPerson4.Text = "" And TPerson5.Text = "" Then Exit Do
            End If
            AnzahlPersonen = 4
            If TPerson4.Text <> "" Then
                If TPerson1.Text = "" Or TPerson2.Text = "" Or TPerson3.Text = "" Then
                    'MsgBox "Sie müssen die Personen lückenlos ausfüllen"
                    MsgBox LoadResString(2103 + Sprache)          'Gerbing 08.11.2005
                    Exit Sub
                End If
                If TPerson5.Text = "" Then Exit Do
            End If
            AnzahlPersonen = 5
            If TPerson5.Text <> "" Then
                If TPerson1.Text = "" Or TPerson2.Text = "" Or TPerson3.Text = "" Or TPerson4.Text = "" Then
                    'MsgBox "Sie müssen die Personen lückenlos ausfüllen"
                    MsgBox LoadResString(2103 + Sprache)          'Gerbing 08.11.2005
                    Exit Sub
                End If
            End If
            Exit Do
        Loop
        Plus1 = " And "
        Plus2 = " And "
        Plus3 = " And "
        Plus4 = " And "
        If Und1.Value = False Then Plus1 = " Or "
        If Und2.Value = False Then Plus2 = " Or "
        If Und3.Value = False Then Plus3 = " Or "
        If Und4.Value = False Then Plus4 = " Or "
    
        'Query.TPersonen.Text = "MEHRERE PERSONEN"                                               'Gerbing 15.06.2008
        Query.TPersonen.Text = LoadResString(1124 + Sprache)
        Query.TPersonen.Enabled = False
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dim SQL As String
    
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then                 'Gerbing 06.12.2005
        Me.Top = Form1.Top                                          'Gerbing 06.12.2006
        Me.Left = Form1.Left
    End If

    Me.Caption = LoadResString(1114 + Sprache)              'Weitere Filter   Gerbing 11.07.2020
    lbl1Person.Caption = LoadResString(1084 + Sprache)    '1.Person:
    lbl2Person.Caption = LoadResString(1085 + Sprache)    '2.Person:
    lbl3Person.Caption = LoadResString(1086 + Sprache)    '3.Person:
    lbl4Person.Caption = LoadResString(1087 + Sprache)    '4.Person:
    lbl5Person.Caption = LoadResString(1088 + Sprache)    '5.Person:
    Label4.Caption = LoadResString(1017 + Sprache)        'Geändert am:
    Label5.Caption = LoadResString(1018 + Sprache)        'von:
    Label6.Caption = LoadResString(1019 + Sprache)        'bis:
    lblBreite.Caption = LoadResString(3160 + Sprache)       'Breite:
    lblHöhe.Caption = LoadResString(3161 + Sprache)         'Höhe:
    lblDauer.Caption = LoadResString(3159 + Sprache)        'Dauer:
    Frame1.Caption = LoadResString(3025 + Sprache)              'Sortier-Filter             Gerbing 11.07.2020
    Frame2.Caption = LoadResString(3022 + Sprache)              'Personen-Filter            Gerbing 11.07.2020
    Frame7.Caption = LoadResString(3028 + Sprache)              'Datums-Filter              Gerbing 11.07.2020
    FrameVideoFilter.Caption = LoadResString(3158 + Sprache)   'Video-Filter                Gerbing 12.06.2016
    Und1.Caption = LoadResString(3023 + Sprache) 'Und
    Und2.Caption = LoadResString(3023 + Sprache) 'Und
    Und3.Caption = LoadResString(3023 + Sprache) 'Und
    Und4.Caption = LoadResString(3023 + Sprache) 'Und
    Und5.Caption = LoadResString(3023 + Sprache) 'Und                                       Gerbing 12.06.2016
    Und6.Caption = LoadResString(3023 + Sprache) 'Und                                       Gerbing 12.06.2016
    SOder.Caption = LoadResString(3024 + Sprache) 'Oder
    OOder.Caption = LoadResString(3024 + Sprache) 'Oder
    LOder.Caption = LoadResString(3024 + Sprache) 'Oder
    SWFOder.Caption = LoadResString(3024 + Sprache) 'Oder
    Oder5.Caption = LoadResString(3024 + Sprache) 'Oder                                     Gerbing 12.06.2016
    Oder6.Caption = LoadResString(3024 + Sprache) 'Oder                                     Gerbing 12.06.2016
    OptNachKomplettemDateiename.Caption = LoadResString(1028 + Sprache)     'Dateiname
    OptNachKomplettemDateiename.tooltipText = LoadResString(3026 + Sprache) 'Alphabetisch aufsteigend sortiert nach dem kompletten Dateiname(Standard)
    OptNurNachDateiname.Caption = LoadResString(1031 + Sprache)             'DateinameKurz
    OptNurNachDateiname.tooltipText = LoadResString(3027 + Sprache)         'Alphabetisch aufsteigend sortiert nach DateinameKurz
    optDatumEinbeziehen.Caption = LoadResString(3029 + Sprache)             'in den Vergleich einbeziehen
    optDatumNichtEinbeziehen.Caption = LoadResString(3030 + Sprache)        'nicht einbeziehen(Standard)
    btnOK.Caption = LoadResString(3001 + Sprache)     '&OK
    btnAbbrechen.Caption = LoadResString(3031 + Sprache)    '&Alle Filter aus
    txtDatumVon = LoadResString(1110 + Sprache) 'beliebig                                   'Gerbing 15.06.2008
    txtDatumBis = LoadResString(1110 + Sprache) 'beliebig
    cmbVG1.AddItem "="
    cmbVG1.AddItem ">"
    cmbVG1.AddItem "<"
    cmbVG1.AddItem ">="
    cmbVG1.AddItem "<="
    cmbVG1.AddItem "<>"
    cmbVG2.AddItem "="
    cmbVG2.AddItem ">"
    cmbVG2.AddItem "<"
    cmbVG2.AddItem ">="
    cmbVG2.AddItem "<="
    cmbVG2.AddItem "<>"
    cmbVG3.AddItem "="
    cmbVG3.AddItem ">"
    cmbVG3.AddItem "<"
    cmbVG3.AddItem ">="
    cmbVG3.AddItem "<="
    cmbVG3.AddItem "<>"
    cmbVG1.ListIndex = 0
    cmbVG2.ListIndex = 0
    cmbVG3.ListIndex = 0
    
    Me.MousePointer = vbHourglass                                                           'Gerbing 29.07.2007
    '---------------------------------------------
'    glngStartMillisek = timeGetTime
    'SQL = "SELECT DISTINCT Fotos.Personen FROM Fotos " & SQLJahreszahl & " AND ((Not (Fotos.Personen)=''))
'
'
    'SQL = SQL & " ORDER BY Personen;" 'Gerbing 08.11.2012
    SQL = "SELECT DISTINCT Fotos." & LoadResString(1027 + Sprache) & " FROM Fotos " & Query.SQLJahresZahl & " AND ((Not (Fotos." & LoadResString(1027 + Sprache) & ")='')) "
    If Query.TSituation.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1024 + Sprache) & "='" & Query.TSituation.Text & "'"
    End If
    If Query.TOrt.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1025 + Sprache) & "='" & Query.TOrt.Text & "'"
    End If
    If Query.TLand.Text <> "*" Then
        SQL = SQL & " AND " & LoadResString(1026 + Sprache) & "='" & Query.TLand.Text & "'"
    End If
    SQL = SQL & " ORDER BY " & LoadResString(1027 + Sprache) & ";"
    'Set rst = db.OpenRecordset(SQL)
    Set adoRs = New ADODB.Recordset
    With adoRs
        .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        '.CursorLocation = Query.enumCursorOrt
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    Set Query.Adodc1.Recordset = adoRs
    Do Until Query.Adodc1.Recordset.EOF
        If Not IsNull(Query.Adodc1.Recordset(LoadResString(1027 + Sprache))) Then
            MP.TPerson1.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
            MP.TPerson2.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
            MP.TPerson3.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
            MP.TPerson4.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
            MP.TPerson5.ComboItems.Add Query.Adodc1.Recordset(LoadResString(1027 + Sprache))
        End If
        Query.Adodc1.Recordset.MoveNext
        'DoEvents       Gerbing auskommentiert 01.01.2008                           'Gerbing 16.09.2004
    Loop
'    glngEndMillisek = timeGetTime
'    Debug.Print "Millisekunden Combobox Personen AddItem" & "=" & (glngEndMillisek - glngStartMillisek)
'    MsgBox "Millisekunden Combobox Personen AddItem" & "=" & (glngEndMillisek - glngStartMillisek)
    Me.MousePointer = vbDefault                                                             'Gerbing 29.07.2007
End Sub

Private Sub Form_Paint()
    On Error Resume Next                    'Gerbing 30.01.2005
    TPerson1.SetFocus
    On Error GoTo 0
End Sub

Private Sub LSWF_Click()

End Sub

Private Sub optDatumEinbeziehen_Click()
    'blnDatumWarSchonMal = True
    If optDatumEinbeziehen.Value = True Then
        txtDatumVon = Date
        txtDatumBis = Date
        btnDatumVon.Enabled = True
        btnDatumBis.Enabled = True
    End If
End Sub

Private Sub optDatumNichtEinbeziehen_Click()
    'blnDatumWarSchonMal = True
    txtDatumVon = LoadResString(1110 + Sprache) 'beliebig
    txtDatumBis = LoadResString(1110 + Sprache) 'beliebig
    btnDatumVon.Enabled = False
    btnDatumBis.Enabled = False
End Sub

Private Sub txtBreite_Validate(KeepFocus As Boolean)
    'MaxLength = 5 in den properties gesetzt
    Dim lenBreite As Long
    Dim i As Long
    
    lenBreite = Len(txtBreite.Text)
    If txtBreite.Text <> "" Then
        For i = 1 To lenBreite
            If Not IsNumeric(Mid(txtBreite.Text, i, 1)) Then
                txtBreite.Text = ""
                KeepFocus = True
                MsgBox LoadResString(2288 + Sprache) 'Sie müssen eine max 5-stellige Zahl eingeben
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub txtDauer_Validate(KeepFocus As Boolean)
    'MaxLength = 5 in den properties gesetzt
    Dim lenDauer As Long
    Dim i As Long
    
    lenDauer = Len(txtDauer.Text)
    If txtDauer.Text <> "" Then
        For i = 1 To lenDauer
            If Not IsNumeric(Mid(txtDauer.Text, i, 1)) Then
                txtDauer.Text = ""
                KeepFocus = True
                MsgBox LoadResString(2288 + Sprache) 'Sie müssen eine max 5-stellige Zahl eingeben
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub txtHöhe_Validate(KeepFocus As Boolean)
    'MaxLength = 5 in den properties gesetzt
    Dim lenHöhe As Long
    Dim i As Long
    
    lenHöhe = Len(txtHöhe.Text)
    If txtHöhe.Text <> "" Then
        For i = 1 To lenHöhe
            If Not IsNumeric(Mid(txtHöhe.Text, i, 1)) Then
                txtHöhe.Text = ""
                KeepFocus = True
                MsgBox LoadResString(2288 + Sprache) 'Sie müssen eine max 5-stellige Zahl eingeben
                Exit For
            End If
        Next i
    End If
End Sub
