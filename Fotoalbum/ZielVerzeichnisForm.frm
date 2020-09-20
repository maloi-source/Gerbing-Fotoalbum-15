VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form ZielVerzeichnisForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Name des Zielverzeichnisses auswählen"
   ClientHeight    =   5112
   ClientLeft      =   6408
   ClientTop       =   2340
   ClientWidth     =   6972
   Icon            =   "ZielVerzeichnisForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5112
   ScaleWidth      =   6972
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtZielOrdner 
      Height          =   492
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6132
      _cx             =   10816
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
      CueBanner       =   "ZielVerzeichnisForm.frx":038A
      Text            =   "ZielVerzeichnisForm.frx":03AA
   End
   Begin VB.CommandButton btnBrowseForFolder 
      Caption         =   "..."
      Height          =   492
      Left            =   6360
      TabIndex        =   5
      Top             =   840
      Width           =   492
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   1932
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1932
   End
   Begin VB.Label lblNameDesZielverzeichnisses 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name des Zielverzeichnisses"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   6132
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"ZielVerzeichnisForm.frx":03CA
      Height          =   1092
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   6612
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bitte treffen Sie diese Entscheidung selbst. Benutzen Sie dafür den Explorer."
      Height          =   612
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   6612
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"ZielVerzeichnisForm.frx":0465
      Height          =   852
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   6612
   End
End
Attribute VB_Name = "ZielVerzeichnisForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Das Formular wurde überarbeitet am 29.06.2011 durch Aufruf von BrowseForFolder
    Dim Prüfdir As String
    Dim msg As String

Private Sub btnBrowseForFolder_Click()
    Dim Prompt As String

    Prompt = ""
    'Prompt = "Standort der Fotos/Videos:"
    'Prompt = LoadResString(1812 + Sprache)
    Prompt = BrowseForFolder("Folder", Prompt, Me.hWnd, False, , True, False)   'ist unicode fähig
    If Prompt <> "" Then                                                                    'Gerbing 17.12.2008
        'MessageBoxW 0, StrPtr(Prompt), StrPtr("GERBING Fotoalbum"), vbInformation
        txtZielOrdner.Text = Prompt
    End If
End Sub

Private Sub btnOK_Click()
    Form1.blnUnloadExportForm = False
    PublicExportTargetFolder = txtZielOrdner.Text
    ExportForm.ZielVerzeichnis = txtZielOrdner.Text
    Call WriteEZV(txtZielOrdner.Text)     'Rückschreiben in fotos.ini
    'msg = "Bitte entscheiden Sie selbst, ob der Inhalt des Zielverzeichnisses gelöscht werden muss." & NL
    msg = LoadResString(2194 + Sprache) & NL
    'msg = msg & "Jetzt ist Gelegenheit dazu zB mit Hilfe des Windows Explorers."
    msg = msg & LoadResString(2195 + Sprache)
    MsgBox msg
    Unload Me
End Sub

Private Sub btnAbbrechen_Click()
    Unload Me
    Form1.blnUnloadExportForm = True
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                       'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then     'Gerbing 06.12.2005
        Me.Top = Form1.Top                              'Gerbing 06.12.2006
        Me.Left = Form1.Left
    End If

    Me.Caption = LoadResString(1102 + Sprache)        'Name des Zielverzeichnisses auswählen
    Label1.Caption = LoadResString(1103 + Sprache)    'Das Programm erzeugt im Zielverzeichnis weitere Unterverzeichnisse. Die Namen der Unterverzeichnisse entsprechen dem Name des zu exportierenden Fotos.
    Label2.Caption = LoadResString(1104 + Sprache)    'Wenn Sie das Zielverzeichnis schon mehrfach für Kopiervorgänge benutzt haben, könnte es notwendig sein den bisherigen Inhalt zu löschen.
    Label3.Caption = LoadResString(1105 + Sprache)    'Bitte treffen Sie diese Entscheidung selbst. Benutzen Sie dafür den Explorer.
    lblNameDesZielverzeichnisses.Caption = LoadResString(2187 + Sprache)    'Name des Zielverzeichnisses
    btnOK.Caption = LoadResString(3001 + Sprache)      '&Ok
    btnAbbrechen.Caption = LoadResString(3013 + Sprache)   '&Abbrechen
    On Error Resume Next
    txtZielOrdner.Text = PublicExportTargetFolder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtZielOrdner.Text = "" Then
        ExportForm.CheckFotos.Value = 0
        ExportForm.txtZielVerzeichnis.Visible = False
    End If
End Sub
