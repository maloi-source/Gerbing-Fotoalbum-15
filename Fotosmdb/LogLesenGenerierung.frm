VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmLogLesenGenerierung 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Datei pruef.log"
   ClientHeight    =   8460
   ClientLeft      =   -12
   ClientTop       =   276
   ClientWidth     =   15000
   Icon            =   "LogLesenGenerierung.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15000
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "Ab&brechen"
      Height          =   372
      Left            =   4800
      TabIndex        =   0
      Top             =   7920
      Width           =   4935
   End
   Begin EditCtlsLibUCtl.TextBox TxtU 
      Height          =   7572
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14772
      _cx             =   26056
      _cy             =   13356
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
      CueBanner       =   "LogLesenGenerierung.frx":038A
      Text            =   "LogLesenGenerierung.frx":03B2
   End
End
Attribute VB_Name = "frmLogLesenGenerierung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim LogFileName As String
    Dim Msg As String
    Dim myStream As TextStream
    Dim sLine As String
    
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(1337 + Sprache)   'Datei Pruef.log
    btnAbbrechen.Caption = LoadResString(1325 + Sprache)                'Abbru&ch
    
    Me.top = 0
    Me.left = 0
    LogFileName = PruefLogFile
    On Error GoTo Fehler
    TxtU.Text = ""
    Screen.MousePointer = vbHourglass
    If (myStream Is Nothing) Then
        ' Open the file for reading.
            Set myStream = PruefFso.OpenTextFile(LogFileName, 1, False, -1)     'unicode
        If (Not myStream Is Nothing) Then
            With myStream
                Do Until myStream.AtEndOfStream
                    sLine = myStream.ReadLine
                    TxtU.Text = TxtU.Text & sLine & vbNewLine
                Loop
                .Close
            End With
            Set myStream = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Fehler:
    'Msg = LogFileName & " kann nicht geöffnet werden"
    Msg = LogFileName & " " & LoadResString(1372 + Sprache)
    MsgBox Msg
    Unload Me
    Exit Sub
End Sub

Private Sub btnAbbrechen_Click()
    Unload Me
End Sub

