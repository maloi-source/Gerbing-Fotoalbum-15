VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmGeschichteDieserSoftware 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Geschichte dieser Software"
   ClientHeight    =   8484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   14040
   Icon            =   "frmGeschichteDieserSoftware.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8484
   ScaleWidth      =   14040
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox TextBox1 
      Height          =   5892
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13932
      _cx             =   24574
      _cy             =   10393
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
      CueBanner       =   "frmGeschichteDieserSoftware.frx":038A
      Text            =   "frmGeschichteDieserSoftware.frx":03AA
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aus Anlaß meines 75. Geburtstages siehe links und bevor ich nicht mehr weiss, was Programmcode ist siehe rechts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   2520
      TabIndex        =   1
      Top             =   5880
      Width           =   9252
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   11760
      Picture         =   "frmGeschichteDieserSoftware.frx":03DA
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2268
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   0
      Picture         =   "frmGeschichteDieserSoftware.frx":21348
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2508
   End
End
Attribute VB_Name = "frmGeschichteDieserSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim strFile As String
        
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
'    If Query.chkFensterGrößeÄnderbar.Value = 1 Then             'Gerbing 06.12.2005
'        Me.Top = Form1.Top                                      'Gerbing 06.12.2006
'        Me.Left = Form1.Left
'        Me.width = Form1.width \ 2
'    End If
    If Sprache = 0 Then
'        TextBox1.LoadFile AppPath & "\Help\Deutsch\ReadmeGeschichteDieserSoftware.rtf", rtfRTF
        strFile = MyReadFile(AppPath & "\Help\Deutsch\ReadmeGeschichteDieserSoftware.txt")

    Else
'        TextBox1.LoadFile AppPath & "\Help\English\ReadmeHistoryOfThisSoftware.rtf", rtfRTF
        strFile = MyReadFile(AppPath & "\Help\English\ReadmeHistoryOfThisSoftware.txt")

    End If
    Me.Caption = LoadResString(1108 + Sprache)                      '1108=Geschichte dieser Software lesen
    TextBox1.Text = strFile
    
    'Label1.Caption = "Aus Anlass meines 75. Geburtstages (siehe links) habe ich beschlossen, die Trennung zwischen Shareware-Version und Professional Version zu streichen. Es gibt ab Version GERBING Fotoalbum 15.0.5 nur noch eine Freeware Vollversion."
    'Label1.Caption = Label1.Caption & "Es kann sein, dass ich einen Zustand erreiche, wo ich nicht mehr weiß, was Compiler und Programmcode ist (siehe rechts)"

    Label1.Caption = LoadResString(1156 + Sprache)
    Label1.Caption = Label1.Caption & vbNewLine & vbNewLine & LoadResString(1157 + Sprache)
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'    TextBox1.width = Me.width - 200
'    TextBox1.height = Me.height - 400
'    On Error GoTo 0
'End Sub

Private Function MyReadFile(ByVal sFilePath As String) As String                        'Gerbing 04.03.2013
    Dim hHandle As Long
    Dim imageData() As Byte
    Dim imageString As String
    Dim bytesRead As Long

    hHandle = GetFileHandle(sFilePath, True)                             'true=read     'versteht unicode filename
    If hHandle <> INVALID_HANDLE_VALUE Then
        If hHandle Then
            bytesRead = GetFileSize(hHandle, ByVal 0&)
            If bytesRead Then
                ReDim imageData(0 To bytesRead - 1)
                ReadFile hHandle, imageData(0), bytesRead, bytesRead, ByVal 0&
                If bytesRead > UBound(imageData) Then
                    imageString = ""
                    imageString = StrConv(imageData, vbUnicode)
                End If
            End If
            CloseHandle hHandle
        End If
    End If
    MyReadFile = imageString
End Function

