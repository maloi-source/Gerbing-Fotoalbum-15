VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConnectSQL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "GERBING Fotoalbum Connect sql server"
   ClientHeight    =   9744
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6540
   Icon            =   "frmConnectSQL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9744
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtStandortFotos 
      Height          =   372
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   5892
      _cx             =   10393
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
      CueBanner       =   "frmConnectSQL.frx":038A
      Text            =   "frmConnectSQL.frx":03AA
   End
   Begin EditCtlsLibUCtl.TextBox txtDatabaseName 
      Height          =   372
      Left            =   2400
      TabIndex        =   18
      Top             =   720
      Width           =   3972
      _cx             =   7006
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
      CueBanner       =   "frmConnectSQL.frx":03CA
      Text            =   "frmConnectSQL.frx":03EA
   End
   Begin EditCtlsLibUCtl.TextBox txtSQLServerName 
      Height          =   372
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   3972
      _cx             =   7006
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
      CueBanner       =   "frmConnectSQL.frx":040A
      Text            =   "frmConnectSQL.frx":042A
   End
   Begin VB.Frame FrameLogin 
      BackColor       =   &H00C0C0C0&
      Height          =   4692
      Left            =   1080
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   4332
      Begin VB.TextBox txtAllowedlicenses 
         Height          =   288
         Left            =   2520
         TabIndex        =   14
         Top             =   3720
         Width           =   732
      End
      Begin VB.TextBox txtNumberOfUsers 
         Height          =   288
         Left            =   2520
         TabIndex        =   13
         Top             =   4200
         Width           =   732
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "Login"
         Height          =   612
         Left            =   1320
         TabIndex        =   11
         Top             =   2880
         Width           =   1932
      End
      Begin MSDataGridLib.DataGrid MyDataGrid1 
         Bindings        =   "frmConnectSQL.frx":044A
         Height          =   2532
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   4212
         _ExtentX        =   7430
         _ExtentY        =   4466
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1031
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAllowedlicenses 
         BackColor       =   &H00C0C0C0&
         Caption         =   "allowed licenses:"
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   2172
      End
      Begin VB.Label lblNumberofusers 
         BackColor       =   &H00C0C0C0&
         Caption         =   "number of users"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   2172
      End
   End
   Begin VB.CommandButton btnBrowseForFolder 
      Caption         =   "..."
      Height          =   372
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "öffnet einen Dialog zur Ordner-Auswahl"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   612
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   1932
   End
   Begin VB.Frame FrameAuthentication 
      BackColor       =   &H00C0C0C0&
      Height          =   1452
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6252
      Begin EditCtlsLibUCtl.TextBox txtUserName 
         Height          =   372
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   1212
         _cx             =   2138
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
         CueBanner       =   "frmConnectSQL.frx":045A
         Text            =   "frmConnectSQL.frx":047A
      End
      Begin VB.OptionButton optWindowsAuthentication 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Windows Authentication"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3372
      End
      Begin VB.OptionButton optSqlserverauthentication 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SQL server authentication"
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3372
      End
      Begin EditCtlsLibUCtl.TextBox txtPassword 
         Height          =   372
         Left            =   4920
         TabIndex        =   20
         Top             =   840
         Width           =   1212
         _cx             =   2138
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
         ReadOnly        =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightMargin     =   -1
         RightToLeft     =   0
         ScrollBars      =   0
         SelectedTextMousePointer=   0
         SupportOLEDragImages=   -1  'True
         TabWidth        =   -1
         UseCustomFormattingRectangle=   0   'False
         UsePasswordChar =   -1  'True
         UseSystemFont   =   0   'False
         CueBanner       =   "frmConnectSQL.frx":049A
         Text            =   "frmConnectSQL.frx":04BA
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H00C0C0C0&
         Caption         =   "user name:"
         Height          =   372
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00C0C0C0&
         Caption         =   "password:"
         Height          =   372
         Left            =   3600
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Label lblStandortFotos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Standort der Fotos/Videos:"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   4332
   End
   Begin VB.Label lblServerName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SQL server name:"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2172
   End
   Begin VB.Label lblDatabasename 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Database name:"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2172
   End
End
Attribute VB_Name = "frmConnectSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If Proversion Then
    Dim Prompt As String
    
Private Sub btnBrowseForFolder_Click()
    'Prompt = "Standort der Fotos/Videos:"
    Prompt = LoadResString(1812 + Sprache)
    Prompt = BrowseForFolder("Folder", Prompt, Me.hWnd, False, , True, False)   'ist unicode fähig
    If Prompt = "" Then
        Exit Sub
    End If
    txtStandortFotos = Prompt
End Sub

Private Sub btnconnect_Click()
    Dim rst1 As ADODB.Recordset
    Dim strSchlüssel As String
    Dim strPrimaryKey As String
    Dim LizenzUnCodiert As String
    Dim rc As Boolean
    Dim msg As String
    Dim strTemp As String
    Dim antwort As Long

    Screen.MousePointer = vbHourglass                                       'Gerbing 27.08.2017
    DoEvents
    
    If txtStandortFotos.Text = "" Then
        'MsgBox lblStandortFotos.Caption & " ist leer."
        MsgBox lblStandortFotos.Caption & LoadResString(1512 + Sprache)
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    If txtSQLServerName.Text = "" Then
        'MsgBox lblServerName.Caption & " ist leer."
        MsgBox lblStandortFotos.Caption & LoadResString(1512 + Sprache)
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    If txtDatabaseName.Text = "" Then
        'MsgBox lblDatabasename.Caption & " ist leer."
        MsgBox lblStandortFotos.Caption & LoadResString(1512 + Sprache)
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    
    'Set DBado = New ADODB.Connection
    Set DBado = CreateObject("ADODB.Connection")
    
    With DBado
        .Provider = "SQLOLEDB.1"
        '.Provider = "SQLNCLI10.1" 'SQL Server Native Client
        .Properties("Persist Security Info").Value = False
        .Properties("Initial Catalog").Value = txtDatabaseName
        .Properties("Data Source").Value = txtSQLServerName
        '   Falls die Windows-Authentifizierung verwendet werden soll, muß "SSPI" benutzt werden
        If optWindowsAuthentication.Value = True Then
            .Properties("Integrated Security").Value = "SSPI"
        Else
            .Properties("User ID").Value = txtUserName
            .Properties("Password").Value = txtPassword
        End If
        On Error Resume Next
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description
            MsgBox msg
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End With
    
    'Set DollarDBado = New ADODB.Connection
    Set DollarDBado = CreateObject("ADODB.Connection")
    
    With DollarDBado
        .Provider = "SQLOLEDB.1"
        '.Provider = "SQLNCLI10.1" 'SQL Server Native Client
        .Properties("Persist Security Info").Value = False
        .Properties("Initial Catalog").Value = "$" & txtDatabaseName
        .Properties("Data Source").Value = txtSQLServerName
        '   Falls die Windows-Authentifizierung verwendet werden soll, muß "SSPI" benutzt werden
        If optWindowsAuthentication.Value = True Then
            .Properties("Integrated Security").Value = "SSPI"
        Else
            .Properties("User ID").Value = txtUserName
            .Properties("Password").Value = txtPassword
        End If
        On Error Resume Next
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description
            MsgBox msg
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End With
    
    Me.MousePointer = vbNormal
    Set rstsql = New ADODB.Recordset
    With rstsql
        .Source = "select * from loggedinusers"
        .ActiveConnection = DBado
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description & vbNewLine & vbNewLine
            
            msg = msg & txtSQLServerName.Text & " " & txtDatabaseName.Text & " missing table LoggedInUsers"
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbInformation
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End With
    Set MyDataGrid1.DataSource = rstsql
    MyDataGrid1.Refresh                                                     'Gerbing 04.03.2013
    MyDataGrid1.Columns(0).width = 3000
    MyDataGrid1.Columns(1).width = 900
    MyDataGrid1.Columns(2).width = 0

    
    'Ermitteln der erlaubten Lizenzen aus der Tabelle LicenseCode
    Set rst1 = New ADODB.Recordset
    With rst1
        .Source = "select * from LicenseCode"
        .ActiveConnection = DBado
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description
            MsgBox msg
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End With
    On Error GoTo 0
    strSchlüssel = rst1.Fields("LicenseCode")
    rst1.Close
    If strSchlüssel = "" Then
        MsgBox LoadResString(1817)    'invalid LicenseCode. The programm will be stopped. Go to the sql server administrator.
        End
    End If
    LizenzUnCodiert = Crypt(Left(strSchlüssel, 10), Mid(strSchlüssel, 24, 5), False)            'Gerbing 19.11.2012
    If LizenzUnCodiert = "error" Or LizenzUnCodiert = "" Then
        MsgBox LoadResString(1817)    'invalid LicenseCode. The programm will be stopped. Go to the sql server administrator.
        End
    End If

    txtAllowedlicenses = Mid(LizenzUnCodiert, 4, 2)
    rc = Kontrolle(strSchlüssel)
    If rc = False Then End                      'Programm beendet bei falschem LicenseCode
    '---------------------------------------------------------------------------------------------------
    gstrAllowedlicenses = Mid(LizenzUnCodiert, 4, 2)
    gintNumberOfUsers = rstsql.RecordCount
    
    'Kontrolle ob mehr Einträge in Tabelle loggedinusers stehen als erlaubte Lizenzen
    If rstsql.RecordCount > txtAllowedlicenses Then
        MsgBox "Zu viele usernames. Das Programm wird beendet. Too many usernames. The program stopps. Go to the sql server administrator."
        End
    End If
    
    gblnSQLServerVersion = True
    gblnSQLServerConnected = True
    'Parameter für erfolgreiche sql server verbindung in fotos.ini eintragen
    PublicLocationFotos = txtStandortFotos.Text
    Call WriteSLF(txtStandortFotos.Text)    'Schreibe [SQL] LocationFotos
    PublicSQLServer = txtSQLServerName.Text
    Call WriteSSRV(txtSQLServerName.Text)   'Schreibe [SQL] Server
    PublicSQLDatabase = txtDatabaseName.Text
    Call WriteSDB(txtDatabaseName.Text)     'Schreibe [SQL] Database
    If optWindowsAuthentication.Value = True Then
        Call WriteSWA("1")                  'Schreibe [SQL] WindowsAuthentication = 1
    Else
        Call WriteSWA("0")                  'Schreibe [SQL] WindowsAuthentication = 0
        PublicSQLServerUserName = txtUserName.Text
        Call WriteSUN(txtUserName.Text)     'Schreibe [SQL] username
        PublicSQLServerPassword = txtPassword.Text
        Call WriteSPW(txtPassword.Text)     'Schreibe [SQL] password
    End If
    txtNumberOfUsers = rstsql.RecordCount
    If txtAllowedlicenses = 99 Then
        Unload Me
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    
    'Frage ob im ersten Satz der Tabelle Fotos der Spaltenname Filename vorkommt
    'SQL = "SELECT * From fotos WHERE not filename Is Null;"
    Set rst1 = New ADODB.Recordset
    On Error Resume Next
    With rst1
        .Source = "SELECT * From fotos WHERE not filename Is Null;"
        .ActiveConnection = DBado
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number = 0 Then
            If Sprache = 0 Then
                Call WriteGlL(1)     'Rückschreiben english in fotos.ini
                msg = "error number=" & Err.Number & vbNewLine
                msg = msg & "error text=" & Err.Description
                msg = "Starten Sie neu. Restart the program."
                MsgBox msg
                End
            End If
        Else
            If Sprache = 3000 Then
                Call WriteGlL(0)     'Rückschreiben deutsch in fotos.ini
                msg = "error number=" & Err.Number & vbNewLine
                msg = msg & "error text=" & Err.Description
                msg = "Starten Sie neu. Restart the program."
                MsgBox msg
                End
            End If
        End If
    End With
    
    'Frage ob es einen ersten Satz in der Tabelle Fotos gibt
    'SQL = "SELECT * From fotos;"
    On Error Resume Next
    rst1.Close
    On Error GoTo 0
    With rst1
        .Source = "select * from Fotos"
        .ActiveConnection = DBado
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If rst1.EOF Then
        'Msg = "Datenbankname:" & " " & PublicSQLServer & " " & PublicSQLDatabase & "  ist leer. Die einzige erlaubte Operation ist das Erzeugen einer neuen Datenbank mit dem Tool Fotosmdb." & vbnewline
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        msg = LoadResString(1806 + Sprache) & " " & PublicSQLServer & " " & PublicSQLDatabase & LoadResString(1836 + Sprache) & vbNewLine
        msg = msg & LoadResString(2159 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then
            End
        Else
            gblnWeiterMitLeererDatenbank = True
            FrameLogin.Visible = True
            Me.MousePointer = vbNormal
            Exit Sub
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    On Error GoTo 0
    strSchlüssel = rst1.Fields(LoadResString(1028 + Sprache))   '1028=Dateiname
    rst1.Close
    
    Set rst1 = DBado.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, "Fotos")) '2529=fotos
    If rst1.EOF Then
        'MsgBox "Die Tabelle fotos enthält keinen Primärschlüssel. Das Programm wird beendet."
        MsgBox LoadResString(1823 + Sprache)
        End
    End If
    strPrimaryKey = rst1.Fields("COLUMN_NAME").Value
    If StrComp(LoadResString(1028 + Sprache), strPrimaryKey, vbTextCompare) <> 0 Then       '1028=Dateiname
        'MsgBox "Die Spalte Dateiname ist nicht der Primärschlüssel. Das Programm wird beendet."
        MsgBox LoadResString(1824 + Sprache)
        End
    End If
    
    'Prüfen ob es den im ersten Satz enthaltenen Dateiname gibt
    strSchlüssel = Replace(strSchlüssel, "+:\", txtStandortFotos.Text & "\")
    
    On Error Resume Next
    strTemp = ""
    On Error GoTo 0
    If file_path_exist(strSchlüssel) = False Then
        Call WriteSLF("")           'Schreibe [SQL] LocationFotos
        'msg = Feldname & " existiert nicht." & vbNewLine
        msg = strSchlüssel & LoadResString(2162 + Sprache) & vbNewLine
        'msg = msg & "Prüfen Sie, ob die Angabe " & txtStandortFotos.Text  & " richtig ist" & vbNewLine & vbNewLine
        msg = msg & LoadResString(1820 + Sprache) & txtStandortFotos.Text & LoadResString(1821 + Sprache) & vbNewLine & vbNewLine
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        msg = msg & LoadResString(2159 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then
            End
        Else
            gblnWeiterMitAnderemFotosStandort = True
        End If
    End If
    FrameLogin.Visible = True
    Screen.MousePointer = vbNormal                                          'Gerbing 07.01.2018
End Sub

Private Sub btnLogin_Click()
    Dim strManagement As String
    Dim msg As String
    Dim antwort As Long

    If Not rstsql.EOF Then
        If rstsql.Fields("LoggedIn") = True Then
            msg = "user " & rstsql.Fields("username") & " already logged in." & vbNewLine
            msg = msg & "Do you want to login with another username. Wollen Sie mit einem anderen username einloggen?"
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbDefaultButton1 + vbYesNo)
            If antwort = vbYes Then
                Exit Sub
            Else
                End
            End If
        End If
    Else
        MsgBox "Es gibt keine Nutzer. Erzeugen Sie Nutzer mit EnterNewUsers.exe. There are no users. Create users with EnterNewUsers.exe."
        End
    End If
    
    'strManagement soll zum Zeitpunkt des Login enthalten IN &Datum&Uhrzeit
    strManagement = "IN " & Now
    strManagement = Crypt(strManagement, rstsql.Fields("username"), True)
    rstsql("LoggedIn") = True
    rstsql("Management") = strManagement
    gstrLoggedInName = rstsql.Fields("username")
    rstsql.Update
    MyDataGrid1.Refresh
    rstsql.Requery
    MyDataGrid1.Columns(0).width = 3000
    MyDataGrid1.Columns(1).width = 900
    MyDataGrid1.Columns(2).width = 0
    
    Unload Me
    frmFotoAlbumWirdGeladen.Show                                        'Gerbing 27.08.2017
    DoEvents                                                            'Gerbing 27.08.2017
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    lblServerName.Caption = LoadResString(1805 + Sprache)               'lbl SQL server name
    lblDatabasename.Caption = LoadResString(1806 + Sprache)             'lbl datase name
    optWindowsAuthentication.Caption = LoadResString(1807 + Sprache)    'Windows Authentication
    optSqlserverauthentication.Caption = LoadResString(1808 + Sprache)   'SQL server authentication
    lblUsername.Caption = LoadResString(1809 + Sprache)                 'user name
    lblPassword.Caption = LoadResString(1810 + Sprache)                 'password
    btnConnect.Caption = LoadResString(1811 + Sprache)                  'Connect
    lblStandortFotos.Caption = LoadResString(1812 + Sprache)            'lbl Standort der Fotos/Videos:
    txtSQLServerName.Text = PublicSQLServer                             'txt SQL server name:
    txtDatabaseName.Text = PublicSQLDatabase                            'txt Database name:
    txtStandortFotos.Text = PublicLocationFotos                         'txt Standort der Fotos/Videos:
    If PublicWindowsAuthentication = "1" Then
        optWindowsAuthentication.Value = True
        txtUserName.Visible = False                                     'Gerbing 26.10.2016
        txtPassword.Visible = False                                     'Gerbing 26.10.2016
    Else
        optSqlserverauthentication.Value = True
        txtUserName.Visible = True
        txtPassword.Visible = True
    End If
    txtUserName.Text = PublicSQLServerUserName
    txtPassword.UsePasswordChar = True                                  'Gerbing 26.10.2016
    txtPassword.Text = PublicSQLServerPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If gstrLoggedInName = "" Then
        If txtAllowedlicenses.Text < 99 Then
            End
        End If
    End If
End Sub

Private Sub MyDataGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim SQL As String
    Dim SQLalt As String
    Dim SQLneu As String
    Dim Pos As Long
    Dim Links As String
    
    SQLalt = rstsql.Source
    '----------------------------------------------------------------------------------------
    Pos = InStr(1, SQLalt, "DESC", vbTextCompare)
    If Pos <> 0 Then
        SQL = " ORDER BY username;"
    Else
        SQL = " ORDER BY username DESC;"
    End If
    Pos = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    If Pos <> 0 Then
        Links = Left(SQLalt, Pos - 1)
    Else
        Links = SQLalt
        'wenn ein Semikolon am Ende steht, dann abschneiden
        Pos = InStr(1, Links, ";")
        If Pos <> 0 Then
            Links = Mid(Links, 1, Pos - 1)
        End If
    End If
    SQLneu = Links & SQL
    rstsql.Close                                                                        'Gerbing 28.03.2016
    With rstsql                                                                         'Gerbing 28.03.2016
        .Source = SQLneu
        .ActiveConnection = DBado                                                       'Gerbing 28.03.2016
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set MyDataGrid1.DataSource = rstsql                                                 'Gerbing 28.03.2016
    MyDataGrid1.Columns(0).width = 3000
    MyDataGrid1.Columns(1).width = 900
    MyDataGrid1.Columns(2).width = 0
End Sub

Private Sub optSqlserverauthentication_Click()
    lblUsername.Visible = True
    lblPassword.Visible = True
    txtUserName.Visible = True
    txtPassword.Visible = True
End Sub

Private Sub optWindowsAuthentication_Click()
    lblUsername.Visible = False
    lblPassword.Visible = False
    txtUserName.Visible = False
    txtPassword.Visible = False
End Sub

Private Function Kontrolle(strSchlüssel As String)
    Dim n As Integer
    Dim i As String
    Dim lngSumme As Long
    Dim lngKontrolle As Long
    Dim lngRest As Long
    Dim strS1 As String
    Dim strS2 As String
    Dim strS3 As String
    Dim strS4 As String
    Dim strS5 As String
    Dim lngAnzahlNumeric As Integer
    Dim strKontrolle As String
    Dim LizenzUnCodiert As String
    Dim strDrei As String
    Dim LinkeZahl As String
    Dim RechteZahl As String
    Dim Buchstabe As String
    Dim ZähleStriche As Long
    Dim lngStart As Long
    Dim Pos As Long
    Dim Buchst As String
    Dim ErsterB As Boolean
    Dim PrüfB As String
    Dim intRest As Integer
    Dim strName As String
    
    Buchst = "ABCDEFGHIJKLMNOPQRSTUVWYXZ"
    'Die Gültigkeit des Freischalteschlüssels muss geprüft werden.
    
    strName = Right(strSchlüssel, 5)
    
    If Len(strSchlüssel) <> 46 Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    If strSchlüssel = "" Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    ZähleStriche = 0
    lngStart = 1
    Do
        Pos = InStr(lngStart, strSchlüssel, "-", vbTextCompare)
        If Pos = 0 Then Exit Do
        ZähleStriche = ZähleStriche + 1
        lngStart = Pos + 1
    Loop
    If ZähleStriche <> 6 Then                                       'ob 6 Bindestriche vorkommen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    
    strKontrolle = Right(strSchlüssel, Len(strSchlüssel) - 11)
    strS1 = Mid(strKontrolle, 1, 5)
    strS2 = Mid(strKontrolle, 7, 5)
    strS3 = Mid(strKontrolle, 13, 5)
    strS4 = Mid(strKontrolle, 19, 5)
    strS5 = Mid(strKontrolle, 25, 5)
    '-------------------------------------------------------------------------------------------------
    'strS1
    lngAnzahlNumeric = 0
    For n = 1 To 5                              'enthaltenen Zahlen aufsummieren
        i = Mid(strS1, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS2
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS2, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS3
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS3, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS4
    lngAnzahlNumeric = 0
    For n = 1 To 5
        i = Mid(strS4, n, 1)
        If IsNumeric(i) Then
            lngAnzahlNumeric = lngAnzahlNumeric + 1
            lngSumme = lngSumme + i
        End If
    Next n
    If lngAnzahlNumeric = 5 Then                'Man darf es nicht mit lauter Zahlen probieren dürfen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    If lngSumme = 0 Then                        'Wenn jemand lauter Nullen probiert würde es klappen
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    '------------------------------------------------------------------------------------------------
    'strS5
    For n = 1 To 5
        i = Mid(strS5, n, 1)
        If IsNumeric(i) Then
            lngKontrolle = lngKontrolle + i
        End If
    Next n
    lngRest = lngSumme Mod 7
    If lngKontrolle <> lngRest Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------
    'Jetzt wird geprüft,ob der Prüfbuchstabe richtig ist
    lngSumme = 0
    'txtS1
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS1, n, 1)
        If Not IsNumeric(i) Then
            Pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + Pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strS2
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS2, n, 1)
        If Not IsNumeric(i) Then
            Pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + Pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtS3
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS3, n, 1)
        If Not IsNumeric(i) Then
            Pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + Pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtS4
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS4, n, 1)
        If Not IsNumeric(i) Then
            Pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + Pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'die 5. Kolonne dient der Kontrolle
    ErsterB = True
    For n = 1 To 5                              'ersten Buchstaben suchen
        i = Mid(strS5, n, 1)
        If Not IsNumeric(i) Then
            If ErsterB = True Then
                PrüfB = i
            Else
                Pos = InStr(1, Buchst, i)
                lngSumme = lngSumme + Pos
            End If
            ErsterB = False
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtName                                    'Gerbing 13.10.2005
    'alle Buchstaben vom Name
    For n = 1 To Len(strName)                   'alle Buchstaben aufsummieren
        i = Mid(strName, n, 1)
        Pos = InStr(1, Buchst, i)
        lngSumme = lngSumme + Pos
    Next n
    '---------------------------------------------------------------------------------------------------
    'jetzt Prüfbuchstabe ausrechnen
    intRest = lngSumme Mod 26
    intRest = intRest + 1
    If PrüfB <> Mid(Buchst, intRest, 1) Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
        
    LizenzUnCodiert = Crypt(Left(strSchlüssel, 10), Mid(strSchlüssel, 24, 5), False)            'Gerbing 19.11.2012
    LinkeZahl = Mid(LizenzUnCodiert, 1, 2)                                              'Gerbing 21.11.2012
    RechteZahl = Mid(LizenzUnCodiert, 4, 2)                                             'Gerbing 21.11.2012
    Buchstabe = Mid(LizenzUnCodiert, 3, 1)                                              'Gerbing 21.11.2012
    If StrComp(Buchstabe, "S", vbBinaryCompare) <> 0 Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    If Not IsNumeric(LinkeZahl) Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    If Not IsNumeric(RechteZahl) Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    If StrComp(LinkeZahl, RechteZahl, vbBinaryCompare) <> 0 Then
        MsgBox LoadResString(1817 + Sprache)    'invalid LicenseCode. The programm will be stopped
        Exit Function
    End If
    Kontrolle = 1
End Function
#End If
