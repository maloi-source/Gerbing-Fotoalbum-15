VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form frmConnectSQL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "GERBING Fotoalbum Connect sql server"
   ClientHeight    =   4872
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6540
   Icon            =   "frmConnectSQL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4872
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtSQLServerName 
      Height          =   372
      Left            =   2400
      TabIndex        =   14
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
      CueBanner       =   "frmConnectSQL.frx":038A
      Text            =   "frmConnectSQL.frx":03AA
   End
   Begin EditCtlsLibUCtl.TextBox txtDatabaseName 
      Height          =   372
      Left            =   2400
      TabIndex        =   13
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
   Begin EditCtlsLibUCtl.TextBox txtStandortFotos 
      Height          =   372
      Left            =   120
      TabIndex        =   10
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
      CueBanner       =   "frmConnectSQL.frx":040A
      Text            =   "frmConnectSQL.frx":042A
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
         TabIndex        =   12
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
         CueBanner       =   "frmConnectSQL.frx":044A
         Text            =   "frmConnectSQL.frx":046A
      End
      Begin EditCtlsLibUCtl.TextBox txtPassword 
         Height          =   372
         Left            =   4920
         TabIndex        =   11
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
         UsePasswordChar =   0   'False
         UseSystemFont   =   0   'False
         CueBanner       =   "frmConnectSQL.frx":048A
         Text            =   "frmConnectSQL.frx":04AA
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
    Dim db As ADODB.Connection
    Dim rst As ADODB.Recordset
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

    Me.MousePointer = vbHourglass
    DoEvents
    Set db = New ADODB.Connection
    With db
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
    
    Me.MousePointer = vbNormal
    Set rst = New ADODB.Recordset
    With rst
        .Source = "select * from loggedinusers"
        .ActiveConnection = db
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description & vbNewLine & vbNewLine
            
            msg = msg & txtSQLServerName.Text & " " & txtDatabaseName.Text & " missing table LoggedInUsers"
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
            Exit Sub
        End If
    End With
    
    'Ermitteln der erlaubten Lizenzen aus der Tabelle LicenseCode
    Set rst1 = New ADODB.Recordset
    With rst1
        .Source = "select * from LicenseCode"
        .ActiveConnection = db
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description
            MsgBox msg
            Exit Sub
        End If
    End With
    On Error GoTo 0
    strSchlüssel = rst1.Fields("LicenseCode")
    rst1.Close
    LizenzUnCodiert = Crypt(Left(strSchlüssel, 10), Mid(strSchlüssel, 24, 5), False)
    rc = Kontrolle(strSchlüssel)
    If rc = False Then End                      'Programm beendet bei falschem LicenseCode
    '---------------------------------------------------------------------------------------------------
    LizenzUnCodiert = Crypt(Left(strSchlüssel, 10), Mid(strSchlüssel, 24, 5), False)
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
    
    'Frage ob im ersten Satz der Tabelle Fotos der Spaltenname Filename vorkommt
    'SQL = "SELECT * From fotos WHERE not filename Is Null;"
    Set rst1 = New ADODB.Recordset
    On Error Resume Next
    With rst1
        .Source = "SELECT * From fotos WHERE not filename Is Null;"
        .ActiveConnection = db
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

    Err.Number = 0                                                                      'Gerbing 14.02.2014
    'Ermitteln des ersten Satzes aus Tabelle Fotos
    'SELECT TOP 1 * FROM fotos
    Set rst1 = New ADODB.Recordset
    With rst1
        .Source = "SELECT TOP 1 * FROM Fotos"
        .ActiveConnection = db
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        If Err.Number <> 0 Then
            msg = "error number=" & Err.Number & vbNewLine
            msg = msg & "error text=" & Err.Description
            MsgBox msg
            Exit Sub
        End If
    End With
    On Error GoTo 0
    If rst1.EOF Then
        'MsgBox "Tabelle fotos ist leer"
        Unload Me
        Exit Sub
    End If
    strSchlüssel = rst1.Fields(LoadResString(1028 + Sprache))   '1028=Dateiname
    rst1.Close
    Set rst1 = db.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, "Fotos")) '2529=fotos
    strPrimaryKey = rst1.Fields("COLUMN_NAME").Value
    If strPrimaryKey = "" Then
        'MsgBox "Die Tabelle fotos enthält keinen Primärschlüssel. Das Programm wird beendet."
        MsgBox LoadResString(1823 + Sprache)
        End
    End If
    If StrComp(LoadResString(1028 + Sprache), strPrimaryKey, vbTextCompare) <> 0 Then       '1028=Dateiname
        'MsgBox "Die Spalte Dateiname ist nicht der Primärschlüssel. Das Programm wird beendet."
        MsgBox LoadResString(1824 + Sprache)
        End
    End If
    'Prüfen ob es den im ersten Satz enthaltenen Dateiname gibt
'    PublicLocationFotos = txtStandortFotos.Text
'    Call WriteSLF(txtStandortFotos.Text)    'Schreibe [SQL] LocationFotos
    strSchlüssel = Replace(strSchlüssel, "+:\", txtStandortFotos.Text & "\")
    On Error Resume Next
    'strTemp = Dir(strSchlüssel)
    If file_path_exist(strSchlüssel) = False Or Err.Number <> 0 Then
    'If strTemp = "" Or Err.Number <> 0 Then
        Call WriteSLF("")    'Schreibe [SQL] LocationFotos
        'msg = Feldname & " existiert nicht." & vbNewLine
        msg = strSchlüssel & LoadResString(2162 + Sprache) & vbNewLine
        'msg = msg & "Prüfen Sie, ob die Angabe " & txtStandortFotos.Text  & " richtig ist" & vbNewLine & vbNewLine
        msg = msg & LoadResString(1820 + Sprache) & txtStandortFotos.Text & LoadResString(1821 + Sprache) & vbNewLine & vbNewLine
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
        End
    End If
    PublicLocationFotos = txtStandortFotos.Text
    Call WriteSLF(txtStandortFotos.Text)    'Schreibe [SQL] LocationFotos
    Unload Me
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
    Else
        optSqlserverauthentication.Value = True
    End If
    txtUserName.Text = PublicSQLServerUserName
    txtPassword.Text = PublicSQLServerPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst.Close
    db.Close
    '   Datenbank-Objekt entfernen
    Set db = Nothing
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
    Dim pos As Long
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
        pos = InStr(lngStart, strSchlüssel, "-", vbTextCompare)
        If pos = 0 Then Exit Do
        ZähleStriche = ZähleStriche + 1
        lngStart = pos + 1
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
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'strS2
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS2, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtS3
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS3, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtS4
    For n = 1 To 5                              'enthaltene Buchstaben aufsummieren
        i = Mid(strS4, n, 1)
        If Not IsNumeric(i) Then
            pos = InStr(1, Buchst, i)
            lngSumme = lngSumme + pos
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
                pos = InStr(1, Buchst, i)
                lngSumme = lngSumme + pos
            End If
            ErsterB = False
        End If
    Next n
    '-------------------------------------------------------------------------------------------------
    'txtName                                    'Gerbing 13.10.2005
    'alle Buchstaben vom Name
    For n = 1 To Len(strName)                   'alle Buchstaben aufsummieren
        i = Mid(strName, n, 1)
        pos = InStr(1, Buchst, i)
        lngSumme = lngSumme + pos
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

