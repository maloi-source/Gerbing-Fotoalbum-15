VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmGPSInDatenbankEintragen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   10584
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   18924
   Icon            =   "frmGPSInDatenbankEintragen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10584
   ScaleWidth      =   18924
   StartUpPosition =   1  'Fenstermitte
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8892
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   18852
      ExtentX         =   33253
      ExtentY         =   15684
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton btnEintragen 
      Caption         =   "Eintragen in die Datenbank"
      Height          =   372
      Left            =   10440
      TabIndex        =   3
      Top             =   1200
      Width           =   3732
   End
   Begin VB.TextBox txtGPS 
      Height          =   372
      Left            =   2160
      TabIndex        =   2
      Text            =   "kopieren Sie hierher die Geo-Position"
      Top             =   1200
      Width           =   7572
   End
   Begin VB.TextBox txtBeschreibung 
      Height          =   972
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Text            =   "frmGPSInDatenbankEintragen.frx":038A
      Top             =   120
      Width           =   18612
   End
   Begin VB.Label lblGPSPosition 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GPS-Position:"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1932
   End
End
Attribute VB_Name = "frmGPSInDatenbankEintragen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEintragen_Click()
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim Zoom As String
    Dim Latitude As String
    Dim Longitude As String
    Dim strTemp As String
    Dim SQL As String
    Dim Msg As String
    Dim rstErsterStart As ADODB.Recordset
    Dim rstsql As ADODB.Recordset
    
    'testweise
    'txtGPS = "https://www.openstreetmap.org/#map=14/50.8452/12.8058"
    
    pos = InStr(1, txtGPS, "=")
    If pos <> 0 Then
        pos1 = InStr(pos, txtGPS, "/")
    End If
    If pos1 <> 0 Then
        Zoom = Mid(txtGPS, pos + 1, pos1 - pos - 1)
        pos2 = InStr(pos1, txtGPS, ".")
        If pos2 <> 0 Then
            pos3 = InStr(pos2, txtGPS, "/")
            If pos3 <> 0 Then
                Latitude = Mid(txtGPS, pos1 + 1, pos3 - pos1 - 1)
            End If
        End If
    End If
    If pos3 <> 0 Then
        Longitude = Mid(txtGPS, pos3 + 1, Len(txtGPS) - pos3)
    End If
    If Not IsNumeric(Zoom) Then
        MsgBox "Zoom='" & Zoom & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(Latitude) Then
        MsgBox "Latitude='" & Latitude & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(Longitude) Then
        MsgBox "Longitude='" & Longitude & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    '----------------------------------------------------------------------------
    If gblnSchreibgeschützt = False Then
        'Jetzt eintragen in die Tabelle Fotos die Felder GPSLatitude und GPSLongitude
        'in gstrFRODN steht der Dateiname des zu updatenden Records                       'Gerbing 04.05.2015
        strTemp = Replace(gstrFRODN, gstrFotosMdbLocation & "\", "+:\")
        strTemp = Replace(strTemp, "'", "''")                                               'Gerbing 23.01.2018
        'Bei Dateinamen mit Hochkomma bringt Open Recordset Laufzeitfehler -> ersetzen durch 2 Hochkommas  'Gerbing 23.01.2018
        'SQL = "SELECT * From Fotos Where Dateiname = " & strTemp
        SQL = "SELECT * From Fotos Where " & LoadResString(1028 + Sprache) & "='" & strTemp & "'"
        Set rstsql = New ADODB.Recordset
        With rstsql
            .Source = SQL
            .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        rstsql.Fields("GPSLongitude") = Longitude
        rstsql.Fields("GPSLatitude") = Latitude
        rstsql.Update
        rstsql.Close
        'Jetzt eintragen in die Tabelle ErsterStart die Felder LetzterGEOPunkt und ZoomListIndex
        On Error Resume Next
        SQL = "select * From ErsterStart;"
        Set rstErsterStart = New ADODB.Recordset
        With rstErsterStart
            .Source = SQL
            .ActiveConnection = DBado                                                   'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        If Err.Number = 0 Then
            'LetzterGEOPunkt zB 51.51558,11.7443
            rstErsterStart.Fields("LetzterGEOPunkt") = Latitude & "," & Longitude
            rstErsterStart.Fields("ZoomListIndex") = Zoom
            rstErsterStart.Update
        End If
        rstErsterStart.Close
        Msg = "Zoom=" & Zoom & vbNewLine
        Msg = Msg & "Latitude=" & Latitude & vbNewLine
        Msg = Msg & "Longitude=" & Longitude
        MsgBox Msg
    End If
End Sub

Private Sub Form_Load()
    Dim strZoom As String
    Dim strGeoPosition As String
    Dim SQL As String
    Dim rstErsterStart As ADODB.Recordset
    
    Call AnpassenNutzerWunsch(Me)
    Me.Caption = LoadResString(1135 + Sprache)                  'Geo-Position in Datenbank eintragen
    'Suchen Sie auf der Karte die gewünschte Geo-Position. Dann klicken Sie mit der rechten Maustaste auf diesen Punkt und wählen Sie 'Adresse anzeigen'.
    'Darauf erscheint auf der linken Seite die Geo-Position dieses Punktes. Klicken Sie jetzt mit der rechten Maustaste auf die Geo-Position und wählen Sie 'Verknüpfung kopieren'.
    'Dann klicken Sie ins Feld Geo-Position und fügen den Inhalt der Zwischenablage ein mit Strg+V
    'Dann klicken Sie auf den Button 1135
    txtBeschreibung.Text = LoadResString(1136 + Sprache)
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1137 + Sprache)
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1138 + Sprache)
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1158 + Sprache) & LoadResString(1140 + Sprache) & "'"
    lblGPSPosition.Caption = LoadResString(1139 + Sprache)      'GPS-Position
    btnEintragen.Caption = LoadResString(1140 + Sprache)        'Eintragen in die Datenbank
    txtGPS.Text = LoadResString(1141 + Sprache)                 'kopieren Sie hierher die Geo-Position
    
    'Die WebBrowser-Anfangseinstellung kommt aus Tabelle ErsterStart
    'Felder LetzterGEOPunkt zB 51.51558,11.74438
    'und
    'ZoomListIndex          zB 14
    'Wenn dort nichts steht startet es mit der Weltkugel
    'https://www.openstreetmap.org/#map=2/0/0
    'ansonsten zB https://www.openstreetmap.org/#map=14/50.8221/12.9084
    
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 14.2.2 gibt es in der Tabelle ErsterStart das Feld LetzterGEOPunkt und ZoomListIndex        'Gerbing 05.09.2016
    'Und in der Tabelle fotos die Felder GPSLatitude und GPSLongitude
    On Error Resume Next
    SQL = "select * From ErsterStart;"
    Set rstErsterStart = New ADODB.Recordset
    With rstErsterStart
        .Source = SQL
        .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If Err.Number <> 0 Then
        'hier existiert das Feld LetzterGEOPunkt nicht
        If gblnSchreibgeschützt = False Then
            If gblnSQLServerVersion = True Then
                'SQL Server
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
        'hier ist Spalte Geoposition nicht vorhanden ich benutze die Weltkugel
        WebBrowser1.Navigate "https://www.openstreetmap.org/#map=2/0/0"
    Else
        'hier existiert das Feld LetzterGEOPunkt und ZoomListIndex
        strGeoPosition = "0/0"
        If rstErsterStart.Fields("LetzterGEOPunkt") = "" Or IsNull(rstErsterStart.Fields("LetzterGEOPunkt")) Then
            'hier ist LetzterGEOPunkt leer ich benutze die Weltkugel
            strZoom = "2"
        Else
            strGeoPosition = rstErsterStart.Fields("LetzterGEOPunkt")
        End If
        If rstErsterStart.Fields("ZoomListIndex") = "" Or IsNull(rstErsterStart.Fields("ZoomListIndex")) Then
            strZoom = "2"                                                               'ich benutze die Weltkugel
        Else
            strZoom = (rstErsterStart.Fields("ZoomListIndex"))
        End If
        strGeoPosition = Replace(strGeoPosition, ",", "/")                              'verwandle 51.51558,11.74438 in 51.51558/11.74438
        'WebBrowser1.Navigate "https://www.openstreetmap.org/#map=2/0/0"
        WebBrowser1.Navigate "https://www.openstreetmap.org/#map=" & strZoom & "/" & strGeoPosition
        'WebBrowser1.Navigate "https://www.openstreetmap.org/#map=14/50.8221/12.9084"
    End If
    rstErsterStart.Close
    On Error GoTo 0
End Sub

