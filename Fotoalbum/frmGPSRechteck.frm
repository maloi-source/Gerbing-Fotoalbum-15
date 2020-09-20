VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmGPSRechteck 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   10584
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   18924
   Icon            =   "frmGPSRechteck.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10584
   ScaleWidth      =   18924
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtEndGPS 
      Height          =   372
      Left            =   2640
      TabIndex        =   6
      Text            =   "kopieren Sie hierher die End-Geo-Position"
      Top             =   1800
      Width           =   7572
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8292
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   18852
      ExtentX         =   33253
      ExtentY         =   14626
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
   Begin VB.CommandButton btnSuchbereich 
      Caption         =   "Suchbereich festlegen"
      Height          =   372
      Left            =   10440
      TabIndex        =   3
      Top             =   1560
      Width           =   3732
   End
   Begin VB.TextBox txtStartGPS 
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      Text            =   "kopieren Sie hierher die Start-Geo-Position"
      Top             =   1320
      Width           =   7572
   End
   Begin VB.TextBox txtBeschreibung 
      Height          =   1092
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Text            =   "frmGPSRechteck.frx":038A
      Top             =   120
      Width           =   18612
   End
   Begin VB.Label lblEndGPSPosition 
      BackColor       =   &H00C0C0C0&
      Caption         =   "End GPS-Position:"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2532
   End
   Begin VB.Label lblStartGPSPosition 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start GPS-Position:"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2532
   End
End
Attribute VB_Name = "frmGPSRechteck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSuchbereich_Click()
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim Zoom As String
    Dim LatitudeStart As String
    Dim LongitudeStart As String
    Dim LatitudeEnd As String
    Dim LongitudeEnd As String
    Dim strTemp As String
    Dim SQL As String
    Dim msg As String
    Dim rstErsterStart As ADODB.Recordset

    'testweise
    'txtStartGPS = "https://www.openstreetmap.org/#map=14/50.8452/12.8058"

    pos = InStr(1, txtStartGPS, "=")
    If pos <> 0 Then
        pos1 = InStr(pos, txtStartGPS, "/")
    End If
    If pos1 <> 0 Then
        Zoom = Mid(txtStartGPS, pos + 1, pos1 - pos - 1)
        pos2 = InStr(pos1, txtStartGPS, ".")
        If pos2 <> 0 Then
            pos3 = InStr(pos2, txtStartGPS, "/")
            If pos3 <> 0 Then
                LatitudeStart = Mid(txtStartGPS, pos1 + 1, pos3 - pos1 - 1)
            End If
        End If
    End If
    If pos3 <> 0 Then
        LongitudeStart = Mid(txtStartGPS, pos3 + 1, Len(txtStartGPS) - pos3)
    End If
    If Not IsNumeric(Zoom) Then
        MsgBox "Zoom='" & Zoom & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(LatitudeStart) Then
        MsgBox "LatitudeStart='" & LatitudeStart & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(LongitudeStart) Then
        MsgBox "LongitudeStart='" & LongitudeStart & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------
        pos = InStr(1, txtEndGPS, "=")
    If pos <> 0 Then
        pos1 = InStr(pos, txtEndGPS, "/")
    End If
    If pos1 <> 0 Then
        Zoom = Mid(txtEndGPS, pos + 1, pos1 - pos - 1)
        pos2 = InStr(pos1, txtEndGPS, ".")
        If pos2 <> 0 Then
            pos3 = InStr(pos2, txtEndGPS, "/")
            If pos3 <> 0 Then
                LatitudeEnd = Mid(txtEndGPS, pos1 + 1, pos3 - pos1 - 1)
            End If
        End If
    End If
    If pos3 <> 0 Then
        LongitudeEnd = Mid(txtEndGPS, pos3 + 1, Len(txtEndGPS) - pos3)
    End If
    If Not IsNumeric(Zoom) Then
        MsgBox "Zoom='" & Zoom & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(LatitudeEnd) Then
        MsgBox "LatitudeEnd='" & LatitudeEnd & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    If Not IsNumeric(LongitudeEnd) Then
        MsgBox "LongitudeEnd='" & LongitudeEnd & LoadResString(1142 + Sprache) 'das ist keine Zahl
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------
    'StartpunktX(longitude=Längengrad) muss kleiner sein als EndpunktX, sonst Fehler
    'StartpunktY(Latitude=Breitengrad) muss größer sein als EndpunktY, sonst Fehler
    If CLng(LongitudeStart) < CLng(LongitudeEnd) Then
        '
    Else
        MsgBox LoadResString(1153 + Sprache) 'Sie haben kein Rechteck definiert. Die Start-Longitude muss kleiner sein als die End-Longitude
    End If
    If CLng(LatitudeStart) > CLng(LatitudeEnd) Then
        '
    Else
        MsgBox LoadResString(1154 + Sprache) 'Sie haben kein Rechteck definiert. Die Start-Latitude muss größer sein als die End-Latitude
    End If
    gstrGEOStartPunkt = ""
    gstrGEOEndPunkt = ""
    'zB     gstrGEOStartPunkt=50.83517,12.81463         GPSLatitude,GPSLongitude   Breite,Länge
    '       gstrGEOEndPunkt=50.83017,12.82298
    gstrGEOStartPunkt = LatitudeStart & "," & LongitudeStart
    gstrGEOEndPunkt = LatitudeEnd & "," & LongitudeEnd
    '---------------------------------------------------------------------------------------
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
        rstErsterStart.Fields("LetzterGEOPunkt") = LatitudeStart & "," & LongitudeStart
        rstErsterStart.Fields("ZoomListIndex") = Zoom
        rstErsterStart.Update
    End If
    rstErsterStart.Close
    msg = "Zoom=" & Zoom & vbNewLine
    msg = msg & "LatitudeStart=" & LatitudeStart & ",LongitudeStart=" & LongitudeStart & vbNewLine
    msg = msg & "LatitudeEnd=" & LatitudeEnd & ",LongitudeEnd=" & LongitudeEnd & vbNewLine
    MsgBox msg
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strZoom As String
    Dim strGeoPosition As String
    Dim SQL As String
    Dim rstErsterStart As ADODB.Recordset
    
    Call AnpassenNutzerWunsch(Me)
    Me.Caption = LoadResString(1148 + Sprache)                      'Geo-Suchbereich festlegen
    'Ihre Aufgabe ist es, ein virtuelles Rechteck zu definieren. Die Start-Geo-Position beginnt links oben, die End-Geo-Position endet rechts unten.
    'Suchen Sie auf der Karte die gewünschte Start-Geo-Position. Dann klicken Sie mit der rechten Maustaste auf diesen Punkt und wählen Sie 'Adresse anzeigen'.
    'Darauf erscheint auf der linken Seite die Geo-Position dieses Punktes. Klicken Sie jetzt mit der rechten Maustaste auf die Geo-Position und wählen Sie 'Verknüpfung kopieren'.
    'Dann klicken Sie ins Feld Start-Geo-Position und fügen den Inhalt der Zwischenablage ein mit Strg+V
    'Wiederholen Sie alles für das Feld End-Geo-Position
    'Zuletzt klicken Sie auf den Button 1145
    txtBeschreibung.Text = LoadResString(1152 + Sprache)                        'Ihre Aufgabe ist es...
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1151 + Sprache)     'Suchen Sie auf der Karte die gewünschte Start-Geo-Position...
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1137 + Sprache)     'Darauf erscheint auf der linken Seite...
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1149 + Sprache)     'Dann klicken Sie ins Feld Start-Geo-Position...
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1150 + Sprache)     'Wiederholen Sie alles für das Feld End-Geo-Position
    txtBeschreibung = txtBeschreibung & " " & LoadResString(1159 + Sprache) & LoadResString(1145 + Sprache) & "'"
    lblStartGPSPosition.Caption = LoadResString(1143 + Sprache)     'Start-GPS-Position
    lblEndGPSPosition.Caption = LoadResString(1144 + Sprache)       'End-GPS-Position
    btnSuchbereich.Caption = LoadResString(1145 + Sprache)          'Suchbereich festlegen
    txtStartGPS.Text = LoadResString(1146 + Sprache)                'kopieren Sie hierher die Start-Geo-Position
    txtEndGPS.Text = LoadResString(1147 + Sprache)                  'kopieren Sie hierher die End-Geo-Position
    
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

