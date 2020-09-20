VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Begin VB.Form AendernForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Ändern oder Löschen des Dateinamens"
   ClientHeight    =   6336
   ClientLeft      =   276
   ClientTop       =   360
   ClientWidth     =   11664
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6336
   ScaleWidth      =   11664
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtNeuerName 
      Height          =   372
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   9012
      _cx             =   15896
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
      CueBanner       =   "AendForm.frx":0000
      Text            =   "AendForm.frx":0020
   End
   Begin EditCtlsLibUCtl.TextBox txtAlterName 
      Height          =   372
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   9012
      _cx             =   15896
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
      CueBanner       =   "AendForm.frx":0040
      Text            =   "AendForm.frx":0060
   End
   Begin VB.PictureBox PictureThumb 
      Height          =   3492
      Left            =   2520
      ScaleHeight     =   3444
      ScaleWidth      =   5484
      TabIndex        =   7
      Top             =   2400
      Width           =   5532
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   8760
      ScaleHeight     =   324
      ScaleWidth      =   324
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnAbbrechen 
      Caption         =   "Ab&brechen"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton btnLöschen 
      Caption         =   "&Löschen"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton btnÄndernDateiname 
      Caption         =   "&Ändern"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   9480
      Top             =   1560
      Width           =   372
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "neuer Name:"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "alter Name:"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2172
   End
End
Attribute VB_Name = "AendernForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NL As String
    Dim msg As String
    Dim outname As String                                                       'Gerbing 06.04.2017
    '----------------------------------------------ab hier für Video Thumbnails-------Gerbing 06.04.2017
    'From Windows SDK header file propkey.h:
    Private Const SCID_THUMBNAILSTREAM As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
    'Requires Windows XP SP2 or later:
    Private Declare Function PropVariantToVariant Lib "propsys" ( _
        ByRef PropVar As Any, _
        ByRef Var As Variant) As Long
    Private ShellObject As shell32.Shell
    Private Const SCID_PerceivedType As String = "{28636AA6-953D-11D2-B5D6-00C04FD918D0},9"
    Private Const SCID_PropStream As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
    
    Private Enum PERCEIVED
        PERCEIVED_TYPE_FIRST = -3
        PERCEIVED_TYPE_CUSTOM = -3
        PERCEIVED_TYPE_UNSPECIFIED = -2
        PERCEIVED_TYPE_FOLDER = -1
        PERCEIVED_TYPE_UNKNOWN = 0
        PERCEIVED_TYPE_TEXT = 1
        PERCEIVED_TYPE_IMAGE = 2
        PERCEIVED_TYPE_AUDIO = 3
        PERCEIVED_TYPE_VIDEO = 4
        PERCEIVED_TYPE_COMPRESSED = 5
        PERCEIVED_TYPE_DOCUMENT = 6
        PERCEIVED_TYPE_SYSTEM = 7
        PERCEIVED_TYPE_APPLICATION = 8
        PERCEIVED_TYPE_GAMEMEDIA = 9
        PERCEIVED_TYPE_CONTACTS = 10
        PERCEIVED_TYPE_LAST = 10
    End Enum
    Private GdipTool As GdipTool


Private Sub btnAbbrechen_Click()
    btnÄndernDateiname.Enabled = True
    btnLöschen.Enabled = True
    On Error Resume Next
    btnÄndernDateiname.Default = True
    btnÄndernDateiname.SetFocus
    Unload Me
End Sub

Private Sub btnÄndernDateiname_Click()
    Dim NeuesVerzeichnis As String
    Dim start As Integer
    Dim Pos1 As Integer
    Dim strTempAlt As String
    Dim strTempNeu As String
    Dim strTempThumb As String
    Dim DateinameFoto As String
    Dim temp As String
    Dim temp1 As String
    Dim AudioDateiNeu As String
    Dim SQL As String
    Dim rc As Boolean
    Dim altThumb As String                                                          'Gerbing 08.12.2016
    Dim ÄndernName As String                                                        'Gerbing 23.01.2018
    Dim neuThumb As String
    Dim strPath As String
    Dim strFile As String
    Dim DateinamenErweiterung As String
    
    'Die Änderung wird nur gemacht, wenn kein neues Verzeichnis angelegt werden muß
    'd.h. wenn das Verzeichnis existiert
    If gblnSQLServerVersion = True Then
        strTempNeu = Replace(txtNeuerName.Text, "+:\", PublicLocationFotos & "\")    'Gerbing 11.04.2005
        strTempAlt = Replace(txtAlterName.Text, "+:\", PublicLocationFotos & "\")    'Gerbing 11.04.2005
    Else
        strTempNeu = Replace(txtNeuerName.Text, "+:\", AppPath & "\")    'Gerbing 11.04.2005
        strTempAlt = Replace(txtAlterName.Text, "+:\", AppPath & "\")    'Gerbing 11.04.2005
    End If
    start = 1
    Do
        Pos1 = InStr(start, strTempNeu, "\")
        If Pos1 = 0 Then Exit Do
        start = Pos1 + 1
    Loop
    NeuesVerzeichnis = Left(strTempNeu, start - 1)
    temp = ""
    ÄndernName = Replace(txtAlterName.Text, "'", "''")          'Gerbing 23.01.2018 Wo ein Hochkomma vorkommt, nach 2 Hochkommas suchen
    On Error Resume Next
    'temp = Dir(NeuesVerzeichnis, vbDirectory)
    If file_path_exist(NeuesVerzeichnis) = False Then
    'If temp = "" Then
'        msg = "In neuer Name ist ein falsches Verzeichnis angegeben" & NL
'        msg = msg & "Das Programm kann keine neuen Verzeichnisse anlegen"
        msg = LoadResString(2401 + Sprache) & NL
        msg = msg & LoadResString(2402 + Sprache)
        MsgBox msg, vbInformation
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------
    'Wenn der zum Dateinamen gehörende Punkt gelöscht wird, soll eine Warnung kommen        'Gerbing 22.01.2008
    'Die Dateinamenerweiterung darf 3,4,5 stellig sein
    'Wenn dort kein Punkt steht, kommt eine Warnung
    If Mid(strTempAlt, Len(strTempAlt) - 3, 1) = "." Then
        If Mid(strTempNeu, Len(strTempNeu) - 3, 1) <> "." Then
            'msg = "Änderung abgelehnt-Sie haben den zum Dateinamen gehörenden Punkt gelöscht"
            msg = LoadResString(2422 + Sprache)
            MsgBox msg
            Exit Sub
        End If
        GoTo PunktGefunden                                                                  'Gerbing 04.10.2012
    End If
    If Mid(strTempAlt, Len(strTempAlt) - 4, 1) = "." Then
        If Mid(strTempNeu, Len(strTempNeu) - 4, 1) <> "." Then
            'msg = "Änderung abgelehnt-Sie haben den zum Dateinamen gehörenden Punkt gelöscht"
            msg = LoadResString(2422 + Sprache)
            MsgBox msg
            Exit Sub
        End If
        GoTo PunktGefunden                                                                  'Gerbing 04.10.2012
    End If
    If Mid(strTempAlt, Len(strTempAlt) - 5, 1) = "." Then
        If Mid(strTempNeu, Len(strTempNeu) - 5, 1) <> "." Then
            'msg = "Änderung abgelehnt-Sie haben den zum Dateinamen gehörenden Punkt gelöscht"
            msg = LoadResString(2422 + Sprache)
            MsgBox msg
            Exit Sub
        End If
    End If
PunktGefunden:                                                                              'Gerbing 04.10.2012
    '---------------------------------------------------------------------------------------
    'txtNeuerName muß in eine Rename-Operation einfließen
    'Name strTempAlt As strTempNeu    'rename altername As neuername     'Gerbing 11.04.2005
    rc = NameAs(strTempAlt, strTempNeu)
    '---------------------------------------------------------------------------------------
    'jetzt untersuchen, ob rename altername As neuername erfolgreich war
    If rc = False Then
        If Err.Number = 52 Then                     '52=Dateiname falsch
            msg = "Errornumber=" & Err.Number & NL
            msg = msg & Err.Description & NL
'            msg = msg & "neuer Name: " & txtNeuerName.Text & NL & NL
'            msg = msg & "Möglicherweise benutzen Sie verbotene Zeichen im Dateiname" & NL
'            msg = msg & "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
'
'            msg = msg & "oder Sie müssen das Programm mit höheren Rechten starten"                     'Gerbing 26.08.2012

            msg = msg & LoadResString(1607 + Sprache) & " " & txtNeuerName.Text & NL & NL
            msg = msg & LoadResString(2403 + Sprache) & NL
            msg = msg & LoadResString(1367 + Sprache) & NL & NL
            
            msg = msg & LoadResString(2423 + Sprache)                                                   'Gerbing 26.08.2012
            'MsgBox Msg, vbInformation
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
            Unload Me
            Exit Sub
        End If
        If Err.Number = 53 Then                     '53=Datei nicht gefunden
            temp = Dir(txtAlterName.Text)
            If temp = "" Then
                msg = "Errornumber=" & Err.Number & NL
                msg = msg & Err.Description & NL
                msg = msg & LoadResString(1606 + Sprache) & " " & txtAlterName.Text & NL & NL
'
'                msg = msg & "Sie können mit der Funktion 'Prüfen1' in FotosMdb.exe kontrollieren, ob noch " & NL
'                msg = msg & "weitere Dateien nicht an ihrem angegebenen Standort stehen"
                
                msg = msg & LoadResString(2404 + Sprache) & NL
                msg = msg & LoadResString(2405 + Sprache)
                'MsgBox Msg, vbInformation
                MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
                Unload Me
                Exit Sub
            Else
                msg = "Errornumber=" & Err.Number & NL
                msg = msg & Err.Description & NL
                msg = msg & LoadResString(1606 + Sprache) & " " & txtAlterName.Text & NL
                msg = msg & LoadResString(1607 + Sprache) & " " & txtNeuerName.Text & NL & NL
    '
    '            msg = msg & "Möglicherweise benutzen Sie verbotene Zeichen im Dateiname" & NL
    '            msg = msg & "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
    '
    '            msg = msg & "oder Sie müssen das Programm mit höheren Rechten starten"                     'Gerbing 26.08.2012
    
                msg = msg & LoadResString(1607 + Sprache) & " " & txtNeuerName.Text & NL & NL
                msg = msg & LoadResString(2403 + Sprache) & NL
                msg = msg & LoadResString(1367 + Sprache) & NL & NL
                
                msg = msg & LoadResString(2423 + Sprache)                                                   'Gerbing 26.08.2012
                'MsgBox Msg, vbInformation
                MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
                Unload Me
                Exit Sub
            End If
        End If
            msg = "Errornumber=" & Err.Number & NL
            msg = msg & Err.Description & NL
            msg = msg & LoadResString(1606 + Sprache) & " " & txtAlterName.Text & NL
            msg = msg & LoadResString(1607 + Sprache) & " " & txtNeuerName.Text & NL & NL
'            msg = msg & "Möglicherweise benutzen Sie verbotene Zeichen im Dateiname" & NL
'            msg = msg & "Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?""<>|"
'
'            msg = msg & "oder Sie müssen das Programm mit höheren Rechten starten"                     'Gerbing 26.08.2012

            msg = msg & LoadResString(1607 + Sprache) & " " & txtNeuerName.Text & NL & NL
            msg = msg & LoadResString(2403 + Sprache) & NL
            msg = msg & LoadResString(1367 + Sprache) & NL & NL
            
            msg = msg & LoadResString(2423 + Sprache)                                                   'Gerbing 26.08.2012
            'MsgBox Msg, vbInformation
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
            Unload Me
            Exit Sub
    End If
    '---------------------------------------------
    'rename altername As neuername war erfolgreich
    'Jetzt eine eventuell vorhandene Datei in ...\GerbingThumbs\... ebenfalls umnennen          'Gerbing 08.12.2016
    'Seit Version 14.2.2 Beim übereinstimmenden Ändern von Datensätzen und Foto, ändere ich auch zugehörige Thumbnails
    'file_split splits a complete file name into directory, file name and extension:
    Call file_split(strTempAlt, strPath, strFile, DateinamenErweiterung)
    altThumb = strPath & "GerbingThumbs\" & strFile & "." & DateinamenErweiterung & ".jpg"
    Call file_split(strTempNeu, strPath, strFile, DateinamenErweiterung)
    neuThumb = strPath & "GerbingThumbs\" & strFile & "." & DateinamenErweiterung & ".jpg"
    On Error Resume Next
    'Name alt As neu    'rename altername As neuername
    rc = NameAs(altThumb, neuThumb)
    If rc = False Then
        'msg = "Fehler beim Umnennen von" & " " & temp & " in " & AudioDateiNeu & NL
        msg = LoadResString(2419 + Sprache) & NL
        msg = msg & altThumb & " in " & NL
        msg = msg & neuThumb & NL
        msg = msg & "Errornumber=" & Err.Number & NL
        msg = msg & "Errortext=" & Err.Description & NL
        'MsgBox Msg
        MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
    End If
    On Error GoTo 0
    '---------------------------------------------
    'rename altername As neuername war erfolgreich
    'Jetzt einen eventuell vorhandenen Audio-Kommentar ebenfalls umnennen   'Gerbing 12.04.2006
    On Error GoTo 0
    DateinameFoto = ErmittleDateiname(strTempAlt)
    AudioDateiNeu = ErmittleDateiname(strTempNeu)
'    Temp = Dir(DateinameFoto & ".mp3")                                     'Gerbing 11.09.2013
'    temp1 = Dir(DateinameFoto & ".wav")
'    If Temp <> "" Or temp1 <> "" Then
    If file_path_exist(DateinameFoto & ".mp3") = True Then
    'If Temp <> "" Then
        temp = DateinameFoto & ".mp3"
    End If
    If file_path_exist(DateinameFoto & ".wav") = True Then
        temp = DateinameFoto & ".wav"
    End If
    If temp <> "" Then
        If temp <> "" Then
            temp = DateinameFoto & ".mp3"
            AudioDateiNeu = AudioDateiNeu & ".mp3"
        Else
            temp = DateinameFoto & ".wav"
            AudioDateiNeu = AudioDateiNeu & ".wav"
        End If
        On Error Resume Next
        'Name temp As AudioDateiNeu    'rename altername As neuername     'Gerbing 12.04.2006
        rc = NameAs(temp, AudioDateiNeu)
        If rc = False Then
            'msg = "Fehler beim Umnennen von" & " " & temp & " in " & AudioDateiNeu & NL
            msg = LoadResString(2419 + Sprache) & NL
            msg = msg & temp & " in " & NL
            msg = msg & AudioDateiNeu & NL
            msg = msg & "Errornumber=" & Err.Number & NL
            msg = msg & "Errortext=" & Err.Description & NL
            'MsgBox Msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
        End If
        On Error GoTo 0
    End If
    '---------------------------------------------------
    'Ändern in der Datenbank
    btnÄndernDateiname.Enabled = False
    btnLöschen.Enabled = False
    'SQL = "select * from fotos where Dateiname = '" & ÄndernName & "'"
    SQL = "select * from Fotos where " & LoadResString(1028 + Sprache) & " = '" & ÄndernName & "'" 'Gerbing 23.01.2018
    On Error Resume Next
    Renam.rstsql.Close
    On Error GoTo 0
    With Renam.rstsql
        .ActiveConnection = Renam.DBado
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    'Renam.rstsql.Fields("Dateiname") = txtNeuerName.Text
    Renam.rstsql.Fields(LoadResString(1028 + Sprache)) = txtNeuerName.Text
    'Neues Feld DateinameKurz                     'Gerbing 10.10.2004
    start = 1
    Do
        Pos1 = InStr(start, txtNeuerName.Text, "\")
        If Pos1 = 0 Then Exit Do
        start = Pos1 + 1
    Loop
    'Renam.rstsql.Fields("DateinameKurz") = Right(txtNeuerName.Text, Len(txtNeuerName.Text) - Start + 1)
    Renam.rstsql.Fields(LoadResString(1031 + Sprache)) = Right(txtNeuerName.Text, Len(txtNeuerName.Text) - start + 1)   'Gerbing 04.02.2020
    Renam.rstsql.Update
    
    btnÄndernDateiname.Enabled = True
    btnLöschen.Enabled = True
    Screen.MousePointer = vbHourglass   'zeige an, daß es eine Weile dauern kann
    On Error Resume Next
    Screen.MousePointer = vbDefault
    btnÄndernDateiname.Default = True
    btnÄndernDateiname.SetFocus
    
    'Der Dateiname im zuletzt aktiven Satz wird gemerkt, dann neu sortiert,
    'dann der zuletzt aktive Satz wieder eingestellt
    Dim Buchmarke, Suchmarke As String
    
    On Error Resume Next
    'Buchmarke = Renam.rsdatagrid.Fields("Dateiname")
    Buchmarke = Renam.rsDataGrid.Fields(LoadResString(1028 + Sprache))
    'SQL = "Select * From Fotos ORDER BY Jahr,Dateiname ASC" 'neu sortieren
    SQL = "Select * From Fotos ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC" 'neu sortieren
'    If Renam.CommandLine = "BGA" Then                     'Gerbing 22.01.2008
'        'Renam.Data1.RecordSource = Renam.BGASQL
'    Else
'        Renam.Data1.RecordSource = SQL
'    End If
    Renam.rsDataGrid.Requery
    Call Renam.SpaltenbreiteWiederherstellen
    'Suchmarke = "Dateiname = '" & Buchmarke & "'"
    Suchmarke = LoadResString(1028 + Sprache) & " = '" & Buchmarke & "'"
    On Error Resume Next
    Renam.rsDataGrid.Find Suchmarke               'zuletzt aktiver Satz wieder eingestellt
    If Err.Number <> 0 Then
        If Renam.CommandLine = "BGA" Then
            'keine Msgbox
        Else
            'msg = "Wenn der zuletzt bearbeitete Satz gelöscht wurde,"
            'msg = msg & "kann das Aktualisieren ihn nicht anzeigen."
            'msg = msg & "In solchen Fällen wird der Beginn der Datenbank gezeigt."
            msg = LoadResString(2406 + Sprache) & vbNewLine
            msg = msg & LoadResString(2407 + Sprache) & vbNewLine
            msg = msg & LoadResString(2408 + Sprache)
            MsgBox msg, vbInformation
        End If
    End If
    Unload Me
End Sub

Private Sub btnHilfe_Click()
    Dim RetVal As Long
    Dim CHMFile As String
    Dim msg As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = Renam.HelpFilePath & "\Help\Deutsch\Renammdb.CHM"                           'Gerbing 23.01.2017
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            msg = CHMFile & vbNewLine
            msg = msg & LoadResString(2544 + Sprache) & vbNewLine
            msg = msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    Else
        CHMFile = Renam.HelpFilePath & "\Help\English\Renammdb.CHM"                           'Gerbing 23.01.2017
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht öffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            msg = CHMFile & vbNewLine
            msg = msg & LoadResString(2544 + Sprache) & vbNewLine
            msg = msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
            Exit Sub
        Else
            RetVal = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If RetVal <= 32 Then
                Call HelpFileErrorMsg(RetVal, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub btnLöschen_Click()
    Dim antwort As Long
    Dim strTemp As String
    Dim DateinameFoto As String
    Dim temp As String
    Dim temp1 As String
    Dim SQL As String
    Dim rc As Boolean
    Dim NewFileName As String                                                       'Gerbing 10.11.2016
    Dim Löschname As String                                                         'Gerbing 23.01.2018
    Dim start As Long
    Dim pos As Long
    
        
    Call Renam.SpaltenbreiteMerken
    On Error Resume Next
    If gblnSQLServerVersion = True Then
        strTemp = Replace(txtAlterName.Text, "+:\", PublicLocationFotos & "\")
    Else
        strTemp = Replace(txtAlterName.Text, "+:\", AppPath & "\")                  'Gerbing 11.04.2005
    End If
    Löschname = txtAlterName.Text                                                   'Gerbing 23.01.2018
    Löschname = Replace(Löschname, "'", "''")                                       'Wo ein Hochkomma vorkommt, nach 2 Hochkommas suchen
    'Renam.DBGridNeu.Columns(6) = Löschname
    'Falls es einen zugehörigen Audio-Kommentar gibt, wird dieser zuerst gelöscht   'Gerbing 12.04.2006
    DateinameFoto = ErmittleDateiname(strTemp)
'    Temp = Dir(DateinameFoto & ".mp3")
'    temp1 = Dir(DateinameFoto & ".wav")
'    If Temp <> "" Or temp1 <> "" Then
    If file_path_exist(DateinameFoto & ".mp3") = True Then
    'If Temp <> "" Then
        temp = DateinameFoto & ".mp3"
    End If
    If file_path_exist(DateinameFoto & ".wav") = True Then
        temp = DateinameFoto & ".wav"
    End If
    If temp <> "" Then                                                              'Gerbing 04.09.2013
        'Kill temp
        rc = file_delete(temp, , True)                                              'Gerbing 04.09.2013
    End If
    'End If
    'Err.Number = 0
'-------------------------------------------------------------------------------------------------------
'   Ich speichere Thumbnails im Ordner ...\GerbingThumbs\...                        'Gerbing 10.11.2016
'   Seit Version 14.2.2 Beim übereinstimmenden Löschen von Datensätzen und Foto, lösche ich auch zugehörige Thumbnails
    NewFileName = strTemp
    start = 1
    Do
        pos = InStr(start, NewFileName, "\")
        If pos = 0 Then Exit Do
        start = pos + 1
    Loop
    NewFileName = Left(NewFileName, start - 1) & "GerbingThumbs\" & Right(NewFileName, Len(NewFileName) - start + 1) & ".jpg" 'Gerbing 08.12.2016
    If file_path_exist(NewFileName) = True Then
        rc = file_delete(NewFileName, False, True) 'ohne Papierkorb, silent
    End If                                                                          'Gerbing 10.11.2016 End
'-------------------------------------------------------------------------------------------------------

    'Kill strTemp                                                           'Gerbing 11.04.2005
    rc = file_delete(strTemp, , True)                                       'Gerbing 04.09.2013
    If rc = False Then Exit Sub
    '----------------------------------------------------------------------------
    'Löschen aus der Datenbank
    'SQL = "select Dateiname from fotos where Dateiname = '" & Löschname & "'"      'Gerbing 23.01.2018
    SQL = "select " & LoadResString(1028 + Sprache) & " from Fotos where " & LoadResString(1028 + Sprache) & " = '" & Löschname & "'" 'Gerbing 26.08.2012
    On Error Resume Next
    Renam.rstsql.Close
    'On Error GoTo 0
    With Renam.rstsql
        .Source = SQL
        .ActiveConnection = Renam.DBado
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Renam.rstsql.Delete
'    If blnHochkomma = True Then
'        End
'    End If
    Renam.rsDataGrid.Requery
    Call Renam.SpaltenbreiteWiederherstellen
    Unload Me
End Sub

Private Sub btnSchließen_Click()
    Unload Me
    On Error Resume Next
End Sub

Private Sub Form_Load()
    Dim SuchString As String
    Dim DateinamenErweiterung As String
    Dim rc As Long
    
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    'MsgBox 9
    Me.Caption = LoadResString(1605 + Sprache)    'Ändern oder Löschen des Dateinamens
    Label1.Caption = LoadResString(1606 + Sprache)      'alter Name:
    Label2.Caption = LoadResString(1607 + Sprache)      'neuer Name:
    btnÄndernDateiname.Caption = LoadResString(1609 + Sprache)          '&Ändern
    btnLöschen.Caption = LoadResString(1608 + Sprache)          '&Löschen
    btnHilfe.Caption = LoadResString(3014 + Sprache)            '&Hilfe
    btnAbbrechen.Caption = LoadResString(1610 + Sprache)        'Ab&brechen
    
    'MsgBox 901
    On Error Resume Next                                        'Gerbing 24.12.2019
    Set ShellObject = CreateObject(CVar("Shell.Application"))   'Gerbing 10.12.2017
    Set GdipTool = New GdipTool                                 'Gerbing 10.12.2017
    'MsgBox 902
    NL = Chr(10) & Chr(13)
    'SuchString = "Dateiname = '" & Renam.GeklickterDateiName & "'"
    SuchString = LoadResString(1028 + Sprache) & " = '" & Renam.GeklickterDateiName & "'"
    txtAlterName = Renam.GeklickterDateiName
    txtNeuerName = Renam.GeklickterDateiName
    'Prüfe, ob die DateinamenErweiterung zu einem Bild gehört               'Gerbing 11.04.2005
    DateinamenErweiterung = UCase(Right(txtAlterName, 3))
    'MsgBox 91
    Select Case DateinamenErweiterung
        Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "WMF"         'Gerbing 29.03.2012
            'nur wenn es tatsächlich eine Bilddatei ist
            Call ThumbnailAnzeigen(txtAlterName, AendernForm.PictureThumb)
        Case "AVI", "MPG", "MOV", "WMV", "ASF", "MP4", "ASX", "MKV", "FLV"  'Gerbing 10.12.2017 bei Videos
            Call VideoThumbnailErzeugen(txtAlterName)
            If outname <> "" Then
                Call ThumbnailAnzeigen(outname, AendernForm.PictureThumb)
            End If
    End Select
    'MsgBox 92
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    On Error Resume Next
End Sub

Private Sub BildFehler(EchterStandort)
    Dim msg As String
    
'    msg = "Bild kann nicht geladen werden" & NL
'    msg = msg & EchterStandort & NL
'    msg = msg & "Prüfen Sie ob diese Datei existiert" & NL
'    msg = msg & "oder ob es sich um einen verbotenen Dateityp handelt."
    msg = LoadResString(2056 + Sprache) & NL
    msg = msg & EchterStandort & NL
    msg = msg & LoadResString(2301 + Sprache) & NL
    msg = msg & LoadResString(2302 + Sprache)
    'MsgBox Msg, vbInformation
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
End Sub

Private Function ErmittleDateiname(Dateiname As String) As String
    Dim pos As Long
    Dim start As Long
    Dim MeinDateiname As String

    'Der Dateiname wird ermittelt durch Suchen ab rechtem Rand bis zum Punkt
    start = Len(Dateiname) - 2
    Do
        pos = InStr(start, Dateiname, ".")
        If pos <> 0 Then
            MeinDateiname = Mid(Dateiname, 1, pos - 1)
            Exit Do
        End If
        start = start - 1
    Loop
    ErmittleDateiname = MeinDateiname
End Function

Private Sub VideoThumbnailErzeugen(FileName As String)
    Dim folder As shell32.folder                                                'Gerbing 06.04.2017
    Dim ShellFolderItem As shell32.ShellFolderItem                              'Gerbing 06.04.2017
    Dim GdipLoader As GdipLoader                                                'Gerbing 06.04.2017
    Dim PropLong As Variant                                                     'Gerbing 06.04.2017
    Dim PropStream As IUnknown                                                  'Gerbing 06.04.2017
    Dim DateinamenErweiterung As String                                         'Gerbing 06.04.2017
    Dim strPath As String                                                       'Gerbing 06.04.2017
    Dim strFile As String                                                       'Gerbing 06.04.2017
    Dim Fotodatei As String                                                     'Gerbing 06.04.2017

    Fotodatei = Replace(FileName, "+:\", AppPath & "\")
    Call file_split(Fotodatei, strPath, strFile, DateinamenErweiterung)
    Set folder = ShellObject.NameSpace(strPath)
    If folder Is Nothing Then
        MsgBox "Folder Is Nothing"
    Else
        Set GdipLoader = New GdipLoader
    End If
    Set ShellFolderItem = folder.ParseName(strFile & "." & DateinamenErweiterung)
    On Error Resume Next                                                        'Gerbing 28.05.2019
    PropLong = ShellFolderItem.ExtendedProperty(SCID_PerceivedType)
    If Not IsEmpty(PropLong) Then
        If PropLong = PERCEIVED_TYPE_VIDEO Then
            On Error Resume Next
            Set PropStream = ShellFolderItem.ExtendedProperty(SCID_PropStream)
            If Err.Number = 0 Then
                If Not PropStream Is Nothing Then
                    outname = AppPath & "\TempThumbs\" & ShellFolderItem.Name & ".jpg"
                    On Error Resume Next
                    If file_path_exist(strPath & "TempThumbs\") = False Then
                        MkDir AppPath & "\TempThumbs\"
                    End If
                    On Error GoTo 0
                    GdipTool.PropStream2PicFileScaled PropStream, outname, PFF_JPEG, , 400, 320
                End If
            Else
                outname = ""
                Exit Sub
            End If
        End If
        On Error GoTo 0
    End If
    On Error GoTo 0                                                             'Gerbing 28.05.2019
End Sub
