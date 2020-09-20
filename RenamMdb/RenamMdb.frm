VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Renam 
   BackColor       =   &H00C0C0C0&
   Caption         =   "RenamMdb"
   ClientHeight    =   5832
   ClientLeft      =   -132
   ClientTop       =   348
   ClientWidth     =   14964
   Icon            =   "RenamMdb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5832
   ScaleWidth      =   14964
   StartUpPosition =   1  'Fenstermitte
   WindowState     =   2  'Maximiert
   Begin EditCtlsLibUCtl.TextBox txtÜberschrift 
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   7692
      _cx             =   13568
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
      HAlignment      =   1
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
      CueBanner       =   "RenamMdb.frx":038A
      Text            =   "RenamMdb.frx":03AA
   End
   Begin EditCtlsLibUCtl.TextBox txtSuchen 
      Height          =   372
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   5532
      _cx             =   9758
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
      CueBanner       =   "RenamMdb.frx":03CA
      Text            =   "RenamMdb.frx":03EA
   End
   Begin VB.TextBox txtFont 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8760
      TabIndex        =   6
      Text            =   "txtFont"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1212
   End
   Begin MSDataGridLib.DataGrid DBGridNeu 
      Height          =   3492
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5412
      _ExtentX        =   9546
      _ExtentY        =   6160
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
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
   Begin VB.ListBox lstSpaltenbreite 
      Height          =   3120
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnWeitersuchen 
      Caption         =   "&Weitersuchen"
      Default         =   -1  'True
      Height          =   372
      Left            =   11520
      TabIndex        =   3
      ToolTipText     =   "Zum Weitersuchen setzen Sie den Zeilenmarkierer auf die nächste Zeile"
      Top             =   120
      Width           =   1572
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "&Hilfe"
      Height          =   372
      Left            =   13320
      TabIndex        =   1
      Top             =   120
      Width           =   1572
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suchen im Feld Dateiname nach:"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5532
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Doppelklicken Sie in die gewünschte Zeile zum Ändern oder Löschen"
      Height          =   552
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10452
   End
End
Attribute VB_Name = "Renam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'29.12.2004 Änderungen in der Spaltenbreite sollen mindestens für die aktuelle Sitzung gespeichert werden
'30.12.2004 Verbesserung beim Löschen oder Namensändern, wenn eine Datei zwar in der Datenbank steht,
'           aber nicht im angegebenen Ordner
'19.01.2005 Schreibgeschützte Fotos kann man manchmal trotzdem Löschen, wenn man vorher das Schreibschutz-
'           Attribut auf vbNormal zurücksetzt
'14.02.2005 Ab Version 10:
'           Einführung nutzerdefinierter Felder.
'           Deshalb können keine anderen Felder mehr geändert werden, als der Dateiname.
'           Änderungen der Spaltenbreite müssen auch für nutzerdefinierte Felder gespeichert werden können.
'           Im Formular AendernForm wird das Bild angezeigt zu Kontrollzwecken
'11.04.2005 Man könnte das Programm fotos.exe so verändern, dass es selbst versucht die Bezeichnung des
'           Fotos-Root-Ordner zu ermitteln, nämlich als App.Path wo fotos.exe steht.
'           Dazu müßte man in die Datenbank eintragen
'           anstelle von zB M:\P7FotoSoundVideo\FOTOS\GG\2005\Ballonfahrt001.jpg
'           +:\2005\Ballonfahrt001.jpg und bei der Ausführung von fotos.exe fotosmdb.exe und renammdb.exe muss
'           +:\ ersetzt werden durch App.Path des entsprechenden Programms.
'
'           Dann entfällt die Funktion Ersetzen im Programm fotosmdb.exe,
'           und alles was zusammenhängt mit 'Fotos-Root-Ordner Festlegen' bei Start von einer CD,
'           und das Feld ErsterStart.ErsterStart wird nicht mehr ausgewertet.
'           aber man muss vom Nutzer verlangen, dass er sämtliche Dateien unterhalb von App.Path von
'           fotos.exe anlegt. Dafür kann er die 3-Einigkeit von fotos.exe, fotos.mdb und Dateien kopieren
'           oder verschieben wohin er will.
'
'           Prüfen der 3-Einigkeit ist in jedem Programm fotos.exe fotosmdb.exe renammdb.exe nötig.
'           Man muss dazu prüfen, ob der erste Satz der Tabelle Fotos, nach Ersetzen des String +:\
'           durch App.Path eine Datei ergibt, die existiert
'           oder bei FotosMdb.exe muß beim Neue Datensätze generieren durch Drag&Drop nach dem
'           strTemp = Replace(AktuellerDateiName, App.Path, "+:" & "\")
'           geprüft werden ob wirklich +: am Anfang von strTemp steht
'16.08.2005 Sortierung nach Spalten durch Klicken in die Spaltenüberschrift abwechseln aufsteigend und
'           absteigend
'10.01.2006 Unterstützung von Sprache siehe fotos.vbp
'           Die Merker-Spalte ist fester Bestandteil der Tabelle Fotos
'12.04.2006 Neue Funktion:
'           Audio-Kommentare benutzen
'           Wenn eine Foto-Datei gelöscht/umgenannt wird, muss falls vorhanden auch eine zugehörige Audio-Datei
'           gelöscht/umgenannt werden.
'09.11.2006 Fehlerkorrektur:
'           Wenn man mitten im Recordset steckt und einen Suchbegriff eintippt, der im ersten Datensatz
'           vorkommt, wird der erste Datensatz nicht gefunden
'14.03.2007 Verbesserung alle Versionen:
'           Neue Hilfe-Dateien im HTML-Format, weil Windows Vista das Winhelp-Format nicht mehr unterstützt
'           zB anstelle RenamMdb.hlp gibt es jetzt RenamMdb.chm
'21.11.2007 Fehlerkorrektur
'           bei falsch oder nicht registrierter dao360.dll
'           kommt die msgbox
'           Errornumber=713
'           Errortext=Objekterstellung durch ActiveX-Komponente nicht möglich
'           You must register the dao360.dll
'           read in http://www.gerbingsoft.de or look for that problem in the internet
'           dann wird das Programm beendet
'22.01.2008 Ich brauche eine geheime Funktion
'           Anstelle der gesamten Tabelle Fotos will ich eine gespeicherte Abfrage benutzen
'           beispielsweise 'DateinameKurz Like 19' um die vierstelligen Jahreszahlen zu finden, die in manchen
'           Dateinamen vorkommen, aber falsch sein können
'           Ich brauche einen Commandline parameter BGA (Benutze Gespeicherte Abfrage)
'16.04.2008 Verbesserung
'           Durch Commandline InhaltDesFeldesDateiname wird das Programm so gestartet, als ob
'           Suchen im Feld Dateiname nach: .... gemacht worden wäre
'17.02.2011 Verbesserung
'           Für Multiuser-Umgebungen ist es notwendig, daß jeder user seine eigene fotos.ini besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'18.02.2011:
'           SpracheFestlegen Abschreiben aus fotos.exe
'           Nicht der Wert in Fotos.ini bestimmt die Sprache, sondern ob es eine Tabelle namens Fotos oder EFotos gibt.
'           daraufhin wird der Wert in Fotos.ini korrigiert
'23.06.2011 13.3.20 Verbesserung:
'           Ich mache die Größe der Fonts für die Controls abhängig von der Einstellung unter 'Eigenschaften von Anzeige' ->
'           Erweitert -> DPI-Einstellungen. Das geschieht automatisch beim Form_Load jedes Formulars.
'           Ich unterscheide normal=96, groß=120, sehr groß>120
'           Das erfordert Bildschirmauflösung mindestens 1024 x 768 bei 96 DPI und
'           mindestens 1280 x 800 bei 120 DPI
'           Der Nutzer soll entscheiden, ob er die Fontgrößen-Anpassung haben will, wenn eine DPI-Einstellung höher als 96
'           gefunden wird, der Wert wird in Fotos.ini gespeichert
'07.11.2011 13.3.21 Verbesserung:
'           Verbesserung für Multi-Nutzer-Umgebung. Vermeidung von overhead, der entsteht bei Benutzung einer fotos.exe vom fremden PC.
'           Jeder PC hat seine lokale fotos.exe und wählt aus, mit welcher fotos.mdb aus einem fremden Ordner oder fremden PC er arbeiten will.
'           Dazu muss der Nutzer beim Start der lokalen fotos.exe die Shift-Taste festhalten. Daraufhin geht ein CommonDialog (ohne ocx) auf zur
'           Auswahl der fotos.mdb
'           Der Ordnername der fotos.mdb steht in gstFotosMdbLocation.
'           Wenn gstFotosMdbLocation leer ist, wird AppPath benutzt. Wenn gstFotosMdbLocation <> "" ist, werden die Tools FotosMdb und Renammdb
'           mit Aufrufparameter gstFotosMdbLocation gestartet.
'14.11.2011 Generelle Entfernung von CommonDialog comdlg32.ocx, weil beim Installieren des MSI-Paketes gemeckert wird
'23.11.2011 13.3.22 Änderung:
'           Ich habe den Winkelmann-Fehler im Windows 7 gefunden. Bei Drücken der Taste F5 kommt ein leeres Grid.
'           und beim Öffnen der Query-Form kommt Fehler-Nr.: -2147467262
'           Ein nackiges Windows 7 ohne Microsoft Office bringt diesen Fehler. Die Installation einer beliebigen Office Komponente
'           ab Office 2003 (probiert mit Word) beseitigt den Fehler. Er tritt auch dann nicht mehr auf, wenn Office wieder deinstalliert
'           wird.
'           Ich muss in frmSprache zu Beginn ermitteln in welchem Betriebssystem ich arbeite.
'           Bei XP und Vista geht es weiter mit der Sprachauswahl.
'           Bei Windows7 und höher, muss ich fragen ob Office 2003 oder höher installiert ist, wenn ja geht es weiter mit der Sprachauswahl.
'           Wenn nein, kommt eine MsgBox mit dem Hinweis, daß erst Office 2003 oder höher installiert werden muss. Dann endet das Programm.
'26.11.2011 13.3.22 Änderung:
'           Anstelle Laufzeitfehler'3050' soll eine vernünftige Ausschrift kommen, abgeschrieben bei fotos.exe
'----------------------------------------------------------------------------------------------------------------------------------------
'29.12.2011 13.4.0 Neue Funktion
'           Unterstützung des SQL-Servers
'           Fotosmdb und Renammdb machen zwar ein Connect siehe frmConnectSQL aber kein Login
'           Wenn Parameter mit CommandLine übergeben werden, erfolgt kein Connect
'           DBado As ADODB.Connection bleibt die gesamte Lebenszeit der session offen
'           rsDataGrid As ADODB.Recordset bleibt die gesamte Lebenszeit der session offen
'           die anderen ADODB.Recordsets werden mehrfach benutzt immer wieder geschlossen und neu geöffnet
'29.12.2011 13.4.0 Neue Funktion wenn Spalte Dateiname nicht der Primärschlüssel ist, geht das Programm garnicht erst los
'29.12.2011 13.4.0 Geänderte SQL-Server-Datenbankstruktur Feld Jahr muss sein nvarchar(4) sonst geht charindex nicht
'29.12.2011 13.4.0 Nur bei der Access-Shareware-Version ist es nötig, daß beim ersten Start von fotos.exe Language = "9" ist
'           nur dann wird msdmo.log erzeugt
'           mit Hilfe des Alters von msdmo.log nerve ich die Shareware-Nutzer mit Einblendung des Shareware-Hinweises
'           Das Datum 30.12.2011 ist das Datum der Fotos.mdb im Auslieferungszustand
'04.03.2012 13.4.1 Fehlerkorrektur
'           Man kann mich austricksen und aus einer Shareware-Version ein SQL-Server-Version machen bei fotos.mdbnichtda
'           frmConnectSQL darf bei gblnPoversion=False nicht erscheinen ab sofort gibt es bedingte Compilierung
'04.03.2012 13.4.2 Verbesserung:
'           nur bei #if Proversion gibt es ein Formular frmConnectSql, sonst wird es bei der Compilierung weggelassen
'05.03.2012 in den Eigenschaften der .exe soll erkennbar sein, ob Proversion=0 oder =-1
'           ich trage ein bei Projekteigenschaften -> Erstellen -> Copyright -> GERBING Software Chemnitz -1 oder 0
'29.03.2012 13.5.0 Verbesserung
'           ThumbnailAnzeigen erfolgt mit GDIPlus, GDI+ ist Bestandteil des Betriebssystems seit XP
'           2 neue native Dateitypen PNG TIF, aber CUR gestrichen
'26.08.2012 Fehlerkorrektur Version 13.5.4
'           nur in englisch-sprachiger Datenbank
'           Der Feldname 'Datenbank' wird nicht gefunden, darum kommt ein Fehler, der aussagt die SQL-Anweisung sei falsch
'26.08.2012 Fehlerkorrektur Version 13.5.4
'           Wenn RenamMdb.exe aus Fotos.exe heraus gestartet wird um einen Dateiname zu ändern, kommt Fehlernr 75
'           Fehler beim Zugriff auf Pfad/Datei ...
'           wenn beim Start nicht mit ausreichenden Rechten gestartet wurde
'26.08.2012 Fehlerkorrektur Version 13.5.4
'           Bisher konnte man im DbGridNeu alle Spalten editieren. das muss schon mal anders gewesen sein.
'           Ab sofort wird DbGridNeu.AllowUpdate = False eingestellt
'27.08.2012 13.5.4 Fehlerkorrektur
'           Kooperationsfehler zwischen Fotos.exe und Renammdb.exe
'           Wenn aus Fotos.exe heraus Renammdb.exe aufgerufen wird und der in Fotos.exe gerade aktuelle Dateiname geändert oder gelöscht werden soll
'           kommt errornumber = 75 Fehler beim Zugriff auf Pfad/Datei beim Ändern
'           kommt errornumber = 70 Zugriff verweigert beim Löschen
'           Die einzige fehlerfreie Lösung die ich gefunden habe, besteht darin nach dem Aufruf von RenamMdb.exe die Fotos.exe zu beenden
'           und vor dem Start von RenamMdb Sleep (1000) einzufügen
'04.10.2012 13.5.4 Fehlerkorrektur
'           Der Dateiname '1960-Wandertag Dieter Knopf, Irmscher, Ullrich Krausse, Guenter Jacob(v. l.).jpg'
'           wird als 5-stellige Dateinamen-Erweiterung erkannt, die zur MsgBox führt
'           "Änderung abgelehnt-Sie haben den zum Dateinamen gehörenden Punkt gelöscht"
'19.11.2012 13.5.5 Fehlerkorrektur SQL Server Version
'           Zum Püfen der ersten Kolonne des LicenseCode wird nicht mehr der Name, sondern die mittlere Kolonne benutzt
'21.11.2012 13.5.5 Fehlerkorrektur SQL Server Version
'           Die bisherige Verschlüsselung der Zahl der Lizenzen ist zu leicht zu knacken durch Probieren
'           Ich verschlüssele jetzt die Zahl an zwei Positionen
'           bisher SQL99
'           jetzt  99S99 und in der Mitte bleibt ein S stehen
'----------------------------------------------------------------------------------------------------------------------------------------
'04.03.2013 Neue Funktion 14.0.0
'           Überraschung: Das DataGrid msdatgrd.ocx ist unicode fähig, Ms Access vermutlich schon lange
'           Unicode-Unterstützung durch die Timosoft Controls und durch FSO
'               geht nicht im XP: exe stürzt ab und auch IDE stürzt ab beim Schließen des Programms, vermutlich weil
'               bei Diashow.Form_Unload set fso=Nothing und das Unload für alle Forms gefehlt hat
'               Geht im Win7
'               Zum Ändern von FontSize muss die Eigenschaft des Timosoft Controls UseSystemFont = False sein
'               Viele Events bei den Timosoft Controls sind standardmäßig disabled. Man muss im gezeichneten Control element rechtsklicken ->
'               Properties -> Häkchen rausnehmen
'           RichTextBox.ocx kann entfallen wird ersetzt durch unicode fähiges Timosoft Text Control
'           INI file wird unicode fähig durch schreiben mit FSO und Benutzen von GetPrivateProfileStringW und
'           WritePrivateProfileStringW
'           Alle 'Name xyz As' ersetzen durch NameAS
'           Alle Dir( ersetzen durch file_path_exist
'           Alle MsgBox wo file names vorkommen ersetzen durch MessageBoxW
'           Für Kill gibt es ein VBA replacement for "Kill(PathName)" with UNICODE support in UnivbzGlobal.bas
'               besser file_delete                                                                                      'Gerbing 04.09.2013
'           Für SetAttr gibt es ein VBA replacement for SetAttr, supports unicode and network in UnivbzGlobal.bas
'           App.Path ersetzen durch getCurrentDir
'08.06.2013 14.0.0
'           Beim normalen Start braucht das Programm keine Administrator-Rechte
'           Das Programm verlangt Administratorrechtezur Bekämpfung von Laufzeitfehler 'Laufzeitfehler '339':
'           Die Komponente CBLCtlsU.ocx oder eine ihrer Abhängigkeiten ist nicht richtig registriert.....
'           Jetzt kann ich aber nicht mehr mit der c:\users\administrator\AppData\Roaming\Gerbing Fotoalbum 14\fotos.ini
'           arbeiten, weil jeder Nutzer der ja jetzt als Administrator starten muss, dieselbe fotos.ini zugeteilt bekommt
'           Ab sofort stehen fotos.ini  und pruef.log im AppPath. Also dort wohin der Nutzer sein GERBING Fotoalbum 14
'           installiert haben wollte. Das ist standardmäßig c:\users\gottfried\Documents\GERBING Fotoalbum 14
'           Von Regprofi.exe muss gerbingsoft.log in c:\Users\Public\Documents\GERBING Fotoalbum 14\gerbingsoft.log gestellt werden
'           Bei der Vollversion steht gerbingsoft.log in c:\Windows\SysWOW64\gerbingsoft.log
'08.06.2013 14.0.0
'           Endlich habe ich es geschafft, daß alle Programme wieder ohne Administrator-Rechte starten dürfen.
'           Das Packen der MSI-Pakete mit COM-Objekten hat zwar die Timosoft-ocx-Dateien installiert, aber Starten ging nur als Administrator
'           Das Packen der MSI-Paket mit den Timosoft-ocx-Dateien als Selfreg=Yes hat den von Anfang an beabsichtigten Effekt gehabt.
'----------------------------------------------------------------------------------------------------------------------------------------
'04.09.2013 Fehlerkorrektur 14.0.1
'           Kill ersetzen durch file_delete
'10.09.2013 Fehlerkorrektur 14.0.1
'           bei Festhalten der Shift-Taste kann der Standort der fotos.mdb ausgewählt werden
'11.09.2013 Fehlerkorrektur 14.0.1
'           Vergessenes Austauschen von Dir( durch file_path_exist
'----------------------------------------------------------------------------------------------------------------------------------------
'14.02.2014 Fehlerkorrektur 14.0.2
'           Bisher kam Fehler '2147217900' ungültiger Spaltennname 'filename' wenn ich Renammdb.exe solo starte (nicht aus fotos.exe heraus)
'----------------------------------------------------------------------------------------------------------------------------------------
'15.02.2014 14.0.3 Verbesserung alle Versionen
'           Die Msgbox zu Programmstart bei falscher fotos.mdb ist überarbeitet worden
'           'msg = Dateiname & " existiert nicht." & vbNewLine
'           'msg = "Datenbank und Fotos passen nicht zueinander" & vbNewLine
'           'msg = msg & "Vermutlich benutzen Sie eine falsche Datenbank-Datei" & vbNewLine
'           'msg = msg & "Benutzen Sie das Tool Fotosmdb um die Datenbank zu überprüfen" & vbNewLine & vbNewLine
'
'           'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
'----------------------------------------------------------------------------------------------------------------------------------------
'21.09.2014 14.0.6 Verbesserung alle Versionen
'           Zustand: Sobald man ein Foto bearbeitet hatte, mußte man zweimal aufs Schließkreuz klicken, um das Programm zu schließen
'                   oder einmal aufs Formular und dann aufs Schließkreuz
'           Lösung: bisher wurde aufgerufen AendernForm.Show 1
'                   jetzt wird aufgerufen AendernForm.Show
'12.10.2014 14.0.6 Verbesserung alle Versionen
'           Zustand: Beim zweiten Aufruf von Renammdb aus fotos.exe heraus kam Laufzeitfehler '5'
'           Lösung: ich muß multiInstance auf 'TRUE' setzen
'20.10.2014 14.0.6 Verbesserung alle Versionen
'           Zustand: Wenn RenamMdb.exe aus fotos.exe heraus aufgerufen wird, 'Öffne RenamMdb für die aktuelle Datei' ist die aktuelle
'                   Datei schwer zu erkennen.
'           Lösung: Ich markiere die ganze Zeile schwarz
'           genauso mach ich es beim Weitersuchen
'31.10.2014 14.0.6 Verbesserung alle Versionen
'           Zustand: SQLERR: ist falsch programmiert, wenn tatsächlich ein Fehler auftritt bleibt es bei errloop hängen
'----------------------------------------------------------------------------------------------------------------------------------------
'06.09.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Bis heute ist mir nicht aufgefallen, dass DbGridNeu keine Scrollbars besitzt
'           Lösung: ich füge per Code Scrollbars hinzu, jetzt funktioniert Resize richtig
'----------------------------------------------------------------------------------------------------------
'07.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: im Windows 10 und Windows 8.1 und vom Standard abweichender DPI-Einstellung zeigt mein Programm verschwommene Schrift
'                   Das kann der Nutzer korrigieren, indem er die exe markiert -> Eigenschaften -> Kompatibilität ->
'                   DPI-Skalierung nicht anwenden
'           Lösung: Ein Programm erklärt sich selbst als DPI-kompatibel. Das geht durch sein Manifest
'----------------------------------------------------------------------------------------------------------
'10.11.2016 14.2.2 Verbesserung alle Versionen
'           Ich speichere Thumbnails im Ordner ...\GerbingThumbs\...
'           Beim übereinstimmenden Löschen von Datensätzen und Foto, lösche ich auch zugehörige Thumbnails
'08.12.2016 Die Dateien im Ordner ...\GerbingThumbs\... heißen zB video1.avi.jpg oder foto33.jpg.jpg
'----------------------------------------------------------------------------------------------------------
'23.01.2017 14.2.2 Verbesserung alle Versionen
'           Zustand: Der Standort der Help-Datei wird angemeckert bei SQL-Server-Version
'           Lösung: Ich brauche den String HelpFilePath, das ist dort wo RenamMdb.exe steht
'----------------------------------------------------------------------------------------------------------
'11.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn 4K-Monitore benutzt werden, muss es möglich sein, die Schriftgröesse besser als bisher anzupassen
'           Lösung1: Es gibt die Schriftgrößen
'                   klein=1
'                   mittel=2
'                   gross=3
'           Die Einstellung wird gespeichert in der ini-Datei   [Adjustments]
'                                                               CheckForDPI 1 oder 2 oder 3
'           Lösung2: oder es genügt die Bildschirmauflösung auf zB 200 DPI einzustellen (Windows 10 kann noch weit höher als 200 DPI)
'----------------------------------------------------------------------------------------------------------
'23.11.2017 15.0.2 Problem mit unicode filename wenn zB GGCnopt\fotos.mdb  Access Datenbank
'           kein Problem mit SQL-Server-Version
'           Zustand: Es kommt 'Kein zulässiger Dateiname' früher ging das schon mal
'                   Vermutlich hat Microsoft daran herumgedreht.
'                   Die Datenbank lässt sich aber mit MS Access öffnen.
'                   Ich komme mit DBsql.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;... darüber hinweg
'                   aber dann spinnen andere Stellen im Code, die mit DAO programmiert sind
'           Lösung: DAO Code durch ADO Code ersetzen
'                   Prüfung ob die Datenbank schreibgeschützt ist mit SQL = "UPDATE FET SET FN = 'test'"
'                   es gibt nur noch DBado für beide Versionen Access oder SQL-Server
'                   Reference auf Microsoft DAO 3.6 Object Library wird nicht mehr gebraucht. dao360.dll wird nicht mehr gebraucht
'----------------------------------------------------------------------------------------------------------
'10.12.2017 15.0.2 Verbesserung alle Versionen
'           Zustand: Bei Videos kommt kein Vorschaubild
'           Lösung: Das Verfahren aus FotosMdb.exe abschreiben
'                   Es wird shell32.dll gebraucht Projekt -> Verweise -> 'Microsoft Shell Controls and Automation'
'                   Im Ordner TempThumbs wird ein Video-Thumbnail erzeugt und danach angezeigt.
'                   Neue Klassenmoduln sind GdipLoader.cls und GdipTools.cls
'                   Bei Programmstart wird der Inhalt vom Ordner \TempThumbs\ gelöscht
'10.12.2017 15.0.2 Verbesserung alle Versionen
'           Zustand: Video-Datei-Typ "MKV" unfd "FLV" wird bisher nicht akzeptiert
'           Lösung: ab sofort ist "MKV" und "FLV" erlaubt
'                   bei "MKV" gibt es keine Vorschaubilder
'23.01.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: Das GERBING Fotoalbum kann keine Dateinamen verarbeiten, wo Hochkomma enthalten sind
'           Schlechte Lösung:
'                   In Fotosmdb.exe sollten ab Version 15.0.2 Hochkommas im Dateinamen durch - ersetzt werden
'                   Was passiert aber mit Dateinamen wo schon Hochkommas enthalten sind:
'                   1.fotos.exe ignoriert derartige Dateinamen zb in F5MehrereZeilen oder in KommentarForm oder in frmGridAndThumb beim
'                       synchronisieren Thumbnail-Ansicht mit Listen-Ansicht
'                   2.Renammdb.exe läßt gemeinsames Löschen im Ordner und der Datenbank zu, danach wird das Programm beendet
'                   3.Man muss alle Dateinamen mit Hochkomma im Fotoalbum-Ordner finden(zB mit Everything) -> in einen Retteordner kopieren->
'                       im Fotoalbum-Ordner löschen ->
'                       Prüfen1 ausführen -> nicht gefundene Datensätze löschen -> im Retteordner umnennen ohne Hochkomma -> umgenannte in den
'                       Fotoalbum-Ordner zurückkopieren -> Prüfen3 ausführen
'
'           Gute Lösung: Wo im Dateiname ein Hochkomma vorkommt, wird beim Aufbau des SQL-Strings nach 2 Hochkommas gesucht
'----------------------------------------------------------------------------------------------------------
'26.03.2018 15.0.3 Verbesserung alle Versionen
'           Zustand: Das Fenster 'Ändern oder Löschen des Dateinamens' verschwindet stets hinter dem Fenster 'RenamMdb'
'           versuchte Lösung: AendernForm.ZOrder hilft nicht
'           Lösung: Rückname der Änderung 21.09.2014, jetzt wieder AendernForm.Show 1
'29.10.2018 15.0.3 Verbesserung alle Versionen
'           Zustand: Weitersuchen nach einem String mit Hochkomma muckert
'           Lösung: Wo im Suchstring ein Hochkomma vorkommt, werden 2 Hochkommas gesucht
'                   Der Nutzer muss zum Finden weiterer Treffer für den Suchstring den Zeilenmarkierer auf die nächste Zeile setzen
'10.11.2018 15.0.3 Verbesserung alle Versionen
'           Endlich habe ich gefunden wie man .Find richtig benutzt zum Suchen und Weitersuchen
'           rsDataGrid.Find strFind, SkipRows, adSearchForward, mark
'           Man braucht im Anfangszustand SkipRows=0 und mark=0 und nach dem ersten Finden SkipRows=1 und mark=rsDataGrid.Bookmark
'           Wenn beim Weitersuchen nichts gefunden wird, muss ich wieder SkipRows=0 und mark=0 setzen
'11.11.2018 15.0.3 Fehlerkorrektur nur im Win10 und nicht in der IDE aber zB when running in C:\...
'           es kommt, bevor Form_Load zum Ende kommt, bereits run time error '13'
'           ich finde nicht warum
'           aber beim Versuch den Fehler zu finden habe ich gemerkt, dass kein Fehler auftritt, wenn ich eine Msgbox ausführe
'           Lösung1: ich lasse eine MsgBox mit dem Startzeitpunkt ausführen
'           Lösung2: ich habe herumprobiert mit Compilieren zu P-Code oder Compilieren zu native code und plötzlich geht es ohne MsgBox fehlerfrei
'           Lösung3: Lösung2 geht doch nicht, also wieder Lösung1
'           27.01.2019 Lösung 4: einfach neu compiliert und es geht ohne run time error
'           - ich vermute meine Umstellung auf Office 2016 war Schuld
'           23.12.2019 Lösung 5: - ich vermute abwechselnd in Win7 und Win10 kompilieren war schuld
'----------------------------------------------------------------------------------------------------------
'29.04.2019 15.0.4 Fehlerkorrektur portable Version
'           Zustand: In der portablen Version kommt runtime error 13 type mismatch in Renam.Form_Resize
'           Lösung: Ich lasse die Prozedur Renam.Form_Resize weg(auskommentiert)
'                   Das hat zur Folge, dass nur ein halbes Fenster kommt
'----------------------------------------------------------------------------------------------------------
'28.05.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Namensänderung eines mp4 videos führt zu Laufzeitfehler '91'
'           Lösung: Einfügen: On Error Resume Next
'==========================================================================================================
'24.12.2019 16.0.0 Fehlerkorrektur Nachbesserung zum 10.12.2017
'           Zustand: es kommt, bevor Form_Load zum Ende kommt, bereits run time error '13'
'               ich finde nicht warum
'               aber beim Versuch den Fehler zu finden habe ich gemerkt, dass kein Fehler auftritt, wenn ich eine Msgbox ausführe
'           Lösung1: ich lasse eine MsgBox mit dem Startzeitpunkt ausführen
'           Lösung2: ich habe herumprobiert mit Compilieren zu P-Code oder Compilieren zu native code und plötzlich geht es ohne MsgBox fehlerfrei
'           Lösung3: Lösung2 geht doch nicht, also wieder Lösung1
'           27.01.2019 Lösung 4: einfach neu compiliert und es geht ohne run time error
'           - ich vermute meine Umstellung auf Office 2016 war Schuld
'           23.12.2019 Lösung 5: - ich vermute abwechselnd in Win7 und Win10 kompilieren war schuld
'           jetzt endlich
'           Ursache: im Windows 7 geht es noch nicht mit Video-Vorschaubild
'           Lösung: on error resume next 'einfügen, dann
'                   Set ShellObject = CreateObject(CVar("Shell.Application"))
'----------------------------------------------------------------------------------------------------------
'04.02.2020 16.0.0 Fehlerkorrektur
'           Zustand: es ist ein Uralt-Fehler
'                   Wenn der Dateiname mit 'Ändern' verändert wird, wird der DateiNameKurz falsch gebildet
'           Lösung: DateiNameKurz richtig bilden






Option Explicit
    Public HelpFileName As String
    Public GeklickterDateiName As String
    
    Dim Col1 As Column
    Dim Col2 As Column
    Dim Col3 As Column
    Dim Col4 As Column
    Dim Col5 As Column
    Dim Col6 As Column
    Dim Col7 As Column
    Dim Col8 As Column
    Dim Col9 As Column
    Dim Col10 As Column
    Dim Col11 As Column
    Dim Col12 As Column
    Dim Col13 As Column
    Dim Col14 As Column
    Dim Col15 As Column
    
    
    Dim msg As String
    
    
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) _
    As Long
    
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    
    Public BGASQL As String                             'Gerbing 22.01.2008
    Public CommandLine As String                        'Gerbing 22.01.2008
    
        Private Type OSVERSIONINFO
      OSVSize         As Long
      dwVerMajor      As Long
      dwVerMinor      As Long
      dwBuildNumber   As Long
      PlatformID      As Long
      szCSDVersion    As String * 128
    End Type
    
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2
    
    ' Auflistung
    Public Enum OfficeVersion
        Office2003 = 11
        Office2007 = 12
        Office2010 = 14
    End Enum
    
    'Hier wird die Registry ausgewertet, ob eine bestimmte Office-Version installiert ist
    ' benötigte API-Deklarationen
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
      Alias "RegOpenKeyExA" ( _
      ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      phkResult As Long) As Long
     
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
      Alias "RegQueryValueExA" ( _
      ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      ByVal lpData As String, _
      lpcbData As Long) As Long
     
    Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
      ByVal hKey As Long) As Long
     
    ' Konstanten
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const ERROR_SUCCESS = 0&
    Private Const REG_SZ = 1
    Private Const KEY_QUERY_VALUE = &H1
    
    Public DBado As ADODB.Connection                                                                'Gerbing 23.11.2017
    Public rstsql As ADODB.Recordset
    Public rsDataGrid As ADODB.Recordset
    
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer  'Gerbing 10.09.2013
    Private Const KeyPressed As Integer = -32767
    Private Const VK_CONTROL = &H11
    Private Const VK_MENU As Long = &H12&
    Private Const VK_SHIFT As Long = &H10&
    Private Const VK_CAPITAL As Long = &H14&
    
    Public HelpFilePath As String                                                                   'Gerbing 23.01.2017
    Dim mark As Variant                                                                             'Gerbing 10.11.2018
    Dim SkipRows As Long                                                                            'Gerbing 10.11.2018

    

Private Sub btnWeitersuchen_Click()
    Dim strFind As String
    Dim msg As String
    Dim strTemp As String
    
    If txtSuchen = "" Then Exit Sub                                     'Gerbing 10.11.2018
    'strFind = "Dateiname like '*" & txtSuchen & "*'"
    strTemp = Replace(txtSuchen, "'", "''")                             'Gerbing 29.10.2018
    strFind = LoadResString(1028 + Sprache) & " like '*" & strTemp & "*'"
    rsDataGrid.Find strFind, SkipRows, adSearchForward, mark            'Gerbing 10.11.2018
    If rsDataGrid.EOF Then
        'msg = "Suchbegriff nicht gefunden. Die Suche beginnt von vorn"
        msg = LoadResString(2411 + Sprache)
        'MsgBox msg, , "Suchen"
        MsgBox msg, , LoadResString(2412 + Sprache)
        rsDataGrid.MoveFirst
        mark = 0
        SkipRows = 0
    Else
        mark = rsDataGrid.Bookmark
        SkipRows = 1
    End If
    On Error GoTo 0
    If DBGridNeu.SelBookmarks.Count = 1 Then                            'Gerbing 20.10.2014
        DBGridNeu.SelBookmarks.Remove 0                                 'Gerbing 20.10.2014
    End If                                                              'Gerbing 20.10.2014
    DBGridNeu.SelBookmarks.Add rsDataGrid.Bookmark                      'Gerbing 20.10.2014
End Sub

Private Sub btnHilfe_Click()
    Dim RetVal As Long
    Dim CHMFile As String
    Dim msg As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = HelpFilePath & "\Help\Deutsch\Renammdb.CHM"                           'Gerbing 23.01.2017
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
        CHMFile = HelpFilePath & "\Help\English\Renammdb.CHM"                           'Gerbing 23.01.2017
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

Public Sub DBGridNeu_BeforeDelete(Cancel As Integer)
    MsgBox "DBGridNeu_BeforeDelete "
End Sub

Private Sub DBGridNeu_DblClick()
    Dim strTemp As String

    GeklickterDateiName = DBGridNeu.Columns(6)
    If gblnSQLServerVersion = True Then
        strTemp = Replace(GeklickterDateiName, "+:\", PublicLocationFotos & "\")
    Else
        strTemp = Replace(GeklickterDateiName, "+:\", AppPath & "\")        'Gerbing 11.04.2005
    End If
    If file_path_exist(strTemp) = False Then
        Call BildFehler(strTemp)
        Exit Sub
    End If
    If GeklickterDateiName <> "" Then
        Call SpaltenbreiteMerken
        Unload AendernForm                                                  'Gerbing 21.09.2014
        AendernForm.Show 1                                                    'Gerbing 21.09.2014 26.03.2018
        AendernForm.ZOrder                                                  'Gerbing 26.03.2018
    Else
        'MsgBox "Das Feld Dateiname darf nicht leer sein. Benutzen Sie die Funktion Prüfen1 im Tool Fotosmdb.exe"
        MsgBox LoadResString(2418 + Sprache)
    End If
End Sub

Private Sub DBGridNeu_HeadClick(ByVal ColIndex As Integer)
    Dim SQL As String
    Dim SQLalt As String
    Dim SQLneu As String
    Dim pos As Long
    Dim Links As String
    Dim Col As Column
    Dim ColCaption As String
    Dim intLeftcol As Integer

    Call SpaltenbreiteMerken
    intLeftcol = DBGridNeu.LeftCol
    Set Col = DBGridNeu.Columns(ColIndex)
    SQLalt = rsDataGrid.Source
    ColCaption = Col.Caption
    pos = InStr(1, SQLalt, "DESC", vbTextCompare)
    If pos <> 0 Then
        SQL = " ORDER BY [" & ColCaption & "];"
    Else
        SQL = " ORDER BY [" & ColCaption & "] DESC;"
    End If
    pos = InStr(1, SQLalt, "ORDER BY", vbTextCompare)
    If pos = 0 Then
        Links = Left(SQLalt, Len(SQLalt) - 3)
    Else
        Links = Left(SQLalt, pos - 1)
    End If
    
    SQLneu = Links & SQL
    ' Recordset erstellen und öffnen
    Set rsDataGrid = New ADODB.Recordset
    With rsDataGrid
        .Source = SQLneu
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
    Call SpaltenbreiteWiederherstellen
    'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
    DBGridNeu.Scroll intLeftcol, 0
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim NL As String
    Dim Feldname As String
    Dim antwort As Long
    Dim strTemp As String
    Dim strFind As String
    Dim msg As String
    Dim pos As Long
    Dim Pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim ParmBGA As String
    Dim RowColChangeName As String
    Dim strVersion As String
    'Public fso As New scripting.FileSystemObject
    Dim oFile As scripting.File
    Dim i As Long

    'MsgBox 8
    mark = 0                                                        'Gerbing 10.11.2018
    SkipRows = 0                                                    'Gerbing 10.11.2018
    DBGridNeu.ScrollBars = dbgBoth                                  'Gerbing 06.09.2015
    'MsgBox "Form_Load"

    init_global                                                     'Gerbing 04.03.2013
    'AppPath = App.Path                  'Gerbing 16.04.2005
    
    'MsgBox "Start time: " & Now                                     'Gerbing 11.11.2018 23.12.2019
    'MsgBox "vor AppPath = getCurrentDir=" & AppPath
    
    AppPath = getCurrentDir
    If Right(AppPath, 1) = "\" Then
        AppPath = Left(AppPath, Len(AppPath) - 1)
    End If
    HelpFilePath = AppPath                                          'Gerbing 23.11.2017

    'MsgBox "nach AppPath = getCurrentDir=" & AppPath


    'Sleep (1000)                                                   'Gerbing 27.08.2012
    gblnProversion = True                                           'Gerbing 04.03.2012
    #If Proversion = 0 Then
        gblnProversion = False
    #End If

    On Error Resume Next
    'gdtDatumFotosMdb = FileDateTime(AppPath & "\fotos.mdb")
    Set oFile = fso.GetFile(AppPath & "\fotos.mdb")                 'Gerbing 04.03.2013
    gdtDatumFotosMdb = oFile.DateLastModified
    On Error GoTo 0
    
    'MsgBox "vor ReadFotosIniFile"
    
    Call ReadFotosIniFile

    'MsgBox "nach ReadFotosIniFile"

    
'    If isAdmin = False Then                                        'Gerbing 08.06.2013
''        Msg = "GERBING Fotoalbum 14 must be started with administrator rights." & vbNewLine
''        Msg = Msg & "You do not find how?" & vbNewLine
''        Msg = Msg & "Remember the installation folder and start RenamMdb.exe from there."
'        msg = LoadResString(2546 + Sprache) & vbNewLine
'        msg = msg & LoadResString(2547 + Sprache) & vbNewLine
'        msg = msg & LoadResString(2552 + Sprache)
'        MsgBox msg
'        End
'    End If
'--------------------------------------------------------------------------------------------------------------------
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    Call AnpassenHeadFont(DBGridNeu)                                'Gerbing 23.06.2011
    
    gblnSubdirectories = True                                       'Gerbing 10.12.2017
    Call RekursiveTempThumbs(AppPath & "\TempThumbs", "*.*")        'Gerbing 10.12.2017
    
    If (GetAsyncKeyState(VK_SHIFT) = KeyPressed) Then               'Gerbing 10.09.2013
        Call FremdeFotosMdb                                         'Gerbing 10.09.2013
    End If
    
    CommandLine = command()                                         'Gerbing 07.11.2011
    If StrComp(CommandLine, "BGA", vbTextCompare) = 0 Then
        ParmBGA = "BGA"
    End If

    gblnCommandLineEmpty = True
'    If CommandLine = "" Then
'        gblnCommandLineEmpty = True
'    End If
    'BGA
    'rowcolchangename=...;
    'fotosmdblocation=...;
    'sqlservername=...;
    'datenbankname=...;
    'WindowsAuthentication=0; heißt nein
    'WindowsAuthentication=1; heißt ja
    'username=...;
    'Password=...;
    'StandortFotos=...;
    pos = InStr(1, CommandLine, "rowcolchangename=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            RowColChangeName = strTemp
            'beispiel - rowcolchangename="+:\2012\DezemberKegelclubRiedelmühle01.jpg";
        End If
    End If
    
    'MsgBox "Rowcolchengename=" & RowColChangeName
    
    pos = InStr(1, CommandLine, "fotosmdblocation=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            'AppPath wird mit Command übergeben
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            AppPath = strTemp
        End If
    End If
    pos = InStr(1, CommandLine, "sqlservername=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            gblnCommandLineEmpty = False
            PublicSQLServer = strTemp
        Else
        
        End If
    End If
    pos = InStr(1, CommandLine, "datenbankname=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            gblnCommandLineEmpty = False
            PublicSQLDatabase = strTemp
        End If
    End If
    pos = InStr(1, CommandLine, "WindowsAuthentication=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            gblnCommandLineEmpty = False
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            If strTemp = "0" Then
                PublicWindowsAuthentication = "0"
            Else
                PublicWindowsAuthentication = "1"
            End If
        End If
    End If
    pos = InStr(1, CommandLine, "username=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            gblnCommandLineEmpty = False
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            PublicSQLServerUserName = strTemp
        End If
    End If
    pos = InStr(1, CommandLine, "Password=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            gblnCommandLineEmpty = False
            PublicSQLServerPassword = strTemp
        End If
    End If
    pos = InStr(1, CommandLine, "StandortFotos=", vbTextCompare)
    If pos <> 0 Then
        Pos1 = InStr(pos, CommandLine, "=", vbTextCompare)
        pos2 = InStr(pos, CommandLine, ";", vbTextCompare)
        If Pos1 <> 0 Then
            strTemp = Mid(CommandLine, Pos1 + 1, pos2 - Pos1 - 1)
            PublicLocationFotos = strTemp
        End If
    End If
    '------------------------------------------------------------------------------------------------------
    If gblnProversion = True Then                                               'Gerbing 04.03.2012
        'Untersuche ob Access-Version oder SQL-Server-Version
        'strTemp = Dir(AppPath & "\Fotos.mdb")
        If file_path_exist(AppPath & "\Fotos.mdb") = False Then
        'If strTemp = "" Then
            gblnSQLServerVersion = True
        End If
    Else
        gblnSQLServerVersion = False
    End If
    If gblnSQLServerVersion = True Then
        AppPath = PublicLocationFotos
    End If

    '------------------------------------------------------------------------------------------------------
CallSpracheFestlegen:
    Call SpracheFestlegen                                                               'Gerbing 18.02.2011
    If PublicLanguage = "" Then
        Sprache = 0                     '0=deutsch
    End If
    If PublicLanguage = "0" Then
        Sprache = 0                     '0=deutsch
    End If
    If PublicLanguage = "1" Then
        Sprache = 3000                  '3000=englisch
    End If
    '-------------------------------------------------------------------------------------------------------
    Me.Caption = "RenamMdb"
    Label1.Caption = LoadResString(1601 + Sprache)  'Doppelklicken Sie in die gewünschte Zeile zum Ändern oder Löschen.
    Label1.ToolTipText = LoadResString(1602 + Sprache) 'Am schnellsten geht es, wenn Sie auf den Datensatzmarkierer doppelklicken."
    Label3.Caption = LoadResString(1603 + Sprache)    'Suchen im Feld Dateiname nach:
    btnWeitersuchen.Caption = LoadResString(1604 + Sprache)       '&Weitersuchen
    btnWeitersuchen.ToolTipText = LoadResString(2563 + Sprache)     'Zum Weitersuchen setzen Sie den Zeilenmarkierer auf die nächste Zeile
    btnHilfe.Caption = LoadResString(1039 + Sprache)                '&Hilfe
    
    NL = Chr(13) & Chr(10)
    If gblnSQLServerVersion = True Then
        txtÜberschrift.Text = PublicSQLServer & " " & PublicSQLDatabase
    Else
        txtÜberschrift.Text = AppPath & "\fotos.mdb"
    End If
    txtÜberschrift.Width = Screen.Width - 200
    DBGridNeu.Width = Screen.Width - 200
    DBGridNeu.DefColWidth = Screen.Width \ 8
    DBGridNeu.AllowRowSizing = False
    'SQL = "Select * From Fotos ORDER BY Jahr,Dateiname ASC" 'neu sortieren
    SQL = "Select * From Fotos ORDER BY " & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & " ASC" 'neu sortieren
    
    DBado.Errors.Clear
    On Error GoTo SQLERR
    
    ' Recordset erstellen und öffnen
    
    'MsgBox 1
    
    Set rsDataGrid = New ADODB.Recordset
    
    'MsgBox 10
    
    With rsDataGrid
        .Source = SQL
        .ActiveConnection = DBado
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    'MsgBox 11
    
    Set DBGridNeu.DataSource = rsDataGrid
    
    'MsgBox 12
    
    DBGridNeu.ReBind
    DBGridNeu.AllowArrows = True
    DBGridNeu.TabAcrossSplits = True
    DBGridNeu.TabAction = dbgGridNavigation
    DBGridNeu.WrapCellPointer = True
    
    On Error GoTo 0
    '---------------------------------------------------------------------------------------------------------
    On Error Resume Next
    If rsDataGrid.EOF = True And rsDataGrid.BOF = True Then
        If gblnSQLServerVersion = True Then
            'MsgBox "Die Datei " & PublicSQLServer & " " & Publicsqldatabase & " ist leer." & NL & "Die einzige mögliche Programmfunktion ist " & "Prüfen3 oder" & NL & "&Neue Datensätze generieren (durch Drag&&Drop vom Windows Explorer)..."
            'MsgBox LoadResString(2145 + Sprache) & " " & PublicSQLServer & " " & PublicSQLDatabase & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
            msg = LoadResString(2145 + Sprache) & " " & PublicSQLServer & " " & PublicSQLDatabase & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), MB_ICONINFORMATION Or MB_TASKMODAL
        Else
            'MsgBox "Die Datei " & AppPath & "\Fotos.mdb" & " "  " ist leer." & NL & "Die einzige mögliche Programmfunktion ist " & "Prüfen3 oder" & NL & "&Neue Datensätze generieren (durch Drag&&Drop vom Windows Explorer)..."
            'MsgBox LoadResString(2145 + Sprache) & AppPath & "\Fotos.mdb" & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
            msg = LoadResString(2145 + Sprache) & AppPath & "\Fotos.mdb" & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), MB_ICONINFORMATION Or MB_TASKMODAL
        End If
    End If
    On Error GoTo 0
    '---------------------------------------------------------------------------------------------------------
    
    'MsgBox 2
    
    If gblnSQLServerVersion = False Then
        '3-Einigkeit überprüfen Gerbing 11.04.2005
        'Feldname = rsDataGrid.Fields("Dateiname")
        If Not rsDataGrid.EOF Then
            Feldname = rsDataGrid.Fields(LoadResString(1028 + Sprache))
        End If
        If Left(Feldname, 3) <> "+:\" Then
            'msg = "Seit Version 12.0.0.0 verlangt das Programm, dass in der Tabelle Fotos" & vbNewLine
            msg = LoadResString(2154 + Sprache) & vbNewLine
            'msg = msg & "das Feld Dateiname generell mit den Zeichen +:\ beginnt" & vbNewLine
            msg = msg & LoadResString(2155 + Sprache) & vbNewLine
            'msg = msg & "Der String +:\ wird vom Programm durch AppPath ersetzt." & vbNewLine
            msg = msg & LoadResString(2156 + Sprache) & vbNewLine
            'msg = msg & "AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
            msg = msg & LoadResString(2157 + Sprache) & vbNewLine
            'msg = msg & "Diese Forderung wurde nicht eingehalten." & vbNewLine & vbNewLine
            msg = msg & LoadResString(2158 + Sprache) & vbNewLine & vbNewLine
            
            'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
            msg = msg & LoadResString(2159 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
            'MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Renammdb"), MB_ICONINFORMATION Or MB_TASKMODAL
            antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbDefaultButton2 + vbYesNo)
            If antwort = vbNo Then
                End
            End If
        End If
        Feldname = Replace(Feldname, "+:\", AppPath & "\")
        'strTemp = Dir(Feldname)
        If file_path_exist(Feldname) = False Then
        'If strTemp = "" Then
            'msg = Feldname & " existiert nicht." & vbNewLine
            msg = Feldname & LoadResString(2162 + Sprache) & vbNewLine
            'msg = "Datenbank und Fotos passen nicht zueinander" & vbNewLine
            msg = msg & LoadResString(2160 + Sprache) & vbNewLine
            'msg = msg & "Vermutlich benutzen Sie eine falsche Datenbank-Datei" & vbNewLine
            msg = msg & LoadResString(2161 + Sprache) & vbNewLine
            'msg = msg & "Benutzen Sie das Tool Fotosmdb um die Datenbank zu überprüfen" & vbNewLine & vbNewLine
            msg = msg & LoadResString(2163 + Sprache) & vbNewLine & vbNewLine
            
            'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
            msg = msg & LoadResString(2159 + Sprache)
            antwort = MessageBoxW(0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbDefaultButton2 + vbYesNo)
            If antwort = vbNo Then
                End
            End If
        End If
    End If
    '-----------------------------------------------------------------------------------------------------------
    
    'MsgBox 3
    
    Call SpaltenBreite
    Call SpaltenbreiteMerken
    
    'MsgBox 4
    
    If RowColChangeName <> "" Then                                  'Gerbing 07.11.2011
        txtSuchen = RowColChangeName                                'Gerbing 16.04.2008
        strTemp = Replace(txtSuchen, "'", "''")                     'Gerbing 29.10.2018
        'strFind = "Dateiname like '*" & txtSuchen & "*'"
        strFind = LoadResString(1028 + Sprache) & " like '*" & strTemp & "*'"
        On Error Resume Next
        rsDataGrid.Find strFind
        If rsDataGrid.EOF Or Err.Number <> 0 Then                   'Gerbing 29.10.2018
            'msg = "Suchbegriff nicht gefunden. Die Suche beginnt von vorn"
            msg = LoadResString(2411 + Sprache)
            'MsgBox msg, , "Suchen"
            MsgBox msg, , LoadResString(2412 + Sprache)
            rsDataGrid.MoveFirst
        End If
        On Error GoTo 0
        If DBGridNeu.SelBookmarks.Count = 1 Then                            'Gerbing 20.10.2014
            DBGridNeu.SelBookmarks.Remove 0                                 'Gerbing 20.10.2014
        End If                                                              'Gerbing 20.10.2014
        DBGridNeu.SelBookmarks.Add rsDataGrid.Bookmark                      'Gerbing 20.10.2014
        Exit Sub
    End If
    
    'MsgBox 5
    
    If rsDataGrid.RecordCount = 0 Then
        'MsgBox "Recordcount = 0"
    Else
        rsDataGrid.MoveFirst
    End If

    
    'MsgBox 6
    
    Exit Sub
SQLERR:                                                                     'Gerbing 31.10.2014
    If DBado.Errors.Count > 0 Then
        For i = 0 To DBado.Errors.Count - 1
            MsgBox "Fehler " & DBado.Errors.Item(i) & vbNewLine
        Next i
    End If
    Resume Next
End Sub

Private Sub Form_Resize()                          'Gerbing 29.04.2019 in der portablen Version weggelassen
    'MsgBox 7
    On Error Resume Next                            'Gerbing 26.03.2018 sonst Laufzeitfehler 380 ungültiger Eigenschaftswert
    DBGridNeu.Height = Me.Height - DBGridNeu.Top - 600  'die 800 damit der Rollbalken über die Task-Leiste kommt    'Gerbing 06.09.2015
    DBGridNeu.Width = Me.Width - 400                                                                                'Gerbing 06.09.2015
    AendernForm.ZOrder                                                      'Gerbing 26.03.2018
    'MsgBox 8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim antwort As Long
    
    If CommandLine = "BGA" Then
        antwort = MsgBox("Wollen Sie das Programm mit einer anderen gespeicherten Abfrage wiederholen?", vbDefaultButton1 + vbYesNo)
        If antwort = vbYes Then
            frmBenutzeGespeicherteAbfrage.Show 1                    'Gerbing 22.01.2008
            Call SpaltenbreiteMerken
            
            With rsDataGrid
                .Source = BGASQL
                .ActiveConnection = DBado
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .CursorLocation = adUseClient
                .Open
            End With
            Set DBGridNeu.DataSource = rsDataGrid

            If rsDataGrid.RecordCount = 0 Then
                'MsgBox "Recordcount = 0"
            Else
                rsDataGrid.MoveFirst
            End If
            Call SpaltenbreiteWiederherstellen
'            'Horizontalen Scrollbalken wieder so einstellen wie vor dem Sortieren
'            DBGridNeu.Scroll intLeftcol, 0
            Cancel = True
        End If
    Else
        Set fso = Nothing
        Unload frmConnectSQL
        Unload Renam
        Unload AendernForm
    End If
End Sub

Public Sub SpaltenBreite()
    Set Col1 = DBGridNeu.Columns(0)
    Set Col2 = DBGridNeu.Columns(1)
    Set Col3 = DBGridNeu.Columns(2)
    Set Col4 = DBGridNeu.Columns(3)
    Set Col5 = DBGridNeu.Columns(4)
    Set Col6 = DBGridNeu.Columns(5)
    Set Col7 = DBGridNeu.Columns(6)
    Set Col8 = DBGridNeu.Columns(7)
    Set Col9 = DBGridNeu.Columns(8)
    Set Col10 = DBGridNeu.Columns(9)
    Set Col11 = DBGridNeu.Columns(10)
    Set Col12 = DBGridNeu.Columns(11)
    Set Col13 = DBGridNeu.Columns(12)
    Set Col14 = DBGridNeu.Columns(13)
    Set Col15 = DBGridNeu.Columns(14)
    
    Col1.Width = 600    'Merker
    Col2.Width = 500    'Jahr
    Col3.Width = 4000   'Situation
    Col4.Width = 2000   'Ort
    Col5.Width = 1000   'Land
    Col6.Width = 4000   'Personen
    Col7.Width = 5000   'Dateiname          'Gerbing 24.01.2005
    Col8.Width = 500    'SWF
    Col9.Width = 1000      'Kommentar
    Col10.Width = 1000     'DateinameKurz
    Col11.Width = 500   'DDatum
    Col12.Width = 500   'BreitePixel
    Col13.Width = 500   'HoehePixel
    Col14.Width = 500   'AudioFileExists
    Col15.Width = 500   'IPTCPresent
End Sub

Public Sub SpaltenbreiteMerken()
    Dim Col As Column
    Dim n As Long
    Dim ColWidth As Long

    'Bei jedem Speichern der Spaltenbreiten wird der bisherige Inhalt der Listbox lstSpaltenbreite zuerst
    'gelöscht, dann werden neue Einträge gemacht
    On Error GoTo 0
    lstSpaltenbreite.Clear
    For n = 0 To DBGridNeu.Columns.Count - 1
        Set Col = DBGridNeu.Columns(n)
        ColWidth = Col.Width
        If Col.Visible = False Then ColWidth = 0
        lstSpaltenbreite.AddItem ColWidth
    Next n
End Sub

Public Sub SpaltenbreiteWiederherstellen()
    Dim Col As Column
    Dim n As Long

    If CommandLine = "BGA" Then
        On Error Resume Next
    End If
    For n = 0 To lstSpaltenbreite.ListCount - 1
        Set Col = DBGridNeu.Columns(n)
        Col.Width = lstSpaltenbreite.List(n)
    Next n
    On Error GoTo 0
End Sub


Private Sub SpracheFestlegen()
    Dim strTemp As String
    Dim strPrimaryKey As String
    Dim SQL As String
    Dim msg As String
    Dim n As Long

    'Untersuche ob Access-Version oder SQL-Server-Version
    If gblnProversion = True Then                                               'Gerbing 04.03.2012
        'im Fall von SQL-Server-Version wird das frmConnect Formular gezeigt
        If gblnSQLServerVersion = True Then
            gblnSQLServerConnected = False
            If gblnCommandLineEmpty = True Then
                frmConnectSQL.Show 1
                If gblnSQLServerConnected = False Then
                    'msgbox "no connection to sql server"
                    MsgBox LoadResString(2460 + Sprache)
                    End
                End If
            End If
        End If
    End If

    Set DBado = New ADODB.Connection
    If gblnSQLServerVersion = True Then
        With DBado
            .Provider = "SQLOLEDB.1"
            '.Provider = "SQLNCLI10.1" 'SQL Server Native Client
            .Properties("Persist Security Info").Value = False
            .Properties("Initial Catalog").Value = PublicSQLDatabase
            .Properties("Data Source").Value = PublicSQLServer
            '   Falls die Windows-Authentifizierung verwendet werden soll, muß "SSPI" benutzt werden
            If PublicWindowsAuthentication = "1" Then
                .Properties("Integrated Security").Value = "SSPI"
            Else
                .Properties("User ID").Value = PublicSQLServerUserName
                .Properties("Password").Value = PublicSQLServerPassword
            End If
            .Open
        End With
        PublicDatagridCaption = PublicSQLServer & " " & PublicSQLDatabase
    Else
        DBado.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AppPath & "\fotos.mdb" 'Gerbing 23.11.2017
        DBado.mode = adModeReadWrite
        On Error Resume Next
        DBado.Open                                                          'Gerbing 04.03.2012 hier kommt runtime error wenn fotos.mdb fehlt
        If Err.Number <> 0 Then                                             'Gerbing 29.03.2012
            msg = "shareware version" & vbNewLine
            msg = msg & Err.Number & vbNewLine
            msg = msg & Err.Description
            MsgBox msg
            End
        End If
        On Error GoTo 0
        '------------------------------------------------------------------------------------------------------
        'Kontrolle ob die Datenbank schreibgeschützt ist                                    'Gerbing 23.11.2017
        On Error Resume Next
        SQL = "UPDATE FET SET FN = 'test'"
        Set rstsql = New ADODB.Recordset
        With rstsql
            .ActiveConnection = DBado                                                       'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            '.CursorLocation = Query.enumCursorOrt
            .Source = SQL
            '     .CacheSize = 2
            .Open
        End With
        If Err.Number <> 0 Then
            Call VierUrsachenFürSchreibsperre                                               'Gerbing 23.11.2017
            End
        End If
        PublicDatagridCaption = AppPath & "\fotos.mdb"
    End If
    
    SQL = "SELECT * From fotos WHERE not filename Is Null;"
    On Error Resume Next
    'On Error GoTo 0
    'On Error GoTo QUERYERR
    If rstsql Is Nothing Then
        Set rstsql = New ADODB.Recordset
    Else
        rstsql.Close
    End If
    Err.Number = 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBado
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    '-------------------------------------------------------------
    If Err.Number <> 0 Then
        Call WriteGlL(0)     'Rückschreiben deutsch in fotos.ini
        Sprache = 0
        Call GlL                                                                        'Gerbing 02.09.2008
        If PublicLanguage = "1" Then                                                    'Gerbing 04.12.2011
            Call VierUrsachenFürSchreibsperre
            End
        End If
    Else
        Call WriteGlL(1)     'Rückschreiben english in fotos.ini
        Sprache = 3000
        Call GlL                                                                        'Gerbing 02.09.2008
        If PublicLanguage = "0" Then                                                    'Gerbing 04.12.2011
            Call VierUrsachenFürSchreibsperre
            End
        End If
    End If
    rstsql.Close
    '-------------------------------------------------------------------------------------
    If gblnSQLServerVersion = False Then
        Set rstsql = DBado.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, "Fotos")) '2529=fotos
        If rstsql.EOF = True Then
            'Msg = "Seit Version 13.4.0 verlangt das Programm in Tabelle 'fotos' Spalte 'Dateiname' einen Primärschlüssel. Dieser wird jetzt erzeugt." & vbnewline
            'msg = msg & "Diese Operation wird nur dann erfolgreich sein, wenn in der Tabelle 'fotos' Spalte 'Dateiname' keine Duplikate vorkommen." & vbnewline
            'msg = msg & "Wenn die Operation nicht erfolgreich ist, müssen Sie zuvor die Duplikate entfernen." & vbnewline
            'msg = msg & "Benutzen Sie dazu eine frühere Version von fotosmdb.exe als 13.3.4"
            msg = LoadResString(1825 + Sprache) & vbNewLine
            msg = msg & LoadResString(1826 + Sprache) & vbNewLine
            msg = msg & LoadResString(1827 + Sprache) & vbNewLine
            msg = msg & LoadResString(1828 + Sprache) & vbNewLine
            MsgBox msg
            'SQL = "Create UNIQUE INDEX Dateiname ON fotos (Dateiname)  WITH PRIMARY"
            SQL = "Create UNIQUE INDEX " & LoadResString(1028 + Sprache) & " ON Fotos (" & LoadResString(1028 + Sprache) & ") WITH PRIMARY"
            On Error Resume Next
            DBado.Execute SQL
            If Err.Number <> 0 Then
                msg = "error number=" & Err.Number & vbNewLine
                msg = msg & "errortext=" & Err.Description
                MsgBox msg
                End
            End If
        Else
            strPrimaryKey = rstsql.Fields("COLUMN_NAME").Value
            If StrComp(LoadResString(1028 + Sprache), strPrimaryKey, vbTextCompare) <> 0 Then       '1028=Dateiname
                'MsgBox "Die Spalte Dateiname ist nicht der Primärschlüssel. Das Programm wird beendet."
                MsgBox LoadResString(1824 + Sprache)
                End
            End If
        End If
    End If
End Sub

Public Sub VierUrsachenFürSchreibsperre()                                                'Gerbing 02.09.2008
    Dim msg As String
    
    'vier mögliche Ursachen
    'Msg = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14\fotos.ini" & vbNewLine
    msg = gstrFotosIniAnwendungsOrdner & "\fotos.ini" & vbNewLine
    'msg = msg & "Die Datei ist schreibgeschützt. Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können." & vbnewline
    msg = msg & LoadResString(2275 + Sprache) & vbNewLine
    'msg = msg & "Es gibt vier mögliche Ursachen für den Lesemodus:" & vbnewline
    msg = msg & LoadResString(2133 + Sprache) & vbNewLine
    'msg = msg & "1. Das Dateiattribut 'Schreibgeschützt' ist gesetzt" & vbnewline
    msg = msg & LoadResString(2134 + Sprache) & vbNewLine
    'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte für Ihren PC" & vbnewline
    msg = msg & LoadResString(2135 + Sprache) & vbNewLine
    'msg = msg & "3. Sie arbeiten mit einer CD oder DVD" & vbnewline
    msg = msg & LoadResString(2136 + Sprache) & vbNewLine
    'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & vbnewline & vbnewline
    msg = msg & LoadResString(2137 + Sprache) & vbNewLine & vbNewLine
    'MsgBox Msg, , LoadResString(1119 + Sprache)
    MessageBoxW 0, StrPtr(msg), StrPtr(LoadResString(1119 + Sprache)), vbInformation
End Sub

Private Sub BildFehler(EchterStandort)
    Dim msg As String
    
'    msg = "Bild kann nicht geladen werden" & vbnewline
'    msg = msg & EchterStandort & vbnewline
'    msg = msg & "Prüfen Sie ob diese Datei existiert" & vbnewline
'    msg = msg & "oder ob es sich um einen verbotenen Dateityp handelt."
    msg = LoadResString(2056 + Sprache) & vbNewLine
    msg = msg & EchterStandort & vbNewLine
    msg = msg & LoadResString(2301 + Sprache) & vbNewLine
    msg = msg & LoadResString(2302 + Sprache)
    'MsgBox Msg, vbInformation
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Renammdb"), vbInformation
End Sub

Private Sub FremdeFotosMdb()                                                                    'Gerbing 10.09.2013
    Dim NetzwerkDir As String
    Dim msg As String
    
begin:
    '(ByVal Filter$, ByVal InitialDir$, ByVal Title$) as String
    NetzwerkDir = ShowOpenUnicodeFotosMdb(Me)    '2458=Standort der fotos.mdb
    'NetzwerkDir = GetOpenName(Filter, AppPath, LoadResString(2458 + Sprache))    '2458=Standort der fotos.mdb
    'Convert the file name to be used
    NetzwerkDir = ConvertFileName(NetzwerkDir)
    If NetzwerkDir = "" Then
        Exit Sub
    End If
    If Mid(NetzwerkDir, Len(NetzwerkDir) - 9, 1) <> "\" Then
        msg = LoadResString(2459 + Sprache)                  '2459=Sie müssen die Datei fotos.mdb auswählen
        MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo begin
    End If
    If StrComp(Right(NetzwerkDir, 9), "fotos.mdb", vbTextCompare) = 0 Then
        AppPath = Mid(NetzwerkDir, 1, Len(NetzwerkDir) - 10)
    Else
        msg = LoadResString(2459 + Sprache)
        MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo begin
    End If
End Sub

Private Function RekursiveTempThumbs(Path As String, SearchStr As String)        'Gerbing 06.04.2017
    Dim FileName As String              ' Walking filename variable...
    Dim DirName As String               ' SubDirectory Name
    Dim dirNames() As String            ' Buffer for directory name entries
    Dim nDir As Long                    ' Number of directories in this path
    Dim i As Long                       ' For-loop counter...
    Dim hSearch As Long                 ' Search Handle
    Dim wfd As WIN32_FIND_DATA
    Dim Cont As Long
    Dim DirCount As Long
    Dim DateinamenErweiterung As String
    Dim rc As Long
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    If gblnSubdirectories = True Then
        Cont = True
        hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(wfd))
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While Cont
            'DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            DirName = RemoveNulls((wfd.cFileName))
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                    dirNames(nDir) = DirName
                    DirCount = DirCount + 1
                    nDir = nDir + 1
                    ReDim Preserve dirNames(nDir)
                End If
            End If
            Cont = FindNextFileW(hSearch, VarPtr(wfd)) 'Get next subdirectory.
            Loop
            Cont = FindClose(hSearch)
        End If
    End If
    ' Walk through this directory.
    hSearch = FindFirstFileW(StrPtr(Path & SearchStr), VarPtr(wfd))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            'Filename = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            FileName = RemoveNulls((wfd.cFileName))
            If (FileName <> ".") And (FileName <> "..") Then
                '---------------------------------------------------
                DateinamenErweiterung = UCase(Right(FileName, 3))
                Select Case DateinamenErweiterung
                    Case "JPG"
                        rc = file_delete(Path & FileName, False, True) 'ohne Papierkorb, silent
                End Select
                '---------------------------------------------------
            End If
            Cont = FindNextFileW(hSearch, VarPtr(wfd)) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    If gblnSubdirectories = True Then
    ' If there are sub-directories...
        If nDir > 0 Then
            ' Recursively walk into them...
            For i = 0 To nDir - 1
                RekursiveTempThumbs = RekursiveTempThumbs(Path & dirNames(i) & "\", SearchStr)
            Next i
        End If
    End If
End Function

Private Function RemoveNulls(OriginalString As String) As String
    Dim pos As Long
    pos = InStr(OriginalString, Chr$(0))
    If pos > 1 Then
        RemoveNulls = Mid$(OriginalString, 1, pos - 1)
    Else
        RemoveNulls = OriginalString
    End If
End Function

