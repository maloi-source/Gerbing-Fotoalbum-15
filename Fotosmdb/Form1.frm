VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FotosMdb"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   15288
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   15288
   StartUpPosition =   1  'Fenstermitte
   Begin EditCtlsLibUCtl.TextBox txtFehlerU 
      Height          =   492
      Left            =   3600
      TabIndex        =   23
      Top             =   6720
      Width           =   11652
      _cx             =   20553
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
      DisabledEvents  =   3072
      DisabledForeColor=   16711680
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
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
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
      CueBanner       =   "Form1.frx":038A
      Text            =   "Form1.frx":03AA
   End
   Begin EditCtlsLibUCtl.TextBox txtArbeitsfortschrittU 
      Height          =   372
      Left            =   3600
      TabIndex        =   22
      Top             =   7440
      Width           =   11652
      _cx             =   20553
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
      CueBanner       =   "Form1.frx":03CA
      Text            =   "Form1.frx":03EA
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
      Height          =   492
      Left            =   10080
      TabIndex        =   21
      Text            =   "txtFont"
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSDataGridLib.DataGrid DBGridNeu 
      Height          =   2892
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   15132
      _ExtentX        =   26691
      _ExtentY        =   5101
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
   Begin VB.Frame FrameEXIFIPTC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIF/IPTC zur�ckschreiben"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   11280
      TabIndex        =   18
      Top             =   4440
      Width           =   3852
      Begin VB.CommandButton btnEXIFIPTC 
         Caption         =   "EXIF/IPTC..."
         Height          =   492
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1572
      End
   End
   Begin VB.CommandButton btnL�scheInhaltFotosMdb 
      Caption         =   "L�sche den Inhalt von fotos.mdb..."
      Height          =   612
      Left            =   8520
      TabIndex        =   17
      Top             =   8040
      Width           =   3252
   End
   Begin VB.CommandButton btn�ffnePruefLog 
      Caption         =   "�ffne die Datei pruef.&log"
      Height          =   612
      Left            =   3600
      TabIndex        =   16
      Top             =   8040
      Width           =   3252
   End
   Begin VB.ListBox lstSpaltenbreite 
      Height          =   240
      Left            =   12360
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnNutzerdefinierteFelderAnlegen 
      Caption         =   "Nutzerdefiniertes Datenbank-&Feld anlegen..."
      Height          =   612
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   7092
   End
   Begin VB.CommandButton btnBeenden 
      Caption         =   "B&eenden"
      Height          =   612
      Left            =   12000
      TabIndex        =   9
      Top             =   8040
      Width           =   3252
   End
   Begin VB.CommandButton btnGenerieren 
      Caption         =   "&Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)..."
      Height          =   612
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7692
   End
   Begin VB.CommandButton btnHilfe 
      Caption         =   "H&ilfe"
      Height          =   612
      Left            =   120
      TabIndex        =   7
      Top             =   8040
      Width           =   3252
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Umkehr-Probe machen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   6360
      TabIndex        =   2
      Top             =   4440
      Width           =   3492
      Begin VB.CommandButton btnPr�fenSAbbrechen 
         Caption         =   "A&bbruch"
         Enabled         =   0   'False
         Height          =   492
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         Width           =   1452
      End
      Begin VB.CommandButton btnPr�fenS 
         Caption         =   "Pr�fen&S"
         Height          =   492
         Left            =   1920
         TabIndex        =   14
         ToolTipText     =   "ob es Differenzen zwischen vorhandenen Audio-Kommentaren und der Spalte 'AudioFileExists' gibt"
         Top             =   480
         Width           =   1452
      End
      Begin VB.CommandButton btnPr�fen3Abbrechen 
         Caption         =   "Abbru&ch"
         Enabled         =   0   'False
         Height          =   492
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1452
      End
      Begin VB.CommandButton btnPr�fen3 
         Caption         =   "Pr�fen&3"
         Height          =   492
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   $"Form1.frx":040A
         Top             =   480
         Width           =   1452
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datenbank fotos.mdb auf g�ltigen Inhalt pr�fen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   4812
      Begin VB.CommandButton btnPr�fen1Abbrechen 
         Caption         =   "Abbru&ch"
         Enabled         =   0   'False
         Height          =   492
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1452
      End
      Begin VB.CommandButton btnReset 
         Caption         =   "&Reset"
         Height          =   492
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "zur�ck zum Inhalt von fotos.mdb"
         Top             =   1200
         Width           =   1452
      End
      Begin VB.CommandButton btnPr�fen2 
         Caption         =   "Pr�fen&2"
         Height          =   492
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "ob die Jahreszahl im Feld 'Jahr' und im Dateiname �bereinstimmt"
         Top             =   480
         Width           =   1452
      End
      Begin VB.CommandButton btnPr�fen1 
         Caption         =   "Pr�fen&1..."
         Height          =   492
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "ob jede im Feld Dateiname eingetragene Foto-Datei  wirklich existiert."
         Top             =   480
         Width           =   1452
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   492
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   6612
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   11663
      _cy             =   868
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pr�fergebnis:"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Falls Fehler auftreten, klicken Sie zum �ffnen der Datei pruef.log auf den Fehlerhinweis"
      Top             =   6840
      Width           =   2412
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arbeitsfortschritt:"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   7560
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 
 '11.02.2004 Dateinamen-Erweiterung "HTM", "PDF", "XLS" wird zugelassen
'16.03.2004 Dateinamen-Erweiterung "WMV" wird zugelassen
'19.03.2004 Option Compare Text muss benutzt werden sonst geht trotz Umwandlung in Ucase der Vergleich
'           von 2 Strings B�rbelGeburtstag - BIHHPanorama falsch aus
'           Schuld ist das �
'22.03.2004 Wenn Pr�fen3 feststellt, da� bestimmte Fotos nicht in der Datenbank fotos.mdb stehen,
'           gibt es jetzt zwei Alternativen:
'           1. Die gefundenen Dateien geh�ren nicht in die angegebenen Ordner -> L�schen
'           2. Die gefundenen Dateien sollen benutzt werden um neue S�tze in der Datenbank fotos.mdb
'           zu generieren
'26.03.2004 Fehlerkorrektur: die beiden Listboxen
'           NachPr�fen3L�schen.lstZus�tzlicheDateien und NachPr�fen3Aufnehmen.lstZus�tzlicheDateien
'           m�ssen vor der Wiederholung eines Laufes Pr�fen3 gel�scht werden
'26.04.2004 Angleichen der Jahreszahl nach Pr�fen2 meldet keinen Fehler, wenn
'           im Zielordner bereits eine Datei mit dem gleichen Namen steht.
'           Ich brauche eine M�glichkeit zum manuellen Eingreifen.
'25.05.2004 Pr�fenH stellt fest, ob es Felder mit doppelten Hochkommas gibt
'           Wenn welche gefunden werden erh�lt der Nutzer Schreibzugriff auf DBGridNeu
'           Bei Reset wird auch der Schreibzugriff zur�ckgesetzt
'02.06.2004 Nach Pr�fenD wird Doppelklick ins DBGridNeu erlaubt, damit kann die Duplikat-Zeile
'           ausgew�hlt werden, die verschwinden soll.
'11.08.2004 miscA55.ico soll nicht als Pr�fen3-Fehler ausgewiesen werden
'           Bei Pr�fen4 und Pr�fen5 fehlt die Sanduhr
'17.09.2004 Es gibt jetzt eine Abfrage f�r Pr�fen2, ob Spalte 'Jahr' und Jahr in Spalte 'Dateiname'
'           �bereinstimmt. Das geht viel schneller.
'30.09.2004 F�r Pr�fen1 funktioniert zwar im Access ein Abfrage, aber die geht nicht unter VB DAO
'           Da will ich wenigstens eine schnellere M�glichkeit als Dir(Dateiname) finden, die gibt es mit
'           Open Fotodatei For Input As #Pr�fDateiNummer
'30.09.2004 Pr�fen3 soll beschleunigt werden durch eine Inkonsistenzabfrage
'           wo zwei Tabellen miteinander verglichen werden
'           ob die wirklichen Dateinamen(Tabelle Temp_Haken) auch in der Tabelle Fotos vorkommen.
'30.09.2004 Pr�fen4 und Pr�fen5 kommen ohne Call Rekursive aus
'04.10.2004 Pr�fen3 verbessern hinsichtlich falsch gew�hltem Fotos-Root-Ordner. Abbruchm�glichkeit
'10.10.2004 Die Arbeit mit der Tabelle 'FotosMitZusatzSpalte' dauert zu lange und au�erdem gehen oft dort
'           gemachte �nderungen verloren, bevor sie in die Tabelle Fotos kopiert werden k�nnen. Ich will
'           auf die Tabelle 'FotosMitZusatzSpalte' ganz verzichten und in die Tabelle Fotos zwei neue
'           Felder aufnehmen:
'           DateinameKurz (Namensanteil von Dateiname)
'           DDatum (Datei Erstellungs Datum)
'           'Pr�fen1' muss eventuelle Ungleichheiten zwischen Dateiname und DateinameKurz korrigieren
'           'Pr�fen1' muss von jeder Datei das Erstellungsdatum lesen und in Spalte DDatum eintragen
'           'Pr�fen1' l��t die Kontrolle Jahr >= 1851 weg
'           'Neue Datens�tze generieren' muss auch die Felder DateinameKurz und DDatum f�llen
'11.10.2004 Die Funktion Ersetzen l��t sich gewaltig beschleunigen, wenn ein Recordset rst benutzt wird,
'           anstelle von Adodc1.Recordset
'11.10.2004 Die Funktion Pr�fen1 l��t sich gewaltig beschleunigen, wenn ein Recordset rst benutzt wird,
'           anstelle von Adodc1.Recordset
'20.10.2004 Fehlerkorrektur:
'           Wenn man Pr�fenD oder Pr�fenH mehrmals nacheinander gemacht hat, wurde die Grid-�berschrift
'           immer l�nger.
'           Wenn man sofort nach Pr�fenD oder Pr�fenH ein Pr�fen4 oder Pr�fen5 gemacht hat,
'           (mit alle S�tze pr�fen), kam Abbruch wegen 'no records in recordset'
'           Gegenma�nahme: Call btnReset_Click
'13.11.2004 Bei der 'Ersetzen-Operation' d�rfen keine Ordner-Namen entstehen, die es auf dem PC
'           garnicht gibt.
'           Gegenma�nahme1: der neue Ordner darf nicht eingetippt werden k�nnen, sondern wird ausgew�hlt
'           Gegenma�nahme2: Undo f�r die zuletzt gemachte Ersetzen-Operation,
'                           daf�r wird mit einer Tabellenerstellungsabfrage die Tabelle TempFotos
'                           erzeugt.
'           Gegenma�nahme3: Komprimieren der Datenbank, weil diese bei Drop Table immer gr��er wird
'29.12.2004 �nderungen in der Spaltenbreite sollen mindestens f�r die aktuelle Sitzung gespeichert werden
'30.12.2004 Das DbGrid war nicht immer nach Dateiname aufsteigend sortiert.
'           Das Einstellen auf das Ende der Datei mit dem Rollbalken war m�hselig, L�sung: Data Control
'           anstelle von unsichtbar sichtbar machen.
'           Das Vorsetzen auf die markierte Datei bei Pr�fen4 oder Pr�fen5 hat zu lange gedauert, L�sung:
'           FindFirst benutzen.
'           Die nach Pr�fen4(Falsch beschnitten) oder Pr�fen5(wei�e Kanten) gefundenen Dateien k�nnen ab
'           jetzt sofort im favoritisierten Bildbearbeitungsprogramm �berarbeitet werden. Dazu in Pruef.log
'           den Dateiname markieren und dann mit der rechten Maustaste daraufklicken.
'02.01.2005 Fehlerkorrektur: im Formular JahrFestlegen war noch ein Hinweis auf Jahr 1850
'22.01.2005 Fehlerkorrektur: im Tooltiptext zu Pr�fen1 war noch ein Hinweis auf Jahr 1850
'           Weglassen Adodc1.Refresh am Ende von Form1.Form_Load weil damit die Standardspaltenbreite
'           wiederkommt
'14.02.2005 Ab Version 10:
'           Das Fotoalbum kann mit nutzerdefinierten Feldern arbeiten
'           Fotosmdb.exe l��t das Anlegen jeweils eine Feldes vom Typ Text zu.
'           Andere Felder sollen mit MS Access angelegt werden.
'           Neue Datens�tze generieren mit Drag&Drop hat gemeckert beim rst.update mit Fehler
'           'AllowZeroLength property is False"
'           Ich habe daraufhin im MS Access Tabellenentwurf die Felder Situation, Ort, Land, Personen,
'           Kommentar mit 'Leere Zeichefolge = Ja' definiert
'           Beim Speichern der Spaltenbreite werden auch die nutzerdefinierten Felder ber�cksichtigt.
'08.03.2005 Es gibt 2 neue Standardfelder in der Tabelle fotos. Das sind BreitePixel und HoehePixel
'           Bei Pr�fen1 und bei 'Neue Datens�tze generieren (durch Drag&Drop vom Windows Explorer)...'
'           werden diese Felder gef�llt.
'           Ich lese den header der Dateien vom Typ AVI ein,
'           f�r die Bilddateien dient Pr�fImage, andere Dateitypen sind nicht ber�cksichtigt
'12.03.2005 Es gibt das neue Feld DatumBreiteHoehe in der Tabelle ErsterStart, dort wird das Datum
'           der Berechnung von BreitePixel und HoehePixel eingetragen. Neuberechnung von
'           BreitePixel und HoehePixel wird nur f�r die Dateien gemacht, deren Dateidatum
'           gleich oder aktueller ist als DatumBreiteHoehe.
'17.03.2005 Fehlerkorrektur: bei 'Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)...'
'           fehlte die Sanduhr
'03.04.2005 irgendwas l�uft falsch bei Wiederholung der Berechnung PixelBreite PixelHoehe
'11.04.2005 Man k�nnte das Programm fotos.exe so ver�ndern, dass es selbst versucht die Bezeichnung des
'           Fotos-Root-Ordner zu ermitteln, n�mlich als AppPath wo fotos.exe steht.
'           Dazu m��te man in die Datenbank eintragen
'           anstelle von zB M:\P7FotoSoundVideo\FOTOS\GG\2005\Ballonfahrt001.jpg
'           +:\2005\Ballonfahrt001.jpg und bei der Ausf�hrung von fotos.exe fotosmdb.exe und renammdb.exe muss
'           +:\ ersetzt werden durch AppPath des entsprechenden Programms.
'           Dann entf�llt die Funktion Ersetzen im Programm fotosmdb.exe,
'           und alles was zusammenh�ngt mit 'Fotos-Root-Ordner Festlegen' bei Start von einer CD,
'           und das Feld ErsterStart.ErsterStart wird nicht mehr ausgewertet.
'           aber man muss vom Nutzer verlangen, dass er s�mtliche Dateien unterhalb von AppPath von
'           fotos.exe anlegt. Daf�r kann er die 3-Einigkeit von fotos.exe, fotos.mdb und Dateien kopieren
'           oder verschieben wohin er will.
'             Pr�fen der 3-Einigkeit ist in jedem Programm fotos.exe fotosmdb.exe renammdb.exe n�tig.
'           Man muss dazu pr�fen, ob der erste Satz der Tabelle Fotos, nach Ersetzen des String +:\
'           durch AppPath eine Datei ergibt, die existiert
'           oder bei FotosMdb.exe mu� beim Neue Datens�tze generieren durch Drag&Drop nach dem
'           strTemp = Replace(AktuellerDateiName, AppPath, "+:" & "\")
'           gepr�ft werden ob wirklich +: am Anfang von strTemp steht
'6.2005     Pr�fen4 und Pr�fen5 soll keinen Fehler liefern, wenn ausgew�hlt wird
'           'beginnend mit markiertem Datensatz' und anstelle des Dateinamens ein anderes Feld markiert ist
'6.2005     �nderungen zum Anbieten einer Light-Version oder einer Professional-Version
'           Die Professional-Version gibts nach Anforderung per E-mail an gerbing.software@freenet.de
'           Daraufhin erh�lt der Kunde die Kontonummer zur �berweisung des Kaufbetrages
'           auf ein noch festzulegendes Konto.
'           Der Kunde bekommt dann einen Freischalteschl�ssel per Post geschickt, mit diesem muss er die
'           Light-Version nochmal installieren.
'           Wenn das Programm ohne Freischalteschl�ssel installiert wird, oder mit FX58A-C3BYE-1FGH3-B3YFG-FX2BA
'           dann l�uft es unbegrenzte Dauer als Light-Version.
'           Ich erkenne die Light-Version am Fehlen der Datei msprivs.log in ...\windows\Systemdirectory.
'           Die datei msprivs.log wird durch RegLight.exe mit einem g�ltigen Professional Schl�ssel erzeugt.
'           Light-Version:
'           -ohne benutzerdefinierte Felder
'10.06.2005 Beim Erzeugen neuer S�tze mit Drag&Drop
'           wird f�r jedes Feld eine Combobox mit den schon vorhandenen Werten angeboten
'19.06.2005 Tooltip beim Hinzuf�gen nutzerdefinierter Felder mittels Combobox
'           Irrt�mlich eingetragene Feldnamen entfernen Sie mit der Return-Taste oder Entf-Taste im numerischen Tastenfeld
'20.07.2005 Verbesserung der Msgbox nach �ffnen von Pr�f.log, nach Pr�fen2,
'           wenn es den Ordner mit dem geforderten Jahr nicht gibt
'20.07.2005 Durch die Einhaltung der Dreieinigkeit kann ich "Pr�fen2" und anschlie�end Verschieben in den
'           richtigen Jahresordner verbessern. Dateien aus beliebigen Ordnern ohne Jahreszahl k�nnen
'           jetzt verschoben werden in Ordner mit Jahreszahl. jetzt besteht nur noch die Gefahr, dass
'           Duplikate entstehen. Der Nutzer soll im Formular Duplikatname die Chance zum Abbrechen
'           bekommen.
'15.08.2005 Fehlerkorrektur:
'           Bisher wurden bei der zweiten Ausf�hrung von Pr�fen3 die bei der ersten Ausf�hrung von
'           Pr�fen3 schon aufgenommenen Dateien immer noch als nicht aufgenommen angezeigt.
'11.12.2005 Verbesserung:
'           Ich werde ab jetzt unterscheiden zwischen Bildern, die ich im native mode anzeigen kann
'           "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF"
'           und Bildern, die ich nur im link mode anzeigen kann.
'           F�r den link mode benutze ich ShellExecute, so wie schon bisher bei Dateityp "HTM", "PDF", "XLS"
'           im link mode kann man dann beispielsweise die Dateitypen
'           "PNG" "PSD" TIF" betrachten
'           F�r "PNG" und "TIF" kann man zB die Windows Bild- und Faxanzeige benutzen,
'           da �ffnet sich f�r jedes neue Bild immer dasselbe Fenster.
'           F�r "PSD" kann man zB den Quicktime Picture Viewer benutzen,
'           da �ffnet sich f�r jedes Bild ein neues Fenster.
'09.01.2006 In Spalte SWF  wird jetzt erlaubt
'           If SWF = "F" Or SWF = "SW" Or SWF = "C" Or SWF = "BW" Then
'           bzw
'           If SWF = "FV" Or SWF = "SV" Or SWF = "CV" Or SWF = "BV" Then
'10.01.2006 Unterst�tzung von sprache siehe fotos.vbp
'           Die Merker-Spalte ist fester Bestandteil der Tabelle Fotos
'16.01.2006 Neue Funktion nach Pr�fen1.
'           Man kann Datens�tze aus der Datenbank entfernen, die zu massenhaft manuell gel�schen
'           Dateien geh�ren. Beispielsweise wenn man diese Bilder tats�chlich nicht mehr in der
'           Datenbank haben will.
'18.01.2006 Pr�fen1 muss 2 neue Pr�fkriterien bekommen
'           Durch �bergang auf Englisch/deutsch gibt es keine Pflichtfelder Jahr und Dateiname mehr
'           Darum muss jetzt das Programm Pr�fen ob in jedem Datensatz Jahr und Dateiname
'           eingetragen ist.
'           Ein Datensatz ohne Dateiname muss sofort gel�scht werden, denn ich finde ihn sonst nicht wieder
'19.01.2006 Verbesserung:
'           bei Pr�fen1 Subfunktion 'BreitePixel und HoehePixel eintragen' lasse ich den Nutzer entscheiden,
'           ob er alle Dateien pr�fen will, oder nur die mit dem h�heren Datum als der letzte
'           Pr�fen1-Vorgang oder garnicht
'           Dazu gibt es ein neues Formular PixelAusrechnen
'22.02.2006 Suche nach dem Problem:
'           warum geht Pr�fen1 im Windows 2000 so langsam
'           am mei�ten Zeit fri�t die Funktion 'Open Fotodatei For Input As #Pr�fDateiNummer'
'           ich benutze stattdessen 'FileDateTime(Fotodatei)'
'14.04.2006 Die Version GERBING Fotoalbum 13.0.1
'           ist ausgeliefert worden mit einer Beispieldatenbank, wo man beim Erzeugen neuer S�tze in der
'           Datenbank fotos.mdb keine Felder leer lassen darf, sonst kommt Laufzeitfehler '3315'.
'           Umgehungsl�sung: Man muss einmal die Sprache wechseln und anschlie�end wieder zur�ckwechseln.
'12.04.2006 Neue Funktion (nur Professional Version):
'           Gesprochener Kommentar (Audio-Datei)
'           Neue Funktion 'Pr�fenS' Datenbankfeld AudioFileExists bereinigen. Priorit�t hat eine
'           vorhandene/nichtvorhandene Audio-Datei.
'           Dabei werden alle Audio-Dateien gel�scht, zu denen es keine zugeh�rige Foto-Datei gibt.
'           Dadurch kann man per Windows Explorer ungew�nschte
'           Audio-Dateien einfach l�schen und danach die Datenbank korrigieren.
'14.05.2006 AudioFileExists geh�rt nicht zu den nutzerdefinierten Feldern
'30.05.2006 Wenn ich Fotosmdb.exe als Tool aus Fotos.exe heraus starten lasse und benutze 'Pr�fen1'
'           dann kommt nach etwa 9500 Datens�tzen laufzeitfehler 3052 'Anzahl der Dateisperrungen �berschritten'
'           = File sharing lock count exceeded
'           Das ist vermutlich eine Macke von DAO 3.6
'           Gegenmassnahme laut Microsoft Knowledge base:
'           increase the maximum number of locks per file in your Registry
'           DBEngine.SetOption dbMaxLocksPerFile, 30000
'           Notl�sung: Fotos.exe beenden und Fotosmdb.exe solo starten
'24.06.2006 DateinameKurz wird ab jetzt immer eingetragen, auch wenn es die Datei nicht gibt
'10.08.2006 Unangenehme Eigenschaft:
'           Nach Pr�fen3 w�chst die Datei fotos.mdb an und manchmal wirkt bei Beenden von fotos.exe
'           das Komprimieren der Datenbank nicht. Der Versuch am Ende alle S�tze aus der Tabelle Temp_Haken
'           zu l�schen hat nichts gebracht.
'09.11.2006 Neue Funktion (alle Versionen):
'           Das alles geht nur mit jpg-Dateien, bei anderen Dateien bleiben die generierten Datenbankfelder leer.
'           Beim Generieren neuer Datenbanks�tze kann der Nutzer EXIF und/oder IPTC Metadaten benutzen.
'           Sowohl bei den Standardfeldern wie bei den nutzerdefinierten Feldern kann ein EXIF oder IPTC-Feld
'           ausgew�hlt werden, das als Datenquelle dient anstelle von manueller Eingabe.
'           F�r den Import des Feldes Jahr gibt es das neue Formular frmOptionJahr zum Ausw�hlen mehrerer
'           Optionen.
'           Wenn die Checkbox chkUnbeaufsichtigt aktiviert ist, muss es m�glich sein eine eventuell gel�schte
'           oder verlorengegangene Datenbank neu zu erzeugen, das geht aber nur mit jpg-Dateien.
'22.11.2006 Im Formular NachPr�fen3Aufnehmen wird ein Vorschaubild angezeigt, wenn der Nutzer
'           auf einen Dateinamen rechtsklickt
'02.12.2006 Verbesserung:
'           Ab sofort kommt ein Thumbnail im Formular JahrFestlegen, wenn beim 'Neue Datens�tze generieren'
'           kein Jahr im Dateiname vorkommt und der Nutzer eine Jahreszahl festlegen soll.
'23.01.2007 Fehlerkorrektur:
'           bisher kam bei schreibgesch�tzter Datenbank und 'Pr�fen3' Laufzeitfehler 3086 beim Versuch
'           alle S�tze derTabelle Temp_Haken zu l�schen
'14.03.2007 Verbesserung alle Versionen:
'           Neue Hilfe-Dateien im HTML-Format, weil Windows Vista das Winhelp-Format nicht mehr unterst�tzt
'           zB anstelle Fotosmdb.hlp gibt es jetzt Fotosmdb.chm
'07.05.2007 Fehlerkorrektur
'           Bisher hat das Anzeigen der EXIF bzw IPTC-Felder stichproben beim Generieren neuer Datens�tze nicht
'           funktioniert, weil tempFilename = Replace(List1.Text, "+:\", AppPath & "\") gefehlt hat
'07.05.2007 Verbesserung alle Versionen:
'           Gasanov EXIF OCX hat einen Fehler, es l��t bei Olympus Fotos frisch aus der Kamera einen Teil der
'           Felder weg. Erst nach Bearbeitung mit PSP9 sind alle Felder da.
'           Es wird ab jetzt ein Klassenmodul clsEFIF.cls benutzt
'07.05.2007 Nutzerdefiniertes Textfeld anlegen hat bisher mit Msgbox geantwortet "Sie m�ssen eine Zahl zwischen
'           1 und 255 eingeben
'10.07.2007 Bei 'NeueDatens�tze generieren' soll in Tabelle ErsterStart, Feld DatumBreiteHoehe das aktuelle
'           Datum eingetragen werden
'29.07.2007 Die folgenden ico Dateien sollen bei Pr�fen3 nicht angemeckert werden
'           FourArrows.ico
'           FourArrowsSave.ico
'           SquareZoom.ico
'23.08.2007 Verbesserung:
'           Bei Pr�fenS wird jede Aktion in pruef.log eingetragen
'01.09.2007 Verbesserung:
'           Geschwindigkeitsverbesserung beim Ermitteln von Bildbreite und Bildh�he bei
'           -Generieren neuer Datens�tze
'           -Pr�fen1
'           durch Benutzung der Funktion GdipGetImageDimension in der gdiplus.dll
'           m�glich mit BMP, DIB, JPG, GIF, PNG, TIF
'           nicht m�glich mit ICO, CUR, PSD
'           Falsches Ergebnis bei EMF, WMF
'24.10.2007 Verbesserung:
'           bei Videos aller Art kann man BreitePixel und HoehePixel ermitteln durch Benutzung des Controls
'           Mediaplayer1 (msdsm.ocx)
'           Wenn der Nutzer das nutzerdefinierte Feld VideoDuration benutzt kann auch Mediaplayer1.Duration
'           eingetragen werden. Als Suchargument ist VideoDuration nur in der Professional Version m�glich.
'21.11.2007 Fehlerkorrektur
'           bei falsch oder nicht registrierter dao360.dll
'           kommt die msgbox
'           Errornumber=429
'           Errortext=Objekterstellung durch ActiveX-Komponente nicht m�glich
'           You must register the dao360.dll
'           read in http://www.gerbingsoft.de or look for that problem in the internet
'           dann wird das Programm beendet
'06.01.2008 Fehlerkorrektur
'           zur IPTC-Anzeige
'           es gibt Felder, die l�nger sein k�nnen als eine Zeile in der Listbox lstExifIptc
'           diese Felder muss man in mehrere Zeilen zerlegen
'09.01.2008 Neue RAW-Datenformate
'           werden bei den Link-Datentypen erlaubt
'           3FR ARW CS1 CS4 CS16 DCS ERF MEF SR2
'09.01.2008 Fehler bei Pr�fen1(alles) wenn Videos im Format 'MOV' Panoramavideo berechnet werden sollen
'           L�sung: bei Fehler weiterarbeiten und Null eintragen
'10.01.2008 Verbesserung
'           Mit FotosMdb.exe Funktion 'IPTC...' kann man den Inhalt der Datenbankfelder in die JPG-Dateien
'           �bertragen. Damit geht man den umgekehrten Weg wie bei der Aufnahme neuer Dateis�tze wo die
'           IPTC-Felder in die Datenbank �bertragen werden k�nnen. Es geht ausschlie�lich mit JPG-Dateien.
'23.01.2008 Fehlerkorrektur
'           zum IPTC-Felder in die Datenbank �bertragen
'           Das Feld Jahr wurde bisher nicht angeboten
'04.02.2008 Verbesserung zu IPTC... - Inhalt der Datenbankfelder in die JPG-Dateien �bertragen.
'           Es soll das nachtr�gliche �bertragen von IPTC Feldern in JPG-Dateien erleichten.
'           Es soll m�glich sein, die Fotos zu finden, in die bisher keine IPTC Felder �bertragen wurden.
'           Das ist dann der Fall wenn das Datenbankfeld IPTCPresent=False ist.
'           IPTCPresent wird belegt durch die neue Funktion Pr�fenIPTC. Es wird True, wenn mindestens ein
'           IPTC Feld im JPG-Foto enthalten ist.
'           IPTCPresent=True wird belegt bei der Funktion 'IPTC...', wenn mindestens ein IPTC Feld entsteht.
'           IPTCPresent=True wird belegt beim 'Neue Datens�tze Generieren' wobei Quelle IPTC ist und wenigstens
'           ein IPTC Feld etwas enth�lt.
'10.03.2008 Button 'Pr�fenIPTC' muss disabled sein, w�hrend eine beliebige andere Pr�ffunktion l�uft.
'02.04.2008 Fehlerkorrektur
'           Wenn man w�hrend 'Pr�fenIPTC' zwischendurch andere Buttons anklickte, z�hlte der Z�hler
'           lblArbeitsfortschritt bis weit h�her als �berhaupt S�tze vorhanden waren.
'08.04.2008 Fehlerkorrektur
'           Bisher konnte man bei Wiederholung der Funktion Pr�fen3 ->
'           Pr�fen3 - Gefundene Dateien in die Datenbank aufnehmen...
'           ein eventuell vorher als letztes gezeigtes Bild sehen und das war irgendeins. Ab jetzt wird
'           in Form NachPr�fen3Aufnehmen gemacht Call BildAnzeigen("")
'29.05.2008 Verbesserung
'           Nach mehrmals Ausf�hren von Pr�fen3 w�chst die fotos.mdb stark an
'           bisher hat CompactDatabase mit der eigenen Datenbank nicht funktioniert
'           Schuld war das Fehlen von
'           dbs.Close
'           im Formular NeueDatens�tzeGenerieren
'01.09.2008 13.3.9 Verbesserung alle Versionen in fotos.exe
'           Zus�tzlich zu Mediaplayer 6.4 (msdxm.OCX) ist ausw�hlbar der aktuelle Windows Media Player (wmp.dll)
'           weil es passiert ist, dass einige Videoclips sich nicht abspielen lassen wollten.
'           Im FotosMdb bleibt Mediaplayer 6.4 (msdxm.OCX) f�r die Berechnung von BreitePixel, HoehePixel,
'           VideoDuration jedoch erhalten.
'02.09.2008 13.3.9 Verbesserung alle Versionen
'           Wenn wegen fehlender Berechtigung pruef.log schreibgesperrt ist, kommt ab jetzt ein Hinweis
'17.10.2008 Wenn Arbeit in einem Benutzerkonto ohne Administratorrechte kommt ein entsprechender Hinweis
'22.11.2008 13.3.11 Verbesserung alle Versionen
'           Pr�fen1 kann jetzt auch ohne irgendeine Aktion abgebrochen werden
'15.12.2008 Verbesserter Hinweis zu Pr�fen1 zusammen mit PixelAusrechnen
'03.01.2009 13.3.11 Verbesserung alle Versionen
'03.01.2009 Bei der Funktion 'IPTC...'
'           Neue Function SchreibePruefLogFehlerBeimZugriffAufDatei beispielsweise bei Dateien mit L�nge=0
'03.01.2009 Nach btnStart_Click der Funktion 'IPTC...'
'           Wenn man da in den Nutzerdefinierten Feldern herumgeklickt hat, kam Laufzeitfehler 3420
'           bei rst1.Movenext 'Objekt nicht mehr festgelegt
'04.01.2009 Verbesserung beim IPTC-Import/Export
'           Ab jetzt wird das Feld SWF ber�cksichtigt
'26.01.2009 13.3.11 Verbesserung alle Versionen
'           Alle Programme, die eine Listbox oder Combobox mit mehr als 32767 Zeilen f�llen, fallen irgendwann auf die Nase.
'           man kan mittels AddItem zwar viel mehr als 32767 Zeilen erzeugen, aber nicht h�her als 32767 Indizieren.
'           Das betrifft Pr�fen1 Pr�fen3 und 'Neue Datens�tze generieren'
'           L�sung:
'           Ich bringe stattdessen bei Pr�fen3 eine MsgBox, dass mehr als 32767 Eintr�ge nicht erlaubt sind, und bei Pr�fen1
'15.10.2009 dass mehr als 32767 Fehler in einem Durchlauf nicht gefunden werden.
'23.02.2009 13.3.11
'           nicht mehr in pruef.log eintragen ....keine JPG-Datei
'10.06.2009 13.3.13
'           Ich muss mitbekommen ob es die Spalte IPTCPresent gibt, wenn die Nachricht kommt
'           "Sie brauchen Administratorrechte und Schreibzugriff auf 'fotos.mdb' um �nderungen vorzunehmen" & vbNewLine
'           "You need administrator rights and write access on 'fotos.mdb' for making changes"
'23.06.2009 13.3.13
'           Beim Schreiben von IPTC-Daten in eine JPG-Datei ist bisher das Datum/Uhrzeit um 2 Stunden nach vorn verschoben worden
'           ab jetzt neue L�sung ohne Zeitverschiebung
'           Beim Synchronisieren mit Total Commander muss auch der rote Ungleich-Button geklickt werden
'03.11.2010 Die Abfrage, ob fotos.mdb vorhanden ist, kam zu sp�t, stattdessen kam Laufzeitfehler 91
'04.11.2010 Es kam bisher Laufzeitfehler 5 beim Eintragen von DDatum, wenn das Dateidatum folgendes Format hatte
'           27.05.2005 00:00:00
'           Dann hat n�mlich die Funktion strTemp = FileDateTime(Fotodatei) als strTemp den Wert 27.05.2005 zur�ckgegeben
'11.11.2010 Bei Pr�fen2 und dem anschlie�enden Verschieben in den richtigen Jahresordner ist
'           eine weitere Rename-Operation n�tig, wenn es eine gleichnamige Sounddatei WAV oder MP3 gibt
'14.02.2011:
'           Man k�nnte Blockierungen im Multi-Nutzer-Betrieb nicht wie bisher mit Laufzeitfehler beenden, sondern eine Msgbox bringen:
'           "Wiederholen Sie diese Funktion, wenn keine anderen Nutzer mit der Datenbank arbeiten"
'           der Fehler tritt auf bei btnPr�fen1_Click beim rst.Edit
'           Es kommt Laufzeitfehler 3260 Aktualisieren nicht m�glich; momentane Sperrung durch Benutzer 'admin' auf Computer 'Elke'
'           oder man k�nnte ausprobieren, ob die Blockierungen durch Benutzung der ADO-Schnittstelle verschwinden
'           es scheint so als g�be es keine Blockierungen mehr durch die ADO-Schnittstelle,
'           aber sobald in adoRS.Fields(....)
'           etwas eingetragen wird, verlangsamt sich die Schleife bis adoRs.EOF auf das hundertfache
'           L�sung:
'           im Multi-Nutzer-Betrieb ist Pr�fen1 verboten es kommt MsgBox
'           msg = msg & "Pr�fen1 muss ausgef�hrt werden, wenn Sie der einzige Nutzer der Datenbank sind" & vbNewLine
'           msg = msg & "Die Namen der anderen Nutzer finden Sie in der Datei " & AppPath & "\fotos.ldb"
'           im Multi-Nutzer-Betrieb ist Pr�fenS verboten
'17.02.2011:
'           F�r Multiuser-Umgebungen ist es notwendig, da� jeder user seine eigene pruef.log (englisch check.log) besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'17.02.2011:
'           F�r Multiuser-Umgebungen ist es notwendig, da� jeder user seine eigene fotos.ini besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'18.02.2011:
'           Korrektur zu 14.02.2011
'18.02.2011:
'           SpracheFestlegen Abschreiben aus fotos.exe
'           Nicht der Wert in Fotos.ini bestimmt die Sprache, sondern ob es eine Tabelle namens Fotos oder EFotos gibt.
'           daraufhin wird der Wert in Fotos.ini korrigiert
'16.03.2011 13.3.20 Fehlerkorrektur alle Versionen:
'           Im Formular Loglesen ist die Funktion Pr�fen3 - Die gefundenen Dateien sind �berfl�ssig -> &l�schen...
'           irgendwann vorlorengegeangen. Dieser Button wurde nicht mehr angeboten.
'           Ab sofort wird er wieder angeboten.
'13.06.2011 13.3.20 Fehlerkorrektur alle Versionen:
'           FileDateTime versteht keine Unicode-Dateinamen es kommt Laufzeitfehler '52' Dateiname oder -nummer falsch
'18.06.2011 13.3.20 Fehlerkorrektur alle Versionen:
'           GetOriginalDateTime versteht keine Unicode-Dateinamen es kommt Laufzeitfehler '52' Dateiname oder -nummer falsch
'           ab sofort wird der Fehler in Pruef.log eingetragen
'23.06.2011 13.3.20 Verbesserung alle Versionen:
'           Ich mache die Gr��e der Fonts f�r die Controls abh�ngig von der Einstellung unter 'Eigenschaften von Anzeige' ->
'           Erweitert -> DPI-Einstellungen. Das geschieht automatisch beim Form_Load jedes Formulars.
'           Ich unterscheide normal=96, gro�=120, sehr gro�>120
'           Das erfordert Bildschirmaufl�sung mindestens 1024 x 768 bei 96 DPI und
'           mindestens 1280 x 800 bei 120 DPI
'           Der Nutzer soll entscheiden, ob er die Fontgr��en-Anpassung haben will, wenn eine DPI-Einstellung h�her als 96
'           gefunden wird, der Wert wird in Fotos.ini gespeichert
'===============================================�bergang auf Win7=======================================================
'26.10.2011 13.3.21 Verbesserung alle Versionen:
'           Weil es im Windows 7 kein msdxm.ocx mehr gibt, entf�llt das Control Mediaplayer
'           Videobreite/H�he/Dauer werden mit mciSendString siehe MovieModule.cls ermittelt
'07.11.2011 13.3.21 Verbesserung alle Versionen:
'           Verbesserung f�r Multi-Nutzer-Umgebung. Vermeidung von overhead, der entsteht bei Benutzung einer fotos.exe vom fremden PC.
'           Jeder PC hat seine lokale fotos.exe und w�hlt aus, mit welcher fotos.mdb aus einem fremden Ordner oder fremden PC er arbeiten will.
'           Dazu muss der Nutzer beim Start der lokalen fotos.exe die Shift-Taste festhalten. Daraufhin geht ein CommonDialog (ohne ocx) auf zur
'           Auswahl der fotos.mdb
'           Der Ordnername der fotos.mdb steht in gstrFotosMdbLocation.
'           Wenn gstrFotosMdbLocation leer ist, wird AppPath benutzt. Wenn gstrFotosMdbLocation <> "" ist, werden die Tools FotosMdb und Renammdb
'           mit Aufrufparameter gstrFotosMdbLocation gestartet.
'           Commandline 'zB fotosmdblocation=H:\FOTOS\GG;
'14.11.2011 Generelles Entfernen von CommonDialog comdlg32.ocx stattdessen Benutzung von standarddialoge.bas
'           weil es Gemecker beim Installieren des MSI-Paketes gibt
'23.11.2011 13.3.22 �nderung alle Versionen:
'           Ich habe den Winkelmann-Fehler im Windows 7 gefunden. Bei Dr�cken der Taste F5 kommt ein leeres Grid.
'           und beim �ffnen der Query-Form kommt Fehler-Nr.: -2147467262
'           Ein nackiges Windows 7 ohne Microsoft Office bringt diesen Fehler. Die Installation einer beliebigen Office Komponente
'           ab Office 2003 (probiert mit Word) beseitigt den Fehler. Er tritt auch dann nicht mehr auf, wenn Office wieder deinstalliert
'           wird.
'           Ich muss in frmSprache zu Beginn ermitteln in welchem Betriebssystem ich arbeite.
'           Bei XP und Vista geht es weiter mit der Sprachauswahl.
'           Bei Windows7 und h�her, muss ich fragen ob Office 2003 oder h�her installiert ist, wenn ja geht es weiter mit der Sprachauswahl.
'           Wenn nein, kommt eine MsgBox mit dem Hinweis, da� erst Office 2003 oder h�her installiert werden muss. Dann endet das Programm.
'26.11.2011 13.3.22 �nderung:
'           Anstelle Laufzeitfehler'3050' soll eine vern�nftige Ausschrift kommen, abgeschrieben bei fotos.exe
'30.11.2011 13.3.22 vergessen bei 13.3.21 Fehlerkorrektur Professional-Version:
'           Im Win7 passiert es, dass die Professional Version sich nicht herstellen l��t. Sie behauptet, sie w�re Shareware-Version.
'           Das kommt von RegProfi.exe, dies bildet sich ein, es schreibt die Datei msprivs.log nach GetSystemDirectoryA (C:\Windows\system32)
'           schreibt aber in Wirklichkeit nach C:\users\vm\AppData\VirtualStore\Windows\System32
'           Das liegt daran, dass RegProfi.exe eigentlich mit Manifest arbeiten m��te, Aber dann kommt Installer-Fehler 1721.
'           Darum schreibe ich die Datei msprivs.log ab sofort in den Pfad von fotos.ini (gstrFotosIniAnwendungsOrdner)
'04.12.2011 13.3.22 Verbesserung alle Versionen:
'           Ursache f�r Laufzeitfehler '13' Typen unvertr�glich gefunden
'29.12.2011 13.4.0 Neue Funktion
'           Unterst�tzung des SQL-Servers
'           Fotosmdb und Renammdb machen zwar ein Connect siehe frmConnectSQL aber kein Login
'           Wenn Parameter mit CommandLine �bergeben werden, erfolgt kein Connect
'           DBsql As ADODB.Connection bleibt die gesamte Lebenszeit der session offen
'           rsDataGrid As ADODB.Recordset bleibt die gesamte Lebenszeit der session offen
'           die anderen ADODB.Recordsets werden mehrfach benutzt immer wieder geschlossen und neu ge�ffnet
'29.12.2011 13.4.0 Gestrichene Funktion Pr�fenD Pr�fen4 und Pr�fen5
'29.12.2011 13.4.0 Neue Funktion wenn Spalte Dateiname nicht der Prim�rschl�ssel ist, geht das Programm garnicht erst los
'29.12.2011 13.4.0 Ge�nderte SQL-Server-Datenbankstruktur Feld Jahr muss sein nvarchar(4) sonst geht charindex nicht
'29.12.2011 13.4.0 Nur bei der Access-Shareware-Version ist es n�tig, da� beim ersten Start von fotos.exe Language = "9" ist
'           nur dann wird msdmo.log erzeugt
'           mit Hilfe des Alters von msdmo.log nerve ich die Shareware-Nutzer mit Einblendung des Shareware-Hinweises
'           Das Datum 30.12.2011 ist das Datum der Fotos.mdb im Auslieferungszustand
'29.12.2011 13.4.0 MDF und LDF sind ab sofort keine erlaubten Dateitypen
'29.12.2011 Versto� gegen die 3-Einigkeit wird wenn gew�nscht in Pruef.log eingetragen
'08.02.2012 13.4.0 Beim L�schen von fotos.mdb kann der Nutzer ab jetzt ausw�hlen, ob er alle *.jpg Fotos l�schen will
'09.02.2012 13.4.0 Ab sofort Kann man " finden aber nicht ' - bisher war es andersrum
'14.02.2012 13.4.0 Aus Form1.txtFont.Fontname=Arial kommt der Fontname f�r alle Controls, Weil ms sans serif schei�e lesbar ist im Windows 7
'03.03.2012 13.4.1 Fehlerkorrektur bei sql server version bei NachPr�fen1L�schen
'04.03.2012 13.4.1 Fehlerkorrektur
'           Man kann mich austricksen und aus einer Shareware-Version eine SQL-Server-Version machen bei fotos.mdbnichtda
'           frmConnectSQL darf bei gblnProversion=False nicht erscheinen ab sofort gibt es bedingte Compilierung
'04.03.2012 13.4.2 Verbesserung:
'           nur bei #if Proversion gibt es ein Formular frmConnectSql, sonst wird es bei der Compilierung weggelassen
'05.03.2012 in den Eigenschaften der .exe soll erkennbar sein, ob Proversion=0 oder =-1
'           ich trage ein bei Projekteigenschaften -> Erstellen -> Copyright -> GERBING Software Chemnitz -1 oder 0
'29.03.2012 13.5.0 Verbesserung
'           ThumbnailAnzeigen erfolgt mit GDIPlus, GDI+ ist Bestandteil des Betriebssystems seit XP
'           2 neue native Dateitypen PNG TIF, aber CUR gestrichen
'20.04.2012 13.5.1 �nderung
'           keine Fehler melden wenn Fehler bei Feld SW (BWC)
'30.10.2012 13.5.4 Verbesserung
'           Beim R�ckschreiben IPTC Hinweis auf m�glicherweise fehlende Administratorrechte
'19.11.2012 13.5.5 Fehlerkorrektur SQL Server Version
'           Zum P�fen der ersten Kolonne des LicenseCode wird nicht mehr der Name, sondern die mittlere Kolonne benutzt
'21.11.2012 13.5.5 Fehlerkorrektur SQL Server Version
'           Die bisherige Verschl�sselung der Zahl der Lizenzen ist zu leicht zu knacken durch Probieren
'           Ich verschl�ssele jetzt die Zahl an zwei Positionen
'           bisher SQL99
'           jetzt  99S99 und in der Mitte bleibt ein S stehen
'18.12.2012 13.5.5 Fehlerkorrektur
'           Das R�ckschreiben der IPTC-Daten dauert entsetzlich lange. Es liegt an aa = Input(LOF(Filenumber), #Filenumber)
'           Ich suche nach Beschleunigung.
'           L�sung:
'           aa = Space$(LOF(Filenumber))                                                        'Gerbing 18.12.2012
'           'Gesamten Inhalt in einem "Rutsch" einlesen                                         'Gerbing 18.12.2012
'           Get #Filenumber, , aa
'28.12.2012 Fehlerkorrektur 13.5.5
'           Bisher konnte man bei Wiederholung der Funktion Pr�fen3 ->
'           Pr�fen3 - Gefundene Dateien in die Datenbank aufnehmen...
'           ein eventuell vorher als letztes gezeigtes Bild sehen und das war irgendeins. Ab jetzt wird
'           in Form LogLesen gemacht Call NachPr�fen3Aufnehmen.BildAnzeigen("")
'-------------------------------------------------------------------------------------------------------------------------------------
'04.03.2013 Fehlerkorrektur 14.0.0
'           Form1.DbGridNeu.AllowUpdate = False in der IDE eingestellt
'04.03.2013 Neue Funktion 14.0.0
'           �berraschung: Das DataGrid msdatgrd.ocx ist unicode f�hig, Ms Access vermutlich schon lange
'           Unicode-Unterst�tzung durch die Timosoft Controls und durch FSO
'               geht nicht im XP: exe st�rzt ab und auch IDE st�rzt ab beim Schlie�en des Programms, vermutlich weil
'               bei Diashow.Form_Unload set fso=Nothing und das Unload f�r alle Forms gefehlt hat
'               Geht im Win7
'               Zum �ndern von FontSize muss die Eigenschaft des Timosoft Controls UseSystemFont = False sein
'               Viele Events bei den Timosoft Controls sind standardm��ig disabled. Man muss im gezeichneten Control element rechtsklicken ->
'               Properties -> H�kchen rausnehmen
'           ListBoxForm.ExLVwu abgeschrieben aus
'           D:\VISUALBA.SIC\VB6BeispielCode\Unicode\Unicode Timosoft\ExplorerListView\samples\VB6\General\OptionListView\
'               Wenn in ListBoxForm das Subclassing abgeschaltet wird, gibt es keine Checkboxen zu sehen
'               Andererseits macht das Subclassing das Programm saulangsam
'               Drag&Drop als Target geht erst wenn man RegisterForOLEDragDrop = True setzt
'               Das Debuggen spinnt in DiashowForm bei eingeschaltetem Subclassing wenn man _OLEDragDrop debuggen will ->
'               probiere ob es mit Alt+F4 weitergeht
'           Die Form.Caption mit unicode sieht man nicht in der IDE sondern erst wenn man die exe startet
'           Alle Datei read/write Operationen f�r Text-Dateien sollte man mit FSO machen. Da muss man vorher testen ob der Dateiname
'           auch nur ein unicode Zeichen enth�lt oder alles ANSI ist.
'           Man muss daraufhin die Datei mit FSO entweder als unicode oder ANSI Datei �ffnen.
'           Alles auswechseln was wie 'Open Path For Binary Access Read As #Handle' aussieht.
'           Achtung bei FSO Gefunden bei Microsoft http://support.microsoft.com/kb/189751/en-us
'               Reads only ASCII data - while the FileSystemObject can create an ASCII or Unicode text file, the FileSystemObject can only
'               read ASCII text files.
'           Die scrrun.dll muss mit ausgeliefert werden Sie ist zust�ndig f�r FSO
'           Chronologie:
'           init_global bei Start des Programms
'           pruef.log als unicode file erzeugen
'           RichTextBox.ocx kann entfallen wird ersetzt durch unicode f�higes Timosoft Text Control
'           NeueDatens�tzeGenerieren alle Listboxen/Comboboxen austauschen wo unicode vorkommen kann
'           Alle FileDateTime unicode f�hig machen durch FSO
'           Alle 'Name xyz As' ersetzen durch NameAS
'           Alle Dir( ersetzen durch file_path_exist
'           Alle MkDir ersetzen durch CreateDirectoryW                                                      'Gerbing 16.11.2015
'           iptcinfo.dll entf�llt das mach ich jetzt selber
'           Drag&Drop realisieren
'           INI file wird unicode f�hig durch schreiben mit FSO und Benutzen von GetPrivateProfileStringW und
'           WritePrivateProfileStringW
'           LoadPicture ersetzen durch LoadPictureW
'           F�r Kill gibt es ein VBA replacement for "Kill(PathName)" with UNICODE support in UnivbzGlobal.bas
'           F�r SetAttr gibt es ein VBA replacement for SetAttr, supports unicode and network in UnivbzGlobal.bas
'           Alle MsgBox wo file names vorkommen ersetzen durch MessageBoxW
'           GERBING Fotoalbum 13 ersetzen durch GERBING Fotoalbum 14
'           App.Path ersetzen durch getCurrentDir
'           chm-files lassen sich in unicode Pfad nicht �ffnen, das hat Microsoft nicht vorgesehen
'           ShellExecute ersetzen durch ShellExecuteW (RunShellExecute)

'10.05.2013 14.0.0 Verbesserung
'           txtFehlerU.Text wird kursiv unterstrichen blau dargestellt
'           bei 'kein Fehler gefunden' soll kein Fenster aufgehen
'04.06.2013 14.0.0 Verbesserung
'           �berarbeitung der EXIF-Felder, es werden jetzt GPS-Felder erkannt
'08.06.2013 14.0.0
'           Beim normalen Start braucht das Programm keine Administrator-Rechte
'           Das Programm verlangt Administratorrechtezur Bek�mpfung von Laufzeitfehler 'Laufzeitfehler '339':
'           Die Komponente CBLCtlsU.ocx oder eine ihrer Abh�ngigkeiten ist nicht richtig registriert.....
'           Jetzt kann ich aber nicht mehr mit der c:\users\administrator\AppData\Roaming\Gerbing Fotoalbum 14\fotos.ini
'           arbeiten, weil jeder Nutzer der ja jetzt als Administrator starten muss, dieselbe fotos.ini zugeteilt bekommt
'           Ab sofort stehen fotos.ini  und pruef.log im AppPath. Also dort wohin der Nutzer sein GERBING Fotoalbum 14
'           installiert haben wollte. Das ist standardm��ig c:\users\gottfried\Documents\GERBING Fotoalbum 14
'           Von Regprofi.exe muss gerbingsoft.log in c:\Users\Public\Documents\GERBING Fotoalbum 14\gerbingsoft.log gestellt werden
'           Bei der Vollversion steht gerbingsoft.log in c:\Windows\SysWOW64\gerbingsoft.log
'08.06.2013 14.0.0
'           Endlich habe ich es geschafft, da� alle Programme wieder ohne Administrator-Rechte starten d�rfen.
'           Das Packen der MSI-Pakete mit COM-Objekten hat zwar die Timosoft-ocx-Dateien installiert, aber Starten ging nur als Administrator
'           Das Packen der MSI-Paket mit den Timosoft-ocx-Dateien als Selfreg=Yes hat den von Anfang an beabsichtigten Effekt gehabt.
'-------------------------------------------------------------------------------------------------------------------------------------
'24.07.2013 Fehlerkorrektur 14.0.1
'           Der Import von IPTC-Feldern hat nicht funktioniert
'05.09.2013 Fehlerkorrektur 14.0.1
'           Kill ersetzen durch file_delete
'           Name ... as ersetzen durch NameAs
'           war vergessen worden bei: NachPr�fen3L�schen, nach Pr�fenS l�schen �berfl�ssige Audiodatei, bei DBEngine.CompactDatabase
'06.09.2013 Fehlerkorrektur 14.0.1
'           Es tritt ein Fehler auf bei Ausf�hrung in unicode-Pfad und englischer Datenbank
'           bei Pr�fen 1 -> Laufzeitfehler '-2147467259 (80004005)' Feld 'Ort' wurde nicht gefunden
'           L�sung: Wenn 'nicht kontrollieren' gew�hlt wird, dann wird kein rstsql.Update gemacht
'10.09.2013 Fehlerkorrektur 14.0.1
'           Im unicode Pfad hunzt die Function SetOriginalDateTime bei IPTC... Datum der Datei erhalten
'           es wird immer das aktuelle Datum eingetragen (Zeitpunkt des Programmablaufes)
'           L�sung: anstelle CreateFile verwenden CreateFileW
'10.09.2013 Fehlerkorrektur 14.0.1
'           bei Festhalten der Shift-Taste kann der Standort der fotos.mdb ausgew�hlt werden
'02.10.2013 Verbesserung 14.0.1
'           Die von Microsoft benutzten Felder in den Eigenschaften -> Details einer JPG-Datei sollen importierbar(ausw�hlbar) sein
'           Beschreibung.Titel als EXIF-XPTitle
'           Beschreibung.Thema als EXIF-XPSubject
'           Beschreibung.Markierungen als EXIF-XPKeywords
'           Beschreibung.Kommentare als EXIF-XPComment
'           Ursprung.Autoren als EXIF-XPAuthor
'-------------------------------------------------------------------------------------------------------------------------------------
'26.10.2013 Verbesserung 14.0.2
'           Die unicode Controls lassen keinen schnellen Neuaufbau einer leeren Datenbank fotos.mdb zu
'           in Form LogLesen dauert TxtU.text = TxtU.text & xyz saulange aber wird garnicht gebraucht
'           ich muss dort wo eine Listbox nur zum Datensammeln aber nicht zum Anzeigen benutzt wird, diese ersetzen durch eine Collection
'           ich muss in NeueDatens�tzeGenerieren nur einmal pro Foto die IPTC-Felder einlesen (bisher 6x)
'           beim L�schen nach Pr�fen3 bei neugefundenen Dateien werden diese ab sofort in den Papierkorb gel�scht
'03.11.2013 Nachbesserung zum 26.10.2013
'           Wenn der Nutzer auf 'Reset' klickt soll auch pruef.log geleert werden
'26.10.2013 Fehlerkorrektur 14.0.2
'           Zustand: Wenn es ein Datenbankfeld gibt, das per Hand ausgef�llt wird und f�r dieses Datenbankfeld ist gleichzeitig ein
'                   Quell-Feld ExifIptc sichtbar, dann wird der Wert aus Quell-Feld ExifIptc genommen.
'                   Mit diesem Zustand kann ich keine manuellen Werte eingeben und gleichzeitig nutzerdefinierte Felder aus einem
'                   Quell-Feld ExifIptc �bernehmen (zB EXIFDateTimeOriginal)
'           L�sung: Wenn gleichzeitig eine manuelle Eingabe vorliegt und ein Quell-Feld ExifIptc sichtbar ist, dann hat manuelle Eingabe
'                   Vorrang
'26.10.2013 Fehlerkorrektur 14.0.2
'           ich hatte vergessen dass jetzt anstelle von msprivs.log gerbingsoft.log abgefragt werden muss
'02.11.2013 �nderung 14.0.2
'           Wenn eine fotos.mdb aus EXIF/IPTC-Feldern neu erstellt wird, bleibt das Feld IPTCPresent = False. Ich habe nachgedacht,
'           ob das ein Programmfehler ist.
'           Nein - es ist kein Programmfehler. IPTCPresent sagt aus, ob schon mal ein �bertragen der IPTC-Felder in die JPG-Fotos
'           mit der Funktion IPTC... stattgefunden hat. In diesem Zusammenhang ist die Funktion Pr�fenIPTC �berfl�ssig und falsch,
'           weil nicht gesagt ist, da� bei gefundenen IPTC-Feldern diese durch die Funktion IPTC... erzeugt worden sind.
'           L�sung: btnPr�fenIPTC und btnPr�fenIPTCAbbrechen werden entfernt
'           IPTCPresent m��te eigentlich hei�en IPTCExported
'06.11.2013 Fehlerkorrektur 14.0.2
'           Bisher habe ich vergessen bei allen Pr�fvorg�ngen die Buttons btnGenerieren, btnNutzerdefinierteFelderAnlegen,
'           btn�ffnePruefLog, btnL�scheInhaltFotosMdb zu disablen
'           Wenn der Nutzer w�hrend eines Pr�fvorganges damit herumspielt kann das Programm abst�rzen
'18.11.2013 private Funktion 14.0.2
'           Ich will Fotosmdb.exe f�r meine privaten W�nsche nutzen, ohne da� die Standard-Nutzer etwas davon merken
'           Ich will das Feld EXIFDateTimeOriginal das es nur in meiner fotos.mdb gibt bei Klick auf Pr�fen3 aktualisieren k�nnen.
'23.11.2013 Nachbesserung zum 26.10.2013
'           gerbingsoft.log muss auch in gblstrSystemDirectory gepr�ft werden
'24.12.2013 Pr�fen1 - Alle berechnen l��t sich m�glicherweise beschleunigen, wenn ich Update nur bei Ungleichheit mache
'18.01.2014 Verbesserung 14.0.2
'           "DOC" und "DOCX" werden zugelassen
'-------------------------------------------------------------------------------------------
'15.02.2014 14.0.3 Verbesserung
'           Die Msgbox zu Programmstart bei falscher fotos.mdb ist �berarbeitet worden
'           'msg = Dateiname & " existiert nicht." & vbNewLine
'           'msg = "Datenbank und Fotos passen nicht zueinander" & vbNewLine
'           'msg = msg & "Vermutlich benutzen Sie eine falsche Datenbank-Datei" & vbNewLine
'           'msg = msg & "Benutzen Sie das Tool Fotosmdb um die Datenbank zu �berpr�fen" & vbNewLine & vbNewLine
'
'           'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
'09.03.2014 14.0.3 Verbesserung alle Versionen
'           "mp4" videos sind ab sofort erlaubt.
'           WMP.dll kann mp4-files abspielen. Genauso gut kann man mp4-files in avi-files umnennen und dann abspielen
'-------------------------------------------------------------------------------------------------------------------------------------
'30.05.2014 14.0.4 Verbesserung
'           Bisher ist Nach Pr�fen3 der Standardbutton 'Pr�fen3 - Die gefundenen Dateien sind �berfl�ssig -> &l�schen...'
'           ich will aber dass der Standardbutton 'Pr�fen3 - Gefundene Dateien in die Datenbank &aufnehmen...' ist
'           Das geht nicht mit Button.Default = True sondern mit Button.Tabindex = 0
'-------------------------------------------------------------------------------------------
'24.06.2014 14.0.5 Fehlerkorrektur
'           Fehler: Bei der Funktion Rekursive ist ein Dateiname von >130 Bytes L�nge bisher ignoriert worden.
'           Das ist aufgetreten seit Version 14.0.0
'           L�sung: Die function FindFirstFileW und FindNextFileW in Module1 sind falsch deklariert, jedoch in UnivbzGlobal richtig
'           Falsch ist      Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'           richtig ist     Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFFData As Long) As Long
'           falsch ist      hSearch = FindFirstFileW(StrPtr(Path & "*"), wfd)
'           richtig ist     hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(wfd))
'           falsch ist      DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
'           richtig ist     DirName = RemoveNulls((wfd.cFileName))
'----------------------------------------------------------------------------------------------------------
'02.07.2014 14.0.5 Fehlerkorrektur
'           Zustand: Es gibt Fotos, bei denen werden die EXIF-Informationen unvollst�ndig angezeigt
'                   bei YCbCrCoefficients ist Schlu�
'           L�sung: in clsEFIF.cls wird abgefragt ob IFD(i).Length = 0 / Bei IFD(i).Length = 0 gab es einen unbehandelten Fehler
'                   und die Prozedur wurde vorzeitig beendet.
'----------------------------------------------------------------------------------------------------------
'09.09.2014 14.0.7 Fehlerkorrektur
'           Zustand: Ein Dateiname mit erlaubter Dateinamen-Erweiterung aber fehlendem Punkt wird gegenw�rtig ins Fotoalbum aufgenommen.
'                   Bei der Anzeige mit fotos.exe wird nach einem Programm gefragt, mit dem diese Datei angezeigt werden soll.
'           L�sung: Beim Generieren neuer Datens�tze werden solche wie 'xyzJPG' nicht aufgenommen, sondern als Fehler angemeckert.
'14.10.2014 14.0.7 Fehlerkorrektur alle Versionen
'           Zustand: Es fehlt die Sanduhr bei 'IPTC...'
'           L�sung: Sanduhr einschalten und wieder ausschalten
'           Zustand: Angebliche Fehler nach 'IPTC...' r�hren von Fehlern aus einem anderen vorher gelaufenen Pr�fvorgang
'           L�sung:  In IPTCGenerieren.Form_Load fehlt 'Form1.FehlerGefunden = False'
'----------------------------------------------------------------------------------------------------------
'18.03.2015 14.0.8 Fehlerkorrektur
'           Zustand: Wenn eine andere Software ein IPTC-Feld bearbeitet hat, kann es passieren dass GERBING Software
'                   keine IPTC-Felder anzeigt (gar keine = leer)
'           L�sung: modIPTC wird korrigiert und fragt nach dem ersten IPTC-Header wenn nicht "Photoshop 3.0" gefunden wird
'                   ob es weiter hinten noch einen IPTC-Header mit "Photoshop 3.0" gibt
'----------------------------------------------------------------------------------------------------------
'19.04.2015 14.1.0 Fehlerkorrektur
'           gewaltige Beschleunigung bei Pr�fen3
'           Millisekunden rstTempHaken.Update f�r eine Datei bei Access mdb 1 bis 2
'           Millisekunden rstTempHaken.Update f�r eine Datei bei SQL-Server 52 mit CursorLocation = adUseClient
'           gewaltige Beschleunigung durch
'           Millisekunden rstTempHaken.Update f�r eine Datei bei SQL-Server 1 bis 2 mit CursorLocation = adUseServer
'           ich hatte irrt�mlich vermutet - es sei eine absichtlich gewollte Verz�gerung beim SQL-Server-Express (der ist kostenlos)
'           aber - bei adUseServer kommt kein Recordcount
'----------------------------------------------------------------------------------------------------------
'22.05.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 18.03.2015
'           Zustand: Die Fotos von Ralph haben ewig lange gebraucht um mit Fotosmdb/Pr�fen3 in die Datenbank aufgenommen zu werden
'                   oder mit Diashow oder Fotoalbum oder WallpaperChanger angezeigt zu werden.
'                   Bei pos = InStr(1, strImageString, IPTCHeader, vbTextCompare) braucht die Programmausf�hrung ewig lange
'                   Scheinbar wird ab einer bestimmten L�nge eines Strings die Function InStr arschlangsam.
'           L�sung: ich muss schreiben pos = InStrB(1, strImageString, IPTCHeader, vbTextCompare)
'                   InStrB geht blitzschnell
'                   aber beim anschlie�enden Vergleich muss die pos korrigiert werden
'           gefunden im Internet:
'                   http://www.aivosto.com/vbtips/stringopt.html und
'                   http://www.vbforums.com/showthread.php?607151-RESOLVED-Experiment-with-InStrB-LenB-LeftB-MidB-RightB
'                   man muss aufpassen da� man geradzahlige Ergebnisse nicht f�r wahr h�lt. Richtig sind nur ungerade Ergebnisse
'                   The code just tends to be quite a bit longer:
'                   lngA = 0
'                   lngA = InStr(strText, strFind)
'                   umwandeln in
'                   Do: lngA = InStrB(lngA + 1, strText, strFind, vbBinaryCompare): Loop Until (lngA And 1) Or (lngA = 0)
'----------------------------------------------------------------------------------------------------------
'01.06.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 22.05.2015
'           Der Fehler wirkte sich so aus, da� bei jedem Neu-Schreiben der IPTC-Felder vorneran ein neuer IPTC-Abschnitt erzeugt wurde.
'           Vom Programm vorgesehen ist aber ein Vermischen bisheriger und neuer IPTC-Felder.
'----------------------------------------------------------------------------------------------------------
'xxxxxxxxxxxxxxxxxxxxxxxxxxx ausgeliefert am 25.06.2015 xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'----------------------------------------------------------------------------------------------------------

'29.09.2015 14.1.2 Fehlerkorrektur Folgeerscheinung vom 01.06.2015
'           Der Fehler wirkte sich so aus, dass bei manchen Fotos durch das R�ckschreiben IPTC... 2 Bytes fehlten. das war der Exif-header
'           'FFE1'
'           dadurch konnte man die Fotos mit keiner Software mehr lesen
'----------------------------------------------------------------------------------------------------------
'30.09.2015 14.1.2 Fehlerkorrektur
'           Zustand: Seit der �nderung von 29.09.2015 meckert WallpaperChanger.exe viele Fotos an, die ich mit IPTC... alle schreiben
'                   im Ordner GG neu geschrieben habe. Ebenso in FotosMdb bei JahrFestlegen und bei NachPr�fen3Aufnehmen, wenn ein
'                   Vorschaubild gezeigt werden soll.
'           L�sung: ich benutze GDI+ CreateThumbnailFromFile anstelle von LoadPictureW
'----------------------------------------------------------------------------------------------------------
'02.10.2015 14.1.2 Verbesserung Folgeerscheinung vom 30.09.2015
'           Zustand: Das Vorschaubild hat miese Qualit�t. Es ist wirklich nur ein Vorschaubild.
'           L�sung: ich benutze GDI+ CreateStdPictureFromFile anstelle von CreateThumbnailFromFile
'----------------------------------------------------------------------------------------------------------
'07.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: im Windows 10 und Windows 8.1 und vom Standard abweichender DPI-Einstellung zeigt mein Programm verschwommene Schrift
'                   Das kann der Nutzer korrigieren, indem er die exe markiert -> Eigenschaften -> Kompatibilit�t ->
'                   DPI-Skalierung nicht anwenden
'           L�sung: Ein Programm erkl�rt sich selbst als DPI-kompatibel. Das geht durch sein Manifest
'----------------------------------------------------------------------------------------------------------
'13.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Ich habe nicht ber�cksichtigt, da� es auch EXIF-Felder mit unicode Inhalt geben kann die werden als ?????
'                   dargestellt
'           L�sung: In Form NeueDatens�tzeGenerieren habe ich txtEXIFInfo durch ein unicode f�higes control ersetzt
'----------------------------------------------------------------------------------------------------------
'14.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Wenn ich den Windows Explorer benutze rechtsklicken -> Eigenschaften -> Details und im Abschnitt Beschreibung
'                   einf�ge Titel, dann wird meine Eingabe sowohl in Titel als auch in Thema eingetragen.
'                   Bei mir intern ist Titel = EXIF-XPTitle und Thema = EXIF-ImageDescription
'                   Wenn ich einen unicode String in Titel eintrage gibt mein Programm unter EXIF-ImageDescription Mist aus
'                   aber unter EXIF-XPTitle ist es richtig
'                   Andere Programme wie ExifToolGUI oder XnViewMP machen es richtig
'           L�sung: Wenn ich einen ascii string finde (IFD.Format=2) dann kommt FromUTF8String dran
'----------------------------------------------------------------------------------------------------------
'xxxxxxxxxxxxx Version 14.1.2 gibt es nicht als ausgelieferte Version xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'----------------------------------------------------------------------------------------------------------
'16.11.2015 14.2.0 Verbesserung alle Versionen
'           Ich will in der Form IPTCGenerieren weitere Felder anbieten, wohin die Datenbank-Felder exportiert werden k�nnen
'           schlie�lich werden diese auch zum Import angeboten.
'           Zum Schreiben dieser Felder ins JPG-Foto hinein benutze ich die Command line version von Exiftool.exe
'           ich muss beachten, da� auch unicode geschrieben werden soll
'           Das sind die Felder, die im Windows Explorer angesteuert werden �ber Rechtsklick auf einen JPG-Dateiname -> Eigenschaften ->
'           Details -> 4 Felder im Abschnitt Beschreibung, ein Feld im Abschnitt Ursprung
'           Beschreibung    Titel           -> EXIF-XPTitle         eignet sich f�r das Feld Situation
'                           Thema           -> EXIF-XPSubject       eignet sich f�r das Feld Personen
'                           Markierungen    -> EXIF-XPKeywords      eignet sich f�r das Feld Ort
'                           Kommentare      -> EXIF-XPComment       eignet sich f�r das Feld Kommentare
'           Ursprung        Autoren         -> EXIF-XPAuthor        eignet sich f�r Land
'
'           Ich stelle die Funktion IPTC... generell um auf Benutzung von Exiftool.exe. Alle zu schreibenden tags werden ab jetzt mit
'           exiftool.exe geschrieben.
'           Meine 28.000 Fotos haben 50 Minuten gebraucht.
'
'           Phil Harvey (exiftool-Autor) schreibt die EXIF-XP... Felder generell als unicode, aber die IPTC-Felder nur, wenn der Parameter
'           -charset IPTC=UTF8 benutzt wird.
'
'           Ich kann jetzt eine Teilmenge schreiben, anstelle von 'F�r die ganze Datenbank' geht
'           'F�r ein einzelnes Jahr' oder
'           'F�r IPTCPresent = False
'           Ich kann auch eine �nderungs-Abfrage mit Parameter benutzen zB
'           PARAMETERS [Jahr] Text ( 255 ); UPDATE Fotos SET IPTCPresent = False Where Jahr Like "*" & [Jahr] & "*"
'
'           Error: Error reading StripOffsets data in IFD0 - E:/Testen ExifTool/Gg/1940/Stadtplan Alle Planquadrate.jpg Gegenma�nahme ->
'           http://u88.n24.queensu.ca/exiftool/forum/index.php?topic=1369.0
'
'           Fehler 'Can't locate PAR.pm in @INC (@INC contains: .) at -e line 860.' kommt wenn AppPath zB wie  ...\Fotosmdb?????
'           unicode Zeichen enth�lt
'
'           L�stige Warning:Multiple Photoshop records entfernt man in einem ganzen Ordner mit
'           exiftool -preserve -overwrite_original -photoshop:all= *.jpg
'           Das entfernt alle Photoshop und alle IPTC tags. Dann muss ich sie neu schreiben.
'
'           Error: JPEG EOI marker not found -> das Foto l��t sich einwandfrei begucken
'
'           Ohne nochmaliges Exportieren mit Version 14.2.0 bekomme ich beim Angucken der EXIF/IPTC-Felder bei allen Umlauten chinesische
'           Krakel zu sehen. Diese ersetzen jeweils 2 oder 3 Zeichen. L�sung ganze Datenbank nochmals exportieren.
'
'16.11.2015 14.2.0 Verbesserung alle Versionen
'           Zustand: Wenn ein IPTC-Feld UTF-8 Code enth�lt (erzeugt durch exiftool mit einem unicode Feld), dann wird Mist angezeigt
'                   Ich sehe zweistellige Zeichen im UTF-8 Code.
'           L�sung: modIPTC.bas function VorhandeneEinzelsegmenteSuchen wird ver�ndert.
'                   Eventuell vorhandene UTF-8 Zeichen werden konvertiert. Es kommt FromUTF8String dran.
'
'           Schreiben in Videos mit exiftool geht zum gegenw�rtigen Zeitpunkt 16.11.2015 nicht
'----------------------------------------------------------------------------------------------------------
'21.12.2015 14.2.1 Verbesserung meine private SQL-Server Datenbank
'           Zustand: bei Pr�fen3 wird nicht nach ExifDateTimeOriginal gefragt
'           L�sung: zwei verschieden Varianten ausf�hren, je nachdem ob mit oder ohne SQL-Server
'----------------------------------------------------------------------------------------------------------
'30.12.2015 14.2.1 Zur�cknehmen der �nderung 'Aktualisieren von ExifDateTimeOriginal'
'           Das wird ab sofort in fotos.exe erledigt
'----------------------------------------------------------------------------------------------------------
'21.03.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: Wenn im Drag&Drop-Container ein anderes Bild angeklickt wird, bleibt der Inhalt vom vorhergehenden Bild in LstU
'                   bzw txtExifInfo erhalten
'           L�sung: Beim Click auf ein Bild im Drag&Drop-Container wird stets der zugeh�rige EXIF bzw IPTC Inhalt bereitgestellt.
'----------------------------------------------------------------------------------------------------------
'02.08.2016 14.2.1 Verbesserung Professional Version
'           Zustand: Wenn ich einmal festgelegt habe, welche nutzerdefinierten Felder aus welchem Quellfeld gef�llt werden sollen,
'                   will ich diese Zuordnung bei jedem nachfolgenden 'Pr�fen3/Neue Datens�tze generieren' standardm��ig anbieten.
'           L�sung: Ich muss die getroffene Zuordnung in der Datenbank speichern in Tabelle UserDefined. Dort gibt es die Felder
'                   FieldName1      Text
'                   SourceField1    Text
'                   FieldName2      Text
'                   SourceField2    Text
'                   FieldName3      Text
'                   SourceField3    Text
'                   FieldName4      Text
'                   SourceField4    Text
'                   FieldName5      Text
'                   SourceField5    Text
'                   Wenn es diese Tabelle gibt und sie ist leer, werden die Zuordnungen abgespeichert(NeueDatens�tzeGenerieren.btnStart_Click).
'                   Wenn es diese Tabelle gibt und sie ist nicht leer, werden die fr�her getroffenen Zuordnungen eingetragen und es wird
'                   chkUnbeaufsichtigt = 1 gesetzt und manuelles �berschreiben wird verhindert(NeueDatens�tzeGenerieren.Form_Load).
'           Will der Nutzer die Zuordnung r�ckg�ngig machen, muss er Rechtsklicken auf den disabled Rahmen
'           Ab Version 14.2.1 muss die Tabelle UserDefined in leerem Zustand ausgeliefert werden, auch f�r SQL-Server
'           f�r alle Nutzer, die von �lteren Versionen upgraden, muss fotosmdb.exe diese Tabelle erzeugen(Access-Version), beim Upgrade kann
'           die SQL-Server-Version ignoriert werden, denn es gibt bisher noch keine ausgelieferte Datenbank
'----------------------------------------------------------------------------------------------------------
'03.08.2016 14.2.1 Verbesserung alle Versionen
'05.08.2016 zugeh�rige Fehlerkorrektur
'           Nach dem Vorbild vom 02.08.2016 k�nnte ich speichern, welche Standard-Datenbank-Felder aus welchen
'           EXIF/IPTC-Feldern aufgef�llt werden sollen. Das ist n�tzlich, wenn ein user andere Felder zum Import benutzen will, als
'           meine Standard-Felder. Dann braucht er sie nicht bei jedem 'Pr�fen3/Neue Datens�tze generieren' neu auszuw�hlen.
'           L�sung: Ich muss die getroffene Zuordnung in der Datenbank speichern in Tabelle DefaultFields. Dort gibt es die Felder
'                   SituationSource Text
'                   LocationSource  Text
'                   CountrySource   Text
'                   PeopleSource    Text
'                   BWCSource       Text
'                   CommentSource   Text
'                   Wenn es diese Tabelle gibt und sie ist leer, werden die Zuordnungen abgespeichert(NeueDatens�tzeGenerieren.btnStart_Click).
'                   Wenn es diese Tabelle gibt und sie ist nicht leer, wird bei Form_Load der Schalter blnDefaultFieldsNotEmpty
'                   eingeschaltet. Wenn ein H�kchen in chkUnbeaufsichtigt gesetzt wird, werden die Zuordnungen aus der Tabelle
'                   DefaultFields benutzt.
'                   Beim Export in die IPTC-Felder von JPG-Fotos �berschreibe ich, wenn die Tabelle DefaultFields nicht leer ist,
'                   die vorher getroffenen Standard-Zuordnungen.(IPTCGenereieren.Form_Load)
'           Will der Nutzer die Zuordnung r�ckg�ngig machen, muss er Rechtsklicken auf NeueDatens�tzeGenerieren.FrameStandardWerte
'27.10.2016 genauso Rechtsklicken auf IPTCGenerieren.FrameStandardWerte
'           Ab Version 14.2.1 muss die Tabelle DefaultFields in leerem Zustand ausgeliefert werden, auch f�r SQL-Server
'           f�r alle Nutzer, die von �lteren Versionen upgraden, muss fotosmdb.exe diese Tabelle erzeugen(Access-Version), beim Upgrade kann
'           die SQL-Server-Version ignoriert werden, denn es gibt bisher noch keine ausgelieferte Datenbank
'----------------------------------------------------------------------------------------------------------
'02.09.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Die Geo-Position in der Grad-Minuten-Sekunden-Darstellung wird manchmal richtig angezeigt und manchmal falsch.
'           L�sung: im Modul clsExif: ich vermeide das Umrechnen in Grad-Minuten-Sekunden sondern benutze gleich eine Dezimalzahl,
'                   so wie sie in frmGEOPosition gebraucht wird
'----------------------------------------------------------------------------------------------------------
'05.09.2016 15.0.0 Verbesserung alle Versionen
'           In fotos.exe kann der Professional Nutzer jetzt auf einer Landkarte markieren welche Fotos mit GEO-Daten er finden will
'           Seit Version 15.0.0 gibt es deshalb in der Tabelle ErsterStart das Feld LetzterGEOPunkt und ZoomListIndex
'           In jeder Datenbank fotos.mdb (Professional Version), wo es diese Felder noch nicht gibt,
'           werden sie erzeugt ohne dass der user es merkt
'----------------------------------------------------------------------------------------------------------
'12.10.2016 15.0.0 Fehlerkorrektur alle Versionen zu 05.09.2016
'           Zustand: Ich habe bisher nicht ber�cksichtigt, dass es negative GPSLatitude und GPSLongitude geben kann
'                   S�dhalbkugel und westliche Hemisph�re
'                   Damit die Vergleiche im SQL-String richtig ablaufen, muss der Datentyp von GPSLatitude und GPSLongitude Double sein
'           L�sung: In Form NeueDatens�tzeGenerieren. Wenn 'GPSLatitudeRef: S' gefunden wird, dann Minus vor GPSLatitude
'                   Wenn 'GPSLongitudeRef: W' gefunden wird, dann Minus vor GPSLongitude
'                   Mit MSAccess Konvertieren von String in Double mit CDbl(...)
'21.10.2016 nochmal korrigiert wenn GPSLatitude = "" dann kam Datentypfehler
'----------------------------------------------------------------------------------------------------------
'27.10.2016 15.0.0 Fehlerkorrektur alle Versionen zu 03.08.2016 und 05.09.2016
'           Zustand: Ich habe bisher nicht unterschieden ob es die Tabelle DefaultFields gibt und sie ist leer oder es gibt sie
'                   und es steht schon 1 Satz drin
'           L�sung: Abfrage verbessern
'----------------------------------------------------------------------------------------------------------
'10.11.2016 15.0.0 Verbesserung alle Versionen
'           Ich speichere Thumbnails im Ordner ...\GerbingThumbs\...
'           Beim L�schen von Datens�tzen, deren Foto nicht gefunden wurde, l�sche ich auch zugeh�rige Thumbnails
'08.12.2016 Die Dateien im Ordner ...\GerbingThumbs\... hei�en zB video1.avi.jpg oder foto33.jpg.jpg
'----------------------------------------------------------------------------------------------------------
'22.11.2016 15.0.0 Verbesserung alle Versionen
'           Ich habe Versuche gemacht die Kommunikation mit exiftool.exe zu verbessern.
'           Ein Beispiel mit CreatePipe f�r Stdin und Stdout hat zwar funktioniert, aber es versteht keine Unicode-Dateinamen.
'           Ich probiere ein neues Verfahren, wo nur CreatePipe f�r Stdout gemacht wird und ich arbeite weiterhin mit argfiles.txt
'           aber ich schreibe max 100 Dateinamen in argfile.txt, dann beginne ich erneut bei 'Starte exiftool.bat'
'           Den Arbeitsfortschritt kann ich beobachten in IPTCGenerieren.txtExifToolOutput
'----------------------------------------------------------------------------------------------------------
'12.12.2016 15.0.0 Fehlerkorrektur alle Versionen
'           Zustand: Manchmal bleibt das Datenbank-Feld 'SWF' leer, wenn ich in chkExif ein H�kchen setze
'           L�sung: wenn f�r das Datenbank-Feld 'SWF' zwar eine IPTC-Quelle angegeben ist, aber dort nichts steht,
'                   nehme ich eine von den 4 Standardangaben
'----------------------------------------------------------------------------------------------------------
'04.01.2017 15.0.0 Fehlerkorrektur alle Versionen
'           Bei 'Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)...' kommt bisher keine Sanduhr
'==========================================================================================================
'26.02.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Nachbesserung zum 12.12.2016
'----------------------------------------------------------------------------------------------------------
'11.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn 4K-Monitore benutzt werden, muss es m�glich sein, die Schriftgr�esse besser als bisher anzupassen
'           L�sung1: Es gibt die Schriftgr��en
'                   klein=1
'                   mittel=2
'                   gross=3
'           Die Einstellung wird gespeichert in der ini-Datei   [Adjustments]
'                                                               CheckForDPI 1 oder 2 oder 3
'           L�sung2: oder es gen�gt die Bildschirmaufl�sung auf zB 200 DPI einzustellen (Windows 10 kann noch weit h�her als 200 DPI)
'----------------------------------------------------------------------------------------------------------
'17.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn ich das Fotoalbum nach C:\ installieren lasse, muss ich als Administrator arbeiten
'                   Wenn ich jetzt neue Fotos mit Pr�fen3 einf�ge, bleiben im Formular NeueDatens�tzeGenerieren
'                   die Felder cmbSituation cmbOrt cmbLand braun (=disabled)
'                   Der Nutzer kann also nichts eingeben
'           L�sung: zB cmbSituationEx.ListIndex = 11 + 49 + 52 + 5 l�st ein Ereignis 'cmbSituationEx_Click' aus
'                   dort wird zB cmbSituation.Enabled = False gesetzt
'                   Ich muss anschlie�end wieder cmbSituation.Enabled = True setzen
'----------------------------------------------------------------------------------------------------------
'26.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn PSP X8 und h�her eine �nderung im Bild macht zB Bildbegradigung(horizontal - vertikal),
'                   l�scht es die von ExifTool gemachten Eintr�ge in IPTC-Feldern aber erzeugt stattdessen den Abschnitt IPTC2
'                   und PSP X8 l�scht aus dem Abschnitt IFD0 alle Felder XPTitle XPKeywords XPAuthor XPSubjects XPComment
'                   und schreibt sie stattdessen in den Abschnitt XMP-photoshop
'                   Die Gerbingsoft-Programme finden dann keine IPTC-Felder oder EXIF-XP...Felder
'           L�sung: Bei 'Pr�fen1' setze ich das Feld IPTCPresent = False,
'                   wenn das DateLastModified der Datei aktueller ist, als das in der Datenbank eingetragene DDatum.
'                   Dadurch kann ich sp�ter mit der Funktion 'EXIF/IPTC...' f�r 'IPTCPresent=False' die fehlenden Felder wieder eintragen.
'----------------------------------------------------------------------------------------------------------
'06.04.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: In der Form NachPr�fen3Aufnehmen kann bei Rechtsklick auf eine Listbox-Zeile ein Vorschaubild gezeigt werden.
'                   Bisher geht das nicht f�r Videos.
'           L�sung: Im Ordner TempThumbs wird ein Video-Thumbnail erzeugt und danach angezeigt.
'                   Neue Klassenmoduln sind GdipLoader.cls und GdipTools.cls
'                   Bei Programmstart wird der Inhalt vom Ordner \TempThumbs\ gel�scht
'                   Es wird shell32.dll gebraucht Projekt -> Verweise -> 'Microsoft Shell Controls and Automation'
'----------------------------------------------------------------------------------------------------------
'08.04.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Nach Pr�fen1 wird der Button Abbruch nicht disabled
'           L�sung: Wenn Pr�fen1 fertig ist, ausf�hren von 'btnPr�fen1Abbrechen.Enabled = False'
'----------------------------------------------------------------------------------------------------------
'08.04.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn die Funktion 'EXIF/IPTC...' nichts zu tun vorfindet kommt kein derartiger Hinweis
'           L�sung: Dann steht in txtArbeitsfortschritt.Text = LoadResString(1130 + Sprache)          'nothing to do Gerbing
'----------------------------------------------------------------------------------------------------------
'16.06.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Bei IPTC-Import aus dem Feld IPTC-Category nach dem Feld SWF kommt Laufzeitfehler '94' Ung�ltige Verwendung von Null
'           L�sung: On error Resume Next
'----------------------------------------------------------------------------------------------------------
'22.07.2017 15.0.1 Fehlerkorrektur SQL-Server-Version, in der Access-Version passiert nichts
'           Zustand: Wenn ein Dateiname ein einfaches Hochkomma enth�lt wie zB d'Ampezzo.jpg dann kommt Laufzeitfehler
'           L�sung: Schuld ist die Duplikatpr�fung, weil diese nicht mit dem Hochkomma zurechtkommt
'                   Jetzt bringe ich lieber einen falschen Hinweis auf Duplikatfehler obwohl keiner vorliegt, als Laufzeitfehler
'                   Der Nutzer soll den Dateinamen ohne Hochkomma bereitstellen
'----------------------------------------------------------------------------------------------------------
'09.08.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Zustand: InnoSetup konnte FotosMdb.exe nicht richtig installieren run time error 430
'           L�sung: f�r Shell32 sp�te Bindung benutzen
'----------------------------------------------------------------------------------------------------------
'25.08.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Delete * From Fotos bringt Fehler bei englischer Version wenn ich fotos.mdb l�schen will
'                   Error on 'DELETE * FROM FotosErrornumber=-2147467259Errortext=Feld 'Kommentar' wurde nicht gefunden.
'           L�sung: 'Delete From Fotos' soll bei SQL Server benutzt werden, geht auch f�r Access
'                   aber ich muss es zweimal nacheinander ausf�hren
'----------------------------------------------------------------------------------------------------------
'03.10.2017 15.0.1 Fehlerkorrektur SQL-Server-Version Nacharbeit zu 22.07.2017
'           Zustand: Wenn ein Dateiname ein Hochkomma enth�lt zB d'Beispiel.JPG, dann wird ein falscher Fehler gemeldet
'                   Es wird Duplikat-Warnung gezeigt.
'           L�sung: In der SQL-Server-Version muss vor der Duplikatpr�fung ein eventuelles Hochkomma im Dateiname ersetzt werden durch '-'
'           Dateinamen mit Hochkomma werden in jeder Version abgewiesen
'----------------------------------------------------------------------------------------------------------
'18.10.2017 15.0.1 Problem CompactDatabase
'           Zustand: Bei Installation nach C:\GERBING Fotoalbum 15\fotos.mdb kann Newfotos.mdb nicht umgenannt werden in fotos.mdb
'                   am Ende ist fotos.mdb ganz verschwunden
'           L�sung: anstelle von 'rename altername, neuername' benutze ich'file_copy(Quellname, Zielname)'
'                   warum rename nicht funktioniert aber file_copy funktioniert, weis ich nicht
'----------------------------------------------------------------------------------------------------------
'06.11.2017 15.0.1 Fehlerkorrektur SQL-Server-Version
'           Zustand: Bisher ist bei jeder Version der Button 'L�sche den Inhalt von fotos.mdb...' sichtbar
'                   Das ist sinnlos bei der SQL-Server-Version, weil es keine fotos.mdb gibt
'           L�sung: Bei der SQL-Server-Version bleibt der Button unsichtbar
'==========================================================================================================
'23.11.2017 15.0.2 Problem mit unicode filename wenn zB GGCnopt\fotos.mdb  Access Datenbank
'           kein Problem mit SQL-Server-Version
'           Zustand: Es kommt 'Kein zul�ssiger Dateiname' fr�her ging das schon mal
'                   Vermutlich hat Microsoft daran herumgedreht.
'                   Die Datenbank l�sst sich aber mit MS Access �ffnen.
'                   Ich komme mit DBsql.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;... dar�ber hinweg
'                   aber dann spinnen andere Stellen im Code, die mit DAO programmiert sind
'           L�sung: DAO Code durch ADO Code ersetzen
'                   Pr�fung ob die Datenbank schreibgesch�tzt ist mit SQL = "UPDATE FET SET FN = 'test'"
'                   Reference auf Microsoft DAO 3.6 Object Library wird nicht mehr gebraucht. ado360.dll wird nicht mehr gebraucht
'                   CompactDatabase wird ganz entfernt. Das macht nur noch fotos.exe
'----------------------------------------------------------------------------------------------------------
'10.12.2017 15.0.2 Verbesserung alle Versionen
'           Zustand: Video-Datei-Typ "MKV" und "FLV" wird bisher nicht akzeptiert
'           L�sung: ab sofort ist "MKV" und "FLV" erlaubt
'                   bei "MKV" gibt es keine Vorschaubilder
'----------------------------------------------------------------------------------------------------------
'07.01.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: Die Formulare �ffnen an unterschiedlichen Positionen, meist mit StartUpPosition=3=Windows-Standard.
'                   Ich will alle in Fenstermitte
'           L�sung: StartUpPosition=1=Fenstermitte
'----------------------------------------------------------------------------------------------------------
'23.01.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: Bei Pr�fenS kommt Laufzeitfehler wenn der Dateiname Hochkomma enth�lt
'           L�sung: Dateinamen mit Hochkomma m�ssen durch 2 Hochkommas ersetzt werden
'----------------------------------------------------------------------------------------------------------
'14.02.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: in VM Win7 und VM Win10 kann Fotosmdb nach Pr�fen3 kein Video-Thumbnail zeigen -> Laufzeitfehler '91'
'           L�sung: On Error Resume Next -> Msgbox 'Folder is nothing'
'==========================================================================================================
'22.11.2018 15.0.3 Verbesserung alle Versionen
'           Zustand: Nachbesserung zum 14.02.2018 bei unicode folder name
'           L�sung MessageBoxW
'==========================================================================================================
'04.04.2019 15.0.4 Fehlerkorrektur vermutlich nur meine Vollversion
'           Zustand: Bei Pr�fen3 -> Neuaufnahme -> Vorschaubild f�r mp4 kommt kein Vorschaubild. In der Sharewareversion kommt es.
'           Nicht funktionierende L�sung: Version 15.0.4 compilieren
'           Ursache gefunden: Es liegt am Name des Ordners mit unicode Zeichen VideoAlbumCnopt
'----------------------------------------------------------------------------------------------------------
'08.04.2019 14.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Obwohl in einem Foto GPS-Daten eingetragen sind(Kontrolle mit Windows-Explorer -> Eigenschaften -> Details -> �berschrift GPS)
'                   kommt in fotos.exe und Diashow.exe eine MsgBox 'Geo positions not available'
'           Ursache: Die GEO-Positionen sind im XMP-Abschnitt des Fotos eingetragen. Das macht zB Geosetter(mit Hilfe von Exiftool),
'                   ich suche sie aber nur im EXIF-Abschnitt.
'                   Andere Software findet diese GEO-Positionen zB ExifToolGUI, PSP 2019, Irfan View, Fotos App, XnViewMP.
'                   Fotos App(Win10) korrigiert sogar selbst�ndig aus dem XMP-Abschnitt in den EXIF-Abschnitt
'           L�sung: Da ich in clsEXIF sowieso jedes JPG-Foto durchsuche, um den EXIF-Abschnitt zu finden, kann ich dort ebenso nach den
'                   XMP-GEO-Positionen suchen
'                   Ich suche nach exif:GPSLatitude und exif:GPSLongitude mit InstrB, weil das rasend schnell geht
'                   Die gefundenen Werte gstrLatXMP und gstrLongXMP muss ich dann noch in ein Format verwandeln, das OpenStreetMap versteht
'                   und in die Datenbankfelder GPSLatitude und GPSLongitude eintragen
'                   zB gstrLatXMP 50,38.7309456N -> 50.64551575
'                   zB gstrLongXMP 11,53.9826786E -> 11.89971130
'----------------------------------------------------------------------------------------------------------
'28.05.2019 14.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Pr�fen3 -> Gefundene in die Datenbank aufnehmen -> rechtsklicken und ziehen -> Laufzeitfehler '91'
'           L�sung: If Not listItem Is Nothing abfragen
'----------------------------------------------------------------------------------------------------------
'28.05.2019 14.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Bei Videos werden bei Pr�fen3 Dateien mit '+:\TempThumbs\... gezeigt, wenn man 2x nacheinander Pr�fen3 macht
'           L�sung: siehe 06.04.2017 nicht nur bei Form_Load sondern auch bei btnPr�fen3_Click Aufrufen von RekursiveTempThumbs
'----------------------------------------------------------------------------------------------------------
'04.07.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Nacharbeiten zum 08.04.2019
'                   Die GPS-Daten in den EXIF-Feldern werden richtig angezeigt, aber sind fehlerhaft in den Datenbank-Feldern
'                   Ursache ist ein Fehler in Form1.GEOKoordinatenUmrechnenXMP
'                   Beispiel: Minuten = 0.287564
'                   Nachkomma = MinutenDouble / 60 'liefert Ergebnis=0
'           L�sung: Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
'                   sonst kommt bei MinutenDouble / 60 Ergebnis=0
'==========================================================================================================
'02.10.2019 15.0.5 Neue Funktion
'           Zustand: Bisher muss ich Fremd-Software zu Hilfe nehmen um die GEO-Position zu einem Foto nachtr�glich
'                   festzulegen. F�r Videos gibt es noch keine brauchbare Fremd-Software.
'                   Bei JPG-Fotos tragen zB Picasa oder GeoSetter die GEO-Position in den EXIF-Abschnitt ein.
'                   Anschlie�end kann ich mit Men� Datei.. -> Feldaktualisierung durch Import-Wiederholung
'                   die Datenbank-Felder GPSLatitude und GPSLongitude auff�llen, das bleibt auch so.
'           L�sung: Neue Funktion in Fotos.exe. Anlegen der Felder GPSLatitude und GPSLongitude, wenn es diese noch nicht gibt.
'                   Jetzt kann jede Datei in der Datenbank mit den Feldern GPSLatitude und GPSLongitude versehen werden.
'                   Daf�r entf�llt der Klimmzug mit MediaInfo.dll
'----------------------------------------------------------------------------------------------------------
'12.10.2019 15.0.5 Neuerung
'           Zustand: Ich will ab sofort keine Shareware-Version und keine Professional Version mehr pflegen. Schade um den Aufwand.
'                   Elke verkauft mehr Leseknochen als ich je Software verkauft habe.
'                   Es soll nur noch eine Freeware Vollversion geben. Die SQL-Server-Version wird nicht kostenlos.
'                   W�re aber kostenlos m�glich mit einer 99-Lizenz.
'           L�sung: �nderungen in der Form IPTCGenerieren
'                   Bei nutzerdefinierten Feldern soll mit Rechtsklick auf den Frame FrameNutzerDefiniert
'                   die Feld-Zuordnung r�cksetzbar sein.
'                   Die Comboboxen cmbFeld1-cmbFeld5 m�ssen aufsteigend sortiert sein
'                   Die Comboboxen cmbFeld1-cmbFeld5 und cmbEx1.cmbEx5 d�rfen nicht DropDown-Liste sein sondern DropDown-Kombinationsfeld
'----------------------------------------------------------------------------------------------------------
'14.11.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Bisher verlange ich vom Nutzer, dass er das Feld EXIFDateTimeOriginal selber erzeugt mit MS Access
'                   Das hat in anderen �hnlichen F�llen das Programm selbst gemacht.
'           L�sung: Falls EXIFDateTimeOriginal nicht angelegt ist, legt fotos.exe es an sowohl bei der Access-Version wie beim SQL Server
'----------------------------------------------------------------------------------------------------------
'14.11.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Bisher verlange ich vom Nutzer, dass er das Feld VideoDuration selber erzeugt mit MS Access
'                   Das hat in anderen �hnlichen F�llen das Programm selbst gemacht.
'           L�sung: Falls VideoDuration nicht angelegt ist, legt fotos.exe es an sowohl bei der Access-Version wie beim SQL Server
'----------------------------------------------------------------------------------------------------------
'15.11.2019 Verbesserung:
'           Zustand: Pr�fen1 -> Alle Dateien berechnen setzt die Spalten Pixelbreite Pixelhoehe VideoDuration auf leer,
'                   wenn in AppPath unicode Zeichen vorkommen zB VideoAlbumCnopt
'           Not-L�sung: den Unicode-AppPath kurz mal umnennen, dann Pr�fen1 ausf�hren
'           L�sung: von fotos.exe abgucken. Dort kann ein Video w�hrend des Abspielens seine Breite H�he Dauer ins DbGridNeu eintragen.
'                   Man braucht WMP.DLL.
'                   Das passiert in WMP_PlayStateChange
'----------------------------------------------------------------------------------------------------------
'18.11.2019 15.0.5 Verbesserung
'           Zustand: Es gibt gegenw�rtig nur eine mir bekannte Software, die die GPS-Angaben eines Smartphone MP4-Videos oder MOV-Videos
'                   von der digitalen Kamera nach dem Editieren unangetastet l��t. Ebenso das Feld 'Encoded Date'.
'                   Das ist im Windows 10 die Fotos App von Microsoft.
'                   Andere Software macht folgendes:
'                   beim Editieren oder Schneiden oder Zusammenf�gen gehen die GPS-Daten eines Smartphone MP4-Videos verloren.
'                   Die GPS-Daten verschwinden sogar beim Zurechtschneiden auf dem Smartphone.
'           L�sung: 1. Ich muss zum Editieren von mp4 oder mov videos im Windows 10 die Fotos App von Microsoft benutzen.
'                   2. Beim Aufnehmen von mp4 oder mov files mit MediaInfo.DLL(must be i386 version, getestet mit Version 18.8.1.0)
'                      nach dem Feld "xyz" suchen,
'                      das ist das GPS-Feld. Was dort steht, wandert in die Datenbank-Felder GPSLatitude und GPSLongitude.
'                      Ich ignoriere die Felder Exif-GPSLatitude und Exif-GPSLongitude, weil dort eh nichts steht.
'                   3. Nach dem Feld 'Encoded Date' suchen. Was dort steht, wandert in das Datenbank-Feld ExifDateTimeOriginal.
'----------------------------------------------------------------------------------------------------------
'22.11.2019 15.0.5 Fehlerkorrektur
'           Zustand: Manchmal steht in der Datei pruef.log ein Fehler aber kein zugeh�riger Dateiname
'                   zB DatensatzNr. xyz Die Datei "" Widerspruch in Dateinamen-Erweiterung und Spalte SWF
'           L�sung: Der Programmierer hat schlecht programmiert manchmal muss es hei�en FotoDatei und manchmal strFotoDatei
'----------------------------------------------------------------------------------------------------------
'24.11.2019 15.0.5 Nachbesserung zum 18.11.2019
'           Zustand: Wenn es in den GPS-EXIF-Feldern keinen Eintrag gibt, enthalten die Felder GPSLatitude und GPSLongitude den Wert '0'
'                   Das ist falsch, es muss der Wert Null sein. Null bedeutet nichts, das Feld enth�lt keinen Wert.
'           L�sung: Zu Beginn der Eintragungen in jeden neuen Datensatz werden die betroffenen Felder auf Null gesetzt
'                   rstsql.Fields(LoadResString(1106 + Sprache)) = Null 'Breite
'                   rstsql.Fields(LoadResString(1107 + Sprache)) = Null 'Hoehe
'                   rstsql.Fields("VideoDuration") = Null
'                   rstsql.Fields("GPSLatitude") = Null
'                   rstsql.Fields("GPSLongitude") = Null
'----------------------------------------------------------------------------------------------------------
'25.11.2019 15.0.5 Verbesserung
'           Zustand: chkExif braucht nicht eingeschaltet zu sein um die Felder GPSLatitude und GPSLongitude zu f�llen
'                   Hauptsache ist, dass in Tabelle UserDefined eine Zuordnung getroffen wurde.
'                   Desselbe soll auch mit ExifDateTimeOriginal m�glich sein.
'           L�sung: cmbFeld1 bis cmbFeld5 abfragen ob in Exif-DateTimeOriginal etwas gefunden wurde
'==========================================================================================================
'22.12.2019 16.0.0 Verbesserung
'           Zustand: Wenn der Dateiname keine Jahreszahl enth�lt wird bisher eine Dummy-Zahl = 9999 benutzt
'           L�sung: Ab sofort benutze ich die aktuelle Jahreszahl
'----------------------------------------------------------------------------------------------------------
'23.12.2019 16.0.0 Fehlerkorrektur
'           Zustand: Manchmal st�rzt Pr�fen1 ab. Vorher kommt ein Fehler �ber nicht vorhandenene Datei.
'           L�sung: Der Programmierer hat schlecht programmiert, manchmal muss es hei�en FotoDatei und manchmal strFotoDatei
'                   In FehlerFotoDatei(DatensatzNr) muss es hei�en
'                   richtig: NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add strFotoDatei
'                   falsch: NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add FotoDatei
'                   Wenn FotoDatei = "" dann st�rzt die exe ab
'----------------------------------------------------------------------------------------------------------
'03.01.2020 16.0.0 Nachbesserung zum 02.10.2019
'           Zustand: In Form IPTCGenerieren funktioniert das �bertragen der Felder GPSLatitude/GPSLongitude in die entsprechenden
'                   EXIF-Felder der JPG-Fotos nur dann, wenn der Nutzer in der gleichen Sitzung die entsprechenden Comboboxen ausgef�llt hat.
'                   Wenn dagegen mit den gespeicherte Werten aus Tabelle 'UserDefined' gearbeitet werden soll, dann funktioniert
'                   die �bertragung nicht.
'           L�sung: cmbEx1_Click bis cmbEx5_Click ausl�sen
'----------------------------------------------------------------------------------------------------------
'04.01.2020 16.0.0 Fehlerkorrektur
'           Zustand: In Form IPTCGenerieren kann btnStart mehrfach geklickt werden danach gibt es den Fehler
'                   'Der Vorgang ist f�r ein ge�ffnetes Objekt nicht zugelassen'
'           L�sung: btnStart.Enabled = False und btnStart.Enabled = True benutzen damit wird verhindert, dass btnStart mehrfach
'                   geklickt werden kann
'----------------------------------------------------------------------------------------------------------
'30.06.2020 16.0.0 Fehlerkorrektur
'           Zustand: Wenn gerbingsoft.log fehlt kann Pr�fenS nicht visible gesetzt werden.
'           L�sung: Abfrage von gerbingsoft.log entfernen









'*********************************************************************************************************
'offenes Problem:
'           Wenn fotosmdb.exe aus fotos.exe heraus gestartet wird und man benutzt 'Pr�fenS'
'           kommt manchmal Laufzeitfehler 3218 zur Datenbank
'           Aktualisierung im Augenblick nicht m�glich weil gesperrt
'       L�sung: m�glicherweise gel�st weil im Multi-Nutzer-Betrieb ist Pr�fenS verboten
'
'offenes Problem:
'           Bei Videos mit unicode filename und manchen anderen videos kann ich die Videosize/Duration nicht ermitteln,
'           ich habe bisher keinen geeigneten sample code gefunden.
'       Umgehungsl�sung: In fotos.exe k�nnen beim Abspielen mit dem internen Mediaplayer diese Werte festgestellt werden.
'           Die trage ich in die Datenbank ein.
'
'nichtl�sbares Problem:
'           Der Wunsch bez�glich 'Pr�fen2' und anschlie�end automatisches Verschieben in den Ordner mit der
'           richtigen 4-stelligen Jahreszahl, egal in welcher Verschachtelungstiefe dieser Ordner gefunden wird,
'           ist nicht l�sbar. Diese Forderung ist nicht erf�llbar, wenn im Verzeichnisbaum die 4-stellige
'           Jahreszahl mehrfach auftaucht (in verschiedener Verschachtelungstiefe), wei� der Automat nicht wohin
'           er verschieben soll.
'           Beispiel es gibt nur zwei Dateien - eine im Jahr 2001 und eine im Jahr 2002
'           C:\P5Daten\VISUALBA.SIC\Foto\Kopie von Fotosmdb\JahrExperimente\Muster1\2001\Musterfoto01.jpg
'           C:\P5Daten\VISUALBA.SIC\Foto\Kopie von Fotosmdb\JahrExperimente\Muster2\2002\Musterfoto02.jpg
'           jetzt wird in Fotos.exe aus 2001 das Jahr 2002 gemacht
'           Woher soll der Automat wissen, da� die Datei nach Muster2\2002 verschoben werden soll
'           und nicht nach Muster1\2002
'           In der Hilfe muss stehen, da� nur in Jahresordner direkt unter AppPath verschoben wird.

'l�sbares Problem mit meiner pers�nlichen englischsprachiger Datenbank:
'           Wenn Fotosmdb S�tze �ndern oder l�schen soll kommt
'           Laufzeitfehler '-2147467259 (80004005)' Feld 'Ort' wurde nicht gefunden
'           Beim �ndern von Feldinhalten in der Datenbank fotos.mdb kommt - Feld 'Ort' wurde nicht gefunden (Fehler 3799)
'           Es sind die G�ltigkeitsregeln schuld, die nur in meiner Datenbank vorkommen zu Ort und Kommentar
'           Wenn ich die Datenbank englischsprachig mache, bleiben doch die deutschsprachigen G�ltigkeitsregeln erhalten
'
'offenes Problem:
'           Ich m�chte gerne die Funktion Pr�fen1 beim Berechnen von BreitePixel/HoehePixel beschleunigen
'           Ein Versuch am 17.06.2016 mit unicodef�higem Ersatz von '...For Binary Access Read As...'
'           ging nicht schneller. Vermutlich gehts nicht besser.
'           Es dauert etwa genauso lange wie eine neue fotos.mdb erzeugen aus den IPTC-Feldern, Dauer etwa 12 Minuten
'
'gel�stes Problem:
'           Bei Ralph Dittrich kam ein Fehler beim Versuch eine Datei mit Pr�fen3 aufzunehmen, deren Dateiname l�nger war als 100 Zeichen
'           Ursache: Tabelle Temp_Haken -> Beim Feld Dateiname war Feldgr��e=100 eingetragen
'           L�sung: Tabelle Temp_Haken -> Beim Feld Dateiname muss Feldgr��e=255 eingetragen werden
'           Meinen Versuch einen Dateinamen mit einer L�nge von �ber 255 Zeichen zu erzeugen hat der Explorer selbst korrigiert
'           Der TotalCommander hat eine entspechende Warnung ausgegeben
'
'nicht reproduzierbares Problem:
'           Der Start von Fotosmdb.exe im Ordner e:\FOTOS\Gg\ mit 30.000 Fotos dauert etwa 10 Sekunden.
'           Der Start von Fotosmdb.exe mit dem SQL-Server dauert etwa 4 Sekunden
'           Bei der Suche nach der Ursache st�rzte die vb6 IDE von Fotosmdb.vbp ab, st�rzte auch fotos.vbp ab, aber nicht diashow.vbp
'           Es passierte mit Shlwapi.DLL und Winmm.DLL und es muss am Datenbankzugriff gelegen haben
'           Nach 6 Stunden herumprobieren und Neuinstallation von AccessDatabaseEngine.exe ging alles wieder
'           und der Start von Fotosmdb.exe im Ordner e:\FOTOS\Gg\ mit 30.000 Fotos dauert etwa 2 Sekunden.











Option Explicit
Option Compare Text
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

    Private Declare Function GetSystemDirectoryA Lib "kernel32" _
       (ByVal lpBuffer As String, ByVal nSize As Long) As Long

    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Dim NL As String
    Public FehlerGefunden As Boolean
    Dim DatensatzNr As Long
    Dim JahresZahl As String
    Public DateiNummer As Long
    Dim Pr�fDateiNummer As Long                         'Gerbing 30.09.2004
    Dim AviNummer As Long
    Private StartVerzeichnis As String
    Public Pr�fen4Alle As Boolean
    Dim Pr�fen1Abbrechen As Boolean                     'Gerbing 06.11.2013
    Dim Pr�fen3Abbrechen As Boolean                     'Gerbing 04.10.2004
    Dim Pr�fenSAbbrechen As Boolean                     'Gerbing 12.04.2006
    Dim Pr�fen4Abbrechen As Boolean
    Dim Pr�fen5Abbrechen As Boolean
    Dim blnMessageAusgeben As Boolean                   'Gerbing 26.01.2009
    Dim MaxPixelWidth As Long
    Dim MaxPixelHeight As Long
    Public Pr�fenNummer As String
    Dim Grenzwert As Integer
    Dim AnteilBetroffenePixelZ�hler As Integer
    Dim AnteilBetroffenePixelNenner As Integer
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
    
    Dim ColWidth1 As Long
    Dim ColWidth2 As Long
    Dim ColWidth3 As Long
    Dim ColWidth4 As Long
    Dim ColWidth5 As Long
    Dim ColWidth6 As Long
    Dim ColWidth7 As Long
    Dim ColWidth8 As Long
    Dim ColWidth9 As Long
    Dim ColWidth10 As Long
    Dim ColWidth11 As Long
    Dim ColWidth12 As Long
    Dim ColWidth13 As Long
    
    Dim DateinamenErweiterung As String
    Public gstrNeuerName
    Public blnMitBH As Boolean
    Public blnNurNeue As Boolean                                'Gerbing 19.01.2006
    Public blnReturn As Boolean                                 'Gerbing 22.11.2008
    
    Public EXF As New clsEXIF                                   'Gerbing 07.05.2007
    'Private iptc As New IPTCInfo.Reader
    
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
    ' ben�tigte API-Deklarationen
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
    
    Public DBsql As ADODB.Connection
    Public rstsql As ADODB.Recordset
    Public rstDBH As ADODB.Recordset
    Public rsDataGrid As ADODB.Recordset
    Public rstTempHaken As ADODB.Recordset
    
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer  'Gerbing 10.09.2013
    Private Const KeyPressed As Integer = -32767
    Private Const VK_CONTROL = &H11
    Private Const VK_MENU As Long = &H12&
    Private Const VK_SHIFT As Long = &H10&
    Private Const VK_CAPITAL As Long = &H14&
'    VK_CLEAR = &HC
'    VK_RETURN = &HD
'    VK_SHIFT = &H10
'    VK_CONTROL = &H11
'    VK_MENU = &H12
'    VK_PAUSE = &H13
'    VK_CAPITAL = &H14
'    VK_ESCAPE = &H1B
'    VK_PRIOR = &H21
'    VK_NEXT = &H22
'    VK_HOME = &H24
'    VK_BAK = &H8
'    VK_TAB = &H9
'    VK_LEFT = &H25
'    VK_UP = &H26
'    VK_RIGHT = &H27
'    VK_DOWN = &H28
'    VK_SELECT = &H29
'    VK_END = &H23
'    VK_SNAPSHOT = &H2C
'    VK_INSERT = &H2D
'    VK_DELETE = &H2E
'    VK_HELP = &H2F
'    VK_F1 = &H70
'    VK_F2 = &H71
'    VK_F3 = &H72
'    VK_F4 = &H73
'    VK_F5 = &H74
'    VK_F6 = &H75
'    VK_F7 = &H76
'    VK_F8 = &H77
'    VK_F9 = &H78
'    VK_F10 = &H79
'    VK_F11 = &H7A
'    VK_F12 = &H7B
'    VK_F13 = &H7C
'    VK_F14 = &H7D
'    VK_F15 = &H7E
'    VK_F16 = &H7F
'    VK_NUMLOCK = &H90
'    VK_SCROLL = &H91
'    VK_WIN = &H5B
'    VK_APPS = &H5D
    Public AnzahlFehlerPr�fen2 As Long
    'Gerbing 04.07.2019
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Const LOCALE_SDECIMAL = &HE                 '  decimal separator
    Public lngVideoDuration As Long                    'Gerbing 15.11.2019
    Public pintBreite As Integer                        'Gerbing 15.11.2019
    Public pintHoehe As Integer                         'Gerbing 15.11.2019
    Public blnMediaPlayerStopped As Boolean             'Gerbing 15.11.2019
    Private Fotodatei As String                         'Gerbing 15.11.2019
    Private strFotoDatei As String                      'Gerbing 15.11.2019
    Public blnMediaPlayerError As Boolean               'Gerbing 15.11.2019
    Private glngStartMillisek As Long                   'Gerbing 15.11.2019
    Private glngEndMillisek As Long                     'Gerbing 15.11.2019

          
Private Sub Form_Initialize()
    Dim retStatus As Status
    Dim returncode As Long

    GdipInitialized = False
    retStatus = Execute(StartUpGDIPlus(GdiplusVersion))             'Gerbing 30.09.2015
    If retStatus = OK Then
        GdipInitialized = True
    Else
        MsgBox "GDI+ not inizialized.", vbOKOnly, "GDI Error"
    End If
    
    InitCommonControls
    Set PruefFso = New FileSystemObject
    Set IniFso = New FileSystemObject
End Sub

Private Sub btnBeenden_Click()
    Unload Me
    Unload NachPr�fen3L�schen
End Sub

Public Sub btnGenerieren_Click()

    Screen.MousePointer = vbHourglass                                                   'Gerbing 04.01.2017
    If gblnSchreibgesch�tzt = True Then
        'MsgBox "Bei einer schreibgesch�tzten Datenbank k�nnen keine neuen S�tze generiert werden."
        MsgBox LoadResString(1371 + Sprache)
        Exit Sub
    End If
    NeueDatens�tzeGenerieren.Hide
    'Unload NeueDatens�tzeGenerieren
    NeueDatens�tzeGenerieren.blnOptGew�hlt = False
    Screen.MousePointer = vbDefault                                                   'Gerbing 04.01.2017
    NeueDatens�tzeGenerieren.Show 1
    'wie Button Reset
    Call SpaltenbreiteMerken
    
    rsDataGrid.Requery
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
    
    Call SpaltenbreiteWiederherstellen
    DBGridNeu.Caption = PublicDatagridCaption
    DBGridNeu.AllowUpdate = False

End Sub

Private Sub btnEXIFIPTC_Click()
    Dim Msg As String
    Dim antwort As Long
    
    If gblnSchreibgesch�tzt = True Then
'        msg = "Sie arbeiten mit einer schreibgesch�tzten Datenbank." & vbNewLine
'        msg = msg & "Wenn auch die Fotos schreibgesch�tzt sind, k�nnen Sie keine IPTC-Felder �bernehmen" & vbNewLine
'        msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = LoadResString(1540 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(1541 + Sprache) & vbNewLine
        Msg = Msg & LoadResString(1542 + Sprache)

        antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
        If antwort = vbNo Then Exit Sub
    End If
    On Error Resume Next                                                                            'Gerbing 02.09.2008
    IPTCGenerieren.Show 1
    On Error GoTo 0                                                                                 'Gerbing 02.09.2008
End Sub


Private Sub btnL�scheInhaltFotosMdb_Click()
    'F�r den Fall, da� der Nutzer die Datenbank mit Hilfe der IPTC-Felder neu aufbauen will, bekommt
    'er hier die M�glichkeit alle Datens�tze der Tabelle Fotos zu l�schen.
    'In der Tabelle ErsterStart ist das Feld 'DatumBreiteHoehe' zu l�schen          'Gerbing 09.11.2006
    'Bei einer Schreibgesch�tzten Datenbank git es nichts zu l�schen
    
    Dim SQL As String
    Dim Msg As String
    Dim antwort As Long
    
'    msg = "Wollen Sie wirklich alle Felder der Tabelle Fotos l�schen?" & NL
'
'    msg = msg & "Diese Funktion sollten Sie nur benutzen, wenn Sie eine Sicherungskopie der Datenbank-Datei besitzen" & NL
'    msg = msg & "und wenn Sie die Felder der Datenbank automatisch aus den IPTC-Feldern reproduzieren k�nnen." & NL
'    msg = msg & "Um die IPTC-Felder zu reproduzieren, wird empfohlen mit dem Tool Fotosmdb und der Funktion Pr�fen3 zu arbeiten"
    Msg = LoadResString(1514 + Sprache) & NL & NL
    Msg = Msg & LoadResString(1515 + Sprache) & NL
    Msg = Msg & LoadResString(1516 + Sprache) & NL
    Msg = Msg & LoadResString(1517 + Sprache)
    antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
    If antwort = vbNo Then Exit Sub
    
    If gblnSchreibgesch�tzt = True Then
        'MsgBox "Bei einer schreibgesch�tzten Datenbank k�nnen Sie nichts l�schen."
        MsgBox LoadResString(1510 + Sprache)
        Exit Sub
    End If
    
'    Msg = "Antworten Sie mit Ja, wenn Sie alle Datens�tze der Datenbank l�schen wollen oder mit Nein, wenn Sie nur die Datens�tze zu JPG-Fotos l�schen wollen."
    Msg = LoadResString(1839 + Sprache)
    antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
    If antwort = vbYes Then
        Screen.MousePointer = vbHourglass
        On Error Resume Next
        ERR.Number = 0                                                  'Gerbing 25.08.2017
        SQL = "DELETE From Fotos"                                       'Gerbing 25.08.2017 Delete * From Fotos bringt Fehler in englischer fotos.mdb
        DBsql.Execute (SQL)
        ERR.Number = 0                                                  'Gerbing 25.08.2017
        DBsql.Execute (SQL)
        If ERR.Number <> 0 Then
            Msg = "Error on 'DELETE FROM Fotos"
            Msg = Msg & "Errornumber=" & ERR.Number & vbNewLine
            Msg = Msg & "Errortext=" & ERR.Description
            MsgBox Msg
            On Error GoTo 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        'nur die Datens�tze zu JPG-Fotos
        Screen.MousePointer = vbHourglass
        On Error Resume Next
        If gblnSQLServerVersion = True Then
            'beim SQL Server muss es hei�en 'Delete from table
            SQL = "DELETE From Fotos where " & LoadResString(1028 + Sprache) & " like '%.jpg'"          'Dateiname=1028
        Else
            SQL = "DELETE * FROM Fotos where " & LoadResString(1028 + Sprache) & " like '%.jpg'"        'Dateiname=1028
        End If
        DBsql.Execute (SQL)
        If ERR.Number <> 0 Then
            Msg = "Error on " & SQL
            Msg = Msg & "Errornumber=" & ERR.Number & vbNewLine
            Msg = Msg & "Errortext=" & ERR.Description
            MsgBox Msg
            On Error GoTo 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If gblnSQLServerVersion = True Then
        'beim SQL Server muss es hei�en 'Delete from table
        SQL = "DELETE From ErsterStart"
        'SQL = "DELETE FROM " & LoadResString(2527 + Sprache)
    Else
        SQL = "DELETE * From ErsterStart"
        'SQL = "DELETE * FROM " & LoadResString(2527 + Sprache)          '2527=ErsterStart
    End If
    DBsql.Execute (SQL)
    If ERR.Number <> 0 Then
        Msg = "Error on 'DELETE * FROM ErsterStart'"
        Msg = Msg & "Errornumber=" & ERR.Number & vbNewLine
        Msg = Msg & "Errortext=" & ERR.Description
        MsgBox Msg
        On Error GoTo 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    MsgBox LoadResString(1007 + Sprache)        'Fertig

    rsDataGrid.Requery
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
End Sub

Private Sub btnNutzerdefinierteFelderAnlegen_Click()
    Dim Msg As String
    
    If gblnProversion = False Then
        Msg = LoadResString(2335 + Sprache) 'F�r diese Funktion ben�tigen Sie die Professional Version.
        MsgBox Msg
        Exit Sub
    End If
    If gblnSchreibgesch�tzt = True Then
        'MsgBox "Bei einer schreibgesch�tzten Datenbank k�nnen keine neuen S�tze generiert werden."
        MsgBox LoadResString(1371 + Sprache)
        Exit Sub
    End If
    '--------
    ND.Show 1
    Call btnReset_Click
End Sub

Private Sub btn�ffnePruefLog_Click()
    Dim retval As Long
    Dim LogFileName As String
    Dim intL�nge As Integer
    Dim ErrorText As String
    Dim Msg As String
    
    LogFileName = PruefLogFile
    retval = RunShellExecute(Me.hWnd, "open", LogFileName, vbNull, vbNull, 1)
    If retval <= 32 Then
        If Mid(LogFileName, Len(LogFileName) - 3, 1) = "." Then                     'Gerbing 25.06.2006
            intL�nge = 3
        End If
        If Mid(LogFileName, Len(LogFileName) - 4, 1) = "." Then
            intL�nge = 4
        End If
        If Mid(LogFileName, Len(LogFileName) - 5, 1) = "." Then
            intL�nge = 5
        End If
        DateinamenErweiterung = Right(LogFileName, intL�nge)
        ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
        Msg = "Errortext=" & ErrorText & vbNewLine
        Msg = Msg & "Errornr=" & retval & vbNewLine & vbNewLine
        
        Msg = Msg & LogFileName & vbNewLine
        'Msg = Msg & "Diese Datei kann nicht ge�ffnet werden." & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
        Msg = Msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
        
        'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
        Msg = Msg & LoadResString(1378 + Sprache) & vbNewLine
        'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verkn�pft." & vbNewLine
        Msg = Msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
        'Msg = Msg & "W�hlen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
        Msg = Msg & LoadResString(2012 + Sprache) & vbNewLine
        'Msg = Msg & "Rechtklicken auf den Dateiname -> �ffnen mit... -> Programm ausw�hlen"
        Msg = Msg & LoadResString(2013 + Sprache)
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
    End If
End Sub

Private Sub btnPr�fen1Abbrechen_Click()
    Dim Mldg, Stil, Titel, Hilfe, Ktxt, antwort, Text1
    'Mldg = "M�chten Sie Pr�fen1 Abbrechen ?"       ' Meldung definieren.           'Gerbing 06.11.2013
    Mldg = LoadResString(1447 + Sprache)            ' Meldung definieren.
    Stil = vbYesNo + vbDefaultButton2
    'Titel = "Pr�fen1 Abbrechen"                 ' Titel definieren.                'Gerbing 06.11.2013
    Titel = LoadResString(1448 + Sprache)                 ' Titel definieren.
    antwort = MsgBox(Mldg, Stil, Titel)
    If antwort = vbYes Then
        Pr�fen1Abbrechen = True
        Screen.MousePointer = vbDefault
        btnPr�fen1Abbrechen.Enabled = False
    Else
        Pr�fen1Abbrechen = False
    End If
End Sub

Private Sub btnPr�fen2_Click()
    Dim Gefunden As String
    Dim Jahr As String
    Dim Msg As String
    Dim SQL As String
    Dim start, pos As Integer
    Dim Erg As Long
    Dim antwort As Long
                
    AnzahlFehlerPr�fen2 = 0                                                    'Gerbing 26.10.2013
    Call SpaltenbreiteMerken
    Call ButtonsDisabled
    'Pr�fenNummer = "Pr�fen2"
    Pr�fenNummer = LoadResString(1444 + Sprache)
    StartVerzeichnis = ""
    txtFehlerU = ""
    FehlerGefunden = False
    '�ffne die Datei pruef.log
    On Error Resume Next
    oStream.Close                                                                           'Gerbing 06.11.2013
    DateiNummer = FreeFile  ' neue Datei-Nr.
    ERR = 0
    'Open PruefLogFile For Output As #DateiNummer
    'object.CreateTextFile(filename[, overwrite[, unicode]])
    Set oStream = PruefFso.CreateTextFile(PruefLogFile, True, True)
    If ERR <> 0 Then
        'Msg = "Die Datei " & PruefLogFile & " kann nicht ge�ffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & PruefLogFile & " " & LoadResString(1372 + Sprache) & NL
        'msg = msg & "Sie m�ssen f�r Schreibrechte sorgen, damit �nderungen an dieser Datei gemacht werden k�nnen." & NL
        Msg = Msg & LoadResString(2276 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number & NL & NL
        
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = Msg & LoadResString(1542 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)                                   'Gerbing 02.09.2008
        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo)
        If antwort = vbNo Then
            LogNichtBenutzbar = False
            Call ButtonsEnabled
            Exit Sub
        Else
            LogNichtBenutzbar = True
        End If
    End If
    '------------------------------------------------------------------------------------------
    On Error GoTo 0
    Msg = Now & "  "
    Msg = Msg & Pr�fenNummer & "  "
    If gblnSQLServerVersion = True Then
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & PublicSQLServer & PublicSQLDatabase
        Msg = Msg & LoadResString(1381 + Sprache) & PublicSQLServer & " " & PublicSQLDatabase
    Else
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & AppPath & "\fotos.mdb"
        Msg = Msg & LoadResString(1381 + Sprache) & AppPath & "\fotos.mdb"
    End If
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    '------------------------------------------------------------------------------------------
    'Die SQL-Anweisung pr�ft ob das Jahr im Dateiname ein anderes ist als in der Spalte Jahr
    '------------------------------------------------------------------------------------------
    If gblnSQLServerVersion = True Then
        'CharIndex hat andere Parameterreihenfolge als InStr
        'SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(jahr,Dateiname,1)=0;"
        SQL = "SELECT Fotos.* From Fotos WHERE CharIndex(" & LoadResString(1023 + Sprache) & "," & LoadResString(1028 + Sprache) & ")=0;" 'Gerbing 08.11.2005
    Else
        'SQL = "SELECT Fotos.* From Fotos WHERE instr(1,Dateiname, jahr)=0;"         'Gerbing 17.09.2004
        SQL = "SELECT Fotos.* From Fotos WHERE instr(1," & LoadResString(1028 + Sprache) & ", " & LoadResString(1023 + Sprache) & ")=0;"     'Gerbing 17.09.2004
    End If
    Set rstsql = New ADODB.Recordset
    With rstsql
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .Source = SQL
        .Open
    End With

    Call SpaltenbreiteWiederherstellen
    If Not rstsql.EOF Then
        'rstsql enth�lt alle Datens�tze bei denen das Jahr im Dateiname ein anderes ist als in der Spalte Jahr
        rstsql.MoveFirst
        Screen.MousePointer = vbHourglass
        Do Until rstsql.EOF
            'Fotodatei = rstsql("Dateiname")
            Fotodatei = rstsql(LoadResString(1028 + Sprache))
            'Jahr = rstsql("Jahr")
            Jahr = rstsql(LoadResString(1023 + Sprache))
            AnzahlFehlerPr�fen2 = AnzahlFehlerPr�fen2 + 1
            Call FehlerPr�fen2(Jahr)    'schreibe den Fehler in loadresstring(1366+sprache)
            rstsql.Movenext
            'txtArbeitsfortschritt.Text = "DatensatzNr." & DatensatzNr
            txtArbeitsfortschrittU.Text = LoadResString(1008 + Sprache) & " " & DatensatzNr
            Erg = DatensatzNr Mod 100
            'Erg = DatensatzNr Mod 10
            If Erg = 0 Then
                DoEvents
            End If
        Loop
    End If
    Screen.MousePointer = vbDefault
    '------------------------------------------------------------------------------------------
    'schlie�e die Datei loadresstring(1366+sprache)
    If FehlerGefunden = False Then
'        Print #DateiNummer, "kein Fehler gefunden"
'        txtfehleru.text = "kein Fehler gefunden"
        On Error Resume Next                                                                'Gerbing 02.09.2008
        'Print #DateiNummer, LoadResString(1382 + Sprache)
        oStream.WriteLine LoadResString(1382 + Sprache)

        On Error GoTo 0                                                                     'Gerbing 02.09.2008
        txtFehlerU.Text = LoadResString(1382 + Sprache)
    Else
        If LogNichtBenutzbar = False Then
            'txtFehlerU.Text = "Fehler siehe " & PruefLogFile
            txtFehlerU.Text = LoadResString(1383 + Sprache) & PruefLogFile
        Else
            'txtFehlerU.Text = PruefLogFile & "nicht benutzbar"
            txtFehlerU.Text = PruefLogFile & LoadResString(2277 + Sprache)
        End If
    End If
    'Close #DateiNummer
    On Error Resume Next
    oStream.Close
    On Error GoTo 0
    txtArbeitsfortschrittU.Text = LoadResString(1384 + Sprache)   'Pr�fen2 beendet
    Call ButtonsEnabled
End Sub


Private Sub btnPr�fen1_Click()
    '1.Pr�fe, ob die DateinamenErweiterung widerspr�chlich ist
    '2.DateinameKurz eintragen
    '3.Pr�fe ob die Datei im Ordner existiert
    '4.f�r Bilddateien oder Video Breite und H�he ermitteln

    Dim Gefunden As String
    Dim Msg As String
    Dim SQL As String
    Dim Erg As Long
    Dim SWF As String
    Dim pos As Long
    Dim pos1 As Long
    Dim start As Long
    Dim strTemp As Date
    Dim antwort As Long
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim MultiUser As Boolean
    Dim NutzerName As String
    Dim MM As New MovieModule
    Dim Error As Long
    Dim lngBreite As Long
    Dim lngH�he As Long
    Dim fs As New Scripting.FileSystemObject
    Dim f
    Dim blnMacheUpdate As Boolean

    If gblnSQLServerVersion = False Then
        'Gerbing 14.02.2011
        'Wer alles Nutzer der Datenbank ist, steht in einer Multi-Nutzer-Umgebung in der Datei fotos.ldb
        'diese gibts nur in einer Multiuser-Umgebung und wird von selbst gel�scht, wenn es nur noch einen Nutzer gibt
        Set cn = CreateObject("ADODB.Connection")                                       'Gerbing 23.11.2017
        cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AppPath & "\fotos.mdb"
        cn.mode = adModeReadWrite
        cn.Open cn.ConnectionString
        Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
        , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
        'Output the list of all users in the current database.
    
    '    Debug.Print Rs.Fields(0).Name, "", Rs.Fields(1).Name, _
    '    "", Rs.Fields(2).Name, Rs.Fields(3).Name
    
        MultiUser = False
        Do
    '        Debug.Print Rs.Fields(0), Rs.Fields(1), _
    '        Rs.Fields(2), Rs.Fields(3)
            If rs.EOF Then Exit Do
            NutzerName = Trim(rs.Fields(0))
            rs.Movenext
            If Not rs.EOF Then
                If NutzerName <> Trim(rs.Fields(0)) Then
                    MultiUser = True
                    Exit Do
                End If
            End If
        Loop
        If MultiUser = True Then
            rs.Close
            cn.Close
            Msg = AppPath & "\fotos.mdb" & vbNewLine
    '        msg = msg & "Pr�fen1 muss ausgef�hrt werden, wenn Sie der einzige Nutzer der Datenbank sind" & vbNewLine
    '        msg = msg & "Die Namen der anderen Nutzer finden Sie in der Datei " & AppPath & "\fotos.ldb"
            Msg = LoadResString(2279 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2280 + Sprache) & AppPath & "\fotos.ldb"
            'MsgBox Msg
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbInformation
            Exit Sub
        End If
        rs.Close                                                                                'Gerbing 18.02.2011
        cn.Close                                                                                'Gerbing 18.02.2011
    End If
    '----------------------------------------------------------------------------------------------------------------
    blnMessageAusgeben = True                                                               'Gerbing 26.01.2009
    NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.RemoveAll
    Call SpaltenbreiteMerken
    '---------------------------------------------------------------------------------------------------------
    SQL = "SELECT DatumBreiteHoehe From ErsterStart;"
    'SQL = "SELECT " & LoadResString(2528 + Sprache) & " From " & LoadResString(2527 + Sprache) & ";"
                                 'Gerbing 12.03.2005
    
    Set rstDBH = New ADODB.Recordset
    With rstDBH
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    '---------------------
    PixelAusrechnen.Show 1
    '---------------------
    If blnReturn = True Then                                                                'Gerbing 22.11.2008
        On Error Resume Next
        rstDBH.Close
        rstsql.Close
        On Error GoTo 0
        txtFehlerU.Text = ""
        Exit Sub                    'Der Nutzer hat gew�hlt Pr�fen1 abbrechen
    End If
    
    Call ButtonsDisabled
    btnPr�fen1Abbrechen = False
    'Pr�fenNummer = "Pr�fen1"
    Pr�fenNummer = LoadResString(1443 + Sprache)
    StartVerzeichnis = ""
    txtFehlerU.Text = ""
    FehlerGefunden = False
    '�ffne die Datei pruef.log
    On Error Resume Next
    oStream.Close                                                                           'Gerbing 06.11.2013
    DateiNummer = FreeFile  ' neue Datei-Nr.
    ERR = 0
    'Open PruefLogFile For Output As #DateiNummer
    'object.CreateTextFile(filename[, overwrite[, unicode]])
    Set oStream = PruefFso.CreateTextFile(PruefLogFile, True, True)
    If ERR <> 0 Then
        'Msg = "Die Datei " & PruefLogFile & " kann nicht ge�ffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & PruefLogFile & " " & LoadResString(1372 + Sprache) & NL
        'msg = msg & "Sie m�ssen f�r Schreibrechte sorgen, damit �nderungen an dieser Datei gemacht werden k�nnen." & NL
        Msg = Msg & LoadResString(2276 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number & NL & NL
        
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = Msg & LoadResString(1542 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)                                   'Gerbing 02.09.2008
        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
        If antwort = vbNo Then
            LogNichtBenutzbar = False
            Call SpaltenbreiteWiederherstellen
            Call ButtonsEnabled
            Exit Sub
        Else
            LogNichtBenutzbar = True
        End If
    End If
    On Error GoTo 0
    Msg = Now & "  "
    Msg = Msg & Pr�fenNummer & "  "
    If gblnSQLServerVersion = True Then
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & PublicSQLServer & PublicSQLDatabase
        Msg = Msg & LoadResString(1381 + Sprache) & PublicSQLServer & " " & PublicSQLDatabase
    Else
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & AppPath & "\fotos.mdb"
        Msg = Msg & LoadResString(1381 + Sprache) & AppPath & "\fotos.mdb"
    End If
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    Call SpaltenbreiteWiederherstellen
    DatensatzNr = 1
    Screen.MousePointer = vbHourglass
    '----------------------------------------------------Gerbing 14.02.2011-------------------------------------
    'SQL = "Select * From Fotos ORDER BY Dateiname"
    SQL = "Select * From Fotos ORDER BY " & LoadResString(1028 + Sprache)
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Source = SQL
        .Open
    End With
    '====================================================================================================
    Pr�fen1Abbrechen = False                                                        'Gerbing 06.11.2013
    btnPr�fen1Abbrechen.Enabled = True                                              'Gerbing 06.11.2013
    Do Until rstsql.EOF
        StartMillisek = timeGetTime
        If IsNull(rstsql.Fields(LoadResString(1028 + Sprache))) Then               'Gerbing 18.01.2006
            FehlerDateinameIsNull (DatensatzNr)
            'Datens�tze ohne Feld Dateiname werden sofort gel�scht
            rstsql.Delete
            GoTo MovenextadoRs 'Gerbing 18.01.2006
        End If
        'strFotoDatei = rstsql.Fields("Dateiname")
        strFotoDatei = rstsql.Fields(LoadResString(1028 + Sprache))
        If gblnSQLServerVersion = True Then
            strFotoDatei = Replace(strFotoDatei, "+:\", PublicLocationFotos & "\")
        Else
            strFotoDatei = Replace(strFotoDatei, "+:\", AppPath & "\")                   'Gerbing 11.04.2005
        End If
        If IsNull(rstsql.Fields(LoadResString(1023 + Sprache))) Then               'Gerbing 18.01.2006
            FehlerJahrIsNull (DatensatzNr)
            GoTo MovenextadoRs 'Gerbing 18.01.2006
        End If
        
        'If Not IsNull(rstsql.Fields("SWF")) Then                                   'Gerbing 11.12.2005
        If Not IsNull(rstsql.Fields(LoadResString(1029 + Sprache))) Then                                   'Gerbing 11.12.2005
            'SWF = rstsql.Fields("SWF")
            SWF = rstsql.Fields(LoadResString(1029 + Sprache))
        End If
        SWF = UCase(SWF)
        '---------------------------------------------
        '1.Pr�fe, ob die DateinamenErweiterung widerspr�chlich ist
        DateinamenErweiterung = UCase(Right(strFotoDatei, 3))
        Select Case DateinamenErweiterung
            Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF"         'Gerbing 09.03.2005
                'in Spalte SWF soll F oder SW stehen
                If SWF = "F" Or SWF = "SW" Or SWF = "C" Or SWF = "BW" Then      'Gerbing 09.01.2006
                    'kein Fehler
                Else
                    Call FehlerDateinamenErweiterungWiderspruch(DatensatzNr)
                End If
            Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"  'Gerbing 10.12.2017
                'in Spalte SWF soll FV oder SV stehen
                If SWF = "FV" Or SWF = "SV" Or SWF = "CV" Or SWF = "BV" Then
                    'kein Fehler
                Else
                    Call FehlerDateinamenErweiterungWiderspruch(DatensatzNr)
                End If
'            Case "HTM", "PDF", "XLS"
'                'kein Fehler
'            Case Else
'                Call FehlerDateinamenErweiterung(DatensatzNr)
            Case Else                                                           'Gerbing 11.12.2005
                'kein Fehler
        End Select
        EndMillisek = timeGetTime
'        Debug.Print "EndMillisec=" & EndMillisek
'        Debug.Print "Millisekunden f�r Schleifenanfang" & "=" & (EndMillisek - StartMillisek)
        '-----------------------------------------------
        '2.DateinameKurz eintragen                                            'Gerbing 24.06.2006 nach oben verlegt
        start = 1
        Do
            pos = InStr(start, strFotoDatei, "\")
            If pos = 0 Then Exit Do
            start = pos + 1
        Loop
        If gblnSchreibgesch�tzt = False Then                                'Gerbing 23.01.2007
            'An diesem rstsql.Edit kommt bei Multi-Nutzer-Umgebung und zwei Nutzer machen Pr�fen1 der Laufzeitfehler 3260
            'rstsql.Edit
        End If
        'rstsql.Fields("DateinameKurz") = Right(strFotoDatei, Len(strFotoDatei) - start + 1)
        If gblnSchreibgesch�tzt = False Then                                'Gerbing 23.01.2007
            rstsql.Fields(LoadResString(1031 + Sprache)) = Right(strFotoDatei, Len(strFotoDatei) - start + 1)
        End If
        '----------------------------------------------
        '3.Pr�fe ob die Datei im Ordner existiert                                           'Gerbing 22.02.2006
        'wird ver�ndert
        'anstelle von Open strFotoDatei For Input As #Pr�fDateiNummer benutze ich
        'FileDateTime(strFotoDatei)
        On Error Resume Next
        'strTemp = FileDateTime(strFotoDatei)
        Set f = fs.GetFile(strFotoDatei)                                        'Gerbing 04.03.2013
        strTemp = f.DateLastModified                                            'zB 03.06.2013 17:35:28

        If ERR <> 0 Or strTemp = "00:00:00" Then                                'Gerbing 04.03.2013
            If blnMitBH = False And blnNurNeue = False Then                     'Gerbing 06.09.2013
                '
            Else
                rstsql.Update                                                   'Gerbing 24.06.2006
            End If
            'das hei�t die Datei gibt es nicht oder sie hat ein ung�ltiges Datum
            On Error GoTo 0
            Call FehlerFotoDatei(DatensatzNr)    'schreibe den Fehler in PruefLogFile
            GoTo MovenextadoRs
        Else
            'DDatum eintragen
            On Error GoTo 0
            pos = InStr(1, strTemp, " ")                                        'Gerbing 04.11.2010
            If pos <> 0 Then
                strTemp = left(strTemp, pos - 1)
            End If                                                              'Gerbing 04.11.2010
            'rstsql.Fields("DDatum") = strTemp
            If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                'Wenn das DateLastModified der Datei aktueller ist, als das in der Datenbank eingetragene DDatum,
                'das tritt auf, wenn mit PSP X8 und h�her eine �nderung im Bild gemacht wurde zB Bildbegradigung(horizontal - vertikal),
                'dann setze ich das Feld IPTCPresent = False                    'Gerbing 26.03.2017
                If DateDiff("d", strTemp, rstsql.Fields(LoadResString(1032 + Sprache))) < 0 Then
                    rstsql.Fields("IPTCPresent") = False
                End If
                rstsql.Fields(LoadResString(1032 + Sprache)) = strTemp
            End If
        End If
        '----------------------------------------------
        If blnMitBH = True Then
            If blnNurNeue = False Then GoTo AlleAusrechnenRs
            'Neuberechnung von BreitePixel und HoehePixel nur f�r neuere Dateien als DatumBreiteHoehe
            'If rst.Fields("DDatum") >= rstDBH.Fields("DatumBreiteHoehe") Then   'Gerbing 12.03.2005
            If rstDBH.EOF Then GoTo AlleAusrechnenRs
            If rstsql.Fields(LoadResString(1032 + Sprache)) >= rstDBH.Fields("DatumBreiteHoehe") Then   'Gerbing 12.03.2005
                '-----------------------------------------------------------------------
                '4.f�r Bilddateien oder Video Breite und H�he ermitteln
                'bei Videos mciSendString siehe MovieModule.cls benutzen            'Gerbing 26.10.2011
                'bei Fotos Call LoadPicBox
AlleAusrechnenRs:
                pintBreite = 0                                                       'Gerbing 15.11.2019
                pintHoehe = 0
                lngVideoDuration = 0
                blnMediaPlayerStopped = False
                blnMediaPlayerError = False
                Select Case DateinamenErweiterung
                    Case "AVI", "MPG", "PEG", "MOV", "MPE", "ASF", "ASX", "WMV", "MP4", "MKV", "FLV"  'Gerbing 10.12.2017
                        Form1.WMP.settings.autoStart = False
                        Form1.WMP.Width = 1
                        Form1.WMP.URL = strFotoDatei
                        Form1.WMP.Visible = True     'erst nach Form1.WMP.URL = ...27.11.2016                                                             'Gerbing 01.09.2008
                        On Error Resume Next
                        ERR = 0
                        Form1.WMP.Controls.play
                        'jetzt muss ich warten bis 'player .playState=1(stopped) kommt
                        'bei Fehlern und wenn ich sage 'ja' bei 'soll der player versuchen den Inhalt wiederzugeben' gibt es keinen Loop
                        'bei Fehlern und wenn ich sage 'nein' bei 'soll der player versuchen den Inhalt wiederzugeben' gibt es einen Loop
                        'nach einer Sekunde beende ich den Loop
                        glngStartMillisek = timeGetTime
                        Do
                            glngEndMillisek = timeGetTime
                            If glngEndMillisek - glngStartMillisek > 1000 Then Exit Do
                            If blnMediaPlayerStopped = True Then Exit Do
                            If blnMediaPlayerError = True Then Exit Do
                            DoEvents
                        Loop
                Case Else
                    'Call LoadPicBox(strFotoDatei, Form1.Pr�fPicture) 'Gerbing 01.09.2007
                    Call LoadPicBox(strFotoDatei) 'Gerbing 01.09.2007
                    pintBreite = gsngPicWidth
                    pintHoehe = gsngPicHeight
                End Select
                On Error Resume Next                                'Gerbing 24.10.2007
                If lngVideoDuration = 0 Then                           'Gerbing 24.10.2007
                    If gblnSchreibgesch�tzt = False Then
                        rstsql.Fields("VideoDuration") = Null
                    End If
                Else
                    If gblnSchreibgesch�tzt = False Then
                        rstsql.Fields("VideoDuration") = lngVideoDuration
                    End If
                End If
                On Error GoTo 0
                'rstsql.Fields("BreitePixel") = intBreite
                If pintBreite = 0 Then
                    If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                        rstsql.Fields(LoadResString(1106 + Sprache)) = Null    'Gerbing 19.01.2006 Dann kann man dieses Feld in fotos.exe manuell editieren
                    End If
                Else
                    If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                        If pintBreite = rstsql.Fields(LoadResString(1106 + Sprache)) Then    'Gerbing 24.12.2013
                            blnMacheUpdate = False
                        Else
                            rstsql.Fields(LoadResString(1106 + Sprache)) = pintBreite
                            blnMacheUpdate = True
                        End If
                    End If
                End If
                'rstsql.Fields("HoehePixel") = intHoehe
                If pintHoehe = 0 Then
                    If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                        rstsql.Fields(LoadResString(1107 + Sprache)) = Null    'Gerbing 19.01.2006 Dann kann man dieses Feld in fotos.exe manuell editieren
                    End If
                Else
                    If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                        If pintHoehe = rstsql.Fields(LoadResString(1107 + Sprache)) Then     'Gerbing 24.1.22013
                            blnMacheUpdate = False
                        Else
                            rstsql.Fields(LoadResString(1107 + Sprache)) = pintHoehe
                            blnMacheUpdate = True
                        End If
                    End If
                End If
            End If
        End If
        If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
            If blnMitBH = False And blnNurNeue = False Then                     'Gerbing 06.09.2013
                '
            Else

                If blnMacheUpdate = True Then                                   'Gerbing 24.12.2013
                    ERR.Number = 0
                    On Error Resume Next
                    rstsql.Update
                    If ERR.Number <> 0 Then
                        rstsql.CancelUpdate                                     'Gerbing 06.09.2013
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
        '----------------------------------------------
MovenextadoRs:
        On Error GoTo ErrorMoveNext                                             'Gerbing 06.09.2013
        rstsql.Movenext
        DatensatzNr = DatensatzNr + 1
        'txtArbeitsfortschrittU.Text = "DatensatzNr." & DatensatzNr
        txtArbeitsfortschrittU.Text = LoadResString(1008 + Sprache) & " " & DatensatzNr
        'Erg = DatensatzNr Mod 100
        Erg = DatensatzNr Mod 10
        If Erg = 0 Then
            DoEvents
        End If
        EndMillisek = timeGetTime
'        Debug.Print "EndMillisec=" & EndMillisek
'        Debug.Print "Millisekunden bis Ende der Schleife" & "=" & (EndMillisek - StartMillisek)
        If Pr�fen1Abbrechen = True Then Exit Do
    Loop
    '===========================================================================================
    Screen.MousePointer = vbDefault
    'schlie�e die Datei pruef.log
    If FehlerGefunden = False Then
'        Print #DateiNummer, "kein Fehler gefunden"
'        txtfehleru.text = "kein Fehler gefunden"
        On Error Resume Next                                                                'Gerbing 02.09.2008
        'Print #DateiNummer, LoadResString(1382 + Sprache)
        oStream.WriteLine LoadResString(1382 + Sprache)

        On Error GoTo 0                                                                     'Gerbing 02.09.2008
        txtFehlerU.Text = LoadResString(1382 + Sprache)
    Else
        If LogNichtBenutzbar = False Then
            'txtFehlerU.Text = "Fehler siehe " & AppPath & "\pruef.log"
            txtFehlerU.Text = LoadResString(1383 + Sprache) & PruefLogFile
        Else
            'txtFehlerU.Text = AppPath & "\pruef.log" & "nicht benutzbar"
            txtFehlerU.Text = PruefLogFile & LoadResString(2277 + Sprache)
        End If
    End If
    'Close #DateiNummer
    On Error Resume Next
    oStream.Close
    On Error GoTo 0
    'txtArbeitsfortschritt.Text = "Pr�fen1 beendet"
    txtArbeitsfortschrittU.Text = LoadResString(1396 + Sprache)
    rstsql.Close
    'Datum des Vorgangs Pr�fen1 mit Berechnung von BreitePixel und HoehePixel in
    'ErsterStart.DatumBreiteHoehe eintragen
    If blnMitBH = True Then                                    'Gerbing 12.03.2005
        SQL = "SELECT DatumBreiteHoehe From ErsterStart;"
        'SQL = "SELECT " & LoadResString(2528 + Sprache) & " From " & LoadResString(2527 + Sprache) & ";"
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .Source = SQL
            '     .CacheSize = 2
            .Open
        End With
        
        If rstsql.EOF Then
            If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                rstsql.AddNew
            End If
        Else
            If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
                'adoRs.Edit
            End If
        End If
        If gblnSchreibgesch�tzt = False Then 'Gerbing 23.01.2007
            rstsql.Fields("DatumBreiteHoehe") = Date
            rstsql.Update
        End If
        rstsql.Close
    End If
    Call SpaltenbreiteWiederherstellen
    Call ButtonsEnabled
    btnPr�fen1Abbrechen.Enabled = False                         'Gerbing 08.04.2017
    Exit Sub
ErrorMoveNext:                                                  'Gerbing 06.09.2013
    rstsql.CancelUpdate
    rstsql.Movenext
    Resume Next
End Sub


Private Sub btnPr�fenS_Click()
    Dim SQL As String
    Dim Msg As String
    Dim start As Long
    Dim pos As Long
    Dim Dateiname As String
    Dim Killname As String
    Dim antwort As Long
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim MultiUser As Boolean
    Dim NutzerName As String
    Dim rc As Boolean
    Dim MyAppPath As String
    
    If gblnSQLServerVersion = True Then                     'Gerbing 05.09.2013
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    On Error GoTo 0
    If gblnSQLServerVersion = False Then
        'Gerbing 14.02.2011
        'Wer alles Nutzer der Datenbank ist, steht in einer Multi-Nutzer-Umgebung in der Datei fotos.ldb
        'diese gibts nur in einer Multiuser-Umgebung und wird von selbst gel�scht, wenn es nur noch einen Nutzer gibt
        Set cn = CreateObject("ADODB.Connection")                                       'Gerbing 23.11.2017
        cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AppPath & "\fotos.mdb"
        cn.mode = adModeReadWrite
        cn.Open cn.ConnectionString
        Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
        , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
        
        'Output the list of all users in the current database.
    
    '    Debug.Print Rs.Fields(0).Name, "", Rs.Fields(1).Name, _
    '    "", Rs.Fields(2).Name, Rs.Fields(3).Name
    
        MultiUser = False
        Do
    '        Debug.Print Rs.Fields(0), Rs.Fields(1), _
    '        Rs.Fields(2), Rs.Fields(3)
            If rs.EOF Then Exit Do
            NutzerName = Trim(rs.Fields(0))
            rs.Movenext
            If Not rs.EOF Then
                If NutzerName <> Trim(rs.Fields(0)) Then
                    MultiUser = True
                    Exit Do
                End If
            End If
        Loop
        If MultiUser = True Then
            rs.Close
            cn.Close
            Msg = AppPath & "\fotos.mdb" & vbNewLine
    '        msg = msg & "Pr�fenS muss ausgef�hrt werden, wenn Sie der einzige Nutzer der Datenbank sind" & vbNewLine
    '        msg = msg & "Die Namen der anderen Nutzer finden Sie in der Datei " & AppPath & "\fotos.ldb"
            Msg = LoadResString(2281 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2280 + Sprache) & AppPath & "\fotos.ldb"
            MsgBox Msg
            Exit Sub
        End If
        rs.Close                                                                                'Gerbing 18.02.2011
        cn.Close                                                                                'Gerbing 18.02.2011
    End If
'----------------------------------------------------------------------------------------------------------------
    If gblnSchreibgesch�tzt = True Then                             'Gerbing 23.01.2007
        'msg = "Bei einer schreibgesch�tzten Datenbank ist diese Funktion nicht m�glich"
        Msg = LoadResString(2421 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
'----------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    SQL = "SELECT AudioFileExists From Fotos;"
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .ActiveConnection = DBsql
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .Source = SQL
        '     .CacheSize = 2
        .Open
    End With
    If ERR.Number <> 0 Then
        On Error GoTo 0
        Msg = LoadResString(2227 + Sprache) & vbNewLine 'Seit Version 13.0.0.3 wird die Spalte AudioFileExists ben�tigt.
        Msg = Msg & LoadResString(2247 + Sprache)     'F�hren Sie fotos.exe aus, um diese Spalte in der Datenbank zu erzeugen
        MsgBox Msg
        Exit Sub
    End If
    On Error GoTo 0
    If rstsql.EOF Then
        'MsgBox "Die Datenbank ist leer"
        MsgBox LoadResString(1840 + Sprache)
        Exit Sub
    End If
'----------------------------------------------------------------------------------------------------------------
    '�ffne die Datei pruef.log                              'Gerbing 23.08.2007
    On Error Resume Next
    oStream.Close
    DateiNummer = FreeFile  ' neue Datei-Nr.
    ERR = 0
    'Open PruefLogFile For Output As #DateiNummer
    'object.CreateTextFile(filename[, overwrite[, unicode]])
    Set oStream = PruefFso.CreateTextFile(PruefLogFile, True, True)
    If ERR <> 0 Then
        'Msg = "Die Datei " & PruefLogFile & " kann nicht ge�ffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & PruefLogFile & " " & LoadResString(1372 + Sprache) & NL
        'msg = msg & "Sie m�ssen f�r Schreibrechte sorgen, damit �nderungen an dieser Datei gemacht werden k�nnen." & NL
        Msg = Msg & LoadResString(2276 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number & NL & NL
        
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = Msg & LoadResString(1542 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)                                   'Gerbing 02.09.2008
        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
        If antwort = vbNo Then
            LogNichtBenutzbar = False
            txtFehlerU.Text = ""
            Exit Sub
        Else
            LogNichtBenutzbar = True
        End If
    End If
'----------------------------------------------------------------------------------------------------------------
    On Error GoTo 0
    Pr�fenNummer = "Pr�fenS"
    Msg = Now & "  "
    Msg = Msg & Pr�fenNummer & "  "
    If gblnSQLServerVersion = True Then
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & PublicSQLServer & PublicSQLDatabase
        Msg = Msg & LoadResString(1381 + Sprache) & PublicSQLServer & " " & PublicSQLDatabase
    Else
        'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & AppPath & "\fotos.mdb"
        Msg = Msg & LoadResString(1381 + Sprache) & AppPath & "\fotos.mdb"
    End If
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008

    Call SpaltenbreiteMerken
    Call ButtonsDisabled
    Pr�fenSAbbrechen = False
    Pr�fenNummer = LoadResString(1468 + Sprache)
    txtFehlerU.Text = ""
    FehlerGefunden = False
    If gblnSQLServerVersion = True Then
        StartVerzeichnis = PublicLocationFotos
    Else
        StartVerzeichnis = AppPath                                   'Gerbing 11.04.2005
    End If
    
    Screen.MousePointer = vbHourglass
    btnPr�fenSAbbrechen.Enabled = True
    If gblnSQLServerVersion = True Then
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 29.12.2011
        'beim SQL Server muss es hei�en 'Delete from table
        SQL = "DELETE From Temp_Haken"
        'SQL = "DELETE FROM " & LoadResString(2523 + Sprache)
    Else
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 30.09.2004
        SQL = "DELETE " & "Temp_Haken.* "
        SQL = SQL & " FROM " & "Temp_Haken;"
        'SQL = "DELETE " & LoadResString(2523 + Sprache) & ".* "
        'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    End If
    DBsql.Execute SQL
    'dann leeres Recordset rstTempHaken �ffnen
    SQL = " SELECT " & "Temp_Haken.*"
    SQL = SQL & " FROM " & "Temp_Haken;"
    'SQL = " SELECT " & LoadResString(2523 + Sprache) & ".*"
    'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    Set rstTempHaken = New ADODB.Recordset
    With rstTempHaken
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Source = SQL
        .LockType = adLockOptimistic
        .Open
    End With
    '-------------------------------------------------------------------------------------------------------
    gblnSubdirectories = True
    Call RekursiveAudio(StartVerzeichnis, "*.*") 'f�lle die Tabelle Temp_Haken mit s�mtlichen wirklichen Audio-Dateinamen einschlie�lich subdirectories
    '-------------------------------------------------------------------------------------------------------
    If Pr�fenSAbbrechen = True Then                                 'Gerbing 04.10.2004
        Call ButtonsEnabled
        Exit Sub
    End If
    '------------------------------------------------------------------------------------
    '1.�ndere in allen Datens�tzen der Tabelle Fotos das Feld AudioFileExists auf 'nein'
    SQL = "UPDATE Fotos SET Fotos.AudioFileExists = 0;"
    DBsql.Execute SQL
    SQL = "SELECT * FROM Temp_Haken;"
    'SQL = "SELECT * FROM " & LoadResString(2523 + Sprache) & ";"
    On Error Resume Next
    rstTempHaken.Close
    On Error GoTo 0
    With rstTempHaken
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Source = SQL
        .LockType = adLockOptimistic
        .Open
    End With
    FehlerGefunden = False
    '-----------------------------------------------------------------------------------
    Do Until rstTempHaken.EOF
        '2.�ndere in allen Datens�tzen der Tabelle Fotos das Feld AudioFileExists auf 'ja', wo ein Satz
        'in Temp_Haken steht
        '3.wenn kein Satz gefunden wird, muss die Audio-Datei gel�scht werden
        FehlerGefunden = True
        'Der Dateiname wird ermittelt durch Suchen ab rechtem Rand bis zum Punkt
        'start = Len(rstTempHaken.Fields("Dateiname")) - 2
        start = Len(rstTempHaken.Fields(LoadResString(1028))) - 2
        Do
            pos = InStr(start, rstTempHaken.Fields(LoadResString(1028)), ".")
            If pos <> 0 Then
                Dateiname = Mid(rstTempHaken.Fields(LoadResString(1028)), 1, pos - 1)
                Exit Do
            End If
            start = start - 1
        Loop
        Dateiname = Replace(Dateiname, MyAppPath, "+:")                                 'Gerbing 05.09.2013
        'Dateinamen mit Hochkomma m�ssen durch 2 Hochkommas ersetzt werden              'Gerbing 23.01.2018
        Dateiname = Replace(Dateiname, "'", "''")                                       'Gerbing 23.01.2018
        'SQL = "SELECT Dateiname, AudioFileExists"
        'SQL = SQL & " From Fotos"
        'SQL = SQL & " WHERE Dateiname Like '+:\2005\Musterfoto01.*';"
        SQL = "SELECT " & LoadResString(1028 + Sprache) & ", AudioFileExists"
        SQL = SQL & " From Fotos"
        SQL = SQL & " WHERE " & LoadResString(1028 + Sprache) & " Like '" & Dateiname & ".%';"
        On Error Resume Next
        rstsql.Close
        On Error GoTo 0
        With rstsql
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .Source = SQL
            .LockType = adLockOptimistic
            .Open
        End With

        'Wenn ein Satz gefunden wird, muss dort AudioFileExists = yes eingetragen werden,
        'wenn kein Satz gefunden wird, muss die Audio-Datei gel�scht werden
        If Not rstsql.EOF Then
            'rstsql.Edit
            rstsql.Fields("AudioFileExists") = vbYes
            rstsql.Update
            'AudioFileExists = yes eingetragen
            Call Pr�fenSKommentar1(rstTempHaken.Fields("Dateiname"))                       'Gerbing 23.08.2007 bei Temp_Haken gibt es keine englische Version
            'Call Pr�fenSKommentar1(rstTempHaken.Fields(LoadResString(1028 + Sprache)))      'Gerbing 23.08.2007
        Else
            Killname = Replace(rstTempHaken.Fields("Dateiname"), "+:\", MyAppPath & "\")  'Gerbing 05.09.2013
            On Error Resume Next
            'Kill Killname
            rc = file_delete(Killname, , True)                                              'Gerbing 05.09.2013
            If rc = False Then
                'msg = "Fehler beim L�schen von" & " " & temp & NL
                Msg = LoadResString(2420 + Sprache) & NL
                Msg = Msg & " " & Killname & NL
                'MsgBox msg
                On Error Resume Next                                                        'Gerbing 02.09.2008
                'Print #DateiNummer, Msg                         'Gerbing 23.08.2007
                oStream.WriteLine Msg
                On Error GoTo 0                                                             'Gerbing 02.09.2008
            Else
                'Audio-Datei gel�scht
                Call Pr�fenSKommentar2(rstTempHaken.Fields("Dateiname"))                   'Gerbing 23.08.2007 bei Temp_Haken gibt es keine englische Version
                'Call Pr�fenSKommentar2(rstTempHaken.Fields(LoadResString(1028 + Sprache)))  'Gerbing 23.08.2007
            End If
            On Error GoTo 0
            rstsql.Close                                                   'Gerbing 10.08.2006
        End If
        rstTempHaken.Movenext
    Loop
    rstTempHaken.Close                                              'Gerbing 10.08.2006
    '-------------------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    'schlie�e die Datei pruef.log
    If FehlerGefunden = False Then
'        txtFehlerU.Text = "kein Fehler gefunden"
        txtFehlerU.Text = LoadResString(1382 + Sprache)
    Else
        'txtFehlerU.Text = "Differenzen wurden bereinigt"
        txtFehlerU.Text = LoadResString(1470 + Sprache)
    End If
    'Close #DateiNummer
    oStream.Close
    'txtArbeitsfortschrittU.Text = "Pr�fenS beendet"
    txtArbeitsfortschrittU.Text = LoadResString(1471 + Sprache)
    btnPr�fenSAbbrechen.Enabled = False
    Call ButtonsEnabled
    '----------------------------------------------------------------------------------------------------
    'wie Button Reset                                               'Gerbing 04.02.2008
    Call SpaltenbreiteMerken
    
    rsDataGrid.Requery
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
    
    Call SpaltenbreiteWiederherstellen
    DBGridNeu.Caption = PublicDatagridCaption
    DBGridNeu.AllowUpdate = False
End Sub

Public Sub btnReset_Click()
    Dim SQL As String
    Dim Msg As String
    Dim antwort As Long
    
    Call SpaltenbreiteMerken
    
    rsDataGrid.Requery
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
    
    Call SpaltenbreiteWiederherstellen
    txtFehlerU.Text = ""
    txtArbeitsfortschrittU.Text = ""
    DBGridNeu.Caption = PublicDatagridCaption
    DBGridNeu.AllowUpdate = False
    
    On Error Resume Next                                                                    'Gerbing 06.11.2013
    oStream.Close                                                                           'Gerbing 06.11.2013
    On Error GoTo 0
End Sub

Private Sub btnHilfe_Click()
    Dim retval As Long
    Dim CHMFile As String
    Dim Msg As String

    If Sprache = 0 Then                             'Gerbing 08.11.2005
        CHMFile = AppPath & "\Help\Deutsch\fotosmdb.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht �ffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING fotosmdb"), vbInformation
            Exit Sub
        Else
            retval = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If retval <= 32 Then
                Call HelpFileErrorMsg(retval, CHMFile)
            End If
        End If
    Else
        CHMFile = AppPath & "\Help\English\fotosmdb.CHM"                           'Gerbing 14.03.2007
        If isUnicodeString(CHMFile) = True Then
            'Msg = "CHM-Help-Dateien lassen sich im Unicode-Pfad nicht �ffnen, das hat Microsoft nicht vorgesehen" & vbNewLine
            'Msg = Msg & "Kopieren Sie die CHM-Help-Dateien in einen Pfad ohne Unicode-Zeichen"
            Msg = CHMFile & vbNewLine
            Msg = Msg & LoadResString(2544 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(2545 + Sprache)
            MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING fotosmdb"), vbInformation
            Exit Sub
        Else
            retval = RunShellExecute(Me.hWnd, "open", CHMFile, vbNull, vbNull, 1)
            If retval <= 32 Then
                Call HelpFileErrorMsg(retval, CHMFile)
            End If
        End If
    End If
End Sub

Private Sub btnPr�fen3Abbrechen_Click()
    Dim Mldg, Stil, Titel, Hilfe, Ktxt, antwort, Text1
    'Mldg = "M�chten Sie Pr�fen3 Abbrechen ?"       ' Meldung definieren.
    Mldg = LoadResString(1457 + Sprache)            ' Meldung definieren.
    Stil = vbYesNo + vbDefaultButton2
    'Titel = "Pr�fen3 Abbrechen"                 ' Titel definieren.
    Titel = LoadResString(1458 + Sprache)                 ' Titel definieren.
    antwort = MsgBox(Mldg, Stil, Titel)
    If antwort = vbYes Then
        Pr�fen3Abbrechen = True
        Screen.MousePointer = vbDefault
        btnPr�fen3Abbrechen.Enabled = False
    Else
        Pr�fen3Abbrechen = False
    End If
End Sub

Private Sub btnPr�fenSAbbrechen_Click()
    Dim Mldg, Stil, Titel, Hilfe, Ktxt, antwort, Text1
    'Mldg = "M�chten Sie Pr�fenS Abbrechen ?"       ' Meldung definieren.
    Mldg = LoadResString(1464 + Sprache)            ' Meldung definieren.
    Stil = vbYesNo + vbDefaultButton2
    'Titel = "Pr�fenS Abbrechen"                 ' Titel definieren.
    Titel = LoadResString(1465 + Sprache)                 ' Titel definieren.
    antwort = MsgBox(Mldg, Stil, Titel)
    If antwort = vbYes Then
        Pr�fenSAbbrechen = True
        Screen.MousePointer = vbDefault
        btnPr�fenSAbbrechen.Enabled = False
    Else
        Pr�fenSAbbrechen = False
    End If
End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim Datei As String
    Dim SQL As String
    Dim antwort As Long
    Dim temp As Long
    Dim Feldname As String
    Dim strTemp As String
    Dim SystemDirectory As String
    Dim cmdline As String
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim strVersion As String
    Dim errLoop As Error
    Dim fso As New Scripting.FileSystemObject
    Dim oFile As Scripting.File
    Dim j As Long
    Dim rstUserDef As ADODB.Recordset                           'Gerbing 02.08.2016
    Dim rstDefFields As ADODB.Recordset
    Dim rstErsterStart As ADODB.Recordset                       'Gerbing 05.09.2016
    Dim rstTemp As ADODB.Recordset                              'Gerbing 05.09.2016

    init_global
    'AppPath = App.Path                                         'Gerbing 16.04.2005
    AppPath = getCurrentDir
    If Right(AppPath, 1) = "\" Then
        AppPath = left(AppPath, Len(AppPath) - 1)
    End If

    Call RekursiveTempThumbs(AppPath & "\TempThumbs", "*.*")         'Gerbing 06.04.2017

    gblnProversion = True                                       'Gerbing 04.03.2012
'    #If Proversion = 0 Then
'        gblnProversion = False
'    #End If
    On Error Resume Next
    'gdtDatumFotosMdb = FileDateTime(AppPath & "\fotos.mdb")
    Set oFile = fso.GetFile(AppPath & "\fotos.mdb")                                  'Gerbing 04.03.2013
    gdtDatumFotosMdb = oFile.DateLastModified

    On Error GoTo 0
    Call ReadFotosIniFile
'----------------------------------------------------------------------------------------------------------------
    Call AnpassenNutzerWunsch(Me)                                   'Gerbing 11.03.2017
    Call AnpassenHeadFont(DBGridNeu)                                'Gerbing 23.06.2011
    
    If (GetAsyncKeyState(VK_SHIFT) = KeyPressed) Then               'Gerbing 10.09.2013
        Call FremdeFotosMdb                                         'Gerbing 10.09.2013
    End If
    cmdline = command()                                             'Gerbing 07.11.2011
    gblnCommandLineEmpty = False
    If cmdline = "" Then
        gblnCommandLineEmpty = True
    End If
    'fotosmdblocation=...;                      'zB StandortFotos=H:\FOTOS\GG;sqlservername=GOTTFRIED;datenbankname=GG;WindowsAuthentication=0;
    'sqlservername=...;
    'datenbankname=...;
    'WindowsAuthentication=0; hei�t nein
    'WindowsAuthentication=1; hei�t ja
    'username=...;
    'Password=...;
    'StandortFotos=...;
    
    pos = InStr(1, cmdline, "fotosmdblocation=", vbTextCompare)     'zB fotosmdblocation=H:\FOTOS\GG;
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            'AppPath wird mit Command �bergeben
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            AppPath = strTemp
        End If
    End If
    pos = InStr(1, cmdline, "sqlservername=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            PublicSQLServer = strTemp
        End If
    End If
    pos = InStr(1, cmdline, "datenbankname=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            PublicSQLDatabase = strTemp
        End If
    End If
    pos = InStr(1, cmdline, "WindowsAuthentication=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            If strTemp = "0" Then
                PublicWindowsAuthentication = "0"
            Else
                PublicWindowsAuthentication = "1"
            End If
        End If
    End If
    pos = InStr(1, cmdline, "username=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            PublicSQLServerUserName = strTemp
        End If
    End If
    pos = InStr(1, cmdline, "Password=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            PublicSQLServerPassword = strTemp
        End If
    End If
    pos = InStr(1, cmdline, "StandortFotos=", vbTextCompare)
    If pos <> 0 Then
        pos1 = InStr(pos, cmdline, "=", vbTextCompare)
        pos2 = InStr(pos, cmdline, ";", vbTextCompare)
        If pos1 <> 0 Then
            strTemp = Mid(cmdline, pos1 + 1, pos2 - pos1 - 1)
            PublicLocationFotos = strTemp
        End If
    End If
    '------------------------------------------------------------------------------------------------------
    If gblnProversion = True Then                                               'Gerbing 04.03.2012
        'Untersuche ob Access-Version oder SQL-Server-Version
        'strtemp = Dir(AppPath & "\Fotos.mdb")
        If file_path_exist(AppPath & "\Fotos.mdb") = False Then
        'If strtemp = "" Then
            gblnSQLServerVersion = True
            btnL�scheInhaltFotosMdb.Visible = False                             'Gerbing 06.11.2017
        End If
    Else
        gblnSQLServerVersion = False
    End If

    'Sprache festlegen bei jedem Start. Wenn die Datenbank fotos.mdb englisch ist, dann Sprache=3000
    '------------------------------------------------------------------------------------------------------
CallSpracheFestlegen:
    Call SpracheFestlegen                                                               'Gerbing 18.02.2011
    PruefLogFile = gstrFotosIniAnwendungsOrdner & "\" & LoadResString(1366 + Sprache)      'Pfad der pruef.log                 Gerbing 17.02.2011
    If PublicLanguage = "" Then
        Sprache = 0                     '0=deutsch
    End If
    If PublicLanguage = "0" Then
        Sprache = 0                     '0=deutsch
    End If
    If PublicLanguage = "1" Then
        Sprache = 3000                  '3000=englisch
    End If
    If gblnSQLServerVersion = False Then
        'Programm beenden, wenn es Fotos.mdb nicht gibt                                     'Gerbing 03.11.2010
        'Datei = Dir(AppPath & "\fotos.mdb", vbNormal + vbHidden + vbReadOnly)
        If file_path_exist(AppPath & "\fotos.mdb") = False Then
        'If Datei = "" Then
    '       msg = "Die Datei " & AppPath & "\fotos.mdb " & vbNewLine
    '       msg = msg & "wurde nicht gefunden."
    '       MsgBox msg, , "Das Programm wird beendet"
            Msg = LoadResString(2145 + Sprache) & " " & AppPath & "\fotos.mdb " & vbNewLine
            Msg = Msg & LoadResString(2413 + Sprache)
            MsgBox Msg, , LoadResString(2139 + Sprache)
            End
        End If
    End If
    On Error GoTo ERR
    'Gerbing 23.11.2017
    Set rstsql = New ADODB.Recordset
    rstsql.Open "SELECT * FROM Fotos", DBsql, _
        adOpenStatic, adLockOptimistic
'----------------------------------------------------------------------------------------------------------
    'Seit Version 14.2.1 gibt es die Tabellen UserDefined und DefaultFields         'Gerbing 02.08.2016
    'die erzeugt das Programm selbst
    On Error Resume Next
    SQL = "select * From UserDefined;"
    Set rstUserDef = New ADODB.Recordset
    With rstUserDef
        .Source = SQL
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then
        'hier existiert die Tabelle UserDefined nicht
        If gblnSchreibgesch�tzt = False Then
            'also wird Tabelle UserDefined erzeugt
            DBsql.Execute _
                "CREATE TABLE UserDefined(" & _
                    "FieldName1 VARCHAR(255) NOT NULL," & _
                    "SourceField1   VARCHAR(255)  NOT NULL," & _
                    "FieldName2 VARCHAR(255) NOT NULL," & _
                    "SourceField2   VARCHAR(255)  NOT NULL," & _
                    "FieldName3 VARCHAR(255) NOT NULL," & _
                    "SourceField3   VARCHAR(255)  NOT NULL," & _
                    "FieldName4 VARCHAR(255) NOT NULL," & _
                    "SourceField4   VARCHAR(255)  NOT NULL," & _
                    "FieldName5 VARCHAR(255) NOT NULL," & _
                    "SourceField5    VARCHAR(255)  NOT NULL)"
        End If
    End If
    rstUserDef.Close
    SQL = "select * From DefaultFields;"
    Set rstDefFields = New ADODB.Recordset
    With rstDefFields
        .Source = SQL
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    If ERR.Number <> 0 Then
        'hier existiert die Tabelle rstDefFields nicht
        If gblnSchreibgesch�tzt = False Then
            'also wird Tabelle rstDefFields erzeugt
                ' Create the DefaultFields table.
                DBsql.Execute _
                    "CREATE TABLE DefaultFields(" & _
                        "SituationSource VARCHAR(255) NOT NULL," & _
                        "LocationSource   VARCHAR(255)  NOT NULL," & _
                        "CountrySource VARCHAR(255) NOT NULL," & _
                        "PeopleSource   VARCHAR(255)  NOT NULL," & _
                        "BWCSource VARCHAR(255) NOT NULL," & _
                        "CommentSource    VARCHAR(255)  NOT NULL)"
        End If
    End If
    rstDefFields.Close
    '----------------------------------------------------------------------------------------------------------
    'Seit Version 14.2.2 gibt es in der Tabelle ErsterStart das Feld LetzterGEOPunkt und ZoomListIndex        'Gerbing 05.09.2016
    'die erzeugt das Programm selbst, wenn es die Professional Version ist
    If gblnProversion = True Then
        On Error Resume Next
        SQL = "select LetzterGEOPunkt From ErsterStart;"
        Set rstErsterStart = New ADODB.Recordset
        With rstErsterStart
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        If ERR.Number <> 0 Then
            'hier existiert das Feld LetzterGEOPunkt nicht
            rstErsterStart.Close
            If gblnSchreibgesch�tzt = False Then
                If gblnSQLServerVersion = True Then
                    'SQL Server
                    DBsql.Execute _
                        "ALTER TABLE ErsterStart ADD LetzterGEOPunkt VARCHAR(255)"          'es hei�t ADD und nicht ADD COLUMN
                    DBsql.Execute _
                        "ALTER TABLE ErsterStart ADD ZoomListIndex VARCHAR(255)"
                Else
                    'Access Version
                    'also wird Feld LetzterGEOPunkt und ZoomListIndex erzeugt
                    DBsql.Execute _
                        "ALTER TABLE ErsterStart ADD COLUMN LetzterGEOPunkt TEXT"
                    DBsql.Execute _
                        "ALTER TABLE ErsterStart ADD COLUMN ZoomListIndex TEXT"
                End If
            End If
        End If
        rstErsterStart.Close
        On Error GoTo 0
    End If
'------------------------------------------------------------------------------------Gerbing 05.9.2016--------
    btnL�scheInhaltFotosMdb.Caption = LoadResString(1509 + Sprache) 'L�sche den Inhalt von fotos.mdb
    btnGenerieren.Caption = LoadResString(1311 + Sprache) '&Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)...
    btnNutzerdefinierteFelderAnlegen.Caption = LoadResString(1312 + Sprache)    'Nutzerdefiniertes Datenbank-Feld anlegen...
    Frame2.Caption = LoadResString(1313 + Sprache)          'Datenbank fotos.mdb auf g�ltigen Inhalt pr�fen
    Frame3.Caption = LoadResString(1314 + Sprache)          'Umkehr-Probe machen
    FrameEXIFIPTC.Caption = LoadResString(1525 + Sprache)   'EXIF/IPTC zur�ckschreiben
    btnPr�fen1.Caption = LoadResString(1315 + Sprache)      'Pr�fen&1
    btnPr�fen2.Caption = LoadResString(1316 + Sprache)      'Pr�fen&2
    btnPr�fen3.Caption = LoadResString(1321 + Sprache)      'Pr�fen&3
    btnPr�fenS.Caption = LoadResString(1466 + Sprache)      'Pr�fen&S
    btnReset.Caption = LoadResString(1322 + Sprache)        '&Reset
    btnPr�fen3Abbrechen.Caption = LoadResString(1325 + Sprache)     'Abbru&ch
    btnPr�fenSAbbrechen.Caption = LoadResString(1467 + Sprache)     'A&bbruch
    Label8.Caption = LoadResString(1328 + Sprache)      'Pr�fergebnis:
    Label7.Caption = LoadResString(1329 + Sprache)      'Arbeitsfortschritt:
    btnHilfe.Caption = LoadResString(1326 + Sprache)        'Hil&fe
    btnBeenden.Caption = LoadResString(1327 + Sprache)      '&Beenden
    btnPr�fen1.ToolTipText = LoadResString(1420 + Sprache)  'ob jede im Feld Dateiname eingetragene Foto-Datei  wirklich existiert.
    btnPr�fen2.ToolTipText = LoadResString(1421 + Sprache)  'ob die Jahreszahl im Feld 'Jahr' und im Dateiname �bereinstimmt
    btnPr�fen3.ToolTipText = LoadResString(1422 + Sprache)  'ob zu allen im AppPath-Ordner und seinen Unter-Ordnern abgelegten Bildern auch ein Eintrag  in der Datenbank fotos.mdb enthalten ist.
    btnPr�fenS.ToolTipText = LoadResString(1469 + Sprache)  'ob es Differenzen zwischen vorhandenen Audio-Kommentaren und der Spalte 'AudioFileExists' gibt
    btnReset.ToolTipText = LoadResString(1427 + Sprache)    'zur�ck zum Inhalt von fotos.mdb
    btnEXIFIPTC.ToolTipText = LoadResString(1526 + Sprache) 'Sie k�nnen festlegen, welche Datenbankfelder in die IPTC-Felder von JPG-Fotos �bertragen werden sollen
    btn�ffnePruefLog.Caption = LoadResString(1508 + Sprache)  '�ffne die Datei pruef.&log
    Label8.ToolTipText = LoadResString(1429 + Sprache)      'Falls Fehler auftreten, klicken Sie zum �ffnen der Datei pruef.log auf den Fehlerhinweis
    
    '----------------------------------------------------------
    'Kontrolle ob Professional-Version vorliegtpr�fen1
'    j = GetSystemDirectoryA("", 0)                      'Gerbing 10.06.2005
'    SystemDirectory = Space(j - 1)
'    Call GetSystemDirectoryA(SystemDirectory, j)
    
'auskommentiert Gerbing 30.06.2020
'    'ob Professional Version wird bei SQL Server Version nicht �berpr�ft
'    If gblnSQLServerVersion = False Then
'        If gblnProversion = True Then                                                                   'Gerbing 04.03.2012
'            'Datei = Dir(gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14" & "\gerbingsoft.log")   'Gerbing 30.11.2011 26.10.2013
'            If file_path_exist(gstrFotosIniAnwendungsOrdner & "\gerbingsoft.log") = True Then
'            'If Datei <> "" Then
'                gblnProversion = True
'            Else
'                j = GetSystemDirectoryA("", 0)                                                          'Gerbing 23.11.2013
'                gblstrSystemDirectory = Space(j - 1)
'                Call GetSystemDirectoryA(gblstrSystemDirectory, j)
'                'Pr�fe gblstrSystemDirectory & "\gerbingsoft.log"
'                'Datei = Dir(gblstrSystemDirectory & "\gerbingsoft.log")   'Gerbing 30.11.2011 26.10.2013
'                If file_path_exist(gblstrSystemDirectory & "\gerbingsoft.log") = True Then              'Gerbing 23.11.2013
'                'If Datei <> "" Then
'                    gblnProversion = True
'                Else
'                    'btnNutzerdefinierteFelderAnlegen.Visible = False
'                    'NeueDatens�tzeGenerieren.FrameNutzerDefiniert.Visible = False
'                    btnPr�fenS.Visible = False                      'Gerbing
'                    btnPr�fenSAbbrechen.Visible = False
'                End If
'            End If
'        Else
'            'btnNutzerdefinierteFelderAnlegen.Visible = False
'            'NeueDatens�tzeGenerieren.FrameNutzerDefiniert.Visible = False
'            btnPr�fenS.Visible = False                      'Gerbing
'            btnPr�fenSAbbrechen.Visible = False
'        End If
'    End If
    '----------------------------------------------------------
    'On Error Resume Next
    'Hier ist bekannt on Professional Version vorliegt
    
    NL = Chr(10) & Chr(13)
    'SQL = "Select * from Fotos ORDER BY dateiname"
    SQL = "Select * from Fotos ORDER BY " & LoadResString(1028 + Sprache)
    
    
    DBsql.Errors.Clear
    On Error GoTo 0
    'DbGridForm.rsDataGrid.Resync
    'DbGridForm.rsDataGrid.Close
    
    ' Recordset erstellen und �ffnen
    Set rsDataGrid = New ADODB.Recordset
    With rsDataGrid
        .Source = SQL
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind
    DBGridNeu.AllowArrows = True
    DBGridNeu.TabAcrossSplits = True
    DBGridNeu.TabAction = dbgGridNavigation
    DBGridNeu.WrapCellPointer = True
    
    Call SpaltenBreite
    Call SpaltenbreiteMerken
    DBGridNeu.Caption = PublicDatagridCaption
    DBGridNeu.AllowRowSizing = False
    
    On Error GoTo Fehler
    On Error GoTo 0
    If gblnSQLServerVersion = False Then
        '------------------------------------------------------------------------------------------------------
        'Kontrolle ob die Datenbank schreibgesch�tzt ist                                    'Gerbing 23.11.2017
        On Error Resume Next
        SQL = "UPDATE FET SET FN = 'test'"
        Set rstTempHaken = New ADODB.Recordset
        With rstTempHaken
            .ActiveConnection = DBsql                                             'Gerbing 23.11.2017
            .CursorType = adOpenDynamic
            '.CursorLocation = Query.enumCursorOrt
            .Source = SQL
            '     .CacheSize = 2
            .Open
        End With
        If ERR.Number <> 0 Then
            gblnSchreibgesch�tzt = True
        Else
            gblnSchreibgesch�tzt = False
        End If
        rstTempHaken.Close                                                      'Gerbing 21.01.2018
        If gblnSchreibgesch�tzt = True Then
            'schreibgesch�tzt
            On Error GoTo 0
            
            Msg = AppPath & "\Fotos.mdb" & NL
            'msg = msg & "Die Datenbank ist schreibgesch�tzt. Sie kann nur im Lesemodus ge�ffnet werden." & NL
            Msg = Msg & LoadResString(2132 + Sprache) & NL
            'msg = msg & "Es gibt vier m�gliche Ursachen f�r den Lesemodus:" & NL
            Msg = Msg & LoadResString(2133 + Sprache) & NL
            'msg = msg & "1. Das Dateiattribut 'Schreibgesch�tzt' ist gesetzt" & NL
            Msg = Msg & LoadResString(2134 + Sprache) & NL
            'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte f�r Ihren PC" & NL
            Msg = Msg & LoadResString(2135 + Sprache) & NL
            'msg = msg & "3. Sie arbeiten mit einer CD" & NL
            Msg = Msg & LoadResString(2136 + Sprache) & NL
            'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & NL & NL
            Msg = Msg & LoadResString(2137 + Sprache) & NL & NL
            
    '        Msg = Msg & "Wenn Sie Schreibzugriff brauchen," & NL
    '        Msg = Msg & "�ndern Sie das Schreibschutz-Attribut," & NL
    '        Msg = Msg & "oder kopieren Sie die Datenbank von CD auf Festplatte" & NL
    '        Msg = Msg & "und �ndern Sie danach das Schreibschutz-Attribut." & NL & NL
            
            'msg = msg & "Wollen Sie im Lesemodus weiterarbeiten?" & NL & NL
            Msg = Msg & LoadResString(2138 + Sprache) & NL & NL
            'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
            If antwort = vbNo Then
                End
            Else
                'Datenbank im Lesemodus �ffnen                  'Gerbing 23.11.2017
                DBsql.Close
                DBsql.mode = adModeRead
                DBsql.Open DBsql.ConnectionString
            End If
        End If
    End If
    If rstsql.EOF = True And rstsql.BOF = True Then
        If gblnSQLServerVersion = True Then
            'MsgBox "Die Datei " & PublicSQLServer & " " & Publicsqldatabase & " ist leer." & NL & "Die einzige m�gliche Programmfunktion ist " & "Pr�fen3 oder" & NL & "&Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)..."
            MsgBox LoadResString(2145 + Sprache) & " " & PublicSQLServer & " " & PublicSQLDatabase & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
        Else
            'MsgBox "Die Datei " & AppPath & "\Fotos.mdb" & " "  " ist leer." & NL & "Die einzige m�gliche Programmfunktion ist " & "Pr�fen3 oder" & NL & "&Neue Datens�tze generieren (durch Drag&&Drop vom Windows Explorer)..."
            MsgBox LoadResString(2145 + Sprache) & AppPath & "\Fotos.mdb" & " " & LoadResString(1512 + Sprache) & NL & LoadResString(1511 + Sprache) & NL & LoadResString(1513 + Sprache) & NL & "'" & LoadResString(1311 + Sprache) & "'"
        End If
        Exit Sub
    End If
    If gblnSQLServerVersion = False Then
        '---------------------------------------------------------------------------------------------------------
        '3-Einigkeit �berpr�fen Gerbing 11.04.2005
        'Der erste Satz in der Datenbank wird kontrolliert
        'Feldname = rstsql.Fields("Dateiname")
        Feldname = rstsql.Fields(LoadResString(1028 + Sprache))
        If left(Feldname, 3) <> "+:\" Then
            'msg = "Seit Version 12.0.0.0 verlangt das Programm, dass in der Tabelle Fotos" & vbNewLine
            Msg = LoadResString(2154 + Sprache) & vbNewLine
            'msg = msg & "das Feld Dateiname generell mit den Zeichen +:\ beginnt" & vbNewLine
            Msg = Msg & LoadResString(2155 + Sprache) & vbNewLine
            'msg = msg & "Der String +:\ wird vom Programm durch AppPath ersetzt." & vbNewLine
            Msg = Msg & LoadResString(2156 + Sprache) & vbNewLine
            'msg = msg & "AppPath ist der Name des Ordners in dem fotos.exe steht." & vbNewLine
            Msg = Msg & LoadResString(2157 + Sprache) & vbNewLine
            'msg = msg & "Diese Forderung wurde nicht eingehalten." & vbNewLine & vbNewLine
            Msg = Msg & LoadResString(2158 + Sprache) & vbNewLine & vbNewLine
            
            'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
            Msg = Msg & LoadResString(2159 + Sprache)
            'antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
            antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton2 + vbYesNo) 'Gerbing 09.09.2014
            If antwort = vbNo Then
                End
            End If
        End If
        Feldname = Replace(Feldname, "+:\", AppPath & "\")
        'strtemp = Dir(Feldname)
        If file_path_exist(Feldname) = False Then
        'If strtemp = "" Then
            'msg = Feldname & " existiert nicht." & vbNewLine
            Msg = Feldname & LoadResString(2162 + Sprache) & vbNewLine
            'msg = "Datenbank und Fotos passen nicht zueinander" & vbNewLine
            Msg = Msg & LoadResString(2160 + Sprache) & vbNewLine
            'msg = msg & "Vermutlich benutzen Sie eine falsche Datenbank-Datei" & vbNewLine
            Msg = Msg & LoadResString(2161 + Sprache) & vbNewLine
            'msg = msg & "Benutzen Sie das Tool Fotosmdb um die Datenbank zu �berpr�fen" & vbNewLine & vbNewLine
            Msg = Msg & LoadResString(2163 + Sprache) & vbNewLine & vbNewLine
            
            'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
            Msg = Msg & LoadResString(2159 + Sprache)
            antwort = MsgBox(Msg, vbDefaultButton2 + vbYesNo)
            If antwort = vbNo Then
                End
            End If
        End If
        '---------------------------------------------------------------------------------------------------------
    End If
    Call SpaltenBreite
    Call SpaltenbreiteMerken
'    Call SpaltenbreiteWiederherstellen
    Me.Show
    Exit Sub
Fehler:
    Msg = "Errornumber=" & ERR & NL
    Msg = Msg & "Errortext=" & Error(ERR) & NL & NL
    If ERR.Number = 429 Then                                                            'Gerbing 21.11.2007
        Msg = Msg & "You must register the dao360.dll" & vbNewLine
    End If
    Msg = Msg & "read in http://www.gerbingsoft.de under news or look for that problem in the internet"
    MsgBox Msg
    End
    Exit Sub
ERR:
    If DBsql.Errors.Count > 0 Then
        For Each errLoop In DBsql.Errors
            MsgBox "Fehler Nr.: " & errLoop.Number & vbCr & _
                errLoop.Description
        Next errLoop
    End If
    Resume Next
End Sub

Private Sub Pr�fenSKommentar1(Dateiname)
    'AudioFileExists = yes eingetragen
    Dim Msg As String
    Dim tmpName As String
    Dim MyAppPath As String
    
    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    FehlerGefunden = True
    tmpName = Replace(Dateiname, "+:\", MyAppPath & "\")
'    Msg = Dateiname & " "
'    Msg = Msg & "wurde erfolgreich als Audio-Datei gekennzeichnet"
    Msg = tmpName & " "
    Msg = Msg & LoadResString(2259 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
End Sub

Private Sub Pr�fenSKommentar2(Dateiname)
    ''Audio-Datei gel�scht
    Dim Msg As String
    Dim tmpName As String
    Dim MyAppPath As String
    
    If gblnSQLServerVersion = True Then                                                     'Gerbing 05.09.2013
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If

    FehlerGefunden = True
    tmpName = Replace(Dateiname, "+:\", MyAppPath & "\")
'    Msg = Dateiname & " "
'    Msg = Msg & "wurde gel�scht, weil es keine dazu passende Fotodatei gibt"
    Msg = tmpName & " "
    Msg = Msg & LoadResString(2260 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
End Sub

Private Sub FehlerJahrIsNull(DatensatzNr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "DatensatzNr." & DatensatzNr
'    Msg = Msg & "Die Datei " & Fotodatei
'    Msg = Msg & "Dieser Datensatz enth�lt kein Jahr. Sie sollten ihn nach dem Pr�fen1 l�schen."
    Msg = LoadResString(1008 + Sprache) & " " & DatensatzNr & " "
    Msg = Msg & LoadResString(2035 + Sprache) & strFotoDatei & " "                          'Gerbing 22.11.2019
    Msg = Msg & LoadResString(2214 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    If NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count > 32766 Then                      'Gerbing 26.01.2009
        If blnMessageAusgeben = True Then
            blnMessageAusgeben = False
            'MsgBox "Das Programm kann in einem Durchlauf maximal 32767 Dateien aufnehmen. Wiederholen Sie die Programmfunktion f�r die noch nicht aufgenommenen Dateien."
            MsgBox LoadResString(2333 + Sprache)
        End If
    Else
        NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add strFotoDatei                 'Gerbing 23.12.2019
    End If                                                                                  'Gerbing 26.01.2009
End Sub

Private Sub FehlerDateinameIsNull(DatensatzNr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "DatensatzNr." & DatensatzNr
'    Msg = Msg & "Dieser Datensatz enth�lt keinen Dateiname. Er wird sofort gel�scht."
    Msg = LoadResString(1008 + Sprache) & " " & DatensatzNr & " "
    Msg = Msg & LoadResString(2213 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    If NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count > 32766 Then                'Gerbing 03.11.2013      'Gerbing 26.01.2009
        If blnMessageAusgeben = True Then
            blnMessageAusgeben = False
            'MsgBox "Das Programm kann in einem Durchlauf maximal 32767 Dateien aufnehmen. Wiederholen Sie die Programmfunktion f�r die noch nicht aufgenommenen Dateien."
            MsgBox LoadResString(2333 + Sprache)
        End If
    Else
        NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add Fotodatei
    End If                                                                                  'Gerbing 26.01.2009
End Sub

Private Sub FehlerFotoDatei(DatensatzNr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "DatensatzNr." & DatensatzNr
'    Msg = Msg & "Die Datei " & Fotodatei
'    Msg = Msg & "  ist nicht vorhanden"
'    Msg = Msg & "  ist nicht vorhanden oder hat ein ung�ltiges Datum"
    Msg = LoadResString(1008 + Sprache) & " " & DatensatzNr & " "
    Msg = Msg & LoadResString(2035 + Sprache) & strFotoDatei & " "                          'Gerbing 22.11.2019
    Msg = Msg & LoadResString(2463 + Sprache)                                               'Gerbing 04.03.2013
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    If NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count > 32766 Then                      'Gerbing 26.01.2009
        If blnMessageAusgeben = True Then
            blnMessageAusgeben = False
            'MsgBox "Das Programm kann in einem Durchlauf maximal 32767 Fehler finden. Wiederholen Sie die Programmfunktion nach der Fehlerbeseitigung f�r die noch nicht gefundenen Fehler."
            MsgBox LoadResString(2334 + Sprache)                                            'Gerbing 15.10.2009
        End If
    Else
        NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add strFotoDatei                 'Gerbing 23.12.2019
    End If                                                                                  'Gerbing 26.01.2009
End Sub

Private Sub FehlerDateinamenErweiterung(DatensatzNr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "DatensatzNr." & DatensatzNr
'    Msg = Msg & "Die Datei " & Fotodatei
'    Msg = Msg & "  Dateinamen-Erweiterung ist nicht erlaubt"
    Msg = LoadResString(1008 + Sprache) & " " & DatensatzNr & " "
    Msg = Msg & LoadResString(2035 + Sprache) & Fotodatei
    Msg = Msg & LoadResString(1397 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    If NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count > 32766 Then                'Gerbing 03.11.2013      'Gerbing 26.01.2009
        If blnMessageAusgeben = True Then
            blnMessageAusgeben = False
            'MsgBox "Das Programm kann in einem Durchlauf maximal 32767 Dateien aufnehmen. Wiederholen Sie die Programmfunktion f�r die noch nicht aufgenommenen Dateien."
            MsgBox LoadResString(2333 + Sprache)
        End If
    Else
        NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Add Fotodatei
    End If                                                                                  'Gerbing 26.01.2009
End Sub

Private Sub FehlerDateinamenErweiterungWiderspruch(DatensatzNr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "DatensatzNr." & DatensatzNr
'    Msg = Msg & "Die Datei " & Fotodatei
'    Msg = Msg & "  Widerspruch in Dateinamen-Erweiterung und Spalte SWF"
    Msg = LoadResString(1008 + Sprache) & " " & DatensatzNr & " "
    Msg = Msg & LoadResString(2035 + Sprache) & strFotoDatei                                'Gerbing 22.11.2019
    Msg = Msg & LoadResString(1398 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
End Sub

Private Sub FehlerPr�fen2(Jahr)
    Dim Msg As String
    
    FehlerGefunden = True
'    Msg = "Jahr: " & Jahr
'    Msg = Msg & "Die Datei " & Fotodatei
'    Msg = Msg & "  Jahreszahl im Feld 'Jahr' und im Dateiname stimmt nicht �berein"
    Msg = LoadResString(1023 + Sprache) & ": " & Jahr & " "
    Msg = Msg & LoadResString(2035 + Sprache) & Fotodatei
    Msg = Msg & LoadResString(1399 + Sprache)
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
End Sub

Private Sub FehlerUmkehrProbe()
    Dim Msg As String
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim pos4 As Long                                                        'Gerbing 29.07.2007
    Dim pos5 As Long
    Dim pos6 As Long
    
    pos = InStr(1, Fotodatei, "Checked.ico", vbTextCompare)
    pos1 = InStr(1, Fotodatei, "Checked1.ico", vbTextCompare)
    pos2 = InStr(1, Fotodatei, "Unchecked.ico", vbTextCompare)        'Gerbing 11.08.2004
    pos3 = InStr(1, Fotodatei, "G.ico", vbTextCompare)              'Gerbing 11.04.2005
    pos4 = InStr(1, Fotodatei, "Unchecked1.ico", vbTextCompare)             'Gerbing 29.07.2007
    
    If pos = 0 And pos1 = 0 And pos2 = 0 And pos3 = 0 And pos4 = 0 Then    'Gerbing 29.07.2007
        FehlerGefunden = True
'        Msg = "Die Datei " & Fotodatei
'        Msg = Msg & "  ist nicht in der Datenbank " & PublicDatagridCaption
'        Msg = Msg & " eingetragen"
        Msg = LoadResString(2035 + Sprache) & Fotodatei
        Msg = Msg & LoadResString(1400 + Sprache) & PublicDatagridCaption
        Msg = Msg & LoadResString(1401 + Sprache)
        On Error Resume Next                                                                'Gerbing 02.09.2008
        'Print #DateiNummer, Msg
        oStream.WriteLine Msg
        On Error GoTo 0                                                                     'Gerbing 02.09.2008
        NachPr�fen3L�schen.KollZus�tzlicheDateien.Add Fotodatei                             'Gerbing 26.10.2013
        NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Add Fotodatei                           'Gerbing 26.10.2013
    End If
End Sub

Private Function Rekursive(Path As String, SearchStr As String)
    Dim Filename As String              ' Walking filename variable...
    Dim DirName As String               ' SubDirectory Name
    Dim dirNames() As String            ' Buffer for directory name entries
    Dim nDir As Long                    ' Number of directories in this path
    Dim i As Long                       ' For-loop counter...
    Dim hSearch As Long                 ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Long
    Dim DirCount As Long
    Dim DateinamenErweiterung As String
    Dim MyAppPath As String
    Dim strTemp As String
    Dim pos As Long

    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    If gblnSubdirectories = True Then
        Cont = True
        hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(WFD))
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While Cont
                'DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
                DirName = RemoveNulls((WFD.cFileName))
                ' Ignore the current and encompassing directories.
                If (DirName <> ".") And (DirName <> "..") Then
                    pos = InStr(1, DirName, "GerbingThumbs", vbTextCompare)                    'Gerbing 10.11.2016
                    If pos = 0 Then
                        ' Check for directory with bitwise comparison.
                        If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                            dirNames(nDir) = DirName
                            DirCount = DirCount + 1
                            nDir = nDir + 1
                            ReDim Preserve dirNames(nDir)
                        End If
                    Else
                        'Dateien in ...\GerbingThumbs\... werden ignoriert
                    End If
                End If
                Cont = FindNextFileW(hSearch, VarPtr(WFD)) 'Get next subdirectory.
            Loop
            Cont = FindClose(hSearch)
        End If
    End If
    ' Walk through this directory.
    hSearch = FindFirstFileW(StrPtr(Path & SearchStr), VarPtr(WFD))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
'            StartMillisek = timeGetTime
            'FileName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            Filename = RemoveNulls((WFD.cFileName))
            If (Filename <> ".") And (Filename <> "..") Then
                DateinamenErweiterung = UCase(Right(Filename, 5))
                Select Case DateinamenErweiterung
                    '5-stellige
                    Case "MRSID"
                        rstTempHaken.AddNew                             'Neue S�tze zu rstTempHaken hinzuf�gen
                        strTemp = Path & Filename
                        strTemp = Replace(strTemp, MyAppPath, "+:")                     'Gerbing 11.04.2005
        '                rstTempHaken!Dateiname = strTemp                                'Gerbing 11.04.2005
        '                rstTempHaken!Merker = False                     'Merker = 0
                        rstTempHaken.Fields(LoadResString(1028)) = strTemp    'Dateiname Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(2524)) = False      'Merker = 0
                        rstTempHaken.Update
                        GoTo Weiter
                End Select
                DateinamenErweiterung = UCase(Right(Filename, 4))
                Select Case DateinamenErweiterung
                    '4-stellige
                    Case "WBMP", "TIFF", "PICT", "QTIF", "JPEG", "FITS", "HPGL", "IW44", "DJVU", "CS16", "DOCX" 'Gerbing 09.01.2008 18.01.2014
                        rstTempHaken.AddNew                             'Neue S�tze zu rstTempHaken hinzuf�gen
                        strTemp = Path & Filename
                        strTemp = Replace(strTemp, MyAppPath, "+:")                     'Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(1028)) = strTemp    'Dateiname Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(2524)) = False      'Merker = 0
                        rstTempHaken.Update
                        GoTo Weiter
                End Select
                DateinamenErweiterung = UCase(Right(Filename, 3))
                Select Case DateinamenErweiterung
                    '3-stellige
                    Case "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF", "AVI", "MPG", "PEG", "MOV", "MP4", "MKV", "FLV", _
                            "MPE", "ASF", "ASX", "WMV", "HTM", "PDF", "XLS", _
                            "ANI", "B3D", "CAM", "CLP", "CPT", "CRW", "CR2", "DCM", "ACR", "IMA", "DCX", "DDS", _
                            "DXF", "DWG", "ECW", "EMF", "EPS", "FPX", "FSH", "ICL", _
                            "ICS", "IFF", "LBM", "IMG", "JP2", "JPC", "J2K", "JPM", "KDC", "LWF", _
                            "MNG", "JNG", "SID", "DNG", "EEF", "NEF", "MRW", "ORF", "RAF", _
                            "DCR", "SRF", "PEF", "X3F", "NLM", "NOL", "NGG", "PBM", "PCD", "PCX", "PGM", "PIC", _
                            "PNG", "PPM", "PSD", "PSP", "RAS", "SUN", "RAW", "RLE", "SFF", "SFW", "SGI", "RGB", _
                            "SWF", "TGA", "TIF", "TTF", "WAD", "WAL", "XBM", "XPM", _
                            "3FR", "ARW", "CS1", "CS4", "DCS", "ERF", "MEF", "SR2", "DOC"        'Gerbing 11.12.2005 und 09.01.2008 18.01.2014 09.03.2014 10.12.2017
                            'Gerbing 09.03.2005
                            rstTempHaken.AddNew                             'Neue S�tze zu rstTempHaken hinzuf�gen
                            strTemp = Path & Filename
                            strTemp = Replace(strTemp, MyAppPath, "+:")                     'Gerbing 11.04.2005
                            rstTempHaken.Fields(LoadResString(1028)) = strTemp    'Gerbing 11.04.2005
                            rstTempHaken.Fields(LoadResString(2524)) = False      'Merker = 0
                            rstTempHaken.Update
                        GoTo Weiter
                End Select
                DateinamenErweiterung = UCase(Right(Filename, 2))
                Select Case DateinamenErweiterung
                    '2-stellige
                    Case "G3"
                        rstTempHaken.AddNew                             'Neue S�tze zu rstTempHaken hinzuf�gen
                        strTemp = Path & Filename
                        strTemp = Replace(strTemp, MyAppPath, "+:")                     'Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(1028)) = strTemp    'Dateiname Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(2524)) = False      'Merker = 0
                        rstTempHaken.Update
                        GoTo Weiter
                End Select
                '---------------------------------------------------
            End If
Weiter:
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) ' Get next file
            
'            EndMillisek = timeGetTime
'            Debug.Print "Millisekunden rstTempHaken.Update f�r Datei " & FileName & "=" & (EndMillisek - StartMillisek)
            
            DoEvents
            If Pr�fen3Abbrechen = True Then Exit Function
        Wend
        Cont = FindClose(hSearch)
    End If
    If gblnSubdirectories = True Then
    ' If there are sub-directories...
        If nDir > 0 Then
            ' Recursively walk into them...
            For i = 0 To nDir - 1
                Rekursive = Rekursive(Path & dirNames(i) & "\", SearchStr)
                txtArbeitsfortschrittU.Text = "Rekursive " & Path & dirNames(i) 'Arbeitsfortschritt
                DoEvents
            Next i
        End If
    End If
End Function

Private Function RekursiveTempThumbs(Path As String, SearchStr As String)        'Gerbing 06.04.2017
    Dim Filename As String              ' Walking filename variable...
    Dim DirName As String               ' SubDirectory Name
    Dim dirNames() As String            ' Buffer for directory name entries
    Dim nDir As Long                    ' Number of directories in this path
    Dim i As Long                       ' For-loop counter...
    Dim hSearch As Long                 ' Search Handle
    Dim WFD As WIN32_FIND_DATA
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
        hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(WFD))
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While Cont
            'DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            DirName = RemoveNulls((WFD.cFileName))
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
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) 'Get next subdirectory.
            Loop
            Cont = FindClose(hSearch)
        End If
    End If
    ' Walk through this directory.
    hSearch = FindFirstFileW(StrPtr(Path & SearchStr), VarPtr(WFD))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            'Filename = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            Filename = RemoveNulls((WFD.cFileName))
            If (Filename <> ".") And (Filename <> "..") Then
                '---------------------------------------------------
                DateinamenErweiterung = UCase(Right(Filename, 3))
                Select Case DateinamenErweiterung
                    Case "JPG"
                        rc = file_delete(Path & Filename, False, True) 'ohne Papierkorb, silent
                End Select
                '---------------------------------------------------
            End If
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) ' Get next file
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

Private Function RekursiveAudio(Path As String, SearchStr As String)
    Dim Filename As String              ' Walking filename variable...
    Dim DirName As String               ' SubDirectory Name
    Dim dirNames() As String            ' Buffer for directory name entries
    Dim nDir As Long                    ' Number of directories in this path
    Dim i As Long                       ' For-loop counter...
    Dim hSearch As Long                 ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Long
    Dim DirCount As Long
    Dim DateinamenErweiterung As String
    Dim MyAppPath As String
    Dim strTemp As String

    If gblnSQLServerVersion = True Then
        MyAppPath = PublicLocationFotos
    Else
        MyAppPath = AppPath
    End If
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    If gblnSubdirectories = True Then
        Cont = True
        hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(WFD))
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While Cont
                'DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
                DirName = RemoveNulls((WFD.cFileName))
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
                Cont = FindNextFileW(hSearch, VarPtr(WFD)) 'Get next subdirectory.
            Loop
            Cont = FindClose(hSearch)
        End If
    End If
    ' Walk through this directory.
    hSearch = FindFirstFileW(StrPtr(Path & SearchStr), VarPtr(WFD))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            'FileName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            Filename = RemoveNulls((WFD.cFileName))
            If (Filename <> ".") And (Filename <> "..") Then
                DateinamenErweiterung = UCase(Right(Filename, 3))
                Select Case DateinamenErweiterung
                    '3-stellige
                    Case "WAV", "MP3"
                        rstTempHaken.AddNew                             'Neue S�tze zu rstTempHaken hinzuf�gen
                        strTemp = Path & Filename
                        strTemp = Replace(strTemp, MyAppPath, "+:")                     'Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(1028)) = strTemp    'Gerbing 11.04.2005
                        rstTempHaken.Fields(LoadResString(2524)) = False      'Merker = 0
                        rstTempHaken.Update
                End Select
                '---------------------------------------------------
            End If
Weiter:
            Cont = FindNextFileW(hSearch, VarPtr(WFD)) ' Get next file
            DoEvents
            If Pr�fenSAbbrechen = True Then Exit Function
        Wend
        Cont = FindClose(hSearch)
    End If
    If gblnSubdirectories = True Then
    ' If there are sub-directories...
        If nDir > 0 Then
            ' Recursively walk into them...
            For i = 0 To nDir - 1
                RekursiveAudio = RekursiveAudio(Path & dirNames(i) & "\", SearchStr)
                txtArbeitsfortschrittU.Text = "Rekursive " & Path & dirNames(i) 'Arbeitsfortschritt
                DoEvents
            Next i
        End If
    End If
End Function

Private Sub btnPr�fen3_Click()
    Dim Verzeichnis As String
    Dim Gefunden As String
    Dim Msg As String
    Dim SQL As String
    Dim Erg As Long
    Dim antwort As Long
    Dim i As Long
    Dim TemprstSQL As ADODB.Recordset                               'Gerbing 21.12.2015
    Dim Temprst As ADODB.Recordset                                  'Gerbing 23.11.2017
    
    Call RekursiveTempThumbs(AppPath & "\TempThumbs", "*.*")        'Gerbing 28.05.2019
    
    blnMessageAusgeben = True                                       'Gerbing 26.01.2009
    If gblnSchreibgesch�tzt = True Then                             'Gerbing 23.01.2007
        'msg = "Bei einer schreibgesch�tzten Datenbank ist diese Funktion nicht m�glich"
        Msg = LoadResString(2421 + Sprache)
        MsgBox Msg
        Exit Sub
    End If
    '----------------------------------------------------------------------------------
    Call SpaltenbreiteMerken
    Call ButtonsDisabled
    Pr�fen3Abbrechen = False                                        'Gerbing 04.10.2004
    'Pr�fenNummer = "Pr�fen3"
    Pr�fenNummer = LoadResString(1459 + Sprache)
    txtFehlerU.Text = ""
    FehlerGefunden = False
    If gblnSQLServerVersion = True Then
        StartVerzeichnis = PublicLocationFotos
    Else
        StartVerzeichnis = AppPath                                   'Gerbing 11.04.2005
    End If
    Debug.Print PublicLocationFotos
    
    Screen.MousePointer = vbHourglass
    '----------------------------------------------------------------------------------
    btnPr�fen3Abbrechen.Enabled = True                              'Gerbing 04.10.2004
    NachPr�fen3L�schen.lstZus�tzlicheDateien.ListItems.RemoveAll
    On Error Resume Next                                                'Gerbing 29.12.2011
    On Error GoTo 0                                                     'Gerbing 29.12.2011
    If gblnSQLServerVersion = True Then
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 29.12.2011
        'beim SQL Server muss es hei�en 'Delete from table
        SQL = "DELETE From Temp_Haken"
        'SQL = "DELETE FROM " & LoadResString(2523 + Sprache)
    Else
        'Zuerst aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 30.09.2004
        SQL = "DELETE " & "Temp_Haken.* "
        SQL = SQL & " FROM " & "Temp_Haken;"
        'SQL = "DELETE " & LoadResString(2523 + Sprache) & ".* "
        'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    End If
    DBsql.Execute SQL
    'dann leeres Recordset rstTempHaken �ffnen
    SQL = " SELECT " & "Temp_Haken.*"
    SQL = SQL & " FROM " & "Temp_Haken;"
    'SQL = " SELECT " & LoadResString(2523 + Sprache) & ".*"
    'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    Set rstTempHaken = New ADODB.Recordset
    With rstTempHaken
        .ActiveConnection = DBsql
        .CursorType = adOpenDynamic
        '.CursorLocation = adUseClient
        .CursorLocation = adUseServer                                   'Gerbing 19.04.2015
        .Source = SQL
        .LockType = adLockOptimistic
        .Open
    End With
    '-------------------------------------------------------------------------------------------------------
    gblnSubdirectories = True
    Call Rekursive(StartVerzeichnis, "*.*") 'f�lle die Tabelle Temp_Haken mit s�mtlichen wirklichen Dateinamen einschlie�lich subdirectories
    '-------------------------------------------------------------------------------------------------------
    rstTempHaken.Close                                                          'Gerbing 10.08.2006
    If Pr�fen3Abbrechen = True Then                                             'Gerbing 04.10.2004
        'Pr�fen3 wurde vom Nutzer abgebrochen
        Call ButtonsEnabled
        If gblnSQLServerVersion = True Then
            'Zuletzt aus der Tabelle Temp_Haken alle S�tze l�schen              'Gerbing 29.12.2011
            'beim SQL Server muss es hei�en 'Delete from table
            SQL = "DELETE From Temp_Haken"
            'SQL = "DELETE FROM " & LoadResString(2523 + Sprache)
        Else
            'Zuletzt aus der Tabelle Temp_Haken alle S�tze l�schen              'Gerbing 30.09.2004
            SQL = "DELETE " & "Temp_Haken.* "
            SQL = SQL & " FROM " & "Temp_Haken;"
            'SQL = "DELETE " & LoadResString(2523 + Sprache) & ".* "
            'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
        End If
        DBsql.Execute SQL
        Exit Sub                                                    'Pr�fen3 wurde vom Nutzer abgebrochen
    End If
    '--------------------------------------------------------------------------------------------------------
    '�ffne die Datei pruef.log
    On Error Resume Next
    ERR = 0
    'Open PruefLogFile For Output As #DateiNummer
    'object.CreateTextFile(filename[, overwrite[, unicode]])
    Set oStream = PruefFso.CreateTextFile(PruefLogFile, True, True)
    If ERR <> 0 Then
        'Msg = "Die Datei " & PruefLogFile & " kann nicht ge�ffnet werden" & NL
        Msg = LoadResString(2035 + Sprache) & " " & PruefLogFile & " " & LoadResString(1372 + Sprache) & NL
        'msg = msg & "Sie m�ssen f�r Schreibrechte sorgen, damit �nderungen an dieser Datei gemacht werden k�nnen." & NL
        Msg = Msg & LoadResString(2276 + Sprache) & NL
        Msg = Msg & "Errortext=" & ERR.Description & NL
        Msg = Msg & "Errornumber=" & ERR.Number & NL & NL
        
        'msg = msg & "Wollen Sie trotzdem weiterarbeiten?"
        Msg = Msg & LoadResString(1542 + Sprache)
        'antwort = MsgBox(Msg, vbDefaultButton1 + vbYesNo)                                   'Gerbing 02.09.2008
        antwort = MessageBoxW(0, StrPtr(Msg), StrPtr("GERBING Fotosmdb"), vbDefaultButton1 + vbYesNo) 'Gerbing 09.09.2014
        If antwort = vbNo Then
            LogNichtBenutzbar = False
            Call ButtonsEnabled
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            LogNichtBenutzbar = True
        End If
    End If
    On Error GoTo 0
    Msg = Now & "  "
    Msg = Msg & Pr�fenNummer & "  "
    'Msg = Msg & "Pr�f-Ergebnis der Datenbank " & PublicDatagridCaption & NL
    Msg = Msg & LoadResString(1381 + Sprache) & PublicDatagridCaption & NL
    On Error Resume Next                                                                    'Gerbing 02.09.2008
    'Print #DateiNummer, Msg
    oStream.WriteLine Msg
    On Error GoTo 0                                                                         'Gerbing 02.09.2008
    '--------------------------------------------------------------------------------------
    'Inkonsistenzabfrage                                                'Gerbing 30.09.2004
    'Die Inkonsistenzabfrage findet alle Dateinamen, die nicht in Tabelle Fotos eingetragen sind
    'SQL = "SELECT Temp_Haken.Dateiname FROM Temp_Haken LEFT JOIN Fotos ON Temp_Haken.Dateiname = Fotos.Dateiname"
    'SQL = SQL & " WHERE (((Fotos.Dateiname) Is Null));"
    SQL = "SELECT Temp_Haken." & LoadResString(1028) & " FROM Temp_Haken LEFT JOIN Fotos ON Temp_Haken." & LoadResString(1028) & " = Fotos." & LoadResString(1028 + Sprache)
    SQL = SQL & " WHERE (((Fotos." & LoadResString(1028 + Sprache) & ") Is Null));"
    On Error Resume Next
    rstsql.Close
    On Error GoTo 0
    With rstsql
        .Source = SQL
        .ActiveConnection = DBsql
        '.CursorType = adOpenStatic
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    NachPr�fen3Aufnehmen.lstZus�tzlicheDateien.ListItems.RemoveAll
    NachPr�fen3L�schen.lstZus�tzlicheDateien.ListItems.RemoveAll
    If Not rstsql.EOF Then
        'Jetzt werden alle Dateinamen, die nicht Tabelle Fotos eingetragen sind nach pruef.log geschrieben
        rstsql.MoveFirst
        Do Until NachPr�fen3L�schen.KollZus�tzlicheDateien.Count = 0                        'Gerbing 26.10.2013
            NachPr�fen3L�schen.KollZus�tzlicheDateien.Remove 1
        Loop
        Do Until NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count = 0                      'Gerbing 26.10.2013
            NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Remove 1
        Loop
        Do Until rstsql.EOF
            Fotodatei = rstsql.Fields("Dateiname")
            'Fotodatei = rstsql.Fields(LoadResString(1028 + Sprache))
            Call FehlerUmkehrProbe      'Fehler in PruefLogFile eintragen
            rstsql.Movenext
            txtArbeitsfortschrittU.Text = Fotodatei                                         'Gerbing 25.10.2013
            DoEvents
        Loop
    End If
    rstsql.Close                                   'Gerbing 10.08.2006
    '--------------------------------------------------------------
    Screen.MousePointer = vbDefault
    'schlie�e die Datei pruef.log
    'schlie�e die Datei PruefLogFile
    If FehlerGefunden = False Then
'        Print #DateiNummer, "kein Fehler gefunden"
'        txtfehleru.text = "kein Fehler gefunden"
        On Error Resume Next                                                                'Gerbing 02.09.2008
        'Print #DateiNummer, LoadResString(1382 + Sprache)
        oStream.WriteLine LoadResString(1382 + Sprache)
        On Error GoTo 0                                                                     'Gerbing 02.09.2008
        txtFehlerU.Text = LoadResString(1382 + Sprache)
    Else
        If LogNichtBenutzbar = False Then
            'txtfehleru.text = "Fehler siehe " & AppPath & "\pruef.log"
            txtFehlerU.Text = LoadResString(1383 + Sprache) & PruefLogFile
        Else
            'txtfehleru.text = AppPath & "\pruef.log" & "nicht benutzbar"
            txtFehlerU.Text = PruefLogFile & LoadResString(2277 + Sprache)
        End If
    End If
    'Close #DateiNummer
    On Error Resume Next
    oStream.Close
    On Error GoTo 0
    'txtArbeitsfortschritt.Text = "Pr�fen3 beendet"
    txtArbeitsfortschrittU.Text = LoadResString(1402 + Sprache)
    btnPr�fen3Abbrechen.Enabled = False
    Call ButtonsEnabled
    If gblnSQLServerVersion = True Then
        'Zuletzt aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 29.12.2011
        'beim SQL Server muss es hei�en 'Delete from table
        SQL = "DELETE From Temp_Haken"
        'SQL = "DELETE FROM " & LoadResString(2523 + Sprache)
    Else
        'Zuletzt aus der Tabelle Temp_Haken alle S�tze l�schen           'Gerbing 30.09.2004
        SQL = "DELETE " & "Temp_Haken.* "
        SQL = SQL & " FROM " & "Temp_Haken;"
        'SQL = "DELETE " & LoadResString(2523 + Sprache) & ".* "
        'SQL = SQL & " FROM " & LoadResString(2523 + Sprache)
    End If
    DBsql.Execute SQL
End Sub

Private Sub Form_Unload(Cancel As Integer)      'Gerbing 13.11.2004
    Dim rc As Boolean
    Dim retStatus As Status

    If GdipInitialized = True Then
        retStatus = Execute(ShutdownGDIPlus)
    End If
    Set oStream = Nothing
    Set PruefFso = Nothing
    Set IniFso = Nothing
    Set EXF = Nothing                           'Gerbing 07.05.2007
    If gblnSQLServerVersion = False Then
        On Error Resume Next
        rstsql.Close
        DBsql.Close
    End If
    End
End Sub

Private Sub ButtonsDisabled()
    btnPr�fen1.Enabled = False
    btnPr�fen2.Enabled = False
    btnPr�fen3.Enabled = False
    btnPr�fenS.Enabled = False
    btnEXIFIPTC.Enabled = False                             'Gerbing 04.02.2008
    btnGenerieren.Enabled = False                       'Gerbing 06.11.2013
    btnNutzerdefinierteFelderAnlegen.Enabled = False    'Gerbing 06.11.2013
    btn�ffnePruefLog.Enabled = False                    'Gerbing 06.11.2013
    btnL�scheInhaltFotosMdb.Enabled = False             'Gerbing 06.11.2013
End Sub

Public Sub ButtonsEnabled()
    btnPr�fen1.Enabled = True
    btnPr�fen2.Enabled = True
    btnPr�fen3.Enabled = True
    btnPr�fenS.Enabled = True
    btnEXIFIPTC.Enabled = True                              'Gerbing 04.02.2008
    btnGenerieren.Enabled = True                        'Gerbing 06.11.2013
    btnNutzerdefinierteFelderAnlegen.Enabled = True     'Gerbing 06.11.2013
    btn�ffnePruefLog.Enabled = True                     'Gerbing 06.11.2013
    btnL�scheInhaltFotosMdb.Enabled = True              'Gerbing 06.11.2013
End Sub

Private Sub SpaltenBreite()
'    'Wenn ich ohne diese Prozedur arbeite bekommen nach jedem Adodc1.Refresh
'    'die Grid.Spalten eine Standardbreite Jahr genauso breit wie Dateiname
    DBGridNeu.Columns(0).Width = 600    'Merker
    DBGridNeu.Columns(1).Width = 500    'Jahr
    DBGridNeu.Columns(2).Width = 1000   'Situation
    DBGridNeu.Columns(3).Width = 1000   'Ort
    DBGridNeu.Columns(4).Width = 1000   'Land
    DBGridNeu.Columns(5).Width = 3000   'Personen
    DBGridNeu.Columns(6).Width = 3000   'Dateiname
    DBGridNeu.Columns(7).Width = 500    'SWF
    DBGridNeu.Columns(8).Width = 1000   'Kommentar      'Gerbing 27.10.2016
    DBGridNeu.Columns(9).Width = 0     'DateinameKurz
    DBGridNeu.Columns(10).Width = 1000   'DDatum
    DBGridNeu.Columns(11).Width = 1000   'BreitePixel
    DBGridNeu.Columns(12).Width = 1000    'HoehePixel

    DBGridNeu.Refresh
End Sub

Public Sub SpaltenbreiteMerken()
    Dim n As Long
    Dim ColWidth As Long

    'Bei jedem Speichern der Spaltenbreiten wird der bisherige Inhalt der Listbox lstSpaltenbreite zuerst
    'gel�scht, dann werden neue Eintr�ge gemacht
    On Error GoTo 0
    lstSpaltenbreite.Clear
    For n = 0 To DBGridNeu.Columns.Count - 1
        ColWidth = DBGridNeu.Columns(n).Width
        If DBGridNeu.Columns(n).Visible = False Then ColWidth = 0
        lstSpaltenbreite.AddItem ColWidth
    Next n
End Sub

Public Sub SpaltenbreiteWiederherstellen()
    Dim n As Long

    For n = 0 To lstSpaltenbreite.ListCount - 1
        DBGridNeu.Columns(n).Width = lstSpaltenbreite.List(n)
    Next n
End Sub

Private Sub FremdeFotosMdb()                                                                    'Gerbing 10.09.2013
    Dim NetzwerkDir As String
    Dim Msg As String
    
Begin:
    '(ByVal Filter$, ByVal InitialDir$, ByVal Title$) as String
    NetzwerkDir = ShowOpenUnicodeFotosMdb(Me)    '2458=Standort der fotos.mdb
    'NetzwerkDir = GetOpenName(Filter, AppPath, LoadResString(2458 + Sprache))    '2458=Standort der fotos.mdb
    'Convert the file name to be used
    NetzwerkDir = ConvertFileName(NetzwerkDir)
    If NetzwerkDir = "" Then
        Exit Sub
    End If
    If Mid(NetzwerkDir, Len(NetzwerkDir) - 9, 1) <> "\" Then
        Msg = LoadResString(2459 + Sprache)                  '2459=Sie m�ssen die Datei fotos.mdb ausw�hlen
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo Begin
    End If
    If StrComp(Right(NetzwerkDir, 9), "fotos.mdb", vbTextCompare) = 0 Then
        AppPath = Mid(NetzwerkDir, 1, Len(NetzwerkDir) - 10)
    Else
        Msg = LoadResString(2459 + Sprache)
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        GoTo Begin
    End If
End Sub

Function GetShellError(lErrorCode As Long) As String
    Const SE_ERR_FNF = 2&, SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&, SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&, SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&, SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&, SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&, ERROR_BAD_FORMAT = 11&

    Select Case lErrorCode
        Case SE_ERR_FNF
            GetShellError = "File not found"
        Case SE_ERR_PNF
            GetShellError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            GetShellError = "Access denied"
        Case SE_ERR_OOM
            GetShellError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            GetShellError = "DLL not found"
        Case SE_ERR_SHARE
            GetShellError = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            GetShellError = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            GetShellError = "DDE Time out"
        Case SE_ERR_DDEFAIL
            GetShellError = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            GetShellError = "DDE busy"
        Case SE_ERR_NOASSOC
            GetShellError = "No association for file extension"
        Case ERROR_BAD_FORMAT
            GetShellError = "Invalid EXE file or error in EXE image"
        Case Else
            GetShellError = "Unknown error"
    End Select
End Function

'Public Function GEOKoordinatenUmrechnenXMP()                                       'Gerbing 08.04.2019
'    zB gstrLatXMP 50,38.7309456N -> 50.64551575
'    zB gstrLongXMP 11,53.9826786E -> 11.89971130
'    Das ist n�tig damit die GEO-Positionen von OpenStreetMap verstanden werden
'    Dim Grad As String
'    Dim Minuten As Double
'    Dim ESWN As String                                                      'East South West Nord
'    Dim Nachkomma As String
'    Dim pos As Integer
'
'    GEOKoordinatenUmrechnenXMP = 0
'    pos = InStr(1, gstrLatXMP, ",")
'    Grad = Mid(gstrLatXMP, 1, pos - 1)
'    Minuten = Mid(gstrLatXMP, pos + 1, Len(gstrLatXMP) - pos - 1)
'    ESWN = Mid(gstrLatXMP, Len(gstrLatXMP), 1)
'    Nachkomma = Minuten / 60
'    pos = InStr(1, Nachkomma, ",")
'    If pos <> 0 Then
'        Nachkomma = Mid(Nachkomma, 1, pos - 1)
'    End If
'    gstrLat = ""
'    If ESWN <> "N" Then
'        gstrLat = "-"                                                       '- auf der S�dhalbkugel
'    End If
'    gstrLat = gstrLat & Grad & "." & Nachkomma
'
'    pos = InStr(1, gstrLongXMP, ",")
'    Grad = Mid(gstrLongXMP, 1, pos - 1)
'    Minuten = Mid(gstrLongXMP, pos + 1, Len(gstrLongXMP) - pos - 1)
'    ESWN = Mid(gstrLongXMP, Len(gstrLongXMP), 1)
'    Nachkomma = Minuten / 60
'    pos = InStr(1, Nachkomma, ",")
'    If pos <> 0 Then
'        Nachkomma = Mid(Nachkomma, 1, pos - 1)
'    End If
'    gstrLong = ""
'    If ESWN <> "E" Then
'        gstrLong = "-"                                                       '- westlich von Greenwich
'    End If
'    gstrLong = gstrLong & Grad & "." & Nachkomma
'End Function

Public Function GEOKoordinatenUmrechnenXMP()                                       'Gerbing 08.04.2019
    'zB gstrLatXMP 50,38.7309456N -> 50.64551575
    'zB gstrLongXMP 11,53.9826786E -> 11.89971130
    'Das ist n�tig damit die GEO-Positionen von OpenStreetMap verstanden werden
    Dim Grad As String
    Dim Minuten As String                                                   'Gerbing 04.07.2019
    Dim MinutenDouble As Double                                             'Gerbing 04.07.2019
    Dim ESWN As String                                                      'East South West Nord
    Dim Ergebnis As String
    Dim pos As Integer
    Dim locale_id As Long                                                   'Gerbing 04.07.2019
    
    GEOKoordinatenUmrechnenXMP = 0
    pos = InStr(1, gstrLatXMP, ",")                                         'das "," kommt in deutscher und englischer Systemsprache
    Grad = Mid(gstrLatXMP, 1, pos - 1)
    Minuten = Mid(gstrLatXMP, pos + 1, Len(gstrLatXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLatXMP, Len(gstrLatXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLat = ""
    If ESWN <> "N" Then
        gstrLat = "-"                                                       '- auf der S�dhalbkugel
    End If
    gstrLat = gstrLat & Ergebnis                                            'Gerbing 04.07.2019
    '---------------------------
    pos = InStr(1, gstrLongXMP, ",")
    Grad = Mid(gstrLongXMP, 1, pos - 1)
    Minuten = Mid(gstrLongXMP, pos + 1, Len(gstrLongXMP) - pos - 1)
    'Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
    'sonst kommt bei MinutenDouble / 60 Ergebnis=0
    If LocaleInfo(locale_id, LOCALE_SDECIMAL) = "," Then
        Minuten = Replace(Minuten, ".", ",")
    End If
    ESWN = Mid(gstrLongXMP, Len(gstrLongXMP), 1)
    MinutenDouble = CDbl(Minuten)
    MinutenDouble = MinutenDouble / 60
    Ergebnis = Grad + MinutenDouble
    Ergebnis = Replace(Ergebnis, ",", ".")                                  'Gerbing 04.07.2019
    gstrLong = ""
    If ESWN <> "E" Then
        gstrLong = "-"                                                       '- westlich von Greenwich
    End If
    gstrLong = gstrLong & Ergebnis
End Function

Private Sub SpracheFestlegen()
    Dim strTemp As String
    Dim strPrimaryKey As String
    Dim SQL As String
    Dim Msg As String
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

    Set DBsql = New ADODB.Connection
    If gblnSQLServerVersion = True Then
        With DBsql
            .Provider = "SQLOLEDB.1"
            '.Provider = "SQLNCLI10.1" 'SQL Server Native Client
            .Properties("Persist Security Info").Value = False
            .Properties("Initial Catalog").Value = PublicSQLDatabase
            .Properties("Data Source").Value = PublicSQLServer
            '   Falls die Windows-Authentifizierung verwendet werden soll, mu� "SSPI" benutzt werden
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
        'DBsql.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath & "\fotos.mdb"
        DBsql.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AppPath & "\fotos.mdb"
        DBsql.mode = adModeReadWrite
        DBsql.Open                                                      'Gerbing 04.03.2012 hier kommt runtime error wenn fotos.mdb fehlt
        PublicDatagridCaption = AppPath & "\fotos.mdb"
    End If
    On Error Resume Next
        SQL = "SELECT * From fotos WHERE not filename Is Null;"

        On Error Resume Next
        'On Error GoTo 0
        'On Error GoTo QUERYERR
        If rstsql Is Nothing Then
            Set rstsql = New ADODB.Recordset
        Else
            rstsql.Close
        End If
        ERR.Number = 0
        With rstsql
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenForwardOnly
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
        '-------------------------------------------------------------
        If ERR.Number <> 0 Then
            Call WriteGlL(0)     'R�ckschreiben deutsch in fotos.ini
            Sprache = 0
            Call GlL                                                                        'Gerbing 02.09.2008
            If PublicLanguage = "1" Then                                                    'Gerbing 04.12.2011
                Call VierUrsachenF�rSchreibsperre
                End
            End If
        Else
            Call WriteGlL(1)     'R�ckschreiben english in fotos.ini
            Sprache = 3000
            Call GlL                                                                        'Gerbing 02.09.2008
            If PublicLanguage = "0" Then                                                    'Gerbing 04.12.2011
                Call VierUrsachenF�rSchreibsperre
                End
            End If
        End If
        rstsql.Close
    '-------------------------------------------------------------------------------------
    If gblnSQLServerVersion = False Then
        'es ist keine SQL-server version - bei SQL server gibt es kein dbHyperlinkField
        On Error GoTo 0
        SQL = "select * from Fotos"
        With rstsql
            .Source = SQL
            .ActiveConnection = DBsql
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With

        'untersuchen ob ein dbHyperlinkField dabei ist
        For n = 0 To rstsql.Fields.Count - 1
            If rstsql.Fields(n).Attributes() = 32770 Then                 '32770=dbHyperlinkField
                'erstes Item in der Collection hat Nummer 1
                'HyperlinkFieldColumns.Add rstsql.Fields(n).Name
                HyperlinkFieldColumns.Add n                               'beispielsweise Spalte 19
            End If
        Next n
        'rstsql.Close
        'On Error GoTo 0
    End If
    If gblnSQLServerVersion = False Then
        Set rstsql = DBsql.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, "Fotos")) '2529=fotos
        If rstsql.EOF = True Then
            'Msg = "Seit Version 13.4.0 verlangt das Programm in Tabelle 'fotos' Spalte 'Dateiname' einen Prim�rschl�ssel. Dieser wird jetzt erzeugt." & vbnewline
            'msg = msg & "Diese Operation wird nur dann erfolgreich sein, wenn in der Tabelle 'fotos' Spalte 'Dateiname' keine Duplikate vorkommen." & vbnewline
            'msg = msg & "Wenn die Operation nicht erfolgreich ist, m�ssen Sie zuvor die Duplikate entfernen." & vbnewline
            'msg = msg & "Benutzen Sie dazu eine fr�here Version von fotosmdb.exe als 13.3.4"
            Msg = LoadResString(1825 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(1826 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(1827 + Sprache) & vbNewLine
            Msg = Msg & LoadResString(1828 + Sprache) & vbNewLine
            MsgBox Msg
            'SQL = "Create UNIQUE INDEX Dateiname ON fotos (Dateiname)  WITH PRIMARY"
            SQL = "Create UNIQUE INDEX " & LoadResString(1028 + Sprache) & " ON Fotos(" & LoadResString(1028 + Sprache) & ") WITH PRIMARY"
            On Error Resume Next
            DBsql.Execute SQL
            If ERR.Number <> 0 Then
                Msg = "error number=" & ERR.Number & vbNewLine
                Msg = Msg & "errortext=" & ERR.Description
                MsgBox Msg
                End
            End If
        Else
            strPrimaryKey = rstsql.Fields("COLUMN_NAME").Value
            If StrComp(LoadResString(1028 + Sprache), strPrimaryKey, vbTextCompare) <> 0 Then       '1028=Dateiname
                'MsgBox "Die Spalte Dateiname ist nicht der Prim�rschl�ssel. Das Programm wird beendet."
                MsgBox LoadResString(1824 + Sprache)
                End
            End If
        End If
    End If
End Sub

Public Sub VierUrsachenF�rSchreibsperre()                                                'Gerbing 02.09.2008
    Dim Msg As String
    
    'vier m�gliche Ursachen
    'Msg = gstrFotosIniAnwendungsOrdner & "\GERBING Fotoalbum 14\fotos.ini" & vbNewLine
    Msg = gstrFotosIniAnwendungsOrdner & "\fotos.ini" & vbNewLine
    'msg = msg & "Die Datei ist schreibgesch�tzt. Sie m�ssen f�r Schreibrechte sorgen, damit �nderungen an dieser Datei gemacht werden k�nnen." & vbnewline
    Msg = Msg & LoadResString(2275 + Sprache) & vbNewLine
    'msg = msg & "Es gibt vier m�gliche Ursachen f�r den Lesemodus:" & vbnewline
    Msg = Msg & LoadResString(2133 + Sprache) & vbNewLine
    'msg = msg & "1. Das Dateiattribut 'Schreibgesch�tzt' ist gesetzt" & vbnewline
    Msg = Msg & LoadResString(2134 + Sprache) & vbNewLine
    'msg = msg & "2. Sie arbeiten mit einem Benutzerkonto ohne Administrator-Rechte f�r Ihren PC" & vbnewline
    Msg = Msg & LoadResString(2135 + Sprache) & vbNewLine
    'msg = msg & "3. Sie arbeiten mit einer CD oder DVD" & vbnewline
    Msg = Msg & LoadResString(2136 + Sprache) & vbNewLine
    'msg = msg & "4. Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte" & vbnewline & vbnewline
    Msg = Msg & LoadResString(2137 + Sprache) & vbNewLine & vbNewLine
    MsgBox Msg, , LoadResString(1119 + Sprache)
End Sub

Private Sub txtFehlerU_Click(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    If LogNichtBenutzbar = True Then Exit Sub                                               'Gerbing 02.09.2008
    If FehlerGefunden = False Then Exit Sub
    
    On Error Resume Next
    If txtFehlerU.Text <> LoadResString(1382 + Sprache) Then                            'Gerbing 10.05.2013 1382=kein Fehler gefunden
        'es wurden Fehler gefunden
        'If Form1.Pr�fenNummer = "Pr�fen3" Then
        If Form1.Pr�fenNummer = LoadResString(1459 + Sprache) Then
            'Pr�fen3 war dran
            If NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count <> 0 Then                  'Gerbing 26.10.2013
                'LogLesen.TxtU.Text = NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count & " Fotos/Videos wurden gefunden"   'Gerbing 26.10.2013
                LogLesen.TxtU.Text = NachPr�fen3Aufnehmen.KollZus�tzlicheDateien.Count & LoadResString(3153 + Sprache)  'Gerbing 26.10.2013
                LogLesen.btnAbbrechen.Default = True
                LogLesen.btnGefundeneAufnehmen.TabIndex = 0                             'Gerbing 30.05.2014
                LogLesen.Show 1
            End If
        End If
        'If Form1.Pr�fenNummer = "Pr�fen2" Then
        If Form1.Pr�fenNummer = LoadResString(1444 + Sprache) Then
            'Pr�fen2 war dran
            If AnzahlFehlerPr�fen2 <> 0 Then
                LogLesen.TxtU.Text = AnzahlFehlerPr�fen2 & LoadResString(3153 + Sprache)                                'Gerbing 26.10.2013
                LogLesen.Show 1
            End If
        End If
        'If Form1.Pr�fenNummer = "Pr�fen1" Then                            'Gerbing 16.01.2006
        If Form1.Pr�fenNummer = LoadResString(1443 + Sprache) Then
            'Pr�fen1 war dran
            If NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count <> 0 Then
                LogLesen.TxtU.Text = NachPr�fen1L�schen.lstZus�tzlicheDateien.ListItems.Count & LoadResString(3153 + Sprache)  'Gerbing 26.10.2013
                LogLesen.Show 1
            End If
        End If
    End If
    
    'wie Button Reset                                               'Gerbing 04.02.2008
    Call SpaltenbreiteMerken

    rsDataGrid.Requery
    Set DBGridNeu.DataSource = rsDataGrid
    DBGridNeu.ReBind

    Call SpaltenbreiteWiederherstellen
    DBGridNeu.Caption = PublicDatagridCaption
    DBGridNeu.AllowUpdate = False
End Sub

Private Function RemoveNulls(OriginalString As String) As String
    Dim pos As Long
    pos = InStr(OriginalString, Chr$(0))
    If pos > 1 Then
        RemoveNulls = Mid$(OriginalString, 1, pos - 1)
    Else
        RemoveNulls = OriginalString
    End If
End Function

' Return a piece of locale information.                             'Gerbing 04.07.2019
Private Function LocaleInfo(ByVal locale As Long, ByVal lc_type As Long) As String
Dim Length As Long
Dim buf As String * 1024

    Length = GetLocaleInfo(locale, lc_type, buf, Len(buf))
    LocaleInfo = left$(buf, Length - 1)
End Function

Private Sub WMP_DeviceSyncError(ByVal pDevice As WMPLibCtl.IWMPSyncDevice, ByVal pMedia As Object)
    MsgBox "WMP_DeviceSyncError"                                    'Gerbing 15.11.2019
End Sub

Private Sub WMP_Error()                             'Gerbing 15.11.2019
    Dim Msg As String
    
    If WMP.URL <> "" Then
        Msg = WMP.URL & NL
'        Msg = Msg & "Es ist ein Fehler beim Abspielen der Datei aufgetreten." & NL
'        Msg = Msg & "Kontrollieren Sie ob die Pfadangabe richtig ist." & NL
'        Msg = Msg & "Kontrollieren Sie, ob sich die Datei au�erhalb von diesem Programm abspielen l��t." & NL & NL
        Msg = Msg & LoadResString(2283 + Sprache) & NL
        Msg = Msg & LoadResString(2284 + Sprache) & NL
        Msg = Msg & LoadResString(2285 + Sprache) & NL & NL
        'MsgBox Msg
        MessageBoxW 0, StrPtr(Msg), StrPtr("GERBING Fotoalbum"), vbInformation
        WMP.Controls.stop                                                       'Gerbing 15.11.2019
        blnMediaPlayerError = True                                              'Gerbing 15.11.2019
    End If
End Sub

Private Sub WMP_MediaError(ByVal pMediaObject As Object)
    MsgBox "MediaPlayer1_MediaError"                                            'Gerbing 15.11.2019
End Sub

Private Sub WMP_ModeChange(ByVal ModeName As String, ByVal NewValue As Boolean)
    MsgBox "MediaPlayer1_ModeChange"
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)                         'Gerbing 15.11.2019
    Dim blnMustUpdate As Boolean

    'player .playState
    'Possible Values
    'This property is a read-only Number (long).
    'Value   State   Description
    '0   Undefined   Windows Media Player is in an undefined state.
    '1   Stopped     Playback of the current media item is stopped.
    '2   Paused      Playback of the current media item is paused. When a media item is paused, resuming playback begins from the same location.
    '3   Playing     The current media item is playing.
    '4   ScanForward The current media item is fast forwarding.
    '5   ScanReverse The current media item is fast rewinding.
    '6   Buffering       The current media item is getting additional data from the server.
    '7   Waiting     Connection is established, but the server is not sending data. Waiting for session to begin.
    '8   MediaEnded  Media item has completed playback.
    '9   Transitioning   Preparing new media item.
    '10  Ready       Ready to begin playing.
    '11  Reconnecting    Reconnecting to stream.

'    Debug.Print "NewState=" & NewState
'    Debug.Print "FotoDatei=" & strFotoDatei
    If NewState = 3 Then                                                        '3=playing
        glngStartMillisek = timeGetTime                                         'Gerbing 30.05.2019
        Form1.pintBreite = WMP.currentMedia.imageSourceWidth
        Form1.pintHoehe = WMP.currentMedia.imageSourceHeight
        lngVideoDuration = WMP.currentMedia.duration
        On Error GoTo 0
        WMP.Controls.stop
    End If
    If NewState = 8 Then                                                        '8=MediaEnded 'Gerbing 07.05.2013
        glngEndMillisek = timeGetTime                                           'Gerbing 30.05.2019
        If glngEndMillisek - glngStartMillisek < 300 Then                       'Gerbing 30.05.2019
            Debug.Print "MedaiEnded nach millisekunden=" & glngEndMillisek - glngStartMillisek 'Gerbing 30.05.2019
            Call WMP_Error                                                      'Gerbing 30.05.2019
        End If                                                                  'Gerbing 30.05.2019
    End If
    If NewState = 1 Then                                                       '1=Stopped 'Gerbing 07.05.2013
        blnMediaPlayerStopped = True
    End If
End Sub

Private Sub WMP_Warning(ByVal WarningType As Long, ByVal Param As Long, ByVal Description As String)
    MsgBox Description                                                          'Gerbing 15.11.2015
End Sub
