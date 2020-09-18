VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "editctlsu.ocx"
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form DiashowForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "GERBING Diashow"
   ClientHeight    =   5856
   ClientLeft      =   120
   ClientTop       =   444
   ClientWidth     =   12744
   Icon            =   "DIASHOW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   5856
   ScaleWidth      =   12744
   Begin VB.CommandButton btnSucheStarten 
      Caption         =   "Suche starten"
      Height          =   612
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1452
   End
   Begin EditCtlsLibUCtl.TextBox txtSearchWords 
      Height          =   372
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   7332
      _cx             =   12933
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
      CueBanner       =   "DIASHOW.frx":038A
      Text            =   "DIASHOW.frx":03AA
   End
   Begin CBLCtlsLibUCtl.ListBox List1UnsortedU 
      Height          =   3972
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5772
      _cx             =   10181
      _cy             =   7006
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048811
      DontRedraw      =   0   'False
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
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertMarkStyle =   1
      IntegralHeight  =   0   'False
      ItemHeight      =   -1
      Locale          =   1024
      MousePointer    =   0
      MultiColumn     =   0   'False
      MultiSelect     =   0
      OLEDragImageStyle=   0
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
   Begin VB.CheckBox chkSortAsDragAndDrop 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sortierfolge wie bei Drag&&Drop"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5772
   End
   Begin VB.CommandButton cmdKeinNeustart 
      Caption         =   "&Container leeren"
      Height          =   612
      Left            =   6000
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdStartAfterDragDrop 
      Caption         =   "S&tart"
      Height          =   612
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6240
      Top             =   480
   End
   Begin CBLCtlsLibUCtl.ListBox List1U 
      Height          =   3972
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5772
      _cx             =   10181
      _cy             =   7006
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048800
      DontRedraw      =   0   'False
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
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertMarkStyle =   1
      IntegralHeight  =   0   'False
      ItemHeight      =   -1
      Locale          =   1024
      MousePointer    =   0
      MultiColumn     =   0   'False
      MultiSelect     =   0
      OLEDragImageStyle=   0
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   -1  'True
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suchen nach:"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Label lblDateinamen 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Container Dateinamen vom Explorer"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   5772
   End
   Begin VB.Menu mnuHilfe 
      Caption         =   "&Hilfe"
   End
   Begin VB.Menu mnuBeenden 
      Caption         =   "&Benden"
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "&Version"
   End
End
Attribute VB_Name = "DiashowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Änderungen:
'13.07.2003 Die Größe der Form DiashowForm darf nicht durch Ziehen an den Seitenkanten verändert werden können
'13.07.2003 Immer wenn die HilfeBoxForm verschwindet muß ein DiashowForm.Image1.Refresh gemacht werden.
'           Nötig ist das für Fotos, die beim ersten Laden größer sind als die Bildschirmgröße.
'13.09.2003 Bei der automatischen Bildanzeige wird eine Zufallsreihenfolge eingeführt
'14.12.2003 Bei F11 wird die Bildbreite x Bildhöhe jeweils um 4 Pixel zu hoch ausgerechnet, also
'           subtrahiere ich jeweils 4
'14.12.2003 Wenn mit einer Steuerdatei (*.dia) gearbeitet wird, bleibt die Sortierreihenfolge so,
'           wie es die Steuerdatei vorgibt, dazu ist eine zusätzliche Listbox mit Sorted=False nötig,
'           die genau über der Listbox mit Sorted=True liegt
'16.12.2003 Immer bei Doppelklick auf eine Zeile in ListBoxForm.ListBoxNeu wird das Bild nicht in der Mitte
'           zentriert. Eine MsgBox würde helfen, ist aber eine Scheißlösung.
'16.12.2003 Im XP ist die Sortierreihenfolge bei Dateinamen, wo Zahlen enthalten sind, anders als bei allen
'           vorangegangenen Windows-Versionen. Das würde mein bisheriges Konzept, der nachträglichen
'           Einfügung ähnlicher Bilder total zerstören. Ab XP Service Pack 1 gibt es einen registry key
'           mit dem die bisher gewohnte Reihenfolge hergestellt werden kann.
'           Das muss unbedingt in der Hilfe-Datei beschrieben werden.
'22.12.2003 Wenn das Bild größer ist als Diashowform.Screenwidth/ScreenHeight
'           oder wenn die Bildbreite >= 800 Pixel ist
'           oder wenn die Bildhöhe >= 600 Pixel ist
'           wird es umgesetzt nach Image1 mit Stretch = True
'10.02.2004 Die Änderung vom 22.12.2003 wird zurückgezogen, weil es zu Qualitätsverlust kommt, wenn ein Bild
'           Von 800x600 auf 1024x768 auseinandergezogen wird. Man sieht die Blöcke.
'08.04.2004 Es geht noch einfacher als bisher, wenn man eine Liste von Dateinamen zB aus dem Explorer auf
'           das Programm Icon von Diashow.exe zieht. Da wird eine Liste von Dateinamen angeboten, wobei
'           A. bei Namen mit Leerzeichen - der Dateiname in doppeltes Hochkomma vorn und hinten eingeschlossen
'           ist und ein Leerzeichen den Dateiname vom nächsten trennt
'           B. bei Namen ohne Leerzeichen - Ein Leerzeichen den Dateiname vom nächsten trennt
'           Leider ist aber bei maximal 2048 Bytes Schluss, dann kommt die Message
'           "Der Zugriff auf das angegebene Gerät, den Pfad oder die Datei wurde verweigert"
'           dann will ich aber wenigstens noch einen Klick beim Starten einsparen.
'           Es zeigt sofort den Drag&Drop-Container.
'           Mit Steuerdatei geht es nur noch über Aufruf-Parameter.
'03.08.2004 Bei Drücken von F10 kann eine neue Funktion eingestellt werden:
'           Alle Bilder mit temporär veränderter Helligkeit anzeigen
'19.03.2005 Auch bei Doppelklick auf einen Namen einer Bilddatei soll Diashow.exe benutzbar sein
'           Dazu muss man die doppelten Hochkomma am Anfang und Ende des Namenstrings entfernen
'25.03.2005 Ergänzung zu 14.12.2003
'           Ein Bild hat dann 4 Pixel mehr Bildbreite bzw Bildhöhe, wenn die Eigenschaft Borderstyle=1
'           gesetzt ist. Also Borderstyle=0 benutzen.
'           Nicht zu fassen, dass mir das erst jetzt auffällt.
'05.05.2005 Es gibt ein neues Formular ImageForm das wird ausschließlich für das Plazieren von
'           Image1 und Picture1 benutzt. Dafür kann ich das Formular DiashowForm mit G.ico darstellen
'           was sonst nicht möglich war.
'05.05.2005 Ab jetzt entfällt die MsgBox nach Doppelklick auf einen Dateiname in ListboxForm
'           Schalter blnListBoxNeuDblClick wird benutzt
'08.11.2005 Beginn der Arbeiten für Englische Version und beliebige Sprachen
'           durch Benutzung einer .res-Datei und LoadResString
'           Die .res-Datei der Entwicklungsumgebung wird bei der Exe
'           zum festen Bestandteil der Exe. Darum müßte man bei Benutzung von verschiedenen .res-Dateien
'           pro Sprache, jedesmal eine neue Exe erzeugen, mit der jeweiligen .res-Datei.
'           Es gibt einen Ausweg, wie man alle Sprachen in eine .res-Datei legen kann.
'           alle Zahlen ab 1000 gelten für deutsche Label.Caption und Form.Caption
'           alle Zahlen ab 2000 gelten für deutsche Msgbox-Texte und Tooltips
'           alle Zahlen ab 3000 gelten für deutsche Commandbuttons, Checkboxen, Options, Frames
'           alle Zahlen ab 4000 entsprechen wie ab 1000 aber englische Label.Caption und Form.Caption
'           alle Zahlen ab 5000 entsprechen wie ab 2000 aber englische Msgbox-Texte und Tooltips
'           alle Zahlen ab 6000 entsprechen wie ab 3000 aber englische Commandbuttons, Checkboxen, Options, Frames
'           LoadResString(1090 + Sprache) benutzen
'           Sprache = 0 bei deutsch
'           Sprache = 3000 bei englisch
'           Die gewählte Sprache wird in die Datei fotos.ini eingetragen
'15.02.2006 Neue Funktion:
'           Beim Fotografieren mit der digitalen Kamera sollen alle gemachten Fotos einer ersten Sichtung mit
'           Diashow.exe unterzogen werden. Die nicht speicherwürdigen Fotos sollen auf einfache Art markiert
'           und dann gelöscht werden können.
'           Lösung: Listbox mit Häkchen benutzen
'           ListBoxForm.ListBoxNeu.Style = 1
'15.02.2006 Verbesserung:
'           Enter-Taste soll wie Doppelklick wirken
'29.07.2006 Neue Funktion:
'           Jedes Bild soll mit individuell vergrößertem/verkleinertem Bildausschnitt auch wiederholt
'           angezeigt werden können. Bisher wird bei jedem Wechsel auf das nächste Bild die Standard-Funktion
'           benutzt (Bild zentriert darstellen, so daß es als Ganzes auf den Bildschirm paßt)
'           Lösung: es wird ein Array eingeführt und ein Private Type, wo jede Strukturzeilennummer zu einer
'           Listboxzeilennummer gehört.
'           Auf das Formular ZoomForm wirkt diese Funktion nicht.
'09.01.2007 Seit das Bild auf Imageform dargestellt wird, hat sich ein Fehler unbemerkt erhalten
'           und zwar Starten mit Diashow-Datei, dann F8 drücken, dann den Commondialog abbrechen,
'           da bleibt ein nichtbedienbares graues Fenster
'09.01.2007 Eine Diashow-Datei als command line argument gibt man so an M:\ABC\dias 1912.dia
'           Ein Bild als command line argument gibt man so an "M:\ABC\Bild.jpg" also in doppelte Hochkomma
'           einschließen
'16.02.2007 Neue Funktionen
'           Abgekupfert aus fotos.exe
'           Im Formular ListBoxForm gibt es nur noch einen Button 'Aktion wählen und ausführen...'
'           daraufhin kommt ein Popup-Menü mit folgenden auswählbaren Aktionen:
'           -Öffne die mit der Dateinamen-Erweiterung verknüpfte Anwendung
'           -Öffne das Druckprogramm für die aktuelle Datei
'           -Öffne das Fenster 'Neue Email senden'
'           -Öffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
'           -Lösche alle mit Häkchen markierten Dateien
'           -Kopiere alle mit Häkchen markierten Dateien. Dabei entstehen keine Duplikate, gleichnamige werden
'            überschrieben
'           Es gibt ab sofort die Checkbox chkFensterGrößeÄnderbar
'20.02.2007 im Formular ListBoxForm werden EXIF/IPTC-Felder angezeigt
'           Für IPTC ist Verweis auf iptc info reader nötig
'           ...\windows\system32\iptcinfo.dll
'           freeware from www.gasanow.net
'           Für EXIF wird der Klassenmodul clsEFIF.cls benutzt
'14.03.2007 Verbesserung alle Versionen:
'           Neue Hilfe-Dateien im HTML-Format, weil Windows Vista das Winhelp-Format nicht mehr unterstützt
'           zB anstelle Diashow.hlp gibt es jetzt Diashow.chm
'19.03.2007 Fehlerkorrektur:
'           nach dem Markieren von Dateien zum Löschen und dem Durchführen des Löschens
'           kam bisher beim Klick ins Schwarze die MsgBox 'Zoom ist nicht möglich'
'07.05.2007 Verbesserung alle Versionen:
'           Gasanov EXIF OCX hat einen Fehler, es läßt bei Olympus Fotos frisch aus der Kamera einen Teil der
'           Felder weg. Erst nach Bearbeitung mit PSP9 sind alle Felder da.
'           Es wird ab jetzt ein Klassenmodul clsEFIF.cls benutzt
'07.05.2007 Fehlerkorrektur:
'           Wenn man das F5-Fenster mit Schließkreuzchen schließt, blockierten bisher alle Funktionstasten
'26.07.2007 erfolgloser Test ob ein Foto beim Ziehen mit der Maus ohne großes Ruckeln verschoben werden kann.
'           Für meinen Geschmack ruckelt es zu stark.
'           weitere Untersuchungen: mit Beispielen aus dem Internet
'           Es ruckelt nicht, wenn man eine Picturebox nimmt und die Methode Picture.Move benutzt
'           aber bei einer Picturebox ist schneller Schluß beim ZoomIn
'           Ergebnis: es bleibt alles beim Alten
'29.07.2007 Verbesserungen:
'           1. In der ListboxForm.Caption steht ab sofort 'Dateiname des angezeigten Bildes' ....
'           2. Man soll am Mouse-Icon erkennen, ob die Rechtecklupe aktiv ist (SquareZoom.ico).
'           Wenn das Bild mit der linken Maustaste verschoben werden kann, soll man das am Mouse-Icon erkennen.
'           Schon beim MouseDown kommt FourArrows.ico.
'           Wenn beim Bildverschieben die verschobenen Positionen gemerkt werden sollen, wird
'           FourArrowsSave.ico benutzt.
'           3. Kenntlich machen, dass EXIF/IPTC nur vom aktuell angezeigten Bild anzeigbar sind
'           Wenn man EXIF/IPTC eingeschaltet hat und wieder Gleichheit zwischen dem aktuell angezeigten Bild und
'           der aktuellen Zeile in ListBoxNeu herstellt, dann soll EXIF/IPTC erneut gelesen werden
'16.08.2007 4. Wenn keine Gleichheit zwischen dem aktuell angezeigten Bild und der aktuellen Zeile in ListBoxNeu
'           herrscht, soll der Nutzer den Hinweis erhalten, dass er Doppelklicken soll oder die Enter-Taste
'           benutzen soll.
'31.08.2007 Horizontale Scrollbar an Listbox List1 anbringen
'23.10.2007 Verbesserung:
'           Wenn im Drag&Drop-Container nur eine Datei steht, soll F2/F3 die nächste/vorhergehende Datei im
'           gleichen Ordner anzeigen. Es wird der gleiche Ordner durchsucht, wie der Dateiname im
'           Drag&Drop-Container. Nur die erlaubten Dateitypen werden untersucht.
'           Wenn im Drag&Drop-Container nur eine Datei steht, ist die Einstellung
'           'Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten' nicht benutzbar
'23.10.2007 Fehlerkorrektur:
'           Wenn man WertxForm schließt, blieben bisher die letzten Änderungen nicht erhalten
'14.11.2007 Neue Funktion: nicht für Windows98
'           Es gibt eine neue Form frmVerknüpfeDateitypen, die wird angezeigt über das Menü Registry
'           Mit Hilfe dieser Form kann man alle oder ausgewählte Dateitypen
'           BMP CUR DIB EMF GIF ICO JPG WMF mit diashow .exe verknüpfen durch Editieren der Registry
'           wie im Rechtsklick-Menü des Windows Explorers.
'28.11.2007 Fehlerkorrektur:
'           siehe auch 23.10.2007
'           Wenn im Drag&Drop-Container nur eine Datei stand (durch Doppelklick) und anschließend nach F8 habe
'           ich den ganzen Ordner in den Drag&Drop-Container gezogen, dann war das Bild aus Schritt 1
'           stets zweimal im Drag&Drop-Container.
'           Lösung: ich muss zum Vergleich strComp benutzen.
'28.11.2007 Verbesserung:
'           Immer wenn nur eine Bilddatei gestartet wird, egal ob über Doppelklick oder durch Ziehen einer
'           einzelnen Bilddatei in den Drag&Drop-Container, dann durchsuche ich den gesamten Ordner nach
'           weiteren Bilddateien.
'11.12.2007 Fehlerkorrektur zu 14.11.2007
'06.01.2008 Fehlerkorrektur zur IPTC-Anzeige
'           es gibt Felder, die länger sein können als eine Zeile in der Listbox lstExifIptc
'           diese Felder muss man in mehrere Zeilen zerlegen
'10.01.2008 Fehlerkorrektur zur IPTC-Anzeige
'           IPTC.OriginatingProgram war vergessen abzufragen
'12.03.2008 Fehlerkorrektur, Fehler tritt nur in der Entwicklungsumgebung auf, wenn mit Befehlszeilenargument
'           und die angegebene Datei wird nicht gefunden
'12.03.2008 Fehlerkorrektur beim Löschen markierter Dateien, wenn ich keine Admin-Rechte habe
'30.07.2008 Verbesserung
'           Wenn das Programm durch Doppelklick auf eine einzelne Datei gestartet wird, soll der Nutzer
'           auswählen können, ob er nur den aktuellen Ordner durchsuchen will, oder auch alle Unterordner
'18.08.2008 In der Listboxform musste man bisher zweimal klicken um zwischen der Anzeige EXIF/IPTC umzuschalten.
'           Ab jetzt genügt ein Klick.
'16.12.2008 Version 13.3.11
'           Bei 'Fenstergröße änderbar' soll es ab jetzt möglich sein
'           Min, Max und Close auszuwählen.
'           Man muss dazu at design time 'ImageForm.Controlbox = True' einstellen
'16.12.2008 Fehlerkorrektur bei Benutzung einer *.dia-Datei
'           Es kam nur ein schwarzer Bildschirm. Erst nach Drücken von F8 ging es weiter.
'24.01.2009 Eine Listbox ist bei 32767 Einträgen am Ende. Ersatz durch eine Collection ist mir zu aufwändig.
'           weil es 4 Listboxen gibt wo ich das ersetzen müßte
'           DiashowForm.lstFilesImGleichenOrdner
'           DiashowForm.List1
'           DiashowForm.List1Unsorted
'           ListBoxForm.ListBoxNeu
'           Ich bringe stattdessen eine MsgBox, dass mehr als 32767 Einträge nicht erlaubt sind.
'16.03.2009 Ich habe aus fotos.exe übernommen
'           Beim Bildladen vergrößern auf Vollbild
'           keine
'           ab Breite 640 Pixel
'           ab Breite 800 Pixel
'           ab Breite 1024 Pixel
'           ab Breite 1024 Pixel oder Höhe 768 Pixel
'03.09.2009 Fehlerkorrektur bei Benutzung von breiten Flachbildschirmen zB 1920x1200
'           wird ein Bild mit Größenverhältnis 4:3 zu einem Ei zusammengedrückt
'24.09.2009 Fehlerkorrektur bei Benutzung von breiten Flachbildschirmen zB 1920x1200
'           Ich muss zusätzlich zum Breiten/Höhen-Verhältnis des Bildes BHV
'           das Breiten/Höhen-Verhältnis des Bildschirms SGVH auswerten
'           If BHV < SGVH Then
'               'das Bild ist zu hoch und zu schmal
'           If BHV > SGVH Then
'               'das Bild ist zu niedrig und zu breit
'07.12.2009 Verbesserung:
'           Es ist ärgerlich, wenn ich Häkchen vor Dateinamen setze, weil ich diese Bilder löschen will, dann F8
'           drücke und alle angemerkte Häkchen gehen verloren. Es soll vor Ausführung von F8 ein Sicherheitshinweis
'           kommen, dass noch Häkchen angemerkt sind, Weiterarbeiten und Häkchen ignorieren, ja oder nein
'19.12.2009 Verbesserung für Nutzer mit breiten Flachbildschirmen:
'           Den Standardwert ändern bei 'Beim Bildladen vergrößern auf Vollbild'
'           bisher:     keine
'           ab jetzt:   ab Bildbreite 1024 Pixel oder Bildhöhe 768 Pixel
'28.12.2009 Verbesserung:
'           In der ListBoxForm wird der Breich für EXIF/IPTC-Daten breiter gezogen, damit fast alle EXIF-Werte reinpassen, ohne
'           den horizontalen Scrollbalken benutzen zu müssen
'28.12.2009 Fehlerkorrektur:
'           Bisher ist bei ListBoxForm.ListBoxNeu_Click
'           txtEXIFInfo.Enabled = False gesetzt worden
'           Dadurch mußte bei Anzeige der EXIF-Daten stets erst nochmal auf die aktuelle Zeile geklickt werden, damit
'           die EXIF-Werte nicht grau dargestellt sind
'22.01.2010 Verbesserung:
'           Es wird eine Checkbox chkSortAsDragAndDrop gebraucht, damit man unterscheiden kann, ob beim Drag&Drop die angebotene
'           Sortierung unverändert bleibt oder ob die angebotenen Dateinamen alphabetisch aufsteigend sortiert werden sollen.
'           Das wird realisiert durch die Anordnung der beiden ListBoxen List1 und List1Unsorted je nach Bedarf an die oberste
'           Position mit zB List1.ZOrder
'22.01.2010 Verbesserung:
'           Beim Programmstart soll erkannt werden, ob dabei eine Taste gedrückt gehalten wurde.
'           Ich will beim Doppelklick auf eine Einzeldatei, damit unterscheiden, ob alle Dateien im gleichen Ordner durchsucht
'           werden sollen, oder ob der Drag&Drop-Container aufgehen soll.
'           Bei Festhalten der Shift-Taste soll der Drag&Drop-Container aufgehen.
'23.01.2010 Verbesserung:
'           Das Löschen von Dateien, die mit eimem Häkchen fürs Löschen vorgemerkt sind, soll ab jetzt in den Papierkorb geschehen.
'           Bisher wird ohne Gnade gekillt.
'           Das Programm lief unter der Entwicklungsumgebung, aber ist abgestürzt als exe.
'           Ich musste bei der Kompilierung jegliche Code-Optimierung ausschalten, dann hat es geklappt
'04.05.2010 Verbesserung:
'           In frmDurchsuchenUnterOrdner kommt ein neuer Button '&Abbrechen'
'03.06.2010 Verbesserung:
'           Es soll schneller gehen, daß ich an die IPTC-Stichworte komme.
'           Lösung: Bei F5-Taste wird ListBoxForm.chkIptcAnzeigen voreingestellt
'03.09.2010 Verbesserung:
'           Die Rechteck-Lupe soll schon beim Zeichnen des Rechtecks das echte Breiten/Höhen-Verhältnis beachten
'           ich muss das Breiten/Höhen-Verhältnis des Bildschirms SGVH auswerten
'17.02.2011 Verbesserung:
'           Für Multiuser-Umgebungen ist es notwendig, daß jeder user seine eigene fotos.ini besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'19.02.2011 Fehlerkorrektur:
'           Bisher kam Laufzeitfehler 9; Index außerhalb des gültigen Bereiches
'           wenn die EXIF-Felder einer Datei mit Dateilänge=0 angezeigt werden sollten
'23.06.2011 13.3.20 Verbesserung:
'           Ich mache die Größe der Fonts für die Controls abhängig von der Einstellung unter 'Eigenschaften von Anzeige' ->
'           Erweitert -> DPI-Einstellungen. Das geschieht automatisch beim Form_Load jedes Formulars.
'           Ich unterscheide normal=96, groß=120, sehr groß>120
'           Das erfordert Bildschirmauflösung mindestens 1024 x 768 bei 96 DPI und
'           mindestens 1280 x 800 bei 120 DPI
'           Der Nutzer soll entscheiden, ob er die Fontgrößen-Anpassung haben will, wenn eine DPI-Einstellung höher als 96
'           gefunden wird, der Wert wird in Fotos.ini gespeichert
'24.11.2011 13.3.22 Änderung:
'           CommonDialog comdlg32.ocx wird aus dem Projekt entfernt und durch Modul standarddialoge.bas ersetzt
'           sysinfo.ocx wird aus dem Projekt entfernt, das hat geprüft ob Windows98 vorliegt.
'26.11.2011 13.3.22 Änderung:
'           Entfernt wird das Verknüpfen von Dateitypen mit diashow.exe über die registry, weil es nicht mehr funktioniert, seit ich nicht mehr in den
'           Ordner C:\Programme installieren lasse
'           Es ist auch nicht mehr möglich Diashow.exe als Standardprogramm bei Rechtsklick auf jpg file auszuwählen
'           Das geht mit 'e:\disks\SystemAllgemein\Shell Extention View\FileTypesMan.exe' im unteren Fenster bei open die Eigenschaft
'           Command-line editieren
'14.02.2012 13.4.0 Änderung:
'           Im Windows 7 sieht die Standardschrift Ms Sans Serif Scheiße aus. Ersetzt durch "Arial" aus DiashowForm.txtFont.FontName
'25.02.2012 13.4.0 Änderung:
'           WertxForm die grauen Flächen sehen Scheiße aus
'25.02.2012 Korrektur zu 23.06.2011 13.3.20 Verbesserung: man muß zuerst ReadIniFile ausführen
'03.03.2012 13.4.1 Verbesserung alle Versionen
'           Mit Tastenkombination Num+Strg+N einschalten der Anzeige eine Zeile Bildbeschreibung ganz oben bei jedem Foto
'           Mit Tastenkombination Num+Strg+M ausschalten der Anzeige eine Zeile Bildbeschreibung ganz oben bei jedem Foto
'29.03.2012 13.5.0 Verbesserung
'           Zoom erfolgt mit GDIPlus antialiasing sämtliche Bild-Darstellung ebenfalls mit GDI+
'           Verbesserung 13.4.1 wird zurückgenommen
'           2 neue native Dateitypen PNG TIF, aber CUR gestrichen
'           Benutzung einer MDIForm
'           Untergeordnete MDIChild Formulare können nicht modal angezeigt werden, Lösung MDIChild = False
'           Beim Laden der MDIForm werden alle MDIChild Formulare geöffnet, Rangordnung kann man mit Form.Zorder 0 oder 1 herstellen
'           Wenn das Schließkreuz der MDIForm unwirksam ist, steht in irgendeiner Form_QueryUnload 'cancel = True'
'29.03.2012 Die Parameter ZoomToFullScreen aus dem Abschnitt [Adjustments] der fotos.ini werden ab jetzt ausgewertet, aber
'           nicht geschrieben
'24.04.2012 Zur Steuerung ob Fenstergröße änderbar ist oder nicht dient
'           MDIFormOhneTitle und MDIFormMitTitle
'           Das Programm startet mit Fenstergröße änderbar
'           MDIFormOhneTitle läßt auch die task bar verschwinden. Bisher weiß ich nicht, wie ich sie zurückholen soll. Es hilft nur
'           neu booten
'           Deshalb wird chkFensterGrößeÄnderbar entfernt
'15.07.2012 13.5.4 Fehlerkorrektur
'           Bei der automatischen Diashow flackern zwei Fenster durch, aber nur, wenn während der automatischen Diashow
'           einmal rechtsgeklickt wurde
'19.12.2012 13.5.5 Fehlerkorrektur
'           Nicht im Windows8 aufgetreten
'           Wenn ListBoxForm sichtbar war und es wurde F3 gedrückt, dann gab es keine Reaktion. F3 mußte nochmals gedrückt werden.
'03.03.2013 13.5.5 Fehlerkorrektur
'           ListBoxForm verkleinern/vergrößern/schließen führt zu Unsinn
'           Lösung: Die Form enthält keine Buttons mehr zum verkleinern/vergrößern/schließen
'-------------------------------------------------------------------------------------------------------------------------------------
'04.03.2013 Neue Funktion 14.0.0
'           Unicode-Unterstützung durch die Timosoft Controls
'           geht nicht im XP und nicht im Win8: exe stürzt ab, vermutlich weil
'           bei Diashow.Form_Unload nicht alle Forms explicit mit Unload entladen wurden
'           Geht im Win7
'           Zum Ändern von FontSize muss die Eigenschaft des Timosoft Controls UseSystemFont = False sein
'           Viele Events bei den Timosoft Controls sind standardmäßig disabled. Man muss im gezeichneten Control element rechtsklicken ->
'           Properties -> Häkchen rausnehmen
'           ListBoxForm.ExLVwu abgeschrieben aus
'               D:\VISUALBA.SIC\VB6BeispielCode\Unicode\Unicode Timosoft\ExplorerListView\samples\VB6\General\OptionListView\
'               Wenn in ListBoxForm das Subclassing abgeschaltet wird, gibt es keine Checkboxen zu sehen
'               Andererseits macht das Subclassing das Programm saulangsam
'               Modul basSubclassing.bas und Klassenmodul ISubclassedWindow.cls einfügen
'           Drag&Drop als Target geht erst wenn man RegisterForOLEDragDrop = True setzt
'           Das Debuggen spinnt in DiashowForm bei eingeschaltetem Subclassing wenn man _OLEDragDrop debuggen will ->
'           probiere ob es mit Alt+F4 weitergeht
'
'           Die formCaption (Form.Caption) mit unicode sieht man nicht in der IDE sondern erst wenn man die exe startet
'           Alle Datei read/write Operationen sollte man mit FSO machen. Da muss man vorher testen ob der Dateiname auch nur ein unicode
'           Zeichen enthält oder alles ansi ist. Man muss daraufhin die Datei mit FSO entweder als unicode oder ansi Datei öffnen.
'           Alles auswechseln was wie 'Open Path For Binary Access Read As #Handle' aussieht
'
'           init_global bei Start des Programms
'           Der Ordner res muss mit ausgeliefert werden, dort stehen die icons für die Häkchen in ListBoxForm
'           Die scrrun.dll muss mit ausgeliefert werden Sie ist zuständig für FSO
'           iptcinfo.dll entfällt das mach ich jetzt selber
'           clsEXIF.cls muss geändert werden damit es unicode filenames versteht
'           INI file wird unicode fähig durch schreiben mit FSO und Benutzen von GetPrivateProfileStringW und
'           WritePrivateProfileStringW
'           Für Kill gibt es ein VBA replacement for "Kill(PathName)" with UNICODE support in UnivbzGlobal.bas aber ich benutze
'               file_delete
'           Alle MsgBox wo file names vorkommen ersetzen durch MessageBoxW
'           App.Path ersetzen durch getCurrentDir
'           Für die MouseIcons wird die .res-Datei benutzt.
'               Hinein kommen sie mit dem VB-Ressourcen-Editor.
'               Heraus kommen sie mit zB Me.MouseIcon = LoadResPicture(105, 1) 105-Das ist Squarezoom.ico
'               Dadurch brauche ich die .ico-Dateien nicht mehr auszuliefern

'15.04.2013 14.0.0 Fehlerkorrektur
'           Bei Fenstergröße änderbar: Format 2534x1584 entspricht 1920x1200 wird nicht exakt angepaßt, ringsherum sind schwarze Ränder
'29.04.2013 14.0.0 Verbesserung
'           'Beim Bildladen vergrößern auf Vollbild' neue Option 'immer'
'28.05.2013 14.0.0 Fehlerkorrektur
'           Fenstergröße nicht änderbar' ist ab sofort nicht mehr auswählbar
'28.05.2013 14.0.0 Fehlerkorrektur
'           Wenn mittels Häkchen ein Bild in den Papierkorb verschoben wurde, hat nachfolgende Taste F3 zum Absturz
'           ohne Mitteilung geführt
'30.05.2013 14.0.0 Fehlerkorrektur
'           Es kam Laufzeitfehler '9' wenn mit der BildPosList gearbeitet werden sollte
'           Ich muss ab sofort ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count) machen und nicht mit einer der anderen Listboxen
'04.06.2013 14.0.0 Verbesserung
'           Überarbeitung der EXIF-Felder, es werden jetzt GPS-Felder erkannt
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
'13.06.2013 14.0.0 Verbesserung
'           Überarbeitung der EXIF-Felder, es werden jetzt weitere EXIF-Felder erkannt
'16.06.2013 14.0.0 Verbesserung
'           Überarbeitung der EXIF-Felder, es werden jetzt weitere EXIF-Felder erkannt
'-------------------------------------------------------------------------------------------------------------------------------------
'07.08.2013 14.0.1 Fehlerkorrektur
'           'ShellExecute print' spinnt mit der standardmäßigen Print- und Faxanzeige
'           Ich biete jetzt mit rundll32.exe einen Dialog an, wo der Nutzer sein gewünschtes Programm zum Drucken des Fotos aussuchen kann
'07.08.2013 14.0.1 Fehlerkorrektur
'           Shellexecute ersetzen durch RunShellExecute(war vergessen worden bei unicode Einführung)
'25.08.2013 14.0.1 Fehlerkorrektur
'           Ich markiere zu löschende Dateien durch ein Häkchen, weil ich sie später mit F5-Menü löschen will. Dann rücke ich mit F3 auf
'           das nächste Foto vor. F3 bringt mich aber an eine falsche Stelle, die meist etwa 5 Fotos zu weit voraus liegt.
'           Nach dem Löschen markierter Fotos wird vorgegangen wie nach F8.
'28.08.2013 14.0.1 Verbesserung
'           Version anzeigen
'-------------------------------------------------------------------------------------------------------------------------------------
'09.06.2014 14.0.5 Verbesserung=Fehlerkorrektur
'           Situation:Das Anzeigen der ListBoxForm nach Drücken von F5 dauert saulange, wenn es sich um große Bilder handelt (etwa 5 MegaBytes)
'           Lösung:Schuld ist ListBoxForm.chkIptcAnzeigen.Value = 1, weil alle IPTC-Felder im Foto gelesen werden
'-------------------------------------------------------------------------------------------
'24.06.2014 14.0.5 Fehlerkorrektur
'           Fehler: Bei der Funktion Rekursive ist ein Dateiname von >130 Bytes Länge bisher ignoriert worden.
'           Das ist aufgetreten seit Version 14.0.0
'           Lösung: Die function FindFirstFileW und FindNextFileW in Module1 sind falsch deklariert, in UnivbzGlobal richtig
'           Falsch ist      Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'           richtig ist     Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFFData As Long) As Long
'           falsch ist      hSearch = FindFirstFileW(StrPtr(Path & "*"), wfd)
'           richtig ist     hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(wfd))
'           falsch ist      DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
'           richtig ist     DirName = RemoveNulls((wfd.cFileName))
'----------------------------------------------------------------------------------------------------------
'02.07.2014 14.0.5 Fehlerkorrektur
'           Zustand: Es gibt Fotos, bei denen werden die EXIF-Informationen unvollständig angezeigt
'                   bei YCbCrCoefficients ist Schluß
'           Lösung: in clsEFIF.cls wird abgefragt ob IFD(i).Length = 0 / Bei IFD(i).Length = 0 gab es einen unbehandelten Fehler
'                   und die Prozedur wurde vorzeitig beendet.
'----------------------------------------------------------------------------------------------------------
'18.03.2015 14.0.8 Fehlerkorrektur
'           Zustand: Wenn eine andere Software ein IPTC-Feld bearbeitet hat, kann es passieren dass GERBING Software
'                   keine IPTC-Felder anzeigt (gar keine = leer)
'           Lösung: modIPTC wird korrigiert und fragt nach dem ersten IPTC-Header wenn nicht "Photoshop 3.0" gefunden wird
'                   ob es weiter hinten noch einen IPTC-Header mit "Photoshop 3.0" gibt
'----------------------------------------------------------------------------------------------------------
'22.05.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 18.03.2015
'           Zustand: Die Fotos von Ralph haben ewig lange gebraucht um mit Fotosmdb/Prüfen3 in die Datenbank aufgenommen zu werden
'                   oder mit Diashow oder Fotoalbum oder WallpaperChanger anzezeigt zu werden.
'                   Bei pos = InStr(1, strImageString, IPTCHeader, vbTextCompare) braucht die Programmausführung ewig lange
'                   Scheinbar wird ab einer bestimmten Länge eines Strings die Function InStr arschlangsam.
'           Lösung: ich muss schreiben pos = InStrB(1, strImageString, IPTCHeader, vbTextCompare)
'                   InStrB geht blitzschnell
'                   aber beim anschließenden Vergleich muss die pos korrigiert werden
'----------------------------------------------------------------------------------------------------------
'31.05.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 22.05.2015
'----------------------------------------------------------------------------------------------------------
'01.06.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 22.05.2015
'----------------------------------------------------------------------------------------------------------
'27.05.2015 14.1.1 Fehlerkorrektur
'           Zustand: bei Dateityp 'PNG' oder 'TIF' führt ein Klick auf die Checkbox chkIptcAnzeigen zur Schleife
'           Lösung: nur bei Dateinamenerweiterung 'JPG' wird ein Klick auf Exif oder IPTC akzeptiert
'----------------------------------------------------------------------------------------------------------
'20.06.2015 14.1.1 Fehlerkorrektur
'           Zustand: Manchmal kommt beim Löschen markierter Dateinamen Laufzeitfehler '91'. Das war bisher nicht reproduzierbar
'                   Heute habe ich einen reproduzierbaren Stand gehabt.
'           Lösung: In der Prozedur LöschenMarkierte sind zwei Stellen wo es zu endlosen Schleifen kommen kann
'                   und in der Schleife tritt irgendwann Laufzeitfehler '91' auf.
'                   Ich arbeite mit On Error Resume Next und beende die Schleife, wenn err.number <> 0 ist
'==========================================================================================================
'27.08.2015 14.1.2 neue Funktion
'           Geocoding: Falls in den EXIF-Felder des Fotos GPS-Positionen vorkommen, sollen sie in einer Landkarte wie bei Google maps
'           angezeigt werden.
'           Zusammenarbeit mit Picasa ist möglich. Picasa kann GPS-Koordinaten ins Bild einfügen.
'           Unter Aktion wählen und ausführen... kommt eine neue Aktion 'Zeige Geo-Position'
'           Bei Klick auf 'Zeige Geo-Position' muß ich untersuchen, ob in den EXIF-Felder des Fotos GPS-Positionen vorkommen.
'           Wenn ja,
'           muss ich die GPS-Koordinaten aus der sexagesimalen Darstellung in die Dezimaldarstellung umrechnen, dann
'           öffnet sich die Form frmGEOPosition
'           Dort macht das user control ucGMap den Hauptteil der Arbeit
'           ucGMap nutzt die http://maps.googleapis.com/maps/api
'           Es heißt, der Aufruf der API ist (limited to 1000 requests per User and Day)
'           Wenn ucGMap breiter/höher als 640 Pixel ist, wird die Schrift unscharf
'----------------------------------------------------------------------------------------------------------
'24.09.2015 14.1.2 Fehlerkorrektur
'           Zustand: Bei einem Bild mit Dateilänge = 0 kommt zuerst eine Fehlermeldung, dann Laufzeitfehler errcode = 6
'           Lösung:  Nach Fehlern mache ich ListBoxForm.Show
'----------------------------------------------------------------------------------------------------------
'07.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: im Windows 10 und Windows 8.1 und vom Standard abweichender DPI-Einstellung zeigt mein Programm verschwommene Schrift
'                   Das kann der Nutzer korrigieren, indem er die exe markiert -> Eigenschaften -> Kompatibilität ->
'                   DPI-Skalierung nicht anwenden
'           Lösung: Ein Programm erklärt sich selbst als DPI-kompatibel. Das geht durch sein Manifest
'----------------------------------------------------------------------------------------------------------
'13.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Ich habe nicht berücksichtigt, daß es auch EXIF-Felder mit unicode Inhalt geben kann die werden als ?????
'                   dargestellt
'           Lösung: In Form F5MehrereZeilen habe ich txtEXIFInfo durch ein unicode fähiges control ersetzt
'----------------------------------------------------------------------------------------------------------
'14.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Wenn ich den Windows Explorer benutze rechtsklicken -> Eigenschaften -> Details und im Abschnitt Beschreibung
'                   einfüge Titel, dann wird meine Eingabe sowohl in Titel als auch in Thema eingetragen.
'                   Bei mir intern ist Titel = EXIF-XPTitle und Thema = EXIF-ImageDescription
'                   Wenn ich einen unicode String in Titel eintrage gibt mein Programm unter EXIF-ImageDescription Mist aus
'                   aber unter EXIF-XPTitle ist es richtig
'                   Andere Programme wie ExifToolGUI oder XnViewMP machen es richtig
'           Lösung: Wenn ich einen ascii string finde (IFD.Format=2) dann kommt FromUTF8String dran
'----------------------------------------------------------------------------------------------------------
'16.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Wenn ein IPTC-Feld UTF-8 Code enthält (erzeugt durch exiftool mit einem unicode Feld), dann wird Mist angezeigt
'                   Ich sehe zweistellige Zeichen im UTF-8 Code.
'           Lösung: modIPTC.bas function VorhandeneEinzelsegmenteSuchen wird verändert.
'                   Eventuell vorhandene UTF-8 Zeichen werden konvertiert. Es kommt FromUTF8String dran.
'==========================================================================================================
'07.03.2016 14.2.1 Fehlerkorrektur
'           Zustand: beim Anzeigen der Geoposition kommt Laufzeitfehler '5'
'           Lösung: Es genügt nicht nach GPSLatitudeRef: zu suchen - Abfragen ob GPSLatitude: gefunden wurde
'----------------------------------------------------------------------------------------------------------
'21.03.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: Beim Wechsel auf einen anderen Bildname in ExLVwU kommt nicht das andere Bild. Ich muß Doppelklicken oder die
'                   Enter-Taste drücken.
'           Lösung: Ich arbeite beim Click so, wie bisher beim Doppelclick
'----------------------------------------------------------------------------------------------------------
'22.03.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: ich brauche 2 neue Buttons für 'Alle markieren' und 'Alle Markierungen entfernen'
'           Lösung: In der Form ListBoxForm realisiert
'==========================================================================================================
'02.09.2016 14.2.2 Verbesserung alle Versionen
'           Zustand: Die Geo-Position in der Grad-Minuten-Sekunden-Darstellung wird manchmal richtig angezeigt und manchmal falsch.
'           Lösung: im Modul clsExif: ich vermeide das Umrechnen in Grad-Minuten-Sekunden sondern benutze gleich eine Dezimalzahl,
'                   so wie sie in frmGEOPosition gebraucht wird
'----------------------------------------------------------------------------------------------------------
'02.09.2016 14.2.2 Verbesserung alle Versionen
'           Zustand: Doppelklick auf die Landkarte wirkt bisher nicht
'           Lösung: Doppelklick aktivieren. Damit wird die Karte auf den Doppelklick-Punkt zentriert
'----------------------------------------------------------------------------------------------------------
'28.09.2016 14.2.2 Verbesserung alle Versionen
'           Zustand: Automatik(zB alle 3 Sekunden) mit 'Beim Bildwechsel individuell vergrößerten/verkleinerten Bildausschnitt beibehalten'
'                   und Zufallsreihenfolge zeigt die Bilder in völlig unerwarteter Position und unerwartetem Zoom
'           Lösung: in der Prozedur Zufallsindex hatte ich vergessen für das bisher gewählte Bild zuerst
'                   ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False zu setzen und nach der Zufallsauswahl
'                   ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
'                   damit die Prozedur AbrufenBildPosList richtig arbeiten kann
'----------------------------------------------------------------------------------------------------------
'28.09.2016 14.2.2 Verbesserung alle Versionen
'           Ich habe Subclass abgeschaltet. Keine veränderte Reaktion festgestellt
'----------------------------------------------------------------------------------------------------------
'11.10.2016 11.03.2016 14.2.2 Fehlerkorrektur alle Versionen
'           Zustand: in Form ListBoxForm kann ich aus EXIF-Feldern Texte kopieren, aber aus IPTC-Feldern nicht
'           Lösung: ich entferne die Listbox LstU und ersetze sie durch die Textbox txtIPTCInfo
'                   kopierte Lösung vom 11.03.2016 aus FotoAlbum.vbp
'----------------------------------------------------------------------------------------------------------
'12.10.2016 14.2.2 Verbesserung alle Versionen
'           Zustand: Bei der Anzeige der IPTC-Felder in ListBoxForm kommt beispielsweise 10 mal die Zeile Caption
'           Lösung : Ich entferne in modIPTC.LeseIPTC alle Loop-Konstruktionen
'==========================================================================================================
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
'17.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Die Mauspositionen werden als Kommazahl angezeigt (F11)
'           Lösung Konvertieren von Double in Long
'----------------------------------------------------------------------------------------------------------
'18.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn es keinen Schreibzugriff für fotos.ini gibt, kommt bisher kein Hinweis
'           Lösung: Msgbox Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können.
'----------------------------------------------------------------------------------------------------------
'12.08.2017 15.0.1 Verbesserung alle Versionen
'           Benutzung der Tastenkombination Strg+C kopiert das aktuelle Bild in die Zwischenablage
'----------------------------------------------------------------------------------------------------------
'13.08.2017 15.0.1 Verbesserung alle Versionen
'           Ab sofort kann im Email-Fenster die aktuelle Datei als Anhang übernommen werden
'           Voraussetzung ist ein installiertes Outlook
'----------------------------------------------------------------------------------------------------------
'19.10.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: es kommt zweimal die MsgBox: Es gibt mit Häkchen markierte Dateien, wollen Sie diese wirklich ignorieren?
'           Lösung: Schalter gblnHäkchenGefunden, dann kommt nur noch eine MsgBox
'----------------------------------------------------------------------------------------------------------
'20.10.2017 15.0.1 Verbesserung alle Versionen
'           Voraussetzungen:
'               Diashow.exe ist im Ordner c:\Program Files (x86)\... installiert, weil es sonst nicht als Standard-Programm für
'               Doppelklick ausgewählt werden kann
'               Diashow.exe ist bereits einmal gestartet
'               Diashow.exe startet ein weiteres mal, es wird keine fotos.ini gefunden
'               und jetzt sind bei Drücken von F5 alle Checkboxen schwarz
'           Zustand: Wenn kein res-Ordner gefunden wird, bleiben die Häkchen in ListBoxForm.ExLVwU schwarz
'           Lösung: MsgBox "Es wurde kein res-Ordner gefunden. Wiederholen Sie die Installation von Diashow"
'21.11.2017 Nachbesserung zum 20.10.2017
'           Anstelle von MsgBox "Es wurde kein res-Ordner gefunden. Wiederholen Sie die Installation von Diashow" kommt
'           MsgBox "Wenn Sie Diashow.exe mehrfach starten, können Sie keine Häkchen im Dateinamen-Fenster setzen"
'----------------------------------------------------------------------------------------------------------
'24.12.2017 Nachbesserung zum 18.03.2017
'           Zustand: wenn in fotos.ini nicht geschrieben werden kann, kommt
'                   Msgbox Sie müssen für Schreibrechte sorgen, damit Änderungen an dieser Datei gemacht werden können
'                   Wollen Sie trotzdem Weiterarbeiten ... ja/nein
'                   Bei Antwort = nein geht es trotzdem weiter
'           Lösung: in Prozedur WriteDPI benutzen 'On Error Resume Next'
'----------------------------------------------------------------------------------------------------------
'18.01.2018 15.0.1 Nachbesserung zum 11.03.2017 alle Versionen
'           Zustand: Verändern der Schriftgröße wirkt erst nach einem Neustart
'           Lösung: Durch Aufruf von 'Call AnpassenNutzerWunsch(frm)' wirkt das Verändern der Schriftgröße sofort
'----------------------------------------------------------------------------------------------------------
'27.01.2018 15.0.1 wird nicht ausgeliefert
'==========================================================================================================
'27.01.2018 15.0.2 Verbesserung
'           Diashow.exe kann von fotos.exe über die Menüleiste Tools aufgerufen werden
'           Dazu soll die Form Diashow in Fenstermitte starten
'----------------------------------------------------------------------------------------------------------
'28.01.2018 15.0.2 Verbesserung
'           Diashow.exe soll Stapel-Dateinamenänderung ausführen können
'           Dafür gibt es die neue Form frmNamenÄndern
'----------------------------------------------------------------------------------------------------------
'28.01.2018 15.0.2 Fehlerkorrektur Nachbesserung zum 12.08.2017
'           Zustand: Strg+C geht nicht bei Unicode Namen -> Laufzeitfehler
'           Lösung: Function LoadPictureWThumb abschreiben aus fotos.exe
'----------------------------------------------------------------------------------------------------------
'29.01.2018 15.0.2 Verbesserung
'           Diashow.exe Form frmBildMitGDIPlus soll eine Menü-Leiste bekommen, die mit der Taste Alt ein/ausgeschaltet werden kann
'----------------------------------------------------------------------------------------------------------
'30.01.2018 15.0.2 Verbesserung
'           Diashow.exe bekommt eine Querverbindung zu Everything.exe
'           Abhängig davon ob chkSortAsDragAndDrop ein/ausgeschaltet ist wird der Button btnSucheStarten sichtbar.
'           Damit werden vom Nutzer eingegebene Suchparameter an Everything.exe übergeben.
'           Das Suchergebnis-Fenster von Everything kann danach mit Drag&Drop in den Diashow-Container gezogen werden.
'==========================================================================================================
'13.03.2018 15.0.3 Nachbesserung zum 28.01.2018
'           Zustand: Beim NamenÄndern kommt Laufzeitfehler wenn 'Zugriff verweigert'
'           Lösung: Weiterarbeiten mit 'on Error resume Next' und MsgBox apiErrorDescription
'----------------------------------------------------------------------------------------------------------
'14.03.2018 15.0.3 Verbesserung
'           Zustand: Ich habe doch tatsächlich vergessen die Änderungen bei F10 -> Was soll beim Bildladen passieren in die fotos.ini
'                   einzutragen. Bei Benutzung von fotos.exe geht es richtig.
'           Lösung: Prozedur WriteAZF aus fotos.exe kopiert
'----------------------------------------------------------------------------------------------------------
'29.09.2018 15.0.3 Verbesserung
'           Zustand:die Google Maps API gehen nur noch mit einem API key zu benutzen, heute noch kostenlos, aber wer weiß wie lange?
'           Lösung: Ich steige um auf OpenStreetMap, da gibts auch ein API
'                   neue Form: frmMap
'                   zoom=14 -> zoom=16
'           in frmMap Form_KeyDown zu erlauben hat im Windows10 zum kommentarlosen Programmende nach 2-3 Fotos geführt
'           Lösung: frmMap Form_KeyDown auskommentiert
'                   gestrichene Form: frmGEOPosition
'----------------------------------------------------------------------------------------------------------
'16.10.2018 15.0.3 Nachbesserung zum 29.09.2018
'           Zustand: In VM Win7 geht die Lösung mit frmMap und OpenStreetMap nicht
'                   Es kommt 'In dem Skript auf dieser Seite ist ein Fehler aufgetreten'
'                   .
'                   URL: https://www.openstreetmap.org/assets/application-.......js
'           Lösung: Ich brauche eine neue Form frmStrgG, wenn Strg + G gedrückt wird
'                   Der Nutzer soll auswählen können ob mit frmMap oder mit ShellExecute
'=========================================================================================================
'08.04.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Obwohl in einem Foto GPS-Daten eingetragen sind(Kontrolle mit Windows-Explorer -> Eigenschaften -> Details -> Überschrift GPS)
'                   kommt eine MsgBox 'Geo positions not available'
'           Ursache: Die GEO-Positionen sind im XMP-Abschnitt des Fotos eingetragen. Das macht zB Geosetter(mit Hilfe von Exiftool),
'                   ich suche sie aber nur im EXIF-Abschnitt.
'                   Andere Software findet diese GEO-Positionen zB ExifToolGUI, PSP 2019, Irfan View, Fotos App, XnViewMP.
'                   Fotos App(Win10) korrigiert sogar selbständig aus dem XMP-Abschnitt in den EXIF-Abschnitt
'           Lösung: Da ich in clsEXIF sowieso jedes JPG-Foto durchsuche, um den EXIF-Abschnitt zu finden, kann ich dort ebenso nach den
'                   XMP-GEO-Positionen suchen
'                   Ich suche nach exif:GPSLatitude und exif:GPSLongitude mit InstrB, weil das rasend schnell geht
'                   Die gefundenen Werte gstrLatXMP und gstrLongXMP muss ich dann noch in ein Format verwandeln, das OpenStreetMap versteht
'                   zB gstrLatXMP 50,38.7309456N -> 50.64551575
'                   zB gstrLongXMP 11,53.9826786E -> 11.89971130
'----------------------------------------------------------------------------------------------------------
'10.05.2019 15.0.4 Verbesserung
'           Zustand: Bei sehr großen JPG-Fotos ist das Drehen mit dem Explorer ohne Auswirkung.
'                   Im Explorer-Icon kann man die erfolgte Drehung sehen.
'                   In Diashow.exe erfolgt keine Drehung, aber jede andere Software macht es.
'                   Die vom Explorer durchgeführte Drehung wird in EXIF.Orientation eingetragen, aber von Diashow.exe nicht ausgewertet
'                   Es gibt 8 mögliche Werte siehe 'Readme EXIF PropertyTagOrientation.docx'
'           Notlösung: Das Foto verkleinern, dann wird es von Diashow.exe korrekt angezeigt
'           Lösung: Das nichtssagende 'mirrored or turned' bzw 'horizontal normal' wird erweitert um die vorangesetzte EXIF.Orientation clsEXIF.cls
'                   und gleichzeitig wird der Wert von EXIF.Orientation gespeichert in gstrEXIFOrientation.
'                   Abhängig von gstrEXIFOrientation wird in frmBildMitGDIPlus.MyDrawImage das Foto gedreht mit GdipImageRotateFlip
'                   Durch das Drehen haben sich Width und Height des Bildes verändert. Mit GdipGetImageDimension muss ich diese neu ermitteln.
'                   Bei Bildern im Hoch-Format muss ich einen dblKorrekturFaktor ausrechnen, weil sonst der untere Bildteil nicht mit gezeigt wird.
'                   Dann kann das Bild wie bisher gezeichnet werden.
'----------------------------------------------------------------------------------------------------------
'04.07.2019 15.0.4 Fehlerkorrektur
'           Zustand: Nacharbeiten zum 08.04.2019
'                   Die GPS-Daten in den EXIF-Feldern sind richtig angezeigt, aber werden falsch angezeigt
'                   Ursache ist ein Fehler in Form1.GEOKoordinatenUmrechnenXMP
'                   Beispiel: Minuten = 0.287564
'                   Nachkomma = MinutenDouble / 60 'liefert Ergebnis=0
'           Lösung: Wenn Komma als Dezimaltrennzeichen verwendet wird, das ist in beispielsweise Deutschland der Fall,
'                   muss der Punkt im String Minuten in Komma verwandelt werden
'                   sonst kommt bei MinutenDouble / 60 Ergebnis=0
'----------------------------------------------------------------------------------------------------------
'30.09.2019 15.0.4 Fehlerkorrektur Nacharbeiten zum 10.05.2019
'           Zustand: Bei Hochkantfotos aus Katjas Handy oder bei Panoramafotos aus Jens Handy
'                   werden diese nur in halber Höhe oder halber Breite dargestellt
'           Lösung: zuerst wird gstrEXIFOrientation ermittelt
'                   dann kommt sofort GdipImageRotateFlip
'                   wegen GdipImageRotateFlip ändert sich Width und Height des Bildes
'                   erst danach X und Y ausrechnen
'=========================================================================================================
'15.04.2020 16.0.0 Verbesserung
'           Zustand: Wenn GPS-Daten vorhanden sind werden sie mit OpenStreetmap angezeigt.
'                   Das geht manchmal sehr schleppend.
'                   Das hat den Nachteil dass nicht auf Satellitenansicht umgeschaltet werden kann, und auch nicht auf Google Street View.
'           Lösung: Als dritte Variante wird Google Maps Anzeigen im Browser angeboten
'----------------------------------------------------------------------------------------------------------
'22.07.2020 16.0.0 Verbesserung
'           Zustand: Das Häkchen ein-oder ausschalten mit der Maus in der Form ListBoxForm ist unbequem.
'                   Ich würde es gern mit einer Tastenkombination machen.
'           Lösung: Strg+E schaltet beim aktuellen Bild das Häkchen ein
'                   Strg+A schaltet beim aktuellen Bild das Häkchen aus
'                   Es darf nicht die Form ListBoxForm zu sehen sein sondern nur das Bild
'----------------------------------------------------------------------------------------------------------
'22.07.2020 16.0.0 Verbesserung
'           Zustand: Beim ersten Bild anzeigen muss ich stets eingreifen und die Form maximieren
'                   Ich will das automatisch die Form maximiert gezeigt wird.
'           Lösung: MDIForm1.WindowState = 2 in der Prozedur DiashowForm.Bildanzeigen


'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
'           Ungelöstes Problem im XP
'           Wenn die Kamera Fujifilm mit USB angeschlossen ist, kann man Diashow.exe nicht anwenden.
'           Es kommt im Windows Explorer ein viel kürzeres Rechtsklick-Menü als bei zB Olympus.
'           Es fehlt 'öffnen mit...'
'-------------------------------------------------------------------------------------------------------------
'           Gelöstes Problem im Windows7
'           Shellexecute "print" geht nicht - genauso wie im Windows Explorer Rechtsklick auf eine Bilddatei -> Drucken nicht geht
'           Windows-Fotoanzeige reagiert nicht
'           Lösung siehe e:\Faq&Lehrmaterial\3 Mein Win7 PC Macken und Lösungen\Windows 7 Windows-Fotoanzeige druckt nicht.txt
'           Bessere Lösung: siehe 07.08.2013
'-------------------------------------------------------------------------------------------------------------
'           Ungelöstes Problem beim NamenÄndern
'           Bei Unicode Namen gibt es beim ersten Versuch einen Abbruch, beim zweiten Versuch geht es


Option Explicit


    Public blnListBoxNeuDblClick As Boolean         'Gerbing 05.05.2005
    Dim Dateiname As String
    Dim NL
    Dim msg As String
    Dim ZeilenNummer
    Dim MsgBoxMussKommen As Boolean
    Dim MouseSichtbar As Boolean
    Public F8Gedrückt As Boolean
    
    Public ZoomProzent As Long
    Public screenWidth As Long
    Public screenHeight As Long
    Public glngDiffX As Long            'Maß der Verschiebung mit der Maus
    Public glngDiffY As Long            'Maß der Verschiebung mit der Maus
    
    Public glngSaveX As Long            'rettet X zum Ausrechnen von DifferenzX für die Rechtecklupe
    Public glngSaveY As Long            'rettet Y zum Ausrechnen von DifferenzX für die Rechtecklupe
    Public RechteckLupeScharf As Boolean
    Public DifferenzX As Long           'muss für die Rechtecklupe mit dem ZoomFaktor multipliziert werden
    Public DifferenzY As Long           'muss für die Rechtecklupe mit dem ZoomFaktor multipliziert werden
    Public ZoomFaktor As Double
    Public gblnWithTitle As Boolean
    Private strOldTitle As String
    Public gblnMouseSichtbar As Boolean

    '----------------------------------------------------------------------------------------------------------------------------
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long
    
    Const WS_THICKFRAME = &H40000
    Const WS_CAPTION = &HC00000
    Const WS_SYSMENU = &H80000
    Const WS_MINIMIZEBOX = &H20000
    Const WS_MAXIMIZEBOX = &H10000
    Const GWL_STYLE = (-16)
    '----------------------------------------------------------------------------------------------------------------------------
    Const SWP_HIDEWINDOW = &H80
    Const SWP_SHOWWINDOW = &H40
    
    Private Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    
    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
        ByVal cy As Long, ByVal wFlags As Long) As Long

    Public CommandLine As String

    Dim BHV As Double       'BreitenHöhenVerhältnis
    Public SGVH As Double      'Screen-Größen-Breiten-Verhältnis                               'Gerbing 24.09.2009
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Public F6TimerGestartet As Boolean
    Public EXF As New clsEXIF
    Private EinzelnesBild As String                                                         'Gerbing 28.11.2007
    Public blnOhneUnterOrdner As Boolean
    
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer  'Gerbing 22.01.2010
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
'
'----------------------------------------04.03.2013-------------------------------------------------------------
    Dim FileNameArray() As String
    Dim m_FileCount As Long
    Private hLastDropTarget As OLE_HANDLE
    Private Const CF_DIBV5 = 17
    Private Const CF_OEMTEXT = 7
    Private Const CF_UNICODETEXT = 13
  
    Private Type GUID
      Data1 As Long
      Data2 As Integer
      Data3 As Integer
      Data4(0 To 7) As Byte
    End Type
    
    Private CFSTR_TARGETCLSID As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
    
    Private Const strCLSID_RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"
    Private CLSID_RecycleBin As GUID
    
    Public gstrFolder As String                                             'Gerbing 28.01.2018
    Dim blnVorhandeneDiashowDateiBenutzen As Boolean                        'Gerbing 30.01.2018



Public Sub BildAnzeigen()
    Dim blnZoomGleich100 As Boolean
    Dim MyGVH As Double               'My-Bild-Größen-Breiten-Verhältnis
    Dim ZoomWidth As Double
    Dim ZoomHeight As Double
    Dim rc As Long
    Dim antwort As Long

    On Error Resume Next    'sonst kann ich nach ZoomForm keine F-Tasten benutzen
    ListBoxForm.Hide
    On Error GoTo 0
    Err = 0
    If MouseSichtbar = True Then
        Me.MousePointer = vbHourglass
    End If
    
    If gblnListBoxNeuDblClick = True Then                                                       'Gerbing 25.08.2013
        rc = LoadPicBox(gblStrAktuellGezeigtesBild)
        If rc <> 0 Then
            'MsgBox ListBoxForm.ListBoxNeu.Text & " Bild kann nicht geladen werden"
            'MsgBox ListBoxForm.ExLVwU.ListItems(1) & " " & LoadResString(2056 + Sprache)
            MessageBoxW 0, StrPtr(gblStrAktuellGezeigtesBild & vbNewLine & LoadResString(2056 + Sprache)), StrPtr("GERBING  Diashow"), MB_TASKMODAL
            Me.MousePointer = vbDefault                                                         'Gerbing 24.09.2015
            ListBoxForm.Show
            gblStrAktuellGezeigtesBild = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)  'Gerbing 24.09.2015
            Exit Sub                                                                            'Gerbing 24.09.2015
        End If
    Else
        rc = LoadPicBox(ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex))
        If rc <> 0 Then
            'MsgBox ListBoxForm.ListBoxNeu.Text & " Bild kann nicht geladen werden"
            'MsgBox ListBoxForm.ExLVwU.ListItems(1) & " " & LoadResString(2056 + Sprache)
            MessageBoxW 0, StrPtr(ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex) & vbNewLine & LoadResString(2056 + Sprache)), StrPtr("GERBING  Diashow"), MB_TASKMODAL
            Me.MousePointer = vbDefault                                                         'Gerbing 24.09.2015
            ListBoxForm.Show
            Exit Sub                                                                            'Gerbing 24.09.2015
        End If
        gblStrAktuellGezeigtesBild = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)
    End If
    
    blnZoomGleich100 = False
    If PublicZoomToFullscreen <> "1" Then                                                       'Gerbing 29.04.2013
        'was schmaler oder niedriger ist als 500 Pixel wird beim ersten Laden nicht vergrößert oder verkleinert
        'ZoomProzent = 100 heißt anzeigen in Originalgröße
        If gsngPicWidth < 500 Or gsngPicHeight < 500 Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth <= DiashowForm.screenWidth And gsngPicHeight <= DiashowForm.screenHeight And WertxForm.OptKeine.Value = True Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 640 And WertxForm.Opt640.Value = True Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 800 And WertxForm.Opt800.Value = True Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 1024 And WertxForm.Opt1024.Value = True Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 1024 And WertxForm.Opt1024Oder768.Value = True Then
            blnZoomGleich100 = True
        End If
        If gsngPicWidth < 768 And WertxForm.Opt1024Oder768.Value = True Then
            blnZoomGleich100 = True
        End If
        'was höher oder breiter ist als Screenwidth bzw Screenheight wird beim ersten Laden vergrößert oder verkleinert
        If gsngPicWidth > DiashowForm.screenWidth Or gsngPicHeight > DiashowForm.screenHeight Then
            blnZoomGleich100 = False
        End If
    End If
    
    If blnZoomGleich100 = False Then
        MyGVH = gsngPicWidth / gsngPicHeight
        If MyGVH >= SGVH Then
            ZoomWidth = screenWidth / gsngPicWidth
            ZoomProzent = 100 * ZoomWidth
        Else
            If gblnWithTitle = False Then
                'Anzeige ohne Titelleiste und Taskbar
                ZoomHeight = screenHeight / gsngPicHeight
            Else
                'Anzeige mit Titelleiste und Taskbar
                'Da muss ich 150 Pixel für die Task Bar im Windows 7 abziehen   'Gerbing 15.04.2013 auskommentiert
                'ZoomHeight = (screenHeight - 150) / gsngPicHeight
                ZoomHeight = screenHeight / gsngPicHeight
            End If
            ZoomProzent = 100 * ZoomHeight
        End If
    Else
        ZoomProzent = 100
    End If
    glngDiffX = 0
    glngDiffY = 0
    RechteckLupeScharf = False
    
    XYPos.lblBildgröße = gsngPicWidth & " x " & gsngPicHeight
    XYPos.Hide                                                                      'Gerbing 24.12.2017
    If MouseSichtbar = True Then
        If RechteckLupeScharf = True Then
            Me.MousePointer = vbCustom
            'Me.MouseIcon = LoadPicture(AppPath & "\SquareZoom.ico")                'Gerbing 29.07.2007
            Me.MouseIcon = LoadResPicture(105, 1)                                   'Gerbing 04.03.2013
        Else
            Me.MousePointer = vbDefault
        End If
    Else
        Me.MousePointer = vbCustom      '99
        'Me.MouseIcon = LoadPicture(AppPath & "\MOUSE01.ICO")                       'Gerbing 29.07.2007
        Me.MouseIcon = LoadResPicture(104, 1)                                       'Gerbing 04.03.2013
    End If
    Call AbrufenBildPosList
    'DiashowForm.Visible = False
    DiashowForm.Hide
    ListBoxForm.chkExifAnzeigen.Enabled = True
    ListBoxForm.chkIptcAnzeigen.Enabled = True
    'Achtung in der IDE wird unicode in Form.Caption nicht angezeigt
    formCaption Me.hwnd, "GERBING Diashow"
    
    Unload frmBildMitGDIPlus                                                        'Gerbing 19.12.2012
    frmBildMitGDIPlus.Show
    MDIForm1.WindowState = 2                                                        'Gerbing 22.07.2020
    If gblnListBoxNeuDblClick = True Then
        'frmBildMitGDIPlus.Caption = gblStrAktuellGezeigtesBild                                          'Gerbing 25.06.2013
        'formCaption frmBildMitGDIPlus.hwnd, gblStrAktuellGezeigtesBild                                  'Gerbing 25.06.2013
    Else
        'frmBildMitGDIPlus.Caption = ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)               'Gerbing 25.06.2013
        'formCaption frmBildMitGDIPlus.hwnd, ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex)       'Gerbing 25.06.2013
    End If
End Sub

Private Sub cmdBeenden_Click()
    End
End Sub
'
'Private Sub chkFensterGrößeÄnderbar_Click()                                                                'Gerbing 30.01.2018
'    If chkFensterGrößeÄnderbar.Value = 0 Then
'        Call MDIFormOhneTitle
'    Else
'        Call MDIFormMitTitle
'    End If
'End Sub

Private Sub btnSucheStarten_Click()
    Dim msg As String
    Dim FileName As String
    Dim antwort As Long
    Dim retval As Long
    Dim pos As Long
    
    If txtSearchWords = "" Then
        'MsgBox "Suchen nach: darf nicht leer sein"
        MsgBox LoadResString(1046 + Sprache)
        Exit Sub
    Else
        pos = InStr(1, txtSearchWords, "!\gerbingthumbs")
        If pos = 0 Then
            txtSearchWords = txtSearchWords & " !\gerbingthumbs"
        End If
    End If
    If file_exist(PublicPathToEverythingExe) = False Or PublicPathToEverythingExe = "" Then
        'msg = "Mit dem Tool Everything.exe können Sie den Computer nach Dateinamen durchsuchen, wo die von Ihnen angegebenen Suchbegriffe enthalten sind" & vbNewLine
        'msg = msg & "Siehe https://www.voidtools.com" & vbNewLine
        'msg = msg & "Wenn Sie Everything.exe benutzen wollen, müssen Sie angeben in welchem Ordner dieses Tool steht." & vbNewLine
        'msg = msg & vbNewLine
        'msg = msg & "Wollen Sie Everything.exe benutzen?"
        
        msg = LoadResString(1047 + Sprache) & vbNewLine
        msg = msg & LoadResString(1048 + Sprache) & vbNewLine
        msg = msg & LoadResString(1072 + Sprache) & vbNewLine
        msg = msg & vbNewLine
        msg = msg & LoadResString(1073 + Sprache)
        
        antwort = MsgBox(msg, vbYesNo, "GERBING Diashow")
        If antwort = vbNo Then
            Exit Sub
        Else
            FileName = ShowOpenUnicodeEveryThing(Me)
            FileName = ConvertFileName(FileName)
            If FileName = "" Then GoTo ErrHandler
            PublicPathToEverythingExe = FileName
            WritePTE (PublicPathToEverythingExe)
            Exit Sub
        End If
    End If
    'Everything.exe mit Aufrufparametern starten
    retval = RunShellExecute(Me.hwnd, "open", PublicPathToEverythingExe, " -s """ & txtSearchWords, vbNull, 1)
    Exit Sub
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub chkSortAsDragAndDrop_Click()
    List1UnsortedU.ListItems.RemoveAll                                                                     'Gerbing 22.01.2010
    List1U.ListItems.RemoveAll                                                                             'Gerbing 22.01.2010
    ListBoxForm.ExLVwUIndex = 0
    If chkSortAsDragAndDrop.Value = 1 Then
        cmdKeinNeustart.Visible = True                                                                          'Gerbing 24.09.2014
        btnSucheStarten.Visible = True                                                                          'Gerbing 30.01.2018
        Label1.Visible = True                                                                                   'Gerbing 30.01.2018
        txtSearchWords.Visible = True                                                                           'Gerbing 30.01.2018
        List1UnsortedU.ZOrder
        List1UnsortedU.Visible = True
        List1U.Visible = False
    Else
        cmdKeinNeustart.Visible = False                                                                          'Gerbing 24.09.2014
        btnSucheStarten.Visible = False                                                                          'Gerbing 30.01.2018
        Label1.Visible = False                                                                                   'Gerbing 30.01.2018
        txtSearchWords.Visible = False                                                                           'Gerbing 30.01.2018
        List1U.ZOrder
        List1UnsortedU.Visible = False
        List1U.Visible = True
    End If
End Sub

Private Sub cmdKeinNeustart_Click()
    If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
        List1UnsortedU.ListItems.RemoveAll
    Else
        List1U.ListItems.RemoveAll
    End If
    ListBoxForm.ExLVwU.ListItems.RemoveAll
    cmdKeinNeustart.Visible = False
    Exit Sub
End Sub

Private Sub cmdStartAfterDragDrop_Click()
    Dim i As Long

    ListBoxForm.ExLVwU.ListItems.RemoveAll
'    If chkFensterGrößeÄnderbar.Value = 0 Then       '0=nicht aktiviert 'Gerbing 27.01.2018
'        Call MDIFormOhneTitle
'    Else
'        Call MDIFormMitTitle
'    End If
    If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
        If List1UnsortedU.ListItems.Count = 0 Then
            'MsgBox "Sie haben keine Dateien ausgewählt"
            MsgBox LoadResString(2303 + Sprache)
            Exit Sub
        End If
    Else
        If List1U.ListItems.Count = 0 Then
            'MsgBox "Sie haben keine Dateien ausgewählt"
            MsgBox LoadResString(2303 + Sprache)
            Exit Sub
        End If
    End If
    If List1U.ListItems.Count = 1 Then                                                         'Gerbing 23.10.2007
'        msg = "Wenn Sie ein einzelnes Bild auswählen," & vbNewLine                     'Gerbing 12.03.2008
'        msg = msg & "durchsucht das Programm den gesamten Ordnern nach weiteren Bildern" & vbNewLine 'Gerbing 28.11.2007
'        msg = msg & "Mit F5 können Sie die Liste aller gefundenen Bilder anzeigen" & vbNewLine   'Gerbing 28.11.2007
'        msg = msg & "Mit F3 vorwärts- mit F2 rückwärtsblättern"                             'Gerbing 28.11.2007
        msg = LoadResString(1234 + Sprache) & vbNewLine                                 'Gerbing 12.03.2008
        msg = msg & LoadResString(1231 + Sprache) & vbNewLine
        msg = msg & LoadResString(1232 + Sprache) & vbNewLine
        msg = msg & LoadResString(1233 + Sprache) & vbNewLine
        MsgBox msg                                                                      'Gerbing 28.11.2007
        Call FüllelstFilesImGleichenOrdner
    End If
    lblDateinamen.Visible = False
    List1UnsortedU.Visible = False
    List1U.Visible = False
    cmdStartAfterDragDrop.Visible = False
    'cmdSpeichern.Visible = False
    cmdKeinNeustart.Visible = False
    If ListBoxForm.ExLVwU.ListItems.Count = 0 Then
        Call KopiereListBox             'die Listbox auf der ListBoxForm anordnen
    End If
    ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count)                               'Gerbing 30.05.2013
    If List1U.ListItems.Count > 0 Then
        For i = 0 To List1U.ListItems.Count - 1
           BildPosList(i).Top = 0
           BildPosList(i).Left = 0
           BildPosList(i).ZoomPercent = 0
        Next i
    End If
    If EinzelnesBild = "" Then
        Call GeheZumErstenBild
    Else
        GeheZumEinzelnenBild                                                            'Gerbing 28.11.2007
    End If
    Call BildAnzeigen
    Screen.MousePointer = vbDefault                                                     'Gerbing 30.07.2008
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'KeyCode = 0, weil sonst aus unbekanntem Grund Keycode zweimal verarbeitet wird
    Dim rc As Long
    Dim retval As Long
    Dim DateinamenErweiterung As String
    Dim CHMFile As String
    Dim ErrorText As String
    Dim msg As String
    Dim ShiftDown As Boolean
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
   
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

    Select Case KeyCode
        Case 33                                         'Bild nach oben geht zum Anfang
            KeyCode = 0
            Call GeheZumErstenBild
            Call BildAnzeigen
        Case vbKeyHome                                  'Pos1-Taste geht zum Anfang
            KeyCode = 0
            Call GeheZumErstenBild
            Call BildAnzeigen
        '--------------------
        Case 34                                         'Bild nach unten geht zum Ende
            KeyCode = 0
            Call GeheZumLetztenBild
            Call BildAnzeigen
        Case vbKeyEnd                                   'Ende-Taste geht zum Ende
            KeyCode = 0
            Call GeheZumLetztenBild
            Call BildAnzeigen
        '--------------------
        Case vbKeyF1                                    'F1
            glngDiffX = 0
            glngDiffY = 0
            RechteckLupeScharf = False
            KeyCode = 0
            If ZoomProzent < 2 Then
                'MsgBox "Minimumzoom erreicht"
            Else
                ZoomProzent = ZoomProzent \ 2
            End If
            Call SpeichernInBildPosList
            frmBildMitGDIPlus.Show
        Case vbKeyF2                                    'F2
            KeyCode = 0
            Call GeheEinBildZurück
            Call BildAnzeigen
        Case vbKeyLeft                                  'Alt und Pfeil nach links wirkt wie F2
            If Shift = vbAltMask Then
                KeyCode = 0
                Call GeheEinBildZurück
                Call BildAnzeigen
            End If
            If Shift = vbShiftMask Then                 'Pfeil nach links und gleichzeitig Shift-Taste
                glngDiffX = glngDiffX - screenWidth / 4
                Call SpeichernInBildPosList
                RechteckLupeScharf = False
                KeyCode = 0
                frmBildMitGDIPlus.Show
            End If
        Case vbKeyF3                                    'F3
            KeyCode = 0
            Call GeheEinBildVorwärts
            Call BildAnzeigen
        Case vbKeyRight                                 'Alt und Pfeil nach rechts wirkt wie F3
            If Shift = vbAltMask Then
                KeyCode = 0
                Call GeheEinBildVorwärts
                Call BildAnzeigen
            End If
            If Shift = vbShiftMask Then                 'Pfeil nach rechts und gleichzeitig Shift-Taste
                glngDiffX = glngDiffX + screenWidth / 4
                Call SpeichernInBildPosList
                RechteckLupeScharf = False
                KeyCode = 0
                frmBildMitGDIPlus.Show
            End If
        Case vbKeyF4                                    'F4
            glngDiffX = 0
            glngDiffY = 0
            RechteckLupeScharf = False
            KeyCode = 0
            ZoomProzent = ZoomProzent * 2
            Call SpeichernInBildPosList
            frmBildMitGDIPlus.Show
        Case vbKeyF6                                    'F6
            F6TimerGestartet = True
            Timer1.Enabled = True
        Case vbKeyF7                                    'F7
            F6TimerGestartet = False
            Timer1.Enabled = False
        Case vbKeyF8                                    'F8
            KeyCode = 0
            Call EsIstF8                                                                'Gerbing 12.03.2008
        Case 38                                         'Pfeil nach oben
            If Shift = vbShiftMask Then                 'Pfeil nach oben und gleichzeitig Shift-Taste
                glngDiffY = glngDiffY - screenWidth / 4
                Call SpeichernInBildPosList
                RechteckLupeScharf = False
                KeyCode = 0
                frmBildMitGDIPlus.Show
            End If
        Case 40                                         'Pfeil nach unten
            If Shift = vbShiftMask Then                 'Pfeil nach unten und gleichzeitig Shift-Taste
                glngDiffY = glngDiffY + screenWidth / 4
                Call SpeichernInBildPosList
                RechteckLupeScharf = False
                KeyCode = 0
                frmBildMitGDIPlus.Show
            End If
        Case vbKeySeparator                         'vbKeySeparator
            'vbKeySeparator gilt als Verschiebung mit der Maus
            'vbKeySeparator gilt als Kennzeichnung für Rechtecklupe
            'vbKeySeparator gilt als als Kennzeichnung für Neuzeichnen anstelle Zoom
            'vbKeySeparator als Tastendruck darf dann nicht aus frmBildMitGDIPlus kommen
            'vbKeySeparator (Enter auf dem numerischen Block)
            frmBildMitGDIPlus.Show
        Case Else
            'ListBoxForm.Hide                                                       'Gerbing 24.09.2015
            ListBoxForm.Show                                                        'Gerbing 24.09.2015
    End Select
End Sub

Private Sub Form_Load()
    Dim msg As String
    Dim pos As Long
    Dim i As Long
    Dim Faktor As Integer
    Dim Pixel As Integer
    Dim rc As Long
    Dim Quelle As String
    Dim Ziel As String
    Dim strTemp As String
    Dim MyDir As String
    Dim AppId As Long
    Dim Handle As Long
    Dim Size As Long
    Dim buffer() As Byte
    Dim lRet As Long
    
    ' this is required to make the Timosoft control work as expected
    'Subclass

    Set fso = New FileSystemObject
    gblnWithTitle = True                                        'Gerbing 29.03.2012
    'MsgBox "Diashowform Form_Load"
    'Horizontale Scrollbar an Listbox List1 anbringen           'Gerbing 31.08.2007
    Faktor = 2
    Pixel = List1U.Width / Screen.TwipsPerPixelX * Faktor
    rc = SendMessage(List1U.hwnd, LB_SETHORIZONTALEXTENT, Pixel, 0)
        
    'Horizontale Scrollbar an Listbox List1Unsort anbringen     'Gerbing 22.01.2010
    Faktor = 2
    Pixel = List1UnsortedU.Width / Screen.TwipsPerPixelX * Faktor
    rc = SendMessage(List1UnsortedU.hwnd, LB_SETHORIZONTALEXTENT, Pixel, 0)

    AppPath = getCurrentDir
    If Right(AppPath, 1) = "\" Then
        AppPath = Left(AppPath, Len(AppPath) - 1)
    End If
    
    'On Error Resume Next
    NL = Chr(10) & Chr(13)
    MsgBoxMussKommen = True
    MouseSichtbar = True
    screenWidth = GetDeviceCaps(Me.hdc, HORZRES)
    screenHeight = GetDeviceCaps(Me.hdc, VERTRES)
    SGVH = screenWidth / screenHeight                               'Gerbing 24.09.2009
    'Sprache festlegen bei jedem Start. Wenn die Datenbank fotos.mdb englisch ist, dann Sprache=3000
    'nicht erfolgt, heißt es gibt keinen Eintrag in fotos.ini
    Call ReadFotosIniFile
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
'-----------------------------------------------------------------------------------------------------------------
    'Me.Show
    lblDateinamen.Caption = LoadResString(1202 + Sprache)   'Container Dateinamen vom Explorer
    cmdStartAfterDragDrop.Caption = LoadResString(3101 + Sprache) 'S&tart
    cmdKeinNeustart.Caption = LoadResString(3102 + Sprache) '&Container leeren
    btnSucheStarten.Caption = LoadResString(1044 + Sprache) 'Suche starten                      'Gerbing 30.01.2018
    btnSucheStarten.ToolTipText = LoadResString(1045 + Sprache)     'Damit starten Sie das Tool Everything.exe mit den von Ihnen formulierten Suchbegriffen
    Label1.Caption = LoadResString(1042 + Sprache)          'suchen nach:                       'Gerbing 30.01.2018
    txtSearchWords.ToolTipText = LoadResString(1043 + Sprache)  'geben Sie hier Suchworte oder Suchfragmente ein(zB Berlin Max) 'Gerbing 30.01.2018
    mnuBeenden.Caption = LoadResString(1040 + Sprache)      'Beenden     'Gerbing 14.11.2007
    mnuHilfe.Caption = LoadResString(1039 + Sprache)        'Hilfe
    chkSortAsDragAndDrop.Caption = LoadResString(1235 + Sprache)    'Sortierfolge wie bei Drag&&Drop 22.01.2010
    List1U.ZOrder
    List1UnsortedU.Visible = False
    List1U.Visible = True
    '----------------------------------------------------------------------------------------------------------
    'Kommandozeilen-Parameter holen
    CommandLine = LCase(command)                            'Gerbing 04.03.2013 unicodefähig
    'If (GetAsyncKeyState(VK_CONTROL) = KeyPressed) Then             'Gerbing 22.01.2010
    If (GetAsyncKeyState(VK_SHIFT) = KeyPressed) Then             'Gerbing 22.01.2010
        CommandLine = ""                                            'Gerbing 22.01.2010
    End If                                                          'Gerbing 22.01.2010
    'MsgBox "commandline = " & CommandLine
    Select Case CommandLine                                 'Gerbing 14.11.2007
        Case ""
            'Kommandozeilen-Parameter vorhanden - nein
            Me.Show
            Me.Width = Screen.Width / 2                     'Gerbing 27.01.2018
            Me.Height = Screen.Height / 2                   'Gerbing 27.01.2018
            If Not F8Gedrückt Then
                If chkSortAsDragAndDrop.Value = 1 Then     'Gerbing 22.01.2010
                    List1UnsortedU.ListItems.RemoveAll
                Else
                    List1U.ListItems.RemoveAll
                End If
                ListBoxForm.ExLVwU.Columns.RemoveAll
            End If
            lblDateinamen.Visible = True
            'cmdSpeichern.Visible = True
            cmdStartAfterDragDrop.Visible = True
            List1UnsortedU.Visible = True
            List1U.Visible = True
        '------------------------------------------------------------------------------------------------------
        Case Else            'Kommandozeilen-Parameter vorhanden - ja
            'Kommandozeilen-Parameter vorhanden - ja
            '"C:\P5Daten\VISUALBA.SIC\Foto\Diashow\Beispielfoto1.jpg"
            'Me.Show                                                                        'Gerbing 28.11.2007
'            MessageBoxW 0, StrPtr(CommandLine), StrPtr(DiashowForm.Caption), MB_ICONINFORMATION Or MB_TASKMODAL
'            MessageBoxW 0, StrPtr(command), StrPtr(DiashowForm.Caption), MB_ICONINFORMATION Or MB_TASKMODAL
            List1U.Visible = False
            List1UnsortedU.Visible = False
            Dateiname = CommandLine
            'MsgBox "Commandline=" & CommandLine
            If Left(CommandLine, 1) = """" Then                                             'Gerbing 19.03.2005
                'Wenn der Dateiname Leerzeichen enthält, müssen die doppelten Hochkomma am Anfang und Ende bleiben
                pos = InStr(1, Dateiname, " ")
                If pos = 0 Then
                    Dateiname = Mid(CommandLine, 2, Len(CommandLine) - 2)
                End If
            End If
            
            Call LeseDiashowDatei
            'ob der Aufrufparameter eine Liste von Dateien war oder eine Diashow-Datei
            '--------------------------------------------
            blnOhneUnterOrdner = True
            If List1U.ListItems.Count = 1 And List1UnsortedU.ListItems.Count = 0 Then                     'Gerbing 12.03.2008
                frmDurchsuchenUnterOrdner.Show 1                                            'Gerbing 30.07.2008
            Else
                KopiereListBox             'die Listbox auf der ListBoxForm anordnen                'Gerbing 16.12.2008
            End If
            '---------------------------------------------
            If List1U.ListItems.Count <> 1 Then
                Call GeheZumErstenBild
            End If
            If EinzelnesBild = "" Then
                Call GeheZumErstenBild
            Else
                GeheZumEinzelnenBild                                                        'Gerbing 28.11.2007
            End If
            Screen.MousePointer = vbDefault                                                 'Gerbing 30.07.2008
            '------------------------------------------------------------------------------- Gerbing 16.12.2008
            Call BildAnzeigen
    End Select
End Sub

Private Sub GeheEinBildVorwärts()
    If ListBoxForm.ExLVwUIndex < ListBoxForm.ExLVwU.ListItems.Count - 1 Then
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False
        ListBoxForm.ExLVwUIndex = ListBoxForm.ExLVwUIndex + 1
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
        Call ZufallsListindex
    Else
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False
        GeheZumErstenBild
    End If
End Sub

Private Sub GeheEinBildZurück()
    If ListBoxForm.ExLVwUIndex > 0 Then
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False
        ListBoxForm.ExLVwUIndex = ListBoxForm.ExLVwUIndex - 1
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
        Call ZufallsListindex
    Else
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False
        GeheZumLetztenBild
    End If
End Sub

Private Sub GeheZumErstenBild()
    ListBoxForm.ExLVwUIndex = 0
    ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
End Sub

Private Sub GeheZumLetztenBild()
    ListBoxForm.ExLVwUIndex = ListBoxForm.ExLVwU.ListItems.Count - 1
    ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
End Sub

Private Sub GeheZumEinzelnenBild()
    Dim n As Long
    Dim Gefunden As Boolean
    
    Gefunden = False
    For n = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1
        ListBoxForm.ExLVwUIndex = n
        If StrComp(EinzelnesBild, ListBoxForm.ExLVwU.ListItems(n), vbTextCompare) = 0 Then
            ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
            Gefunden = True
            Exit For                                                                    'Gerbing 13.06.2013
        Else
            ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False
        End If
    Next n
    If Gefunden = False Then                                                            'Gerbing 25.08.2013
        ListBoxForm.ExLVwUIndex = 0
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True
    End If
End Sub

Private Sub IniWerteMerken(FileData)
    Dim msg As String
    
    'MessageBoxW 0, StrPtr(FileData), StrPtr(DiashowForm.Caption), MB_ICONINFORMATION Or MB_TASKMODAL
    ZeilenNummer = ZeilenNummer + 1
    If FileData <> "" Then
        If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
            List1UnsortedU.ListItems.Add FileData
        Else
            List1U.ListItems.Add FileData
        End If
    Else
        Call WerteFehler(FileData)
    End If
End Sub

Private Sub Form_Resize()
'    Static WindowState As FormWindowStateConstants
'    ' see if window state has changed
'    If WindowState <> Me.WindowState Then
'        ' use custom caption updating in MDI form
'        MDIForm1.CaptionW
'        ' remember current window state
'        WindowState = Me.WindowState
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim antwort As Long
    
    '---------------------------------------------------------------'Gerbing 07.12.2009
    If gblnHäkchenGefunden = False Then                             'Gerbing 19.10.2017
        For i = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1
                If ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 2 Or ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 4 Then
                    gblnHäkchenGefunden = True
                End If
            Next
        If gblnHäkchenGefunden = True Then
            'msg = "Es gibt mit Häkchen markierte Dateien, wollen Sie diese wirklich ignorieren?"
            msg = LoadResString(3149 + Sprache)
            antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
            If antwort = vbYes Then
                Exit Sub
            Else
                Cancel = True
                Exit Sub
            End If
        End If
    End If

    Set fso = Nothing
    If gblnWithTitle = False Then
        Call MDIFormMitTitle
    End If
    Set DiashowForm.EXF = Nothing
    Unload frmDurchsuchenUnterOrdner
    Unload HilfeBoxForm
    Unload frmBildMitGDIPlus
    Unload WertxForm
    Unload XYPos
    Unload ListBoxForm
    Unload frmMap                                                   'Gerbing 29.09.2018
    End
End Sub

Private Sub List1U_AbortedDrag()
    List1U.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, List1U.ListItems(0)
    End Sub

Private Sub List1U_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyEscape Then
        If Not (List1U.DraggedItems Is Nothing) Then List1U.EndDrag True
    End If
End Sub

Private Sub List1U_OLEDragDrop(ByVal Data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim files() As String
    Dim KeineÜbereinstimmung As Boolean
    Dim DateinamenErweiterung As String
    Dim i As Long
    Dim n As Long
    Dim insertAfter As Variant

    'On Error GoTo NoFiles
    files = Data.GetData(vbCFFiles)
    For n = LBound(files) To UBound(files)
        KeineÜbereinstimmung = True
        'Sämtliche Einträge in List1 prüfen, ob sie übereinstimmen
        'mit Data.Files(n)
        If chkSortAsDragAndDrop.Value = 1 Then     'Gerbing 22.01.2010
            If List1UnsortedU.ListItems.Count <> 0 Then
                KeineÜbereinstimmung = True
                'List1Unsortedu.ListIndex = 0
                For i = 0 To List1UnsortedU.ListItems.Count - 1
                    'If List1Unsorted.List(i) = Data.Files(n) Then
                    If StrComp(List1UnsortedU.ListItems(i), files(n), vbTextCompare) = 0 Then 'Gerbing 28.11.2007
                        KeineÜbereinstimmung = False
                        Exit For
                    End If
                Next i
            End If
        Else
            If List1U.ListItems.Count <> 0 Then
                KeineÜbereinstimmung = True
                'List1.ListIndex = 0
                For i = 0 To List1U.ListItems.Count - 1
                    'If List1.List(i) = Data.Files(n) Then
                    If StrComp(List1U.ListItems(i), files(n), vbTextCompare) = 0 Then    'Gerbing 28.11.2007
                        KeineÜbereinstimmung = False
                        Exit For
                    End If
                Next i
            End If
        End If
        '------------------------------------------------
        'Nur Einträge mit keineÜbereinstimmung übernehmen
        If KeineÜbereinstimmung = True Then
            DateinamenErweiterung = Right(files(n), 3)
            DateinamenErweiterung = UCase(DateinamenErweiterung)
            Select Case DateinamenErweiterung
                Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "WMF"           'Gerbing 29.03.2012
                    If chkSortAsDragAndDrop.Value = 1 Then     'Gerbing 22.01.2010
                        'List1UnsortedU.ListItems.Add Data.files(n)
                        Set insertAfter = List1UnsortedU.ListItems.Add(files(n), , 1)
                    Else
                        'List1U.ListItems.Add Data.files(n)
                        Set insertAfter = List1U.ListItems.Add(files(n), , 1)
                    End If
                Case Else
                    If MsgBoxMussKommen = True Then
        '                            msg = "Nur Dateinamen-Erweiterungen "
        '                            msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL   'Gerbing 29.03.2012
        '                            msg = msg & "sind erlaubt." & NL
        '                            msg = msg & "Nur die gültigen Dateinamen-Erweiterungen werden übernommen." & NL
        '                            msg = msg & "Dieser Hinweis wird nicht wiederholt."
                        msg = LoadResString(2304 + Sprache)
                        msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL    'Gerbing 29.03.2012
                        msg = msg & LoadResString(2305 + Sprache) & NL
                        msg = msg & LoadResString(2306 + Sprache) & NL
                        msg = msg & LoadResString(2094 + Sprache)
                        MsgBox msg
                        MsgBoxMussKommen = False
                    End If
            End Select
        End If
        If n > 32766 Then                                                               'Gerbing 24.01.2009
            'MsgBox "Sie haben mehr als 32767 Bilder angeboten. Das Programm hat die Anzahl von Bildern auf 32767 begrenzt."
            MsgBox LoadResString(2332 + Sprache)
            Exit For
        End If
    Next n
    Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
    ReDim BildPosList(List1U.ListItems.Count)                   'Gerbing 30.05.2013
    If List1U.ListItems.Count > 0 Then
        For i = 0 To List1U.ListItems.Count - 1
           BildPosList(i).Top = 0
           BildPosList(i).Left = 0
           BildPosList(i).ZoomPercent = 0
        Next i
    End If
'NoFiles:
End Sub

Private Sub List1U_RecreatedControlWindow(ByVal hwnd As Long)
    'InsertItems
End Sub

Private Sub List1UnsortedU_AbortedDrag()
    List1UnsortedU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, List1U.ListItems(0)
End Sub

Private Sub List1UnsortedU_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyEscape Then
        If Not (List1UnsortedU.DraggedItems Is Nothing) Then List1UnsortedU.EndDrag True
    End If
End Sub

Private Sub List1UnsortedU_OLEDragDrop(ByVal Data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
    Dim files() As String
    Dim KeineÜbereinstimmung As Boolean
    Dim DateinamenErweiterung As String
    Dim i As Long
    Dim n As Long
    Dim insertAfter As Variant

    'On Error GoTo NoFiles
    files = Data.GetData(vbCFFiles)
    For n = LBound(files) To UBound(files)
        KeineÜbereinstimmung = True
        'Sämtliche Einträge in List1 prüfen, ob sie übereinstimmen
        'mit Data.Files(n)
        If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
            If List1UnsortedU.ListItems.Count <> 0 Then
                KeineÜbereinstimmung = True
                'List1Unsortedu.ListIndex = 0
                For i = 0 To List1UnsortedU.ListItems.Count - 1
                    'If List1Unsorted.List(i) = Data.Files(n) Then
                    If StrComp(List1UnsortedU.ListItems(i), files(n), vbTextCompare) = 0 Then 'Gerbing 28.11.2007
                        KeineÜbereinstimmung = False
                        Exit For
                    End If
                Next i
            End If
        Else
            If List1U.ListItems.Count <> 0 Then
                KeineÜbereinstimmung = True
                'List1.ListIndex = 0
                For i = 0 To List1U.ListItems.Count - 1
                    'If List1.List(i) = Data.Files(n) Then
                    If StrComp(List1U.ListItems(i), files(n), vbTextCompare) = 0 Then    'Gerbing 28.11.2007
                        KeineÜbereinstimmung = False
                        Exit For
                    End If
                Next i
            End If
        End If
        '------------------------------------------------
        'Nur Einträge mit keineÜbereinstimmung übernehmen
        If KeineÜbereinstimmung = True Then
            DateinamenErweiterung = Right(files(n), 3)
            DateinamenErweiterung = UCase(DateinamenErweiterung)
            Select Case DateinamenErweiterung
                Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "WMF"           'Gerbing 29.03.2012
                    If chkSortAsDragAndDrop.Value = 1 Then     'Gerbing 22.01.2010
                        'List1UnsortedU.ListItems.Add Data.files(n)
                        Set insertAfter = List1UnsortedU.ListItems.Add(files(n), , 1)
                    Else
                        'List1U.ListItems.Add Data.files(n)
                        Set insertAfter = List1U.ListItems.Add(files(n), , 1)
                    End If
                Case Else
                    If MsgBoxMussKommen = True Then
        '                            msg = "Nur Dateinamen-Erweiterungen "
        '                            msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL   'Gerbing 29.03.2012
        '                            msg = msg & "sind erlaubt." & NL
        '                            msg = msg & "Nur die gültigen Dateinamen-Erweiterungen werden übernommen." & NL
        '                            msg = msg & "Dieser Hinweis wird nicht wiederholt."
                        msg = LoadResString(2304 + Sprache)
                        msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL    'Gerbing 29.03.2012
                        msg = msg & LoadResString(2305 + Sprache) & NL
                        msg = msg & LoadResString(2306 + Sprache) & NL
                        msg = msg & LoadResString(2094 + Sprache)
                        MsgBox msg
                        MsgBoxMussKommen = False
                    End If
            End Select
        End If
        If n > 32766 Then                                                               'Gerbing 24.01.2009
            'MsgBox "Sie haben mehr als 32767 Bilder angeboten. Das Programm hat die Anzahl von Bildern auf 32767 begrenzt."
            MsgBox LoadResString(2332 + Sprache)
            Exit For
        End If
    Next n
    Me.MousePointer = vbDefault                                                         'Gerbing 29.07.2007
    ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count)                   'Gerbing 30.05.2013
    If List1U.ListItems.Count > 0 Then
        For i = 0 To List1U.ListItems.Count - 1
           BildPosList(i).Top = 0
           BildPosList(i).Left = 0
           BildPosList(i).ZoomPercent = 0
        Next i
    End If
'NoFiles:
End Sub

Private Sub mnuBeenden_Click()
    End
End Sub

Private Sub mnuHilfe_Click()
    Dim retval As Long
    Dim DateinamenErweiterung As String
    Dim CHMFile As String
    Dim ErrorText As String
    Dim msg As String

    If Sprache = 0 Then
        'CommonDialog1.HelpFile = AppPath & "\help\deutsch\diashow.HLP"
        CHMFile = AppPath & "\Help\Deutsch\diashow.CHM"                         'Gerbing 14.03.2007
        retval = ShellExecute(Me.hwnd, "open", CHMFile, vbNull, vbNull, 1)
        If retval <= 32 Then
            DateinamenErweiterung = "CHM"
            ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
            msg = "Errortext=" & ErrorText & vbNewLine
            msg = msg & "Errornr=" & retval & vbNewLine & vbNewLine
            
            msg = msg & CHMFile & vbNewLine
            'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
            msg = msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
            
            'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
            msg = msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
            
            'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
            msg = msg & LoadResString(1378 + Sprache) & vbNewLine
            'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
            msg = msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
            'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
            msg = msg & LoadResString(2012 + Sprache) & vbNewLine
            'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
            msg = msg & LoadResString(2013 + Sprache)
            'MsgBox msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
        End If

    Else
        'CommonDialog1.HelpFile = AppPath & "\help\english\diashow.HLP"
        CHMFile = AppPath & "\Help\English\diashow.CHM"                           'Gerbing 14.03.2007
        retval = ShellExecute(Me.hwnd, "open", CHMFile, vbNull, vbNull, 1)
        If retval <= 32 Then
            DateinamenErweiterung = "CHM"
            ErrorText = GetShellError(retval)           'Gerbing 20.08.2008
            msg = "Errortext=" & ErrorText & vbNewLine
            msg = msg & "Errornr=" & retval & vbNewLine & vbNewLine
            
            msg = msg & CHMFile & vbNewLine
            'Msg = Msg & "Diese Datei kann nicht geöffnet werden." & vbNewLine & vbNewLine
            msg = msg & LoadResString(1376 + Sprache) & vbNewLine & vbNewLine
            
            'Msg = Msg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
            msg = msg & LoadResString(2208 + Sprache) & vbNewLine & vbNewLine
            
            'Msg = Msg & "oder es ist keine Anwendung mit der" & vbNewLine
            msg = msg & LoadResString(1378 + Sprache) & vbNewLine
            'Msg = Msg & "Dateinamen-Erweiterung(Datei-Typ) " & DateinamenErweiterung & " verknüpft." & vbNewLine
            msg = msg & LoadResString(1379 + Sprache) & DateinamenErweiterung & LoadResString(1380 + Sprache) & vbNewLine
            'Msg = Msg & "Wählen Sie selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
            msg = msg & LoadResString(2012 + Sprache) & vbNewLine
            'Msg = Msg & "Rechtklicken auf den Dateiname -> Öffnen mit... -> Programm auswählen"
            msg = msg & LoadResString(2013 + Sprache)
            'MsgBox msg
            MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
        End If
    End If
End Sub

Private Sub mnuVersion_Click()
    Dim msg As String
    
    'Versions-Informationen ermitteln                               'Gerbing 28.08.2013
    msg = "Version " & GetDiashowExeVersion                         'Gerbing 28.08.2013
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbInformation
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    HilfeBoxForm.Hide                                                                       'Gerbing 15.07.2012
    On Error Resume Next
    On Error GoTo 0
    Call GeheEinBildVorwärts
    Unload frmBildMitGDIPlus
    Call BildAnzeigen
    Timer1.Enabled = True
End Sub

Private Sub WerteFehler(FileData)
    Dim msg
    
'    msg = "Fehlerhafte Zeile in " & DateiName & NL
'    msg = msg & FileData & NL
'    msg = msg & "ZeilenNummer: " & ZeilenNummer
    msg = LoadResString(2307 + Sprache) & Dateiname & NL
    msg = msg & FileData & NL
    msg = msg & LoadResString(2308 + Sprache) & ZeilenNummer
    'MsgBox msg, vbCritical
    MessageBoxW 0, StrPtr(msg), StrPtr("GERBING Diashow"), vbCritical
    End
End Sub

Public Sub KopiereListBox()
    Dim n As Long
    
    'ListBoxForm.ExLVwU.hImageList(ilSmall) = hImgLst                                   'auskommentiert sonst kommt calendar.ico noch zusätzlich
    ListBoxForm.ExLVwU.hImageList(ilState) = hStateImgLst
    
    Screen.MousePointer = vbHourglass                                                       'Gerbing 24.01.2009
    ListBoxForm.ExLVwU.Columns.RemoveAll
    If List1U.ListItems.Count = 1 Then                                                             'Gerbing 23.10.2007
        'es ist ein einzelner filename
        EinzelnesBild = List1U.ListItems(0)
        List1U.ListItems.Count                                                                     'Gerbing 28.11.2007
        'For n = 0 To lstFilesImGleichenOrdner.ListItems.Count - 1                               'Gerbing 28.11.2007
        For n = 0 To UBound(FileNameArray)
            ListBoxForm.ExLVwU.ListItems.Add(FileNameArray(n), , 0, 0, n).StateImageIndex = 1
            'List1U.ListItems.Add FileNameArray(n)
            If n Mod 10 = 0 Then
                On Error Resume Next
                formCaption ListBoxForm.hwnd, FileNameArray(n)                              'Arbeitsfortschritt anzeigen alle 10 files
                On Error GoTo 0
            End If
            DoEvents
        Next n
        Screen.MousePointer = vbNormal                                                       'Gerbing 24.01.2009
        List1U.ZOrder                                                                        'Gerbing 22.01.2010
        Exit Sub
    End If
    '--------------------------------
    'es ist mehr als ein einzelner filename
    Screen.MousePointer = vbHourglass
    If chkSortAsDragAndDrop.Value = 1 Then     'Gerbing 22.01.2010
        For n = 0 To List1UnsortedU.ListItems.Count - 1
            ListBoxForm.ExLVwU.ListItems.Add(List1UnsortedU.ListItems(n), , 0, 0, n).StateImageIndex = 1
        Next n
        List1UnsortedU.ZOrder                                                                'Gerbing 22.01.2010
    Else
        For n = 0 To List1U.ListItems.Count - 1
            ListBoxForm.ExLVwU.ListItems.Add(List1U.ListItems(n), , 0, 0, n).StateImageIndex = 1
        Next n
        List1U.ZOrder                                                                        'Gerbing 22.01.2010
        Screen.MousePointer = vbNormal                                                       'Gerbing 24.01.2009
    End If
End Sub

Private Sub LeseDiashowDatei()
    'Beispiel für Dateinamen-Liste
    'D:\FotosGG\1851\BurgRabenstein03.jpg "D:\FotosGG\1851\Chemnitz 1761.jpg"
    Dim FNum As Long
    Dim msg As String
    Dim start As Long
    Dim pos As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim EinzelName As String
    Dim DiaDatei As CEncodedFile
    Dim FileData As String
    Dim rc As Boolean
    Dim oStream As TextStream
    
    pos = InStr(1, Dateiname, ".dia", vbTextCompare)
    If pos <> Len(Dateiname) - 3 And pos <> Len(Dateiname) - 4 Then
'        'Wenn nicht 'dia' ganz rechts steht, ist es eine Dateinamen-Liste
'        'Für Testzwecke die angebotenen Liste in Parameter.txt schreiben
'        FNum = FreeFile
'        Open App.Path & "\Parameter.txt" For Output As FNum
'        Print #FNum, Dateiname
'        Close FNum
    
        'Es wird vermutet, daß der Aufrufparameter eine Liste von Dateien beeinhaltet, die durch Drag&Drop
        'auf das Programm Icon entstanden ist.
        'Es geht noch einfacher als bisher, wenn man eine Liste von Dateinamen zB aus dem Explorer auf
        'das Programm Icon von Diashow.exe zieht.
        'Da wird eine Liste von Dateinamen angeboten, wobei
        'A. bei Namen mit Leerzeichen - der Dateiname in doppeltes Hochkomma vorn und hinten eingeschlossen
        'ist und ein Leerzeichen den Dateiname vom nächsten trennt
        'B. bei Namen ohne Leerzeichen - Ein Leerzeichen den Dateiname vom nächsten trennt
        'Die Einzelnamen werden übernommen nach List1
        start = 1
        Do
            pos = InStr(start, Dateiname, " ")  'gibt es ein Leerzeichen zwischen Start und String-Ende
            If pos = 0 Then
                'nein
                'es ist der letzte Einzelname ohne Leerzeichen
                EinzelName = Mid(Dateiname, start, Len(Dateiname) - start + 1)
                Call ÜbernehmeNurGültige(EinzelName)
                Exit Do
            End If
            'Gibt es doppelte Hochkomma sowohl von Start bis Leerzeichen wie nach dem Leerzeichen
            pos1 = InStr(start, Dateiname, """")
            pos2 = InStr(pos, Dateiname, """")
            If pos1 < pos And pos2 <> 0 Then
                'ja, es ist ein EinzelName mit Leerzeichen
                EinzelName = Mid(Dateiname, pos1 + 1, pos2 - pos1 - 1)
                Call ÜbernehmeNurGültige(EinzelName)
                start = pos2 + 1
                If start > Len(Dateiname) Then Exit Do
                start = start + 1
            Else
                'es ist ein Einzelname ohne Leerzeichen
                EinzelName = Mid(Dateiname, start, pos - start)
                Call ÜbernehmeNurGültige(EinzelName)
                start = pos + 1
            End If
        Loop
        If List1U.ListItems.Count = 0 Then
'            MsgBox "Sie haben per Drag&Drop keine gültigen Dateinamen-Erweiterungen angeboten"
            MsgBox LoadResString(2309 + Sprache)
            End
        End If
        Exit Sub
    '-------------------------------------------------------------------------------------------------------------------
    Else
        'wenn als Aufrufparameter eine dia-Datei angegeben ist, dann muß sie existieren
        'Wenn der Dateiname doppelte Hochkomma am Anfang und Ende enthält, müssen diese verschwinden
        pos = InStr(1, Dateiname, """")
        If pos = 1 Then
            Dateiname = Mid(Dateiname, 2, Len(Dateiname) - 2)
        End If
        If file_exist(Dateiname) = False Then
'            msg = "Sie haben folgende Dia-Show-Datei angegeben:" & NL
'            msg = msg & DateiName & NL
'            msg = msg & "Diese Datei existiert nicht"
            msg = LoadResString(2310 + Sprache) & NL
            msg = msg & Dateiname & NL
            msg = msg & LoadResString(2311 + Sprache)
            MessageBoxW 0, StrPtr(msg), StrPtr(DiashowForm.Caption), MB_ICONINFORMATION Or MB_TASKMODAL
            End
        End If
        On Error GoTo 0
        blnVorhandeneDiashowDateiBenutzen = True
        ZeilenNummer = 0
        start = 1
        'dia-Datei mit FSO öffnen
        If (oStream Is Nothing) Then
            ' Open the file for reading.
            If IsStringUnicode(Dateiname) = True Then                       'ist es ein Dateiname mit unicode
                Set oStream = fso.OpenTextFile(Dateiname, 1, False, -1)     'unicode
            Else
                Set oStream = fso.OpenTextFile(Dateiname, 1, False, 0)      'ansi
                'jetzt können trotzdem in der Datei unicode zeichen sein
                If (Not oStream Is Nothing) Then
                    With oStream
                        'die ersten beiden Bytes auslesen
                        pos = 2
                        FileData = oStream.Read(pos)
                        '.Close
                    End With
                    If FileData = "ÿþ" Then
                        Set oStream = Nothing
                        'die ersten beiden Bytes enthalten das unicode Kennzeichen
                        Set oStream = fso.OpenTextFile(Dateiname, 1, False, -1)     'unicode
                    End If
                End If
            End If
            If (Not oStream Is Nothing) Then
                With oStream
                    Do Until oStream.AtEndOfStream
                        FileData = oStream.ReadLine
                        Call IniWerteMerken(FileData)
                    Loop
                    .Close
                End With
                Set oStream = Nothing
            End If
        End If
    End If
    If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
        If List1UnsortedU.ListItems.Count = 0 Then
'            MsgBox "Formatfehler in " & DateiName, vbCritical
            'MsgBox LoadResString(2312 + Sprache) & Dateiname, vbCritical
            MessageBoxW 0, StrPtr(LoadResString(2312 + Sprache) & Dateiname), StrPtr("GERBING Diashow"), vbCritical
            End
        End If
    Else
        If List1U.ListItems.Count = 0 Then
'            MsgBox "Formatfehler in " & DateiName, vbCritical
            'MsgBox LoadResString(2312 + Sprache) & Dateiname, vbCritical
            MessageBoxW 0, StrPtr(LoadResString(2312 + Sprache) & Dateiname), StrPtr("GERBING Diashow"), vbCritical
            End
        End If
    End If
End Sub

Private Sub ZufallsListindex()
    Dim sngWert As Single
    Dim lngWert As Long
    Dim Obergrenze As Long
    Dim Untergrenze As Long

    If F6TimerGestartet = True And WertxForm.OptAlphabetisch = False Then
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = False              'Gerbing 28.09.2016
        'Obergrenze = 999999
        Obergrenze = ListBoxForm.ExLVwU.ListItems.Count
        Untergrenze = 1
        Randomize
        If ListBoxForm.ExLVwU.ListItems.Count = 0 Then Exit Sub
        sngWert = (Obergrenze - Untergrenze + 1) * Rnd + Untergrenze
        lngWert = Round(sngWert)
        
        If lngWert >= ListBoxForm.ExLVwU.ListItems.Count Then
            ListBoxForm.ExLVwUIndex = ListBoxForm.ExLVwU.ListItems.Count - 1
        Else
            ListBoxForm.ExLVwUIndex = lngWert
        End If
        ListBoxForm.ExLVwU.ListItems(ListBoxForm.ExLVwUIndex).Selected = True              'Gerbing 28.09.2016
    End If
End Sub

Private Sub ÜbernehmeNurGültige(EinzelName As String)
    Dim DateinamenErweiterung As String
    Dim Punkt As String

    'MessageBoxW 0, StrPtr(EinzelName), StrPtr("GERBING  Diashow"), MB_TASKMODAL
    DateinamenErweiterung = Right(EinzelName, 3)
    DateinamenErweiterung = UCase(DateinamenErweiterung)
    Select Case DateinamenErweiterung
        Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "WMF"           'Gerbing 29.03.2012
            Punkt = Mid(EinzelName, Len(EinzelName) - 3, 1)
            If Punkt = "." Then
                List1U.ListItems.Add EinzelName
            End If
        Case Else
            If MsgBoxMussKommen = True Then
'                    msg = "Nur Dateinamen-Erweiterungen "
'                    msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL 'Gerbing 29.03.2012
'                    msg = msg & "sind erlaubt." & NL
'                    msg = msg & "Nur die gültigen Dateinamen-Erweiterungen werden übernommen." & NL
'                    msg = msg & "Dieser Hinweis wird nicht wiederholt."
                    msg = LoadResString(2304 + Sprache)
                    msg = msg & "'BMP', 'DIB', 'EMF', 'GIF', 'ICO', 'JPG', 'PNG', 'TIF', 'WMF'" & NL 'Gerbing 29.03.2012
                    msg = msg & LoadResString(2305 + Sprache) & NL
                    msg = msg & LoadResString(2306 + Sprache) & NL
                    msg = msg & LoadResString(2094 + Sprache)
                MsgBox msg
                MsgBoxMussKommen = False
            End If
    End Select
End Sub

Private Sub NeuStart()
    Dim i As Long
    Dim antwort As Long
    Dim Filter As String
    Dim flags As Long

    If blnVorhandeneDiashowDateiBenutzen = True Then 'vorhandene Diashow-Datei benutzen, List1Unsorted benutzen
        'Show the open dialog and get the file path as Unicode in a VB String
        Dateiname = ShowOpenUnicode(Me)
            If Left$(Dateiname, 1) = Chr(0) Then
            'MsgBox "Das Programm wird beendet"
            MsgBox LoadResString(2139 + Sprache)
            Set fso = Nothing
            If gblnWithTitle = False Then
                Call MDIFormMitTitle
            End If
            End
        End If
        'Convert the file name to be used
        Dateiname = ConvertFileName(Dateiname)
        DiashowForm.blnListBoxNeuDblClick = True                'Gerbing 16.12.2008 siehe 05.05.2005
        DiashowForm.Caption = ""
        If blnVorhandeneDiashowDateiBenutzen = True Then
            List1UnsortedU.ListItems.RemoveAll
            ListBoxForm.ExLVwU.ListItems.RemoveAll
        Else
            List1U.ListItems.RemoveAll
            ListBoxForm.ExLVwU.ListItems.RemoveAll
        End If
        Call LeseDiashowDatei
        Call KopiereListBox             'die Listbox auf der ListBoxForm anordnen
        If blnVorhandeneDiashowDateiBenutzen = False Then
            ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count)                   'Gerbing 30.05.2013
            If List1U.ListItems.Count > 0 Then
                For i = 0 To List1U.ListItems.Count - 1
                   BildPosList(i).Top = 0
                   BildPosList(i).Left = 0
                   BildPosList(i).ZoomPercent = 0
                Next i
            End If
        Else
            ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count)                   'Gerbing 30.05.2013
            If List1UnsortedU.ListItems.Count > 0 Then
                For i = 0 To List1UnsortedU.ListItems.Count - 1
                   BildPosList(i).Top = 0
                   BildPosList(i).Left = 0
                   BildPosList(i).ZoomPercent = 0
                Next i
            End If
        End If
        Call GeheZumErstenBild
        Screen.MousePointer = vbDefault                                                             'Gerbing 16.12.2008
        Call BildAnzeigen
    '------------------------------------------------------------------------------------------------------------------
    Else    'dann war Drag&Drop ausgewählt
        If Not F8Gedrückt Then
            If chkSortAsDragAndDrop.Value = 1 Then    'Gerbing 22.01.2010
                List1UnsortedU.ListItems.RemoveAll
            Else
                List1U.ListItems.RemoveAll
                ListBoxForm.ExLVwU.ListItems.RemoveAll
            End If
        End If
        lblDateinamen.Visible = True
        'cmdSpeichern.Visible = True
        cmdStartAfterDragDrop.Visible = True
        List1UnsortedU.Visible = True
        List1U.Visible = True
        If chkSortAsDragAndDrop.Value = 1 Then                                                          'Gerbing 22.01.2010
            List1UnsortedU.ZOrder
        End If
    End If
End Sub

Private Sub AbrufenBildPosList()                          'Gerbing 29.07.2006
    Dim i As Long
    
    'Wenn WertxForm.CheckSpeichernBildPosition eingeschaltet ist, wird die zu ListBoxForm.ExLVwU.ListItems(i).Text
    'gehörende Array-Zeile aus BildPosList
    'zur Bildpositionierung benutzt
    'aber nicht, wenn alle 3 Einträge Null sind
    
    On Error GoTo 0
    If WertxForm.CheckSpeichernBildPosition.Value = 1 Then
        For i = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1             'Gerbing 30.05.2013
            If ListBoxForm.ExLVwU.ListItems(i).Selected = True Then
                Exit For
            End If
        Next i
        
'        i = ListBoxForm.ExLVwUIndex
        If i <= UBound(BildPosList) Then
            If BildPosList(i).Left = 0 And BildPosList(i).Top = 0 And BildPosList(i).ZoomPercent = 0 Then
            '
            Else
                glngDiffX = BildPosList(i).Left
                glngDiffY = BildPosList(i).Top
                ZoomProzent = BildPosList(i).ZoomPercent
            End If
        End If
    End If
End Sub

Public Sub SpeichernInBildPosList()                          'Gerbing 29.07.2006
    Dim i As Long

    'Wenn WertxForm.CheckSpeichernBildPosition eingeschaltet ist, wird die zu ListBoxForm.ListBoxNeu.Text
    'gehörende Array-Zeile in BildPosList
    'zur Bildpositionierung gespeichert
    
    If WertxForm.CheckSpeichernBildPosition.Value = 1 Then
        i = ListBoxForm.ExLVwUIndex
        'On Error Resume Next
        BildPosList(i).Left = glngDiffX
        BildPosList(i).Top = glngDiffY
        BildPosList(i).ZoomPercent = ZoomProzent
        'On Error GoTo 0
    End If
End Sub

Public Sub EsIstF8()                                               'Gerbing 12.03.2008
    Dim blnHäkchenGefunden As Boolean
    Dim i As Long
    Dim antwort As Long
    
    '---------------------------------------------------------------'Gerbing 07.12.2009
    blnHäkchenGefunden = False
    For i = 0 To ListBoxForm.ExLVwU.ListItems.Count - 1
           If ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 2 Or ListBoxForm.ExLVwU.ListItems(i).StateImageIndex = 4 Then
                blnHäkchenGefunden = True
           End If
        Next
    If blnHäkchenGefunden = True Then
        'msg = "Es gibt mit Häkchen markierte Dateien, wollen Sie diese wirklich ignorieren?"
        msg = LoadResString(3149 + Sprache)
        antwort = MsgBox(msg, vbDefaultButton1 + vbYesNo)
        If antwort = vbNo Then
            frmBildMitGDIPlus.Show                                  'Gerbing 24.04.2012
            Exit Sub
        End If
    End If
    '----------------------------------------------------------------------------------------
    Timer1.Enabled = False      'Timer anhalten
    F8Gedrückt = True
    cmdKeinNeustart.Visible = True
    'DiashowForm.Caption = "GERBING Diashow"    08.11.2005
    DiashowForm.Caption = LoadResString(1201 + Sprache)
    Call NeuStart
    If blnVorhandeneDiashowDateiBenutzen = False Then 'vorhandene Diashow-Datei benutzen    'Gerbing 09.01.2007
        DiashowForm.Show
    End If
    Me.MousePointer = vbDefault                                                 'Gerbing 29.07.2007
End Sub


Public Sub MDIFormMitTitle()
    Dim l As Long

    'title bar mitsamt ihren boxes wiederherstellen
    MDIForm1.Caption = strOldTitle
    l = GetWindowLong(MDIForm1.hwnd, GWL_STYLE)
    l = l Or (WS_SYSMENU)
    l = l Or (WS_THICKFRAME)
    l = l Or (WS_MINIMIZEBOX)
    l = l Or (WS_MAXIMIZEBOX)
    l = l Or WS_CAPTION
    l = SetWindowLong(MDIForm1.hwnd, GWL_STYLE, l)
        
    'Show taskbar
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    
    'MDIForm in voller Größe minus taskbar
    MDIForm1.WindowState = 0
    MDIForm1.Height = screenHeight * Screen.TwipsPerPixelY - 600
    MDIForm1.Width = Screen.Width * Screen.TwipsPerPixelX
    MDIForm1.Left = 0
    MDIForm1.Top = 0
    gblnWithTitle = True
End Sub

Public Sub MDIFormOhneTitle()
    Dim l As Long
        
    'title bar mitsamt ihren boxes verschwinden lassen
    strOldTitle = MDIForm1.Caption
    MDIForm1.Caption = ""
    l = GetWindowLong(MDIForm1.hwnd, GWL_STYLE)
    l = l And Not (WS_SYSMENU)
    l = l And Not (WS_THICKFRAME)
    l = l And Not (WS_MINIMIZEBOX)
    l = l And Not (WS_MAXIMIZEBOX)
    l = l And Not WS_CAPTION
    l = SetWindowLong(MDIForm1.hwnd, GWL_STYLE, l)
        
    'Hide taskbar
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    
    'MDIForm in voller Größe full screen zeichnen
    MDIForm1.WindowState = 0
    MDIForm1.Height = screenHeight * Screen.TwipsPerPixelY
    MDIForm1.Width = Screen.Width * Screen.TwipsPerPixelX
    MDIForm1.Left = 0
    MDIForm1.Top = 0
    gblnWithTitle = False
End Sub

Public Sub FüllelstFilesImGleichenOrdner()
    Dim start As Long
    Dim pos As Long
    Dim pos1 As Long
    Dim strTemp As String
    Dim n As Long
    
    Screen.MousePointer = vbHourglass
    DoEvents
    start = 1
    Do
        'hintersten backslash suchen, was davor steht ist der folder
        pos1 = InStr(start, DiashowForm.List1U.ListItems(0), "\")
        If pos1 = 0 Then Exit Do
        start = pos1 + 1
        pos = pos1
    Loop
    gstrFolder = Left(DiashowForm.List1U.ListItems(0), pos)
'    strTemp = Dir(strFolder, vbReadOnly)
'    If strTemp <> "" Then
        m_FileCount = 0
        ReDim FileNameArray(m_FileCount)
        Call Rekursive(gstrFolder, "*.*")
'    End If
    Call Aufräumen
End Sub

Private Function Rekursive(path As String, SearchStr As String)
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
    
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    If gblnSubdirectories = True Then
        Cont = True
        hSearch = FindFirstFileW(StrPtr(path & "*"), VarPtr(wfd))
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While Cont
            'DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            DirName = RemoveNulls((wfd.cFileName))
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
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
    hSearch = FindFirstFileW(StrPtr(path & SearchStr), VarPtr(wfd))
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
'                StartMillisek = timeGetTime
'                EndMillisek = timeGetTime
'                Debug.Print "EndMillisec=" & EndMillisek
'                Debug.Print "Millisekunden für StripNulls" & "=" & (EndMillisek - StartMillisek)

            'FileName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
            FileName = RemoveNulls((wfd.cFileName))
            If (FileName <> ".") And (FileName <> "..") Then
                '---------------------------------------------------
'                    StartMillisek = timeGetTime
                DateinamenErweiterung = UCase(Right(FileName, 3))
'                    EndMillisek = timeGetTime
'                    Debug.Print "EndMillisec=" & EndMillisek
'                    Debug.Print "Millisekunden für DateinamenErweiterung" & "=" & (EndMillisek - StartMillisek)
                Select Case DateinamenErweiterung
                    Case "BMP", "DIB", "EMF", "GIF", "ICO", "JPG", "PNG", "TIF", "WMF"           'Gerbing 29.03.2012
'                                StartMillisek = timeGetTime
                            ReDim Preserve FileNameArray(m_FileCount)
                            FileNameArray(m_FileCount) = path & FileName
                            m_FileCount = m_FileCount + 1

'                                EndMillisek = timeGetTime
'                                Debug.Print "EndMillisec=" & EndMillisek
'                                Debug.Print "Millisekunden für lstFilesImGleichenOrdner.ListItems.Add" & "=" & (EndMillisek - StartMillisek)
                End Select
                '---------------------------------------------------
            End If
'                StartMillisek = timeGetTime
            Cont = FindNextFileW(hSearch, VarPtr(wfd)) ' Get next file
'                EndMillisek = timeGetTime
'                Debug.Print "EndMillisec=" & EndMillisek
'                Debug.Print "Millisekunden für FindNextFileW" & "=" & (EndMillisek - StartMillisek)
        Wend
        Cont = FindClose(hSearch)
    End If
    If gblnSubdirectories = True Then
    ' If there are sub-directories...
        If nDir > 0 Then
            ' Recursively walk into them...
            For i = 0 To nDir - 1
                Rekursive = Rekursive(path & dirNames(i) & "\", SearchStr)
            Next i
        End If
    End If
End Function

Private Sub Aufräumen()
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    Call DiashowForm.KopiereListBox             'die Listbox auf der ListBoxForm anordnen
    ReDim BildPosList(ListBoxForm.ExLVwU.ListItems.Count)                   'Gerbing 30.05.2013
    Screen.MousePointer = vbHourglass
    DoEvents
    If DiashowForm.List1U.ListItems.Count > 0 Then
        For i = 0 To DiashowForm.List1U.ListItems.Count - 1
           BildPosList(i).Top = 0
           BildPosList(i).Left = 0
           BildPosList(i).ZoomPercent = 0
        Next i
    End If
    'Unload Me
End Sub




