VERSION 5.00
Begin VB.Form HistoryDesProgramms 
   Caption         =   "Kommentar"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "HistoryDesProgramms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�nderungen:
'11.06.2003 Fehler 'falscher Inhalt von Screen.Width, Screen.Height', wenn man nach dem Programmstart
'           die Bildschirmaufl�sung ver�ndert. Dieser Fehler l��t sich umgehen durch Verwendung der
'           GetDeviceCaps API Funktion.
'13.07.2003 Die Gr��e der Form Query darf nicht durch Ziehen an den Seitenkanten ver�ndert werden k�nnen.
'13.07.2003 Immer wenn die Form HilfeBx verschwindet mu� ein Form1.Image1.Refresh gemacht werden.
'           N�tig ist das f�r Fotos, die beim ersten Laden gr��er sind als die Bildschirmgr��e.
'14.12.2003 Bei Strg+B wird die Bildbreite x Bildh�he jeweils um 4 Pixel zu hoch ausgerechnet, also
'           subtrahiere ich jeweils 4
'16.12.2003 Immer bei Doppelklick auf eine Zeile in DbGridForm.DbGridNeu wird das Bild nicht in der Mitte
'           zentriert. Eine MsgBox w�rde helfen, ist aber eine Schei�l�sung.
'16.12.2003 Im XP ist die Sortierreihenfolge bei Dateinamen, wo Zahlen enthalten sind, anders als bei allen
'           vorangegangenen Windows-Versionen. Das w�rde mein bisheriges Konzept, der nachtr�glichen
'           Einf�gung �hnlicher Bilder total zerst�ren. Ab XP Service Pack 1 gibt es einen registry key
'           mit dem die bisher gewohnte Reihenfolge hergestellt werden kann.
'           Das muss unbedingt in der Hilfe-Datei beschrieben werden.
'22.12.2003 Wenn das Bild gr��er ist als ScreenWidth/ScreenHeight
'           oder wenn die Bildbreite >= 800 Pixel ist
'           oder wenn die Bildh�he >= 600 Pixel ist
'           wird es umgesetzt nach Image1 mit Stretch = True
'           das mache ich weil ich anfangs meine Fotos mit mit 800 x 600 Pixeln gespeichert hatte
'           und erst sp�ter zu 1024 x 768 �bergegeangen bin
'23.12.2003 Bei �nderung des Wertes Video DisplaySize durch den Nutzer (durch Rechtsklick aufs Video und
'           Wahl einer Zoomstufe fehlt das Zentrieren des Videos. Ich benutze daf�r ab jetzt einen Timer
'           der alle 100 Millisekunden drankommt
'23.12.2003 Fehler bei der Kopierfunktion
'           in $fotos.mdb steht immer derselbe Dateiname
'29.12.2003 Wahlweise kann bei der Suche nach Personen gew�hlt werden ob es voller Name sein soll
'           oder ob es Wortbestandteil sein soll
'13.01.2004 Neues Verfahren beim W�hlen des Fotos-Root-Ordner bei Start von CD
'           neuer Modul bFolder.bas
'23.01.2004 Wenn ein Benutzer von Windows 2000 oder XP keine Administrator-Rechte auf dem eigenen
'           PC hat, dann ist die Datenbankdatei fotos.mdb schreibgesch�tzt. Man mu� das dem Nutzer mitteilen.
'           Dann fragen, ob er mit nur Lesezugriff weiterarbeiten will, ansonsten mu� er daf�r sorgen,
'           da� er unter Administrator-Rechten arbeitet.
'24.01.2004 Die �nderung vom 22.12.2003 wird zur�ckgezogen. Die Bildqualit�t leidet bei automatischer
'           Vergr��erung von 800x600 auf 1024 x768
'09.02.2004 Ich brauche eine neue Funktion f�r Sortierreihenfolge, wenn ich Fotos in der Reihenfolge
'           betrachten will, wie zB in einem Buch angeordnet waren(nach Buchseiten)
'           und die Ber�cksichtigung von Jahreszahlen dabei au�er acht lassen will.
'           F�r diese Funktion wird die neue Tabelle 'FotosMitZusatzSpalte' der Datenbank 'fotos.mdb' benutzt
'10.02.2004 Ich brauche einen neuen Button 'weitere Filter...' dahinter steckt auch eine neue Funktion
'           Filtern nach Fotos, deren Datum (Ge�ndert am) einem bestimmten Zeitraum entsprechen muss.
'           Das ist notwendig, wenn ich nur die zuletzt hinzugef�gten Fotos wiederfinden will, weil ich sie
'           unter Schreibzugriff /WRITE mit Stichworten versehen will.
'           F�r diese Funktion wird die Tabelle 'FotosMitZusatzSpalte' der Datenbank 'fotos.mdb' benutzt
'11.02.2003 Es sind auch die Dateinamen-Erweiterungen "HTM", "PDF", "XLS" erlaubt.
'           Wenn diese gefunden werden, startet die mit der jeweiligen Dateinamen-Erweiterungen verbundene
'           Anwendung durch Benutzung von ShellExecute
'12.02.2004 Wenn mit /WRITE gestartet worden ist und das Kommentarfenster wird vom Nutzer beschrieben, dann
'           wird die �nderung ins Feld 'Kommentar' der Datenbank zur�ckgeschrieben. Auf diese Weise k�nnen auch
'           sehr lange Kommentare mittels Kopieren/Einf�gen erzeugt werden.
'25.02.2004 Auch bei Taste F8 mu� StichworteUpdate gemacht werden
'           StichworteUpdate hei�t, da� die �nderungen, die m�glicherweise garnicht in der Tabelle 'Fotos'
'           sondern in der Tabelle 'FotosMitZusatzSpalte' gemacht worden sind,
'           auf die Tabelle 'Fotos' zu �bertragen sind
'27.02.2004 Es fehlt bisher die M�glichkeit, nach Fotos zu suchen, wo in irgendeinem Feld nichts
'           eingetragen ist, weil im Zuge einer Stichwort-Pflege genau dieser Zustand �berpr�ft und
'           eventuell korrigiert werden soll.
'           Das geht jetzt durch Angabe von NULL.
'27.02.2004 Fehlerkorrektur: Bei 'Fenstergr��e �nderbar' wird das Bild oben links angeordnet. Bisher ist das
'           aber bei ZoomIn oder ZoomOut vers�umt worden.
'01.03.2004 Die �nderung vom 16.12.2003 wird zur�ckgezogen, weil jetzt nach Doppelklick auf ein anderes Bild,
'           dann F5, dann Benutzung der Funktion /WRITE nicht sofort dorthin geschrieben werden kann wo
'           der Cursor steht, sondern erst noch ein zus�tzlicher Klick notwendig ist
'08.03.2004 Wenn die Funktion /WRITE nicht aktiviert ist, kann man keinen Text aus dem F5-Fenster markieren
'           und in die Zwischenablage kopieren, weil jeder Tastendruck zum Verschwinden des Fensters f�hrt.
'           Es kann aber sehr w�nschenswert sein, einen Text zu kopieren und diesen dann als Suchbegriff zu
'           benutzen. Zu diesen Zweck gibt es einen neuen Button
'           'Kopiere den markierten Text in die Zwischenablage'
'09.03.2004 Ich will, da� beim Installieren ein tats�chlich funktionierendes Beispiel im Installationsordner
'           installiert wird. Es werden 3 Beispielsfotos mitgeliefert. Bei jedem Start von fotos.exe
'           wird die Spalte ErsterStart in der Tabelle Spaltenbreite abgefragt. Wenn dort ein H�kchen steht,
'           ist es der erste Start und die 3 Dateinamen in fotos.mdb werden an App.Path angepa�t. Nach dem
'           ersten Start wird das H�kchen aus der Spalte ErsterStart in der Tabelle Spaltenbreite entfernt.
'15.03.2004 Fehlerkorrektur:
'           Wenn mit 'weitere Filter...' gearbeitet wird und der Benutzer editiert die Schlagworte, dann wird
'           die letzte Zeile, mit der der Benutzer gearbeitet hat, nicht in die Datenbank zur�ckgeschrieben,
'           wenn F8 ged�ckt wird. Das wird jetzt korrigiert.
'26.03.2004 Fehlerkorrektur:
'           Wenn nicht mit /WRITE gearbeitet wird, kann man die Merkerspalte nicht editieren. Das wird
'           jetzt ge�ndert und geht immer, egal ob /WRITE oder nicht.
'16.04.2004 Fehlerkorrektur:
'           Bei vergr��ertem/verkleinertem Bild wird ein gew�nschter Kommentar nicht gezeigt. Das wird jetzt
'           korrigiert.
'16.04.2004 Fehlerkorrektur:
'           Bei Videos ist F9 unwirksam (Mauszeiger sichtbar/unsichtbar)
'22.05.2004 Wenn beim Suchbegriff, der in jedem Feld gesucht werden soll, hintendran mehrere Leerzeichen
'           stehen, wird im Extremfall garnichts gefunden. Ich mu� vor dem Suchen die �berfl�ssigen
'           Leerzeichen abschneiden.
'           Genauso bei allen anderen Suchfeldern.
'22.05.2004 Beim Exportieren die voraussichtliche Datenmenge in Bl�cke von 3 Bytes auftrennen dazwischen
'           sollen Punkte stehen.
'25.05.2004 Wenn in irgendeinem Feld ein Hochkomma gespeichert wird (zB d'Artagnan), dann st�rzt das
'           Programm beim Wiederfinden ab. Ein wohlbekanntes Problem, aber es hat noch niemand gemerkt.
'           Ich mu� als Suchstring mit doppeltem Hochkomma arbeiten. Dann sind einfache Hochkommas in den
'           Datenfeldern erlaubt, aber doppelte Hochkommas verboten. In fotosmdb.exe mu� es eine Suchfunktion
'           nach doppelten Hochkommas geben.
'08.06.2004 Export der Datenbanks�tze hat Fehler gemeldet
'           Feld Kommentar darf keinen String = "" enthalten
'           Darum setze ich das Feld auf Null, wenn "" gefunden wird
'09.06.2004 Verbesserung der Fehlerbeschreibung, wenn ein Video nicht gespielt werden kann
'09.06.2004 Wenn ein H�kchen gesetzt ist bei SQL nacharbeiten, aber zum zweitenmal auf den Button 'Fotos finden'
'           geklickt worden ist, dann mu� drankommen GoTo SQLWurdeBearbeitet
'09.06.2004 Wenn alle Felder auf 'Beliebig' gesetzt werden sollen, mu� auch das Feld SQLText auf den
'           Ausgangswert "Select * From Fotos ORDER BY Dateiname" gesetzt werden
'13.06.2004 Seltsamerweise stellt sich nach der Taste F5 die Liste der gefundenen Dateien immer auf den Anfang
'           anstelle auf den zuletzt aktuell gewesenen Satz. Als Gegenma�nahme habe ich den F5Timer erfunden
'           aber sp�ter wieder entfernt.
'21.06.2004 Ich brauche Funktionen um das Bild tempor�r zu sch�rfen oder unscharf zu machen.
'           Es ist nicht vorgesehen die �nderungen �ber diese Funktionen speicherbar zu machen.
'           Wenn der Nutzer die �nderungen speichern will, soll er �ber F5 gehen.
'           Ich habe umfangreichen Code zu diesem zweck aus imgproc2 kopiert
'           vbAccelerator Image Processing Sample (imgproc2.zip)
'           Copyright � 1998 Steve McMahon (steve@dogma.demon.co.uk)
'           Das sind 2 Formulare: frmScharfUnscharf, frmImage
'           1 Module: mHLSRGB
'           5 Klassenmodule: cDIBSection, cImageProcessDIB, cMRUFileList, cPalette, cRegistry
'22.06.2004 siehe 13.06.2004. Auch nach F5 und �ffnen der mit 'jpg' verkn�pften Anwendung und anschlie�end
'           Dr�cken von F2 oder F3 oder F6 bin ich am Anfang der gefundenen Dateien.
'05.07.2004 siehe 22.06.2004. Auch nach F5 und mit der Maus auf dem Fenster herumfahren geht es manchmal
'           an den Anfang der gefundenen Dateien
'06.07.2004 Suche nach IS NULL muss erweitert werden durch OR Feldxyz = ""
'04.08.2004 Bei Dr�cken von F12 kann eine neue Funktion eingestellt werden:
'           Alle Bilder mit tempor�r ver�nderter Helligkeit anzeigen.
'           Solange wie diese Funktion aktiv ist, wird mit einem auff�lligen Mauszeiger darauf hingewiesen.
'16.09.2004 Nach Doppelklick auf eine Zeile im F5-Fenster ist bisher ein eventuell ge�ffnetes Hilfefenster
'           ge�ffnet geblieben. Ab jetzt wird das Hilfefenster geschlossen.
'16.09.2004 DbGridForm.btnStichwort wird nur dann Visible=True, wenn in DbGridForm.Adodc1.RecordSource
'           "fotosmitzusatzspalte" vorkommt.
'16.09.2004 siehe 22.06.2004. Wenn nach F5 als erste Aktion der Scrollbalken bewegt wird und dann mit der Maus
'           an den Formularoberrand gefahren wird, geht es an den Anfang der gefundenen Dateien. Dagegen
'           hilft jetzt ein Picture1-Element das �ber das gesamte Formular gelegt wird.
'16.09.2004 DbGridForm wird jetzt von Anfang an an den linken oberen Rand platziert.
'16.09.2004 Die Tasten Bild auf und Bild ab sind ab sofort nur noch im F5-Fenster(DbGridForm) wirksam.
'16.09.2004 Das Formular Query ist ab jetzt von Anfang an sichtbar. Alle Buttons au�er 'Beenden'
'           werden disabled bis das Formular fertig geladen ist
'20.09.2004 Es gibt eine Fehlerkontrolle auf Differenzen in der Spalte 'Jahr' und dem Jahr in der Spalte
'           'Dateiname'. Die betreffenden Bilder werden gesucht und k�nnen begutachtet werden.
'10.10.2004 Das �ndern des Dateinamens mu� bei /WRITE verboten sein. MsgBox mit Hinweis auf RenamMdb bringen.
'10.10.2004 Verbesserung:
'           Die Arbeit mit der Tabelle 'FotosMitZusatzSpalte' dauert zu lange und au�erdem gehen oft dort
'           gemachte �nderungen verloren, bevor sie in die Tabelle Fotos kopiert werden k�nnen. Ich will
'           auf die Tabelle 'FotosMitZusatzSpalte' ganz verzichten und in die Tabelle Fotos zwei neue
'           Felder aufnehmen:
'           DateinameKurz (Namensanteil von Dateiname)
'           DDatum (Datei Erstellungs Datum)
'           1.Wenn fotos.exe startet muss eine Abfrage gemacht werden, ob alle Felder Dateiname und
'           DateinameKurz �bereinstimmen. Wenn nicht: Hinweis auf Ausf�hrung von Pr�fen1
'19.10.2004 Fehlerkorrekrur:
'           Beim Start des Programms von CD sind die Tasten F2 und F3 unwirksam bzw machen was sie wollen.
'           Ich habe die �nderungen 22.06.2004 und 13.06.2004 zur�ckgezogen (F5-Timer)
'19.10.2004 Verbesserung:
'           Es gibt eine CheckBox zur Kennzeichnung ob 'Weitere Filter...' aktiv sind
'           Ein-/Ausschalten geht nur wenn auf den Button 'Weitere Filter..' geklickt wird
'19.10.2004 Verbesserung:
'           Im F5-Fenster kann man jetzt durch Klick in die Spalten�berschrift die Sortierfolge dieser
'           Spalte �ndern, abwechselnd in aufsteigend/absteigend
'19.10.2004 Verbesserung:
'           Wenn ShellExecute ausgef�hrt werden soll und mit der Dateinamen-Erweiterung ist keine Anwendung
'           verkn�pft, dann kommt bisher nur ein schwarzer Bildschirm mit Sanduhr. Besser w�re ein Hinweis
'           dass der Nutzer die fehlende Verkn�pfung selbst herstellen soll.
'19.10.2004 Fehlerkorrekrur:
'           Da war ein Uralt-Fehler drin. Es gibt zwei Formulare, wo Datum-von und Datum-bis ausgew�hlt werden
'           kann. Wenn �berhaupt schon ein Datum ausgew�hlt war, dann sollte dieses Datum auch wieder beim
'           �ffnen des Kalenders kommen. Bei dieser Gelegenheit war von und bis verdreht worden.
'03.11.2004 Fehlerkorrektur:
'           Die Tasten 'Bild nach oben' und 'Bild nach unten' sollen in DbGridForm.DbGridNeu
'           nach oben oder nach unten bl�ttern, das war bisher nur bei '/WRITE' realisiert.
'03.11.2004 Die Taste F6 wird h�ufig aus Versehen gedr�ckt, weil eigentlich F5 gemeint ist.
'           Von jetzt ab muss Strg+F6 gedr�ckt werden
'04.11.2004 Zur�cknahme der �nderung vom 08.03.2004 ich brauche daf�r keinen Button, sondern erlaube bei
'           Tastenkombination Strg + C, dass DbGridform sichtbar bleibt.
'07.11.2004 Diverse Korrekturen:
'           Combobox SW/F soll nicht beschreibbar sein. Locked=True.
'           Nur bei SQL <> "" darf man GoTo SQLWurdeBearbeitet ausf�hren.
'           Leeres Kriterium Situation, Ort, Land, Personen muss verboten sein.
'           In $fotos.mdb die neuen Felder DateinameKurz und DDatum �bernehmen.
'           Bei Benutzung von $fotos.mdb darauf hinweisen, dass mehrere Funktionen nicht m�glich sind:
'           - keine Stichwortpflege
'           - kein Strg+E
'           - kein Strg+K
'12.11.2004 Bisher wurde nach Clicken des btnRefresh die SQL-Nachbearbeitung nicht zur�ckgesetzt.
'12.11.2004 Wenn nicht /WRITE vorlag, konnte man bisher nur ein Merker-H�kchen setzen
'           weitere waren nicht m�glich
'15.11.2004 In jedem Suchfeld au�er Jahr und SWF und Personen (weil es da bereits mehrere Personen gibt)
'           kann man ab jetzt mehrere Suchbegriffe eingeben, die
'           entweder mit OR (Trennzeichen %%%) oder mit AND (Trennzeichen &&&) verkn�pft werden.
'           Mischen von %%% und &&& in einem Suchfeld ist verboten.
'02.01.2005 Fehlerkorrektur und Vereinfachung:
'           Merkerspalte in jeden Datensatz ein/ausschalten wechselweise bei jedem Click
'20.01.2005 Anstelle der Taste F5 (Listenfenster �ffnen) soll Taste F5 + Taste Umsch
'           den Inhalt der aktuellen Zeile der Datenbank als Fenster mit mehreren Zeilen zeigen
'21.01.2005 Nur bei /Write tritt folgendes Problem auf:
'           Die Spalte Kommentar enth�lt Inhalte der Form
'           {\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}
'           Das ist unvermeidbar, wenn man �ber das Kommentar-Fenster den Inhalt der Spalte 'Kommentar'
'           editiert. Es ist vermeidbar, wenn man in der DbGridForm.DbGridNeu in der Kommentarspalte editiert.
'           Und es ist vermeidbar, wenn sich der Inhalt im Kommentar-Fenster nicht ge�ndert hat.
'           Wenn man diese Inhalte w�hrend der Bearbeitung mit /Write entfernen will, muss das Programm
'           die Eigenschaft AllowRowSizing = True zulassen, damit die Kommentarspalte mehrere Zeilen anzeigen
'           kann.
'23.01.2005 Verbesserung:
'           Der Hinweis 'Weitere Filter sind aktiv' bleibt unsichtbar, solange keine weiteren Filter aktiv sind.
'           Bisher war nur die Checkbox nicht gesetzt.
'24.01.2005 Es d�rfen jetzt 2 Instanzen derselben fotos.exe gestartet werden, wenn im Formular Query ein
'           H�kchen bei chkFensterGr��e�nderbar gesetzt ist
'30.01.2005 Fehlerkorrektur Laufzeitfehler im Formular MP:
'           wenn Weitere Filter... -> Ge�ndert am einbeziehen -> von ist sp�ter als bis
'09.02.2005 Bei Fehlerkontrolle auf Differenzen in Jahr und Dateiname wird die Anzahl der gefunden S�tze
'           weggelassen, weil sie stets einen falschen Wert von 34 oder 35 anzeigt
'13.02.2005 Wenn in Dbgridform.DbgridNeu eine Spalte so zusammengeschoben ist, dass sie unsichtbar ist,
'           dann ist ihr Wert Col.Width nicht etwa Null sondern der alte Wert. Beim Speichern der Spaltenbreiten
'           muss man also nach Col.Visible fragen.
'14.02.2005 Ab Version 10:
'           Das Fotoalbum kann mit nutzerdefinierten Feldern arbeiten.
'           Fotosmdb.exe kann jeweils ein Feld vom Typ Text mit max L�nge 255 erzeugen, Felder mit anderen
'           Datentypen m�ssen mit MS Access erzeugt werden.
'           Im Formular DbGridForm kann man nach den nutzerdefinierten Feldern sortieren und die
'           nutzerdefinierten Felder editieren.
'           Man kann maximal 5 nutzerdefinierte Feldnamen und Feldwerte in die Suche einbeziehen dazu dient das
'           neue Formular ND(NutzerdefinierteFelder) jedoch d�rfen mehr als 5 nutzerdefinierte Felder
'           angelegt werden.
'           F�r das Speichern der Spaltenbreiten m�ssen die Tabellen Spaltenbreite und ErsterStart
'           neu entworfen werden.
'           Bei 'Suche in jedem Feld' wird in den nutzerdefinierten Feldern nicht gesucht, weil die Suche
'           mit LIKE *Begriff* gemacht wird und bei nutzerdefinierten Feldern vom Typ Zahl oder W�hrung
'           dadurch kein Suchergebnis kommt.
'22.02.2005 Fehlerkorrekturen:
'           Nach F10 (Kommentarfenster einblenden) und wenn das Kommentarfenster aktiv ist, wirkten bisher keine
'           Funktionstasten.
'           Bei Umsch + F5-Taste konnte man keinen Text zum Kopieren markieren
'           Bei DbGridNeu_HeadClick muss man den Feldnamen in [] einschlie�en f�r den Fall, dass der
'           Feldname Sonderzeichen enth�lt.
'           Verbesserung:
'           Wenn bei 'Suche Begriff in jedem Feld' ein Datum angegeben wird, soll eine Warnung kommen
'07.03.2005 Verbesserung:
'           Ich will bei 'Nur den ersten Treffer pro Jahr erlauben' ohne die Datei $fotos.mdb auskommen.
'           Daf�r erfinde ich die tempor�re Tabelle Fotos_ErsterTreffer.
'           Jetzt kann man auch mit Merkerspalte und Export f�r 'Nur den ersten Treffer pro Jahr erlauben'
'           arbeiten. Nach wie vor ist kein �ndern der Stichworte nach Taste F5 m�glich.
'           F�r das Exportieren ist nach wie vor die Datenbank $fotos.mdb notwendig.
'           Bei Beenden des Programms wird die Datenbank fotos.mdb komprimiert.
'11.03.2005 Verbesserung:
'           Es gibt 2 neue Standardfelder in der Datenbank fotos.mdb. Das sind BreitePixel und HoehePixel
'           Bei Pr�fen1 und bei 'Neue Datens�tze generieren (durch Drag&Drop vom Windows Explorer)...'
'           in fotosmdb.exe werden diese Felder gef�llt.
'11.03.2005 Fehlerkorrektur:
'           Wenn man ein Video ausgef�hrt hat, ging anschlie�end Strg + B nicht mehr.
'11.03.2005 Fehlerkorrektur:
'           Wenn mit Merkerspalte gearbeitet wurde, ging sortieren nach Spalte 'Dateiname' falsch
'           wegen Inner Join mit Temp_Haken und unklare Anweisung ORDER BY Dateiname (welcher Dateiname)
'13.03.2005 Verbesserung:
'           Nach Wechsel der Sortierfolge einer Spalte den Horizontalen Scrollbalken wieder so einstellen
'           wie vor dem Sortieren
'24.03.2005 Verbesserung:
'           An der linken Kante eines Bildes, das den Bildschirm voll ausf�llt, zappelte bisher ein schwarzer
'           Cursor. Das ist die linke Kante der Bedienelementzeile von Mediaplayer1. L�sung: ich schiebe
'           Mediaplayer1 einfach 30 Twips nach links au�erhalb der Form.
'24.03.2005 Verbesserung:
'           Bei Videos blieb eine �nderung der Videogr��e von zB 100% auf 200% nicht erhalten sondern wurde
'           beim n�chsten Video stets wieder zur�ckgesetzt.
'           L�sung: Displaysize = MediaPlayer1.Displaysize und umgekehrt
'25.03.2005 Erg�nzung zu 14.12.2003
'           Ein Bild hat dann 4 Pixel mehr Bildbreite bzw Bildh�he, wenn die Eigenschaft Borderstyle=1
'           gesetzt ist. Also Borderstyle=0 benutzen.
'           Nicht zu fassen, dass mir das erst jetzt auff�llt.
'26.03.2005 Verbesserung:
'           Bei den Tasten Umsch + F5 kam bisher unter Kommentar: RichTextBox1, wenn es keinen Kommentar gab
'           In Zukunft wird das Feld auf "" gel�scht
'28.03.2005 Bei 'Suche in jedem Feld' wird ab jetzt eine Combobox TBegriff benutzt worin alle bisher benutzten
'           Stichworte alphabetisch geordnet gespeichert werden
'29.03.2005 Verbesserung:
'           Beim Button Reset wird alles auf den Anfangswert zur�ckgesetzt wie nach Programmstart.
'           Bei 'erster Treffer pro Jahr' wird nach Jahr,Dateiname sortiert
'31.03.2005 Nachbesserung zum 28.03.2005
'           Combobox TBegriff darf bei Programmstart nicht sichtbar sein
'           Neue Begriffe nur aufnehmen, wenn sie nicht bereits aufgenommen sind
'11.04.2005 Man k�nnte das Programm fotos.exe so ver�ndern, dass es selbst versucht die Bezeichnung des
'           Fotos-Root-Ordner zu ermitteln, n�mlich als App.Path wo fotos.exe steht.
'           Dazu m��te man in die Datenbank eintragen
'           anstelle von zB M:\P7FotoSoundVideo\FOTOS\GG\2005\Ballonfahrt001.jpg
'           +:\2005\Ballonfahrt001.jpg und bei der Ausf�hrung von fotos.exe fotosmdb.exe und renammdb.exe muss
'           +:\ ersetzt werden durch App.Path des entsprechenden Programms.
'
'           Dann entf�llt die Funktion Ersetzen im Programm fotosmdb.exe,
'           und alles was zusammenh�ngt mit 'Fotos-Root-Ordner Festlegen' bei Start von einer CD,
'           und das Feld ErsterStart.ErsterStart wird nicht mehr ausgewertet.
'           aber man muss vom Nutzer verlangen, dass er s�mtliche Dateien unterhalb von App.Path von fotos.exe
'           anlegt. Daf�r kann er die 3-Einigkeit von fotos.exe, fotos.mdb und Dateien kopieren oder
'           verschieben wohin er will.
'
'           Pr�fen der 3-Einigkeit ist in jedem Programm fotos.exe fotosmdb.exe renammdb.exe n�tig.
'           Man muss dazu pr�fen, ob der erste Satz der Tabelle Fotos, nach Ersetzen des String +:\ durch
'           App.Path eine Datei ergibt, die existiert.
'16.04.2005 App.Path generell durch AppPath ersetzen und wenn App.Path mit "\" endet, den "\" wegschneiden
'           Ursache: im Windows XP passiert nichts, wenn ein Pfad zB I:\\fotos.mdb entsteht, aber Windows 98
'           kann mit 2 aufeinanderfolgenden "\\"  nicht umgehen.
'17.04.2005 Bisher war bei Dr�cken der Taste F8 ein eventuell vorher sichbares Formular XYPos weiterhin
'           sichtbar. Das wird ab jetzt verhindert.
'17.04.2005 Ver�nderter Text, wenn ein Bild nicht geladen werden kann
'20.04.2005 Ver�nderter Text, wenn fotos.mdb schreibgesch�tzt ist, es kommt eine vierte M�glichkeit dazu
'           Sie arbeiten mit Daten auf einem Netzwerk-PC und haben keine Schreibrechte
'22.04.2005 Verbesserung:
'           Beschleunigung beim Start beim F�llen der Comboboxen. Die Comboboxen vom Formular MP werden erst
'           aufgef�llt, wenn das Formular MP ge�ffnet wird.
'06.05.2005 Fehlerkorrektur:
'           Nach F5 Button -> '�ffnen der mit jpg verkn�pften Anwendung' hat nicht funktioniert
'*************************************************************************************************************
'Beginn der Arbeiten an Shareware oder Professional Version
'*************************************************************************************************************
'23.05.2005 �nderungen zum Anbieten einer Shareware-Version
'           Die Vollversion gibts nach Anforderung per E-mail an information@gerbingsoft.de
'           Ich erkenne die Shareware-Version am
'           Vorhandensein der Datei msdmo.log in ...\windows\Systemdirectory und am Fehlen der Datei
'           msplugin.log in ...\windows\Systemdirectory. Die datei msdmo.log wird nicht bei der
'           Installation erzeugt, sondern beim ersten Start, wenn keine Freischaltedatei msplugin.log gefunden
'           wird und keine Datei msdmo.log. Die Datei msdmo.log bekommt als Datum das Datum von heute - 100,
'           damit sie nicht so leicht zu entdecken ist. Die Shareware-Version bringt das
'           Shareware-Hinweis-Fenster immer h�ufiger, je �lter die Datei msdmo.log ist.
'           Wenn das Programm mit Freischalteschl�ssel installiert wird, wird w�hrend der Installation die
'           G�ltigkeit des Freischalteschl�ssels gepr�ft (das ist die Aufgabe von registrieren.exe) und bei
'           G�ltigkeit eine Freischaltedatei msplugin.log im Ordner ...\windows\Systemdirectory erzeugt. Die
'           Freischaltedatei hat das Datum von heute - 100, damit sie nicht so leicht zu entdecken ist. in der
'           Freischaltedatei steht codiert der Freischalteschl�ssel. Der Freischalteschl�ssel muss
'           personengebunden vergeben werden.
'           Ich f�hre dazu eine Datenbank, da k�nnte ich bei unberechtigter Weitergabe des
'           Freischalteschl�ssels erkennen, wer diesen weitergegeben hat. Der Freischalteschl�ssel darf nicht
'           durch Probieren erzeugt werden k�nnen.
'           Beispiel: FX58A-C3BYE-1FGH3-B3YFG-FX2BA-GGERBING
'           Die Zahlen in den ersten 4 Kolonnen werden summiert 5+8+3+1+3+3=23
'           Die Summe wird durch 7 geteilt und ergibt Rest 2. In der letzten Kolonne muss der Rest stehen
'           und richtig sein.
'           "Fehler - 2305" kommt, wenn Open SystemDirectory & "\msdmo.log" For Output As #Dateinummer mi�lingt
'           "Fehler - 2205" kommt, wenn Open SystemDirectory & "\msplugin.log" For Input As FNum mi�lingt
'
'           F�r Freunde und Bekannte gibt es auch eine Vollversion, f�r die braucht man keinen
'           Freischalteschl�ssel. Da wird bei der Installation msplugin.log in den
'           Ordner ...\windows\Systemdirectory gestellt
'06.06.2005 Fehlerkorrektur zu Demo-Version
'           Nach Ablauf der G�ltigkeit hat das Klicken im Listenfenster auf eine Spalten�berschrift
'           zu Laufzeitfehler 5 gef�hrt, weil kein Anteil ORDER BY im SQL String enthalten war.
'06.06.2005 es gibt irgendein Problem, wenn ich in der DbGridForm einen Doppelklick auf eine Spalten�berschrift
'           ausf�hre. Das Image1 verrutscht au�erhalb der Zentrierung. Es sieht so aus, als w�rde der Doppelklick
'           gleichzeitig als Verschiebeklick aufgefasst.
'           Ich erfinde den Schalter gblnDbGridFormDblClick
'10.06.2005 �nderungen zum Anbieten einer Shareware-Version oder einer Professional-Version
'           Die Professional-Version gibts nach Anforderung per E-mail an information@gerbingsoft.de
'           Ich erkenne die Shareware-Version am Fehlen der Datei msplugin.log in ...\windows\Systemdirectory
'           und am Fehlen der Datei msprivs.log in ...\windows\Systemdirectory.
'           Ich erkenne die Professional-Version am Vorhandensein der Datei msplugin.log
'           in ...\windows\Systemdirectory
'           und am Vorhandensein der Datei msprivs.log in ...\windows\Systemdirectory
'           Die datei msprivs.log wird durch RegProfi.exe mit einem g�ltigen Professional Schl�ssel erzeugt.
'           Shareware-Version:
'           -ohne benutzerdefinierte Felder
'           -ohne Trefferauswahl alle oder 'erster Treffer'
'           -ohne Eingeben von Suchkriterien Profi
'10.06.2005 Verbesserung bei der Suche nach nutzerdefinierten Feldern
'           Combobox mit allen Feldinhalten anbieten
'           Man muss irrt�mlich gew�hlte Feldnamen auch wieder l�schen k�nnen. Dazu die Tasten
'           Return oder Entf auf dem numerischen Tastenfeld benutzen
'16.06.2005 Neue Funktion (nur in der Professional Version)
'           Arbeit mit gespeicherten Abfragen
'16.06.2005 Fehlerkorrektur:
'           Die Anzahl der gefundenen Fotos im Listenfenster (F5) ist manchmal anfangs falsch
'           wird aber nach dem ersten Doppelklick auf irgendein Foto richtig, wenn nicht auf das erste
'           geklickt wird. Ursache: es hat ein recordset.Movelast und dann wieder recordset.Movefirst gefehlt
'16.06.2005 �berarbeitung von DbGridForm.DbGridNeu_HeadClick
'19.07.2005 Fehlerkorrektur:
'           Im Zusammenhang mit tempor�r scharf/unscharf und wenn das diesbez�gliche Fenster nicht geschlossen,
'           sondern minimiert wird, dann konnte man bisher die Query-Form beim Beenden nicht schlie�en.
'19.07.2005 Verbesserung zum Kommentar-Fenster siehe 21.01.2005:
'           Wenn in der Spalte Kommentar Inhalte der folgenden Form auftauchen
'           {\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}
'           das passiert, wenn ich Editieren im Kommentar-Fenster zulasse
'           (gegenw�rtig ist Editieren im Kommentar-Fenster nicht m�glich),
'           dann k�nnte man die Zeilendarstellung mehrzeilig machen, wie das auch in der MS Access Umgebung
'           funktioniert.
'           Gegenw�rtig wird jeder Datensatz durch eine einzelne Zeile dargestellt.
'           oder ich lasse den Nutzer im kommentar-Fenster ausw�hlen, ob er Kommentar editieren will
'           und weise darauf hin, dass im Listenfenster nur die erste Zeile kommt
'--------------------------------------------------------------------------------------------------------------
'19.07.2005 Fehlerkorrektur:
'           bezieht sich auf Versionen < 12.0.50
'           Ich habe im Ordner Reglight Alt eine reglight.exe ausgeliefert, wo der ausgelieferte
'           Registrierschl�ssel durch einen Buchstaben ver�ndert einen g�ltigen Freischaltecode ergibt, mit dem
'           eine Datei msprivs.log erzeugt wird.
'           Jeder, der mit dieser reglight.exe arbeitet, kann mich ganz einfach austrixen und eine
'           Professional Version erzeugen.
'           Ich muss Schl�sselGenerieren.exe und reglight.exe �berarbeiten und gleich beim Start von fotos.exe
'           den Inhalt von msprivs.log kontrollieren und msprivs.log l�schen, wenn ein falscher
'           Freischalteschl�ssel drinsteht.
'           "123" wird zugelassen -> Voll-Version
'           g�ltiger Freischalteschl�ssel -> Professional-Version
'           "Fehler - 2007" kommt, wenn Open SystemDirectory & "\msprivs.log" For Input As FNum mi�lingt
'20.07.2005 Verbesserung:
'           Bei Videos hat bisher die Taste F10(Kommentarfenster einblenden) garnichts bewirkt.
'           Das war schlicht und einfach vergessen.
'21.07.2005 Fehlerkorrektur:
'           Bisher bei Videos und /WRITE und sichtbarer DbGridForm f�hrte F2 oder F3 d�cken zwar zum Start des
'           vorigen/n�chsten Videos aber DbGridForm blieb oben.
'22.07.2005 Bei F10 soll das Kommentarfenster sofort aufgehen. Bisher geht es erst beim n�chsten Foto auf.
'25.07.2005 Ab jetzt ist nicht nur das �ndern von 'Dateiname DateinameKurz DDatum' verboten, sondern auch
'           'BreitePixel HoehePixel'
'02.09.2005 Ab jetzt wird bei Taste Umsch+F5 der Dateiname so angezeigt, wie er in der Datenbank steht
'02.09.2005 Verbesserung:
'           Die Form 'Query' soll eine Men�-Leiste bekommen
'           Tools Reset Hilfe Beenden
'           daf�r entfallen ganz unten die Buttons 'Hilfe' und 'Beenden'
'13.10.2005 Anstelle des Begriffs Demo-Version benutze ich jetzt Shareware-Version
'13.10.2005 f�r die Professional-Version gibt es eine sechste Kolonne mit dem Familienname zB
'           0AV7Q-8B34W-28L05-N007V-2JGRC-MEIER
'           die Shareware-Version ben�tigt keinen Lizenzcode, aber
'           nach 90 Tagen - F8 verursacht das Auftauchen der Shareware-Form
'           nach 180 Tagen - F5 verursacht das Auftauchen der Shareware-Form
'           nach 365 Tagen - F2, F3 und deren �quivalente verursachen das Auftauchen der Shareware-Form
'           nach 545 Tagen darf die Tabelle Fotos maximal 5000 S�tze enthalten, sonst Beendigung des Programms
'08.11.2005 Verbesserung:
'           Gegen Cracker der Sharewareversion
'           Das Wort Shareware wird mit Chr(xy).... gebildet
'           die Form Shareware hei�t jetzt Copy
'           das Wort msplugin.log wird mit Chr(xy).... gebildet
'           das Wort msprivs.log wird mit Chr(xy).... gebildet
'           das Wort msdmo.log wird mit Chr(xy).... gebildet
'08.11.2005 Beginn der Arbeiten f�r Englische Version und beliebige Sprachen
'           durch Benutzung einer .res-Datei und LoadResString
'           Die .res-Datei der Entwicklungsumgebung wird bei der Exe
'           zum festen Bestandteil der Exe. Darum m��te man bei Benutzung von verschiedenen .res-Dateien
'           pro Sprache, jedesmal eine neue Exe erzeugen, mit der jeweiligen .res-Datei.
'           Es gibt einen Ausweg, wie man alle Sprachen in eine .res-Datei legen kann.
'           alle Zahlen ab 1000 gelten f�r deutsche Label.Caption und Form.Caption
'           alle Zahlen ab 2000 gelten f�r deutsche Msgbox-Texte und Tooltips
'           alle Zahlen ab 3000 gelten f�r deutsche Commandbuttons, Checkboxen, Options, Frames
'           alle Zahlen ab 4000 entsprechen wie ab 1000 aber englische Label.Caption und Form.Caption
'           alle Zahlen ab 5000 entsprechen wie ab 2000 aber englische Msgbox-Texte und Tooltips
'           alle Zahlen ab 6000 entsprechen wie ab 3000 aber englische Commandbuttons, Checkboxen, Options, Frames
'           LoadResString(1090 + Sprache) benutzen
'           Sprache = 0 bei deutsch
'           Sprache = 3000 bei englisch
'           Die gew�hlte Sprache wird in die Datei fotos.ini eingetragen
'           Ich muss eine deutsche fotos.mdb ausliefern und alle Abfragen entfernen, weil die bisherigen
'           Anwender nur eine deutsche Datenbank mit deutschen Feldnamen haben.
'           Es braucht nicht 2 MSI-Files zu geben, wenn ich beim Wechsel der Sprache eine
'           Tabellenerstellungsabfrage mache, wo entweder aus Tabelle Fotos Tabelle EFotos entsteht oder
'           aus EFotos entsteht Fotos. Nach Sprachwechsel ist Neustart n�tig.
'           Das MSI-File enth�lt sowohl die deutschen wie die englischen Help-Dateien, beim Installieren
'           werden die Help-Files in zwei verschiedene Ordner abgelegt.
'           Ich brauche dann lediglich 2 Web-Auftritte deutsch und englisch.
'
'06.12.2005 Ich habe gek�mpft wie ein Ochse, damit die Query Form nicht von der Taskleiste verschwindet.
'           Immer bei Query.Hide verschwindet sie und kommt durch Query.Show 1 nicht zur�ck.
'           Ohne Erfolg.
'           Bei 'Fenstergr��e �nderbar' sollen alle xxxForm.Width so gro� sein, wie Form1.Width
'10.12.2005 Verbesserung:
'           Drag&Drop f�r Export/Import
'11.12.2005 Verbesserung:
'           Ich werde ab jetzt unterscheiden zwischen Bildern, die ich im native mode anzeigen kann
'           "BMP", "CUR", "DIB", "EMF", "GIF", "ICO", "JPG", "WMF"
'           und Bildern, die ich nur im link mode (Link-Dateitypen) anzeigen kann.
'           F�r den link mode benutze ich ShellExecute, so wie schon bisher bei Dateityp "HTM", "PDF", "XLS"
'           im link mode kann man dann beispielsweise die Dateitypen
'           "PNG" "PSD" TIF" betrachten.
'           F�r "PNG" und "TIF" kann man zB die Windows Bild- und Faxanzeige benutzen,
'           da �ffnet sich f�r jedes neue Bild immer dasselbe Fenster.
'           F�r "PSD" kann man zB Irfan View benutzen,
'           da �ffnet sich f�r jedes Bild ein neues Fenster.
'14.12.2005 Ich benutze ab jetzt die bedingte Compilierung
'           man muss eintragen unter Projekt -> Eigenschaften -> Registerkarte 'Erstellen'
'           Argumente f�r bedingte Kompilierung
'           Proversion = -1      erzeugt die Professional Version
'           Proversion = 0       erzeugt die Shareware Version
'24.12.2005 Probleme mit SpeichernSpaltenBreite und SetSpaltenBreite wurden beseitigt
'29.12.2005 Ich will f�r alle link mode formats die Automatic ausser Kraft setzen
'           d.h. wenn ich mit MoveNext auf eine Datei mit Link-Filetyp stosse,
'           dann ignoriere ich diese Datei und gehe zur n�chsten
'29.12.2005 Ich verzichte auf die Kontrolle, ob das Programm schon einmal gestartet wurde
'29.12.2005 Fehlerkorrektur:
'           Formular ND Nutzerdefinierte Felder
'           Wenn man zum zweitenmal auf den OK-Button klickte, wurde ein manuell eingegebener oder per Maus
'           ausgesuchter Wert stets durch den ersten Wert der Combobox �berschrieben
'30.12.2005 Ich kann nicht finden warum manchmal scheinbar die Taste F11 (Kommentarfenster soll verschwinden)
'           wirkungslos ist. Offenbar enth�lt 'AppActivate FotoAlbumTitle' etwas falsches.
'           Es passiert nie bei H�kchen in 'Fenstergr��e �nderbar'
'           und es passiert immer erst nach F8
'02.01.2006 nicht korrigierbar, deshalb nicht ausf�hrbar
'           Bei DbGridForm.DbGridNeu_HeadClick auf die Spalte Kommentar verschwindet der komplette Inhalt des
'           DBGrid, wenn im Feld Kommentar komplexer Inhalt steht. ZB Kopie einer Excel-Tabelle oder
'           sehr viel formatierter Text. Bei erneutem Click kommt der Inhalt wieder, manchmal auch erst nach
'           HeadClick auf eine andere Spalte.
'           Deshalb wird Sortieren der Kommentar-Spalte abgewiesen
'04.01.2006 Hurra mit dem �bergang auf ADO und DataGrid geht das Sortieren der Kommentar-Spalte
'04.01.2006 Problem mit der Umstellung auf englisch
'           Nach dem Neuerstellen der Tabellen nach Sprachwechsel mittels Tabellenerstellungsabfrage, haben
'           alle Tabellenfelder vom Typ Text die Eigenschaft 'Leere Zeichenfolge nein'.
'           vorher hatte ich manuell eingestellt 'Leere Zeichenfolge ja', weil ich sonst an einer Macke von
'           DBGrid scheitere. Man kann den Inhalt einer DBGrid-Spalte, die ein Textfeld darstellt nicht
'           von Inhalt ja auf Inhalt nein ver�ndern, wenn Feldeigenschaft='Leere Zeichenfolge nein'.
'           Ich muss die Verwendung der Controls Data1 und DBGrid1 ver�ndern.
'           Data1 wird Adodc1 (Microsoft ADO Data Control 6.0 (SP6) (OLEDB)=msadodc.ocx
'           ADO DataGrid1 (Microsoft DataGrid Control 6.0 (SP6) (OLEDB)=MSDatGrd.ocx wird unter dem bisher
'           benutzten Name DBGrid1 weiterbenutzt.
'           F�r die Arbeit mit der Merkerspalte m�sste ich aber weiterhin Data1 und DBGrid1 benutzen
'           was ich nicht will und jede Prozedur wie zB GeheEinBildVorw�rts m�sste ich doppelt f�hren
'           Ich �bernehme darum die Merkerspalte fest an den Anfang der Tabelle Fotos
'           da gibt es zwar Probleme bei gleichzeitiger Benutzung der Merkerspalte durch mehrere Nutzer
'           das muss man dann zur Administratorarbeit machen, wo vorher alle anderen Nutzer sich abzumelden haben
'           bei �bergang auf Version 13.0.0.0 wird das vom Programm selbst ausgef�hrt
'           Wozu dient die Tabelle Temp_Haken?
'           Sie wird nur gebraucht f�r den �bergang von Version 12.50.0 auf Version 13.0.0 um in der Tabelle
'           Fotos als erstes Feld das Feld Merker zu erzeugen.
'           -------------------------------------
'           Pfeil-Tasten-Navigation bei ADO DataGrid:
'           Wenn man im DataGrid einen Wert einer Zelle ver�ndert hat, muss man zum Beenden die Enter-Taste
'           dr�cken, bevor man mit den Pfeil-Tasten in ein anderes Feld wechseln kann, oder man muss mit der
'           Maus ein anderes Feld anklicken. Solange man die Enter-Taste nicht gedr�ckt hat, kann man eine
'           begonnene �nderung mit der Esc-Taste r�ckg�ngig machen.
'           -------------------------------------
'           Formulierung der SQL-Anweisung 'LIKE':
'           bisher wurde formuliert zB ... situation Like '*Muster*'  ...
'           mit ADO wird formuliert zB ... situation Like '%Muster%'  ...
'10.01.2006 Das Fenster F5MehrereZeilen braucht ab jetzt nicht mit Schlie�kreuz geschlossen zu werden, sondern
'           das passiert auch bei jeder Funktionstaste
'13.01.2006 Fehlerkorrektur:
'           Nach der Taste F8 schien manchmal die Tasteneingabe zu blockieren, dann kam man nur weiter
'           wenn man die rechten Maustaste und das Formular Hilfebx benutzt hat.
'           Ab jetzt blende ich nach F8 generell das Formular Hilfebx ein.
'           Msgbox hat zwar auch geholfen, aber das ist eine schreckliche Kr�cke.
'26.01.2006 Verbesserung:
'           zu Export/Import per Drag&Drop mit schreibgesch�tzter Datenbank bzw bei Abbruch
'           und zu Export ohne Drag&Drop
'           Bei leerer Datenbank ist die einzige erlaubte Operation Import mit Drag&Drop
'27.01.2006 Verbesserung:
'           Beim Export in ein Zielverzeichnis wo noch keine Fotos/Videos stehen
'           wird ab jetzt nur noch jedes Unterverzeichnis angelegt
'           das unterhalb von AppPath liegt. Damit l��t sich 3-Einigkeit in einem beliebigen anderen Ordner
'           erzeugen. Zur guten letzt braucht man nur noch $fotos.mdb vom Quellordner in
'           fotos.mdb im Zielordner zu verwandeln.
'31.01.2006 Verbesserung:
'           Der herk�mmliche Export in ein Zielverzeichnis ist so verbessert worden, dass sich damit ein
'           neuer AppPath-Ordner erzeugen l�sst
'           darum wird der herk�mmliche Import ganz gestrichen
'05.02.2006 Fehlerkorrektur:
'           Bei H�kchen setzen und wieder entfernen bei
'           'Fehlerkontrolle auf Differenzen in Jahr und Dateiname' und bei
'           'gespeicherte Abfragen benutzen'
'           wurden bisher manchmal die H�kchen falsch sichtbar oder unsichtbar gemacht bei
'           'weitere Filter sind aktiv' und bei
'           'Suche nach nutzerdefinierten Feldern ist aktiv'
'22.02.2006 es entf�llt ab sofort die G�ltigkeit von "123" in msprivs.log
'           Die Vollversion f�r Freunde bekommt eine Datei msprivs.log mit g�ltigem Schl�ssel
'           Die Professional Version wird mit einer leeren Datei msprivs.log installiert, wenn regprofi.exe
'           nicht ausgef�hrt wird, bleibt es eine leere Datei.
'           Die Shareware-Version hat als Argument f�r bedingte Kompilierung 'Proversion = 0'
'           Die Professional/Voll-Version hat als Argument f�r bedingte Kompilierung 'Proversion = -1'
'           dudurch wird bestimmter Code in der exe nur bei der Vollversion compiliert
'---------------------------------------------------------------------------------------------------------
'nach Version 13.0.1 aufgetreten
'nicht korrigiert in der Sharewareversion, aber korrigiert in der Professional Version
'                                          und Vollversion f�r Freunde
'
'21.03.2006 Mehrere Fehler:
'           Das Editieren eines Feldes geht nicht, wenn man nicht in eine andere Zeile wechselt, der
'           Wechsel in ein anderes Feld der gleichen Zeile gen�gt nicht.
'           Nach F8-Taste kommt dumme MsgBox 'Sie d�rfen in die Merkerspalte nur 0 oder 1 eintragen,
'           wenn man den Cursor vorher in eine leere Spalte gestellt hatte.
'           Mit krummen Mitteln (unvermutete Tasten-Benutzung) kann man das Feld Jahr auf leer l�schen,
'           solche Datens�tze werden dann von fotos.exe nicht mehr angezeigt,
'           aber von fotosmdb.exe und Renammdb.exe.
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.3
'12.04.2006 Neue Funktion (nur Professional Version):
'           Gesprochener Kommentar (Audio-Datei)
'           0.Es gibt eine Checkbox wo definiert wird, ob mit Audio-Kommentaren gearbeitet wird oder nicht
'           Das ist sinnvoll f�r das Aussehen des Kommentar-Fensters, weil am oberen Rand allerhand Platz
'           gebraucht wird, wenn mit Audio-Kommentaren gearbeitet wird
'           1.Die Audio-Datei tr�gt denselben Dateiname wie ein zugeh�riges Foto. Die Dateinamen-Erweiterung
'           ist entweder wav oder mp3. Der Name der Audio-Datei steht nicht in der Datenbank.
'           2.Es gibt ein neues Feld in der Datenbank AudioFileExists Datentyp ja/nein
'           Bei Nichtvorhandensein ist dieses Feld per Programm zu erzeugen und zwar hintenran egal ob es
'           nutzerdefinierte Felder gibt oder nicht.
'           Dieses Feld muss als Suchkriterium gesucht werden k�nnen.
'           3.Im Kommentar-Fenster wird die M�glichkeit zum Aufnehmen und Abspielen angeordnet.
'           3.1.Abspielen geschieht mit msdxm.ocx entweder Fotoname.wav oder Fotoname.mp3
'           3.2.Aufzeichnen geschieht mit wav siehe WaveRecorder
'           Es kommt keine Warnung wenn kein Mikrophon-Pegel gefunden wird, aber es wird ein Mikrophon-Pegel
'           benutzt.
'           3.3.Umwandeln in mp3 geht mit ACM, falls auf dem PC ein mp3-Encoder vorhanden ist, sonst bleibt
'           es wav.
'           3.4.Das Kommentar-Fenster muss bei F10 immer aufgehen, egal ob das Feld 'Kommentar' leer ist oder
'           nicht, aber bei einem Video bleibt FrameAudio unsichtbar.
'           3.5.Button 'Audio-Kommentar l�schen'
'           4.Neue Funktion 'Pr�fenS' Datenbankfeld AudioFileExists bereinigen. Priorit�t hat eine
'           vorhandene/nichtvorhandene Audio-Datei. Dadurch kann man per Windows Explorer ungew�nschte
'           Audio-Dateien einfach entfernen und danach die Datenbank korrigieren.
'           Ab sofort kann ich aber nicht mehr das Programm PlayMP3 ausliefern, das w�rde Kommentare abspielen.
'16.04.2006 siehe 30.12.2005
'           Wenn ich nach 'KommentarForm.Show' einf�ge 'KommentarForm.SetFocus'
'           dann verschwindet die Blockierung bei sichtbarer Kommentarform, (leider auch nicht erfolgreich)
'           Aber manchmal reagiert die Leertaste als ob ich ins Fenstertitelsymbol links geklickt h�tte.
'02.05.2006 Fehlerkorrektur:
'           Die Merkerspalte l��t sich auf leer l�schen.
'           Als Folgefehler werden �nderungen in anderen Spalten wieder auf den Wert vor der �nderung
'           zur�ckgesetzt.
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.4
'08.05.2006 Fehlerkorrektur nur in Professional Version aufgetreten:
'           Die Funktion 'erster Treffer pro Jahr' meldet immer Errorcode=3021
'           Errortext=..kein aktueller Datensatz
'           Ursache: bei einer Tabellenerstellungsabfrage wird das % in LIKE '%Mustermann%'
'           nicht verkraftet, ich muss % ersetzen durch *
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.5
'13.05.2006 Fehlerkorrektur in allen Versionen:
'           nach F10 und H�kchen bei 'Kommentar soll editiert werden' kann man keine Leerzeichen einf�gen
'           Es gab eine Sharewareversion im Internet vom 29.04.2006 bis 13.05.2006 wo dieser Fehler drin war
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.6
'14.05.2006 Fehlerkorrektur in allen Versionen:
'           Wenn man die Spalte 'Jahr' editiert und klickt in das Feld Jahr der darunterliegenden Zeile,
'           dann wird das Feld 'Jahr' der darunterliegenden Zeile zum aktuellen Wert gemacht.
'           Das passiert nicht in der Entwicklungsumgebung, sondern nur in der exe
'           Ausweg: man muss dem Nutzer mitteilen, das er das Nachbarfeld der Zeile klicken muss in der das
'           Jahr ge�ndert wurde.
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.7
'15.05.2006 Fehlerkorrektur in allen Versionen:
'           Vermeiden von Laufzeitfehlern bei schreibgesch�tzter Datenbank oder bei schreibgesch�tztem Ordner
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.8
'30.05.2006 Fehlerkorrektur in allen Versionen:
'           Beim Dr�cken von F10. Ein Bild ohne Kommentar bekommt als Kommentar einen falschen Inhalt, den des
'           zuletzt aktuellen Bildes gezeigt.
'30.05.2006 Fehlerkorrektur in Professional Version:
'           Gespeicherte Abfragen d�rfen nicht formuliert sein "Like '*elke*'... sondern  m�ssen formuliert
'           werden durch Replace "Like '%elke%'...
'20.06.2006 Fehlerkorrektur in Professional Version: weitere Korrekturen zum 30.05.2006
'           Wenn man keine gespeicherte Abfrage ausgew�hlt hat kam Laufzeitfehler 5
'           Bei gespeicherten Abfragen ohne ORDER BY kam Laufzeitfehler 5 beim Spalten sortieren
'           Bei gespeicherten Abfragen ist Having Count(*) verwandelt worden in Having Count(%)
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.9
'24.04.2006 in Professional Version
'           Ich habe Versuche gemacht mit 50.000 Datens�tzen zu Geschwindigkeit und ob noch alles geht
'           Geschwindigkeit: vertretbar au�er bei 'erster Treffer pro Jahr'
'           Wenn die Datenbank Duplikats�tze enth�lt, geht Funktion 'erster Treffer pro Jahr' nicht.
'           Es kommt Fehler: Die Schl�sselspalteninformationen sind ungen�gend oder inkorrekt. Es sind zuviele
'           Zeilen von der Aktualisierung (es handelte sich um Delete) betroffen. Das ist verst�ndlich, wenn
'           2 S�tze identisch sind, welchen soll er dann l�schen.
'           Neue L�sung: Die Tabelle Fotos_ErsterTreffer entf�llt aber es gibt eine neue Tabelle FET
'           mit GROUP BY Jahr ORDER BY Jahr, dann kommt eine Abfrage mit zwei verkn�pften Tabellen Fotos und FET
'           -----------------------------------------------------------------------------------------
'           Beim Sortieren nach Spalten�berschrift (etwa 5 Sekunden) wollte ich den Mauszeiger in die Sanduhr
'           verwandeln, aber das geht nicht, weil der Mauszeiger auf 'Pfeil abw�rts' fest eingestellt bleibt
'24.06.2006 Fehlerkorrektur in allen Versionen
'           Komprimieren der Datenbank hat nicht funktioniert bei Button 'T�r zu', aber hat
'           funktioniert bei Men� beenden oder Klick aufs Schlie�-Kreuz vom Formular Query
'25.06.2006 Im Button DbGridForm.btn�ffneAnwendung wird bisher die Dateinamenerweiterung immer 3-stellig
'           angezeigt. Ab jetzt 3- oder 4- oder 5-stellig
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.10
'26.07.2006 Fehlerkorrektur in allen Versionen:
'           Beim Kopieren von Teilen des Fotoalbums mittels Drag&Drop und Auswahl der Option alle per
'           Merkerspalte markierten Datens�tze, wurden bisher alle S�tze wo die Merkerspalte eingeschaltet war,
'           nach Datei $fotos.mdb Tabelle Fotos kopiert, ohne Ber�cksichtigung der Suchkriterien.
'           Die Korrektur muss im Formular DbGridForm.btnMerkerspalteEinschalten_Click erfolgen,
'           weil hier bisher alle S�tze der Tabelle Fotos ein/ausgeschaltet wurden.
'           Jetzt werden die Suchkriterien Query.SQL ber�cksichtigt ab Abschnitt WHERE
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.11
'29.07.2006 Neue Funktion in allen Versionen:
'           Die Funktion ist nur benutzbar, wenn das Suchergebnis weniger als 100 S�tze betr�gt.
'           Jedes Bild soll mit individuell vergr��ertem/verkleinertem Bildausschnitt auch wiederholt
'           angezeigt werden k�nnen. Bisher wird bei jedem Wechsel auf das n�chste Bild die Standard-Funktion
'           benutzt (Bild zentriert darstellen, so da� es als Ganzes auf den Bildschirm pa�t)
'           L�sung: es wird ein Array eingef�hrt und ein Private Type einschlie�lich Dateiname.
'           Sowohl beim Speichern ins Array als auch beim Abrufen muss das Array durchsucht werden, ob es einen
'           Eintrag mit dem Dateiname des aktuellen Satzes gibt
'           Auf das Formular ZoomForm wirkt diese Funktion nicht.
'           Einstellbar durch eine CheckBox auf der WertxForm.
'10.08.2006 Neue Funktion in allen Versionen:
'           Bei Videos
'           MediaPlayer1.ShowStatusBar = True   'da wird Spieldauer und aktuelle Spielsekunde gezeigt
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.12
'20.08.2006 Neue Funktion in allen Versionen:
'           Im Listenfenster (F5) sollen anstelle von bisher nur '�ffnen der mit 'jpg' verkn�pften Anwendung'
'           weitere Funktionen m�glich sein, die sich auf die aktuelle Datei beziehen, mit ShellExecute.
'           Das sind jetzt insgesamt 4
'           -�ffnen der mit 'jpg' verkn�pften Anwendung f�r die aktuelle Datei (im Win 2000 gehts nicht mit diashow.exe)
'           -�ffne das Druckprogramm f�r die aktuelle Datei
'           -�ffne das Fenster 'Neue Email senden'
'           -�ffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist (markieren geht nicht im Windows 98)
'21.08.2006 Fehlerkorrektur:
'           Bei Dateinamenerweiterung "JPEG" hat bisher das Programm gedacht es handelt sich um "MPEG"
'----------------------------------------------------------------------------------------------------------
'ab Version 13.0.13
'22.08.2006 Fehlerkorrektur in allen Versionen:
'           siehe 29.07.2006 Wenn das Formular WertxForm durch Schlie�kreuzchen verlassen wird, gehen bisher
'           alle Einstellungen verloren
'28.08.2006 Verbesserung:
'           5. Funktion f�r die aktuelle Datei
'           -Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
'30.08.2006 Fehlerkorrektur: Bisher hat nicht funktioniert Export in ein Zielverzeichnis mit den mit
'           Merkerspalte markierten S�tzen, da kam Msgbox 'Programmierfehler bei FileCopy'
'04.10.2006 Fehlerkorrektur schwer reproduzierbar:
'           obwohl Rechteck-Zoom nicht eingeschaltet war, konnte manchmal ein Rechteck gezeichnet werden
'06.11.2006 alt: Das Fenster F5MehrereZeilen l��t sich nur durch Klick aufs Schlie�kreuz schlie�en.
'           neu: Jetzt schlie�t es sich auch durch Klick aufs Bild
'06.11.2006 Korrektur:
'           wenn ohne Audio-Kommentar gearbeitet wird, gibts auch keine Vorbereitungen f�r Mediaplayer
'06.11.2006 Hilfe zu den Bedientasten enth�lt bei Taste F5 nicht mehr den Hinweis 'zur�ck mit Esc'
'----------------------------------------------------------------------------------------------------------
'ab Version 13.2.1
'09.11.2006 Verbesserung alle Versionen:
'           Neue Funktionen - Ber�cksichtigung von EXIF und IPTC
'           Das Formular F5MehrereZeilen kann jetzt entweder EXIF oder IPTC Informationen zeigen
'           der Nutzer muss ein H�kchen in der entsprechenden Checkbox setzen
'           es �ffnet sich eine Listbox mit Feld & " - " & Feldinhalt
'           wenn es keinen Feldinhalt gibt wird das Feld nicht angezeigt
'26.11.2006 Verbesserung alle Versionen:
'           im Formular DbGridForm wird ganz oben links ein Thumbnail zur aktuellen Gridzeile gezeigt
'           falls es ein Bildformat 'native' ist
'----------------------------------------------------------------------------------------------------------
'ab Version 13.2.2 alle Versionen
'08.12.2006 siehe 26.11.2006 der Thumbnail hat zu starker Programmverz�gerung gef�hrt, wenn man exportiert
'----------------------------------------------------------------------------------------------------------
'ab Version 13.2.3 nur in Professional Version
'19.01.2007 Ab sofort kann auch der SQL-Text einer gespeicherten Abfrage nachbearbeitet werden
'           Ab sofort funktioniert auch bei gespeicherten Abfragen die Aktion Weiterselektieren der mit
'           Merkerspalte vorgemerkten Bildern
'           Fehlerkorrektur bei gespeicherten Abfragen:
'           Bisher kam es zum Fehler, wenn im SQL-Text zb 'Fotos.BreitePixel' formuliert wurde.
'           Richtig muss formuliert werden 'BreitePixel'
'           Das Programm ersetzt 'Fotos.' durch ""
'23.01.2007 13.2.3 alle Versionen
'           neue 6. Aktion in DbGridForm
'           -L�schen markierte Dateien(Merkerspalte) in Datenbank und Standort
'----------------------------------------------------------------------------------------------------------
'24.01.2007 13.2.4 alle Versionen
'           Fehlerkorrektur alle Versionen:
'           Bei H�kchen in 'Fenstergr��e �nderbar' und
'           H�kchen bei 'beim Bildwechsel...Bildausschnitt beibehalten' ruckt bisher das Bild erst auf die
'           linke obere Ecke, dann auf seine gespeicherte Stelle
'02.02.2007 Fehlerkorrektur alle Versionen:
'           Formular F5MehrereZeilen ist bisher nicht in der Lage die maximal 255 Zeichen anzuzeigen, die jedes
'           Standardfeld haben kann. Bei 2 Zeilen ist bisher Schlu�.
'           Ab jetzt wird �berall eine vertikale Scrollbar hinzugef�gt.
'02.02.2007 Fehlerkorrektur Professional Version:
'           Bisher wurde nach Entfernen des H�kchens aus 'gespeichert Abfragen benutzen' die Checkbox
'           'Suche nach nutzerdefinierten Felder ist aktiv' sichtbar. Das ist ab jetzt unterdr�ckt.
'04.02.2007 Fehlerkorrektur Professional Version:
'           Bei der Suche nach nutzerdefinierten Feldern gab es Probleme, wenn in nutzerdefinierten Feldern
'           gesucht wird, wo alle Datens�tze (is Null) sind. Und es gab Probleme bei verschiedenen Datentypen
'           die richtigen Vergleichsoperanden anzubieten.
'           Wenn in einem Wert=W�hrung ein Komma auftaucht muss es im SQLString durch . ersetzt werden
'           Wenn in einem Wert=Zahl ein Komma auftaucht muss es im SQLString durch . ersetzt werden
'04.02.2007 Neue Funktion Professional Version:
'           Zus�tzliche Aktion in DbGridForm - Gehe zum Hyperlink
'           zB #http://www.gerbingsoft.de# da �ffnet sich der Standardbrowser mit der angegebenen URL
'           zB #M:\P7FotoSoundVideo\FOTOS\GG\2007\Februar01.JPG# geht genau wie Doppelklick im Windows Explorer
'           zB #+:\2007\Februar01.JPG# geht genau wie Doppelklick im Windows Explorer
'09.02.2007 Verbesserung alle Versionen:
'           Im Formular DbGridForm die Combobox Combo1 als Popup-Men� darstellen. Das ist besser als die
'           bisherige Auswahl einer Zeile aus der Combobox und anschlie�end Klicken auf den Button
'           'Aktion ausf�hren'. Weil ich h�ufig zwar die Zeile gew�hlt hatte, dann aber vergessen habe auf den
'           Button 'Aktion ausf�hren' zu klicken. Dann wartet man und es passiert nichts.
'09.02.2007 Verbesserung:
'           Sprache wechseln nur mit einem Hinweis, da� eine Sicherungskopie von fotos.mdb gemacht werden soll
'14.03.2007 Verbesserung alle Versionen:
'           Neue Hilfe-Dateien im HTML-Format, weil Windows Vista das Winhelp-Format nicht mehr unterst�tzt
'           zB anstelle Fotos.hlp gibt es jetzt Fotos.chm
'----------------------------------------------------------------------------------------------------------
'07.05.2007 13.2.5 alle Versionen
'           Verbesserung
'           Gasanov EXIF OCX hat einen Fehler, es l��t bei Olympus Fotos frisch aus der Kamera einen Teil der
'           Felder weg. Erst nach Bearbeitung mit PSP9 sind alle Felder da.
'           Es wird ab jetzt ein Klassenmodul clsEFIF.cls benutzt
'10.07.2007 13.2.6 alle Versionen
'           bei der Sharewareversion hat bisher bei der Anzeige der EXIF-Werte der Scroll-Balken unten gefehlt
'----------------------------------------------------------------------------------------------------------
'29.07.2007 13.3.0 alle Versionen
'           Verbesserung
'           Man soll am Mouse-Icon erkennen, ob die Rechtecklupe aktiv ist (SquareZoom.ico)
'           Wenn das Bild mit der linken Maustaste verschoben werden kann, soll man das am Mouse-Icon erkennen.
'           Schon beim MouseDown kommt FourArrows.ico
'           Wenn beim Bildverschieben die verschobenen Positionen gemerkt werden sollen, wird
'           FourArrowsSave.ico benutzt
'31.07.2007 Fehlerkorrektur
'           Ab sofort ist es nicht mehr m�glich in der Suchmaske im Formular Query das Sternchen zu l�schen und
'           durch Leerzeichen zu ersetzen. Es h�tte die Meldung kommen m�ssen: das Feld darf nicht leer sein.
'           Ab sofort wird Trim benutzt.
'02.08.2007 Fehlerkorrektur
'           Bei Export/Import mit Drag&Drop
'           und man will die Funktion mehrfach wiederholen, kommt der Hinweis: es mu� eine zweite Anwendung
'           (Datenbank fotos.mdb) ge�ffnet sein und man muss erst die Import-Anwendung schlie�en und wieder
'           �ffnen damit es geht.
'           Ursache: ich frage die Titelzeilen der ge�ffneten Fenster ab, ob dort "FotoAlbum-" steht.
'           Nach einem Import steht dort aber der Arbeitsfortschritt.
'           L�sung: Nach Ende des Importier-Vorgangs
'           Form1.Caption = Form1.FotoAlbumTitle
'06.08.2007 Verbesserung
'           Nach 'Fotos finden' dauert es zu lange bis die Sanduhr zu sehen ist, da ist man versucht nochmal zu
'           klicken und dann kommt manchmal eine Fehlernachricht mit angeblich falscher SQL-Anweisung:
'           Ab jetzt kommt sofort die Sanduhr.
'21.08.2007
'           Verbesserung
'           Das Fenster 'Einstellungen' soll auch �bers Men� erreichbar sein.
'           Zu den Einstellungen kommt eine neue hinzu.
'           Beim Bildladen vergr��ern auf Vollbild (4 Optionen)
'           o keine
'           o ab Breite 640 Pixel
'           o ab Breite 800 Pixel
'           o ab Breite 1024 Pixel
'22.08.2007 Folgende 3 Einstellungen werden in der Datei fotos.ini gespeichert
'           [Adjustments]
'           AutomaticInterval Standard = 3
'           BackgroundColor   Standard = black
'           ZoomToFullscreen  Standard = 0
'----------------------------------------------------------------------------------------------------------
'16.09.2007 13.3.1 alle Versionen
'           in der Prozedur Kommentarform.Form_Resize
'           fehlt ein Kommentarform.Richtextbox1.Refresh
'           weil sonst die Kommentarform-H�hen�nderung durch Ziehen mit der Maus unber�cksichtigt bleibt
'25.09.2007 Beim Button 'weitere Filter' soll die Sanduhr zu sehen sein. Beim ersten Programmstart dauert es
'           ziemlich lange bis Form MP sich �ffnet
'26.09.2007 Beim Kopieren von Fotos und wenn es eine gleichnamige .wav oder .mp3 Datei gibt wie Quellname
'           dann wird diese mitkopiert
'04.10.2007 Fehlerkorrektur
'           im Zusammenhang mit 'Das Fenster 'Einstellungen' soll auch �bers Men� erreichbar sein.'
'           Es kam Laufzeitfehler 9, wenn vor dem Button 'Fotos finden' schon in den Einstellungen eingestellt
'           war 'Beim Bildwechsel individuell vergr��erten verkleinerten Bildausschnitt beibehalten.
'           L�sung: Das darf nur mit maximal 99 Fotos gehen.
'19.10.2007 Verbesserung:
'           zur L�sung des folgenden Fehlers
'           Videos vom Typ .mpg (�F69)k�nnen nicht im Fullscreenmodus hintereinander abgespielt werden
'           Windows Media Player 6.4 (mplayer2.exe) bringt ebenfalls einen Fehler am Clip-Ende
'           aber auf dem neuen PC von Betti geht alles richtig
'           L�sung:
'           Der Nutzer kann entscheiden, ob er mit der bisherigen L�sung(msdxm.ocx) arbeiten will,
'           oder ob er eine Command line zum Start des wmplayers benutzen will
'           Beispiel command line:
'           wmplayer "Dateiname" doppelte Hochkomma sind n�tig, wenn Dateiname zB Leerzeichen enth�lt
'           Ich muss nicht erfahren, ob wmplayer fertig ist, weil eine neue command line ein neues Video
'           startet, egal wie weit das bisherige ist. F�r das kontinuierliche Abspielen mit dem wmplayer
'           benutze ich mediaplayer1.Duration. Nach Ablauf von Duration wird das n�chste Video gestartet.
'----------------------------------------------------------------------------------------------------------
'07.11.2007 13.3.2 Professional Version
'           Verbesserung und Fehlerkorrektur
'           Das Programm macht die Korrektur ab jetzt selbst
'           msg = "Vermeiden Sie im SQL-String Formulierungen wie Select *, Feldname FROM..." & vbNewLine
'           msg = msg & "Dadurch w�rden die gespeicherten Feldbreiten auf falsche Felder angewendet." & vbNewLine
'           msg = msg & "Formulieren Sie besser Select * FROM..."
'           msg = msg & "Das Programm wird diese Korrekturen jetzt selbst vornehmen."
'           zwischen 'Select... und ...FROM' muss ein '*' stehen, sonst kann man den String 'Select * FROM...'
'           nicht herstellen. Es kommt ein Warnhinweis
'12.11.2007 13.3.3 Fehlerkorrektur alle Versionen
'           1. Bisher kam ein Fehler falls bei 'Explorer-Fenster �ffnen' im Dateiname ein Komma enthalten ist.
'           L�sung man muss den Dateinamen in doppelte Hochkomma einschlie�en.
'           2. Auf dem Form1-Hintergrund bei 'Videos mit externem Mediaplayer abspielen' kam bisher stets
'           ein Hinweis 'Video wird geladen. Dr�cken Sie F5 zur Anzeige des Listenfensters' auch bei Fotos.
'           Ab sofort wird der Hinweis bei Fotos unterdr�ckt.
'20.11.2007 13.3.3 Verbesserung alle Versionen
'14.11.2007 Neue Funktion:
'           gleich beim ersten Start von fotos.exe will ich die Sprache festlegen
'           Das geschieht durch Auslieferung der fotos.ini mit dem Wert
'           Language=9     ;not selected yet
'21.11.2007 13.3.3 Fehlerkorrektur alle Versionen
'           Wenn ich ohne /WRITE arbeitete, bekam ich nach Spaltensortierung trotzdem Schreibzugriff.
'           Ich muss nach der Spaltensortierung so verfahren, wie bei Query_btnOK_Click.
'21.11.2007 13.3.3 Fehlerkorrektur alle Versionen
'           bei falsch oder nicht registrierter dao360.dll
'           kommt die msgbox
'           Errornumber=429
'           Errortext=Objekterstellung durch ActiveX-Komponente nicht m�glich
'           You must register the dao360.dll
'           read in http://www.gerbingsoft.de or look for that problem in the internet
'           dann wird das Programm beendet
'------------------------------------------------------------------------------------------------------------
'24.12.2007 13.3.4 Verbesserung alle Versionen
'           Der MSI-Installer hat einen Schwachpunkt. Es gibt nur per-user-Installationen,
'           nicht per-machine-Installationen. Beim ersten Aufruf nach der Installation soll der Nutzer die
'           M�glichkeit erhalten das Desktop-Symbol und das Startmen� vom aktuellen Nutzer in den Ordner
'           C:\Dokumente und Einstellungen\All Users zu kopieren.
'           Einzig sicher ist das Kopieren von C:\Dokumente und Einstellungen\Nutzer\Desktop
'           nach C:\Dokumente und Einstellungen\All Users\Desktop weil es keine R�ckkopplung gibt, welchen
'           Installationsordner der Nutzer genommen hat.
'           Dazu wird frmSprache aufgebohrt.
'27.12.2007 13.3.4 Verbesserung Professional-Version
'           Beim Editieren eines Hyperlinks muss kontrolliert werden ob er das richtige Format hat.
'           Ein Hyperlink muss in # eingeschlossen sein. Beispielsweise #http://www.gerbingsoft.de#
'           Leider kann man bei ADO.Fields die Datentyp-Eigenschaft nicht feststellen.
'           Das geht nur bei DAO.Fields.Attributes. Gleich bei der ersten Benutzung von DAO untersuche ich die
'           Tabelle Fotos bzw EFotos ob es Felder mit Fields.Attributes Hyperlink gibt. Die Spaltennummer
'           dieser Felder wird in der Collection HyperLinkColumns gespeichert.
'28.12.2007 13.3.4 Fehlerkorrektur alle Versionen
'           siehe 21.11.2007 bei falsch oder nicht registrierter dao360.dll
'30.12.2007 13.3.4 Verbesserung Professional-Version
'           Bisher waren nach 'erster Treffer pro Jahr' keine Stichwort�nderungen m�glich.
'           L�sung: Alle S�tze mit dem betreffenden Dateiname in Tabelle (E)Fotos suchen, falls es wieder
'           Erwarten Duplikate gibt. Dann die Stichwort�nderung in Tabelle (E)Fotos nachziehen
'31.12.2007 13.3.4 Fehlerkorrektur alle Versionen
'           Bisher gab es keine Kontrolle ob beim Editieren in Spalte SWF falsche Werte eingetragen werden
'30.12.2007 13.3.4 Verbesserung alle Versionen
'           Die Spalte AudioFileExist darf garnicht editiert werden k�nnen.
'           Bei BreitePixel und HoehePixel war die Pr�fung falsch. Sie d�rfen garnicht editiert werden k�nnen.
'01.01.2008 13.3.4 Verbesserung Professional-Version
'           In Query.FrageObNurErstenTreffer hat irgendwas nicht gestimmt
'           bei Kombination von 'erster Treffer pro Jahr' und 'Suche Begriff in jedem Feld'
'           manchmal wurde zB zu Andreas nichts gefunden weil Tabelle FET leer war
'06.01.2008 13.3.4 Fehlerkorrektur alle Versionen
'           Fehlerkorrektur zur IPTC-Anzeige
'           es gibt Felder, die l�nger sein k�nnen als eine Zeile in der Listbox lstExifIptc
'           diese Felder muss man in mehrere Zeilen zerlegen
'09.01.2008 13.3.4 Verbesserung alle Versionen
'           Neue RAW-Datenformate
'           werden bei den Link-Datentypen erlaubt
'           3FR ARW CS1 CS4 CS16 DCS ERF MEF SR2
'10.01.2008 13.3.4 Verbesserung alle Versionen
'           Mit FotosMdb.exe Funktion 'IPTC' kann man den Inhalt der Datenbankfelder in die JPG-Dateien
'           �bertragen. Damit geht man den umgekehrten Weg wie bei der Aufnahme neuer Dateis�tze wo die
'           IPTC-Felder in die Datenbank �bertragen werden k�nnen. Es geht ausschlie�lich mit JPG-Dateien.
'           Vorhandene IPTC-Felder bleiben vorhanden, sofern nicht Datenbankfelder angegeben werden, die zum
'           �berschreiben benutzt werden sollen.
'------------------------------------------------------------------------------------------------------------
'19.01.2008 13.3.5 Fehlerkorrektur Professional-Version
'           siehe Gerbing 10.06.2005
'           Das Entfernen irrt�mlich gew�hlter Feldnamen war auskommentiert
'26.01.2008 13.3.5 Fehlerkorrektur alle Versionen
'           sehr sp�t gemerkter Fehler
'           Wenn ich bisher SQL nachbearbeiten wollte (H�kchen gesetzt) und schnell nacheinander auf den
'           Button 'Fotos finden' geklickt habe, da komme ich nicht zum SQL nachbearbeiten sondern es zeigt
'           mir sofort die Fotos an
'           Ursache: ich habe bisher vergessen den Button 'Fotos finden' zwischendurch zu disablen
'04.02.2008 13.3.5 Fehlerkorrektur alle Versionen
'           Verbesserung in FotosMdb zu IPTC siehe 10.01.2008
'           Datenbankfeld IPTCPresent erlaubt einfaches Nachziehen von Fotos, in die bisher keine
'           IPTC Felder �bertragen worden sind.
'------------------------------------------------------------------------------------------------------------
'10.03.2008 13.3.6 Fehlerkorrektur alle Versionen
'           Wenn man bei 'SQL Nachbearbeiten' das H�kchen rausnimmt, muss Button 'Fotos finden' enabled werden
'------------------------------------------------------------------------------------------------------------
'16.04.2008 13.3.7 Verbesserung alle Versionen
'           Es gibt eine neue 'Aktion w�hlen...'
'           �ffne RenamMdb f�r die aktuelle Datei
'16.04.2008 13.3.7 Fehlerkorrektur alle Versionen
'           �nderungen am Feld IPTCPresent sind verboten
'15.06.2008 13.3.7 Verbesserung alle Versionen
'           Wenn weitere Filter aktiv sind, steht in Personen 'WEITERE FILTER'. Es soll aber nur wenn mehrere
'           Personen ausgew�hlt sind, dort stehen 'MEHRERE PERSONEN'.
'15.06.2008 13.3.7 Verbesserung
'           So wie es bei Audio-Kommentar -> Aufnahme Start schon der Fall ist, soll auch bei
'           Audio-Kommentar l�schen, zur Sicherheit nochmal nachgefragt werden, ob wirklich gel�scht werden
'           soll.
'------------------------------------------------------------------------------------------------------------
'25.06.2008 13.3.8 Verbesserung alle Versionen
'           Bei Programmstart soll ab jetzt die Spalte Merker auf 0 gel�scht werden, sonst kann es passieren,
'           dass man beim L�schen von S�tzen, die durch Merkerspalte ausgew�hlt sind, diejenigen �bersieht,
'           die noch von einem vorherigen Programmlauf markiert sind.
'           Gleichzeitig soll die Msgbox 'Wollen Sie wirklich alle mit der Merkerspalte markierten Dateien aus
'           der Datenbank und an ihrem Standort l�schen?' die Anzahl Recordcount enthalten
'30.06.2008 13.3.8 Verbesserung
'           Wenn man die Form WertxForm anzeigt durch den Men�-Punkt Einstellungen in der Form Query,
'           dann m�ssen die Buttons zur 'Automatik starten' und alle Einstellungen zur 'ver�nderten Helligkeit'
'           unsichtbar sein.
'09.07.2008 13.3.8 Fehlerkorrektur
'           Wenn es nur eine Datei als Suchergebnis gibt und man benutzt F5 und anschlie�end Aktionen:
'           '�ffne Explorer-Fenster wo die aktuelle Datei markiert ist" dann �ffnet sich der Explorer und zeigt
'           den Arbeitsplatz
'03.08.2008 13.3.8 Fehlerkorrektur alle Versionen
'           irgendwie hatte sich der Fehler eingeschlichen, dass man mit einer schreibgesch�tzten fotos.mdb
'           auch im Lesemodus nicht weiterarbeiten konnte.
'----------------------------------------------------------------------------------------------------------
'           Version 13.3.8 ist die letzte, die Win98 und Win2000 unterst�tzt, weil wmp.dll sich nicht registrieren l��t
'----------------------------------------------------------------------------------------------------------
'22.08.2008 13.3.9 Verbesserung alle Versionen
'           Man sollte die Datenbank gleich beim Start Komprimieren, weil bei Pr�fen3
'           und bei Arbeit mit 'Nur den ersten Treffer pro Jahr erlauben'
'           der Umfang immer gr��er wird
'26.08.2008 13.3.9 Verbesserung Professional Version
'           Bei 'erster Treffer pro Jahr' und nicht vorhandener Tabelle FET soll es trotzdem weitergehen
'26.08.2008 13.3.9 Verbesserung Professional Version
'           Bei 'erster Treffer pro Jahr' und �ndern eines Spalteninhaltes sah es f�r den Anfang so aus, als
'           ob die �nderung gemacht w�rde, aber beim Klick in eine andere Zeile wird der alte Inhalt
'           wiederhergestellt.
'           Das liegt daran weil ein Recordset mit 'Inner join' cannot be updated. Der Cursortype bleibt auf
'           OpenStatic stehen.
'           L�sung: Nach �bernahme der �nderung in die Tabelle Fotos rufe ich die Prozedur
'           FrageObNurErstenTreffer erneut auf.
'           Wenn nach der �nderung eines Feldinhaltes F2 oder F3 oder �quivalent gedr�ckt wird, bleibt das
'           Programm stehen bzw unbehandelte Ausnahme in Vb6.exe(Oledb32.dll)... Access Violation
'           Ursache nicht erkennbar. L�sung: Diese Tasten werden bei 'erster Treffer pro Jahr'
'           ignoriert.
'           Aber: Wenn man vorher au�erhalb des Grid klickt, dann geht F2/F3
'01.09.2008 13.3.9 Verbesserung alle Versionen
'           Zus�tzlich zu Mediaplayer 6.4 (OCX) ist ausw�hlbar der aktuellen Windows Media Player (wmp.dll)
'           weil es passiert ist, dass einige Videoclips sich nicht abspielen lassen wollten.
'           deshalb siehe auch 19.10.2007 �bergang zu externem Mediaplayer. Eigentlich sollte ab jetzt der
'           externe Mediaplayer �berfl�ssig sein.
'           Mediaplayer1 = wmp.dll
'           Mediaplayer2 = msdxm.ocx
'02.09.2008 13.3.9 Verbesserung alle Versionen
'           Wenn der Benutzer mit einem eingeschr�nkten Benutzerkonto arbeitet, hat er bisher keinen Hinweis
'           bekommen, wenn die Dateien fotos.ini und pruef.log f�r den Schreibzugriff gesperrt sind.
'----------------------------------------------------------------------------------------------------------
'10.10.2008 13.3.10
'           Es kommt eine Msgbox zum Fehler -2147221164 Klasse nicht registriert
'           wenn XP SP3 vor GERBING Fotoalbum 13 installiert wird
'----------------------------------------------------------------------------------------------------------
'20.11.2008 13.3.11 Verbesserung alle Versionen
'           Bei 'Fenstergr��e �nderbar' soll ab jetzt Form1 ganz oben links angeordnet werden.
'               Form1.Width soll = Screenwidth/2 sein, Form1.Height soll = ScreenHeight sein
'               und durch Form1.Controlbox = True soll es m�glich sein Min, Max und Close auszuw�hlen.
'               Wenn bereits ein zweites Fotoalbum gestartet wurde, soll Form1 am rechten Bildschirmrand
'               angeordnet werden.
'20.11.2008 Nicht realisierter Versuch:
'           Das Formular Query soll nicht mehr aus der Taskleiste verschwinden und es soll einen Min-Button
'               in der Titelleiste haben. Dazu darf man nicht Query.Hide benutzen, sondern muss in einer
'               Schleife solange warten, bis die Suchkriterien eingetippt und der Button 'Fotos finden' geklickt
'               wurde. siehe auch 06.12.2005
'               Das Weglassen von Query.Hide hat aber ungewollte Nebeneffekte. Jetzt gibt es zwei Eintr�ge in
'               der Taskleiste zu GERBING Fotoalbum. Bildanzeigen kommt manchmal zweimal dran und manchmal
'               garnicht.
'           Realisierte L�sung:
'           Benutzung der Prozedur Query.FormInTaskbar (gefunden im Internet)
'23.11.2008 13.3.11 Verbesserung alle Versionen
'           Wenn Gespeicherte Abfragen dran war und wieder zur�ckgewechselt wird, muss intern ein Aufruf von
'           Refresh gemacht werden
'03.01.2009 13.3.11 Verbesserung alle Versionen
'           Mehrere Korrekturen in FotosMdb
'23.02.2009 13.3.11 Neue Freeware Komponente
'           Batch Histogram Correction
'----------------------------------------------------------------------------------------------------------
'11.03.2009 13.3.12 alle Versionen
'           Fehler tritt mal auf und mal nicht
'           H�kchen setzen bei 'Fenstergr��e �nderbar' -> Kopieren -> Mit Drag&Drop -> Die korrekte Msgbox
'           'es mu� eine zweite Anwendung (Datenbank fotos.mdb) ge�ffnet sein....
'           ist hinter anderen Fenstern versteckt oder kommt garnicht.
'           L�sung: Private Const vbMsgBoxTopMost As Long = &H40000 benutzen
'                   neuer Algorithmus zum Suchen aller Fenstertitel
'           Achtung es darf kein DoEvents drin sein
'13.03.2009 13.3.12 alle Versionen
'           Fehler beim Abspielen von Videos
'           Laufzeitfehler '438' wenn man nach Start des ersten Videos auf F5 und danach wieder aufs Video klickt
'16.03.2009 13.3.12 alle Versionen
'           Bei Funktionstaste F12 gibt es eine neue Option
'           ab Bildbreite 1024 oder ab Bildh�he 768 Pixel
'23.03.2009 13.3.12 alle Versionen
'           Seit Benutzung eines neuen Flachbildschirms mit Aufl�sung 1600x1200 Pixel habe ich ver�ndert
'           Anzeige -> Einstellungen -> Erweitert -> 120 DPI und
'           Anzeige -> Darstellung -> Schriftgrad:gro� da wird die Schrift sch�rfer besonders bei PDF-Dateien
'           aber im DbGridForm.DataGrid ist dadurch der Cursor unsichtbar geworden, weil RowHeight zu niedrig war, DBGridNeu.RowHeight = 180
'           ab jetzt DBGridNeu.RowHeight = 220
'25.03.2009 13.3.12 alle Versionen
'           siehe 19.01.2007
'           Es hat immer noch nicht funktioniert den SQL-Text einer gespeicherten Abfrage zu editieren
'06.05.2009 13.3.12 alle Versionen
'           alt:LoadResString(3081 + Sprache)          'Videos abspielen mit internem Mediaplayer 10
'           neu:LoadResString(3081 + Sprache)          'Videos abspielen mit internem Mediaplayer 7 oder aufw�rts
'----------------------------------------------------------------------------------------------------------
'10.06.2009 13.3.13 alle Versionen
'           manchmal hat die Spaltensortierung nach Klick auf Headline nicht funktioniert. Schuld war ein DoEvents
'           bei dem es h�ngen blieb.
'13.08.2009 13.3.13 alle Versionen
'           In ExportForm DoEvents auskommentiert weil Blockierung auftritt, beim Exportieren, egal ob Drag&Drop oder
'           �ber auszuw�hlenden Ordner.
'03.09.2009 13.3.13 alle Versionen
'           Bei Benutzung von breiten Flachbildschirmen zB 1920x1200
'           wird ein Bild mit Gr��enverh�ltnis 4:3 zu einem Ei zusammengedr�ckt
'----------------------------------------------------------------------------------------------------------
'14.09.2009 13.3.14 alle Versionen
'           AboutForm anzeigen aus der Form Query
'24.09.2009 13.3.14 alle Versionen
'           Ich muss zus�tzlich zum Breiten/H�hen-Verh�ltnis des Bildes BHV
'           das Breiten/H�hen-Verh�ltnis des Bildschirms SGVH auswerten
'           If BHV < SGVH Then
'               'das Bild ist zu hoch und zu schmal
'           If BHV > SGVH Then
'               'das Bild ist zu niedrig und zu breit
'29.09.2009 13.3.14 alle Versionen
'           Manchmal beim Exportieren bleibt es beim 20. Bild h�ngen. Schuld war ein DoEvents
'23.11.2009 Formale Korrektur bei Query.Label1.Caption Der Inhalt des Feldes 'gespeicherte Abfragen in fotos.mdb'
'           war nicht komplett lesbar weil das Feld nicht breit genug war
'----------------------------------------------------------------------------------------------------------
'09.12.2009 13.3.15 Professional Version
'           Verbesserung:
'           Falls zu einem Bild gleichnamige Sound-Dateien vorhanden sind, will ich ab jetzt zwei Varianten anbieten,
'           wie die Sound-Dateien gestartet werden k�nnen.
'           - Sound-Dateien, falls vorhanden, sofort mit der Bildanzeige automatisch starten, neues Formular frmStartSoundAutomatisch,
'               dieses Formular bleibt stets unsichtbar
'           - Sound-Dateien, falls vorhanden, �bers Kommentar-Fenster manuell starten
'           Wenn in WertxForm (Werte einstellen) nicht optmanuell ausgew�hlt wird, kann der Nutzer keine Sound-Kommentare bearbeiten
'----------------------------------------------------------------------------------------------------------
'11.03.2010 13.3.16 alle Versionen
'           Anpassung an Vista und Windows7:
'           In alle exe-Dateien wird ein Manifest eingef�gt mit requestedExecutionLevel = requireAdministrator
'           Dazu dient ManifestEinf�gen.exe. Dadurch wird in Vista und Windows7 der Nutzer aufgefordert als Administrator zu starten.
'           Die /WRITE-L�sung wird entfernt. Sie stammt aus Windows95/98-Zeiten, als es mit Mitteln des Betriebssystems noch schwierig
'           war einem Nutzer den Zugriff auf Datenbankdaten zu verwehren. Heute ist es mit NTFS-Rechten leicht, zu differenzieren, ob ein
'           Nutzer den Zugriff auf Datenbankdaten erhalten oder nicht erhalten soll.
'           Ich setze gstrCommandLine = "/WRITE" anstelle die command line einzulesen
'----------------------------------------------------------------------------------------------------------
'28.03.2010 13.3.17 alle Versionen
'           Verbesserung:
'           Wenn nur Fotos.exe und fotos.mdb in einem Ordner stehen, kam bisher Laufzeitfehler '13' Typen unvertr�glich
'----------------------------------------------------------------------------------------------------------
'07.05.2010 13.3.17 alle Versionen
'           Verbesserung:
'           Man muss das H�kchen �berarbeiten der SQL-Anweisung nach jeder Suche ausschalten
'----------------------------------------------------------------------------------------------------------
'03.09.2010 13.3.18 Verbesserung alle Versionen:
'           Die Rechteck-Lupe soll schon beim Zeichnen des Rechtecks das echte Breiten/H�hen-Verh�ltnis beachten
'           ich muss das Breiten/H�hen-Verh�ltnis des Bildschirms SGVH auswerten
'27.09.2010 13.3.18 Verbesserung alle Versionen:
'           Wenn in der Entwicklungsumgebung die Datei fotos.ini fehlt, kommt Laufzeitfehler '13' Typen unvertr�glich
'           Wenn in einer aus msi-Datei installierten Version die Datei fotos.ini fehlt, wird die Installation wiederholt und die fehlende
'           Datei dabei erzeugt.
'22.11.2010 13.3.18 Verbesserung alle Versionen:
'           Die automatische Diashow kann ab jetzt mit zwei Sortierreihenfolgen arbeiten
'           zuf�llig
'           alphabetisch aufsteigend
'22.11.2010 13.3.18 Fehlerkorrektur alle Versionen:
'           Laufzeitfehler 402 'Das oberste Formular muss zuerst geschlossen oder ausgeblendet werden'
'           im Formular WertxForm, wenn man gleich zu Beginn im Formular WertxForm 'Automatik zuf�llig' einstellt,
'           danach alle Fotos finden
'           danach Tasten Strg+F6
'05.12.2010 13.3.18 Verbesserung alle Versionen:
'           Bei Dr�cken von F5 soll die Thumbnailansicht 1000 x 750 Twips gro� sein
'20.12.2010 13.3.18 Verbesserung alle Versionen:
'           F�r Multinutzer-Umgebungen wird in DbGridform ein btnRefresh eingef�hrt. Damit kann ein Multiuser leichter feststellen,
'           ob seine �nderungen gemacht wurden, oder ob es Konflikte mit anderen usern gegeben hat.
'           Mit btnShowUsers kann man f�r Testzwecke sehen, wieviel user die Datenbank ge�ffnet haben, als Debug.Print
'20.12.2010 13.3.18 Verbesserung alle Versionen:
'           F�r Multiuser-Umgebungen ist es notwendig, da� jeder user seine eigene fotos.ini besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'           Beim Packen des Installationspaketes muss die fotos.ini eingeordnet werden nach
'           x:\Dokumente und Einstellungen\user\Anwendungsdaten\GERBING Fotoalbum 13
'----------------------------------------------------------------------------------------------------------
'01.01.2011 13.3.19 Verbesserung Professional Version:
'           Bei 'erster Treffer pro Jahr' wird trotz guter Absicht immer noch keine Stichwort�nderung durchgef�hrt, weil
'           DbGridForm.DBGridNeu.AllowUpdate = False gesetzt wurde
'17.02.2011 13.3.19 Verbesserung alle Versionen:
'           In Fotosmdb.exe
'           F�r Multiuser-Umgebungen ist es notwendig, da� jeder user seine eigene pruef.log (englisch check.log) besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'           In Diashow.exe Fotosmdb.exe Renammdb.exe
'           F�r Multiuser-Umgebungen ist es notwendig, da� jeder user seine eigene fotos.ini besitzt.
'           das geschieht mit GetSpecialFolder(CSIDL_APPDATA)              'x:\Dokumente und Einstellungen\user\Anwendungsdaten
'----------------------------------------------------------------------------------------------------------
'18.02.2011 13.3.20 Verbesserung alle Versionen:
'           Sprache wechseln wird verboten, wenn mehr als 1 Nutzer mit der Datenbank arbeitet. Bisher kommt auch schon ein Hinweis
'           mit error 3356, danach Laufzeitfehler und Programmende
'19.02.2011 13.3.20 Fehlerkorrektur alle Versionen:
'           Bisher kam Laufzeitfehler 9; Index au�erhalb des g�ltigen Bereiches
'           wenn die EXIF-Felder einer Datei mit Dateil�nge=0 angezeigt werden sollten
'23.06.2011 13.3.20 Verbesserung alle Versionen:
'           Ich mache die Gr��e der Fonts f�r die Controls abh�ngig von der Einstellung unter 'Eigenschaften von Anzeige' ->
'           Erweitert -> DPI-Einstellungen. Das geschieht automatisch beim Form_Load jedes Formulars.
'           Ich unterscheide normal=96, gro�=120, sehr gro�>120
'           Das erfordert Bildschirmaufl�sung mindestens 1024 x 768 bei 96 DPI und
'           mindestens 1280 x 800 bei 120 DPI
'           Der Nutzer soll entscheiden, ob er die Fontgr��en-Anpassung haben will, wenn eine DPI-Einstellung h�her als 96
'           gefunden wird im Formular WertxForm und der Wert wird in Fotos.ini gespeichert
'29.06.2011 im Zuge der Fontgr��en-Anpassung gefundene Fehler:
'           Form ExportForm - �nderung vom 13.08.2009 �berarbeitet DoEvents in ExportForm ersetzen durch Control.Refresh
'           Beim Export mit Drag&Drop kommt in der Exportvorbereitung falsch immer rst1.Recordcount = 1
'           Form ND - Nach Schlie�en mit Schlie�kreuz kam MsgBox angeblich keine nutzerdefinierten Felder vorhanden
'           Form KommentarForm - beim Durchbl�ttern mit F3 kommt Laufzeitfehler, wenn eine Datei einen Audiokommentar enth�lt
'           Form MP - nach Schlie�en mit Schlie�kreuz steht da 'weitere Filter sind aktiv' aber in Wirklichkeit sind keine aktiv
'           Form ZielverzeichnisForm - �berarbeitet f�r die Benutzung von BrowseForFolder
'----------------------------------------------------------------------------------------------------------
'===========�bergang zu Windows7 64bit===================================================================================================
'26.10.2011 13.3.21 �nderung alle Versionen:
'           Ich habe jetzt Windows 7 (64 bit) und lasse VB6 dort laufen. Konsequenzen sind:
'           msdxm.ocx (version6.4) l�uft nicht unter Windows 7 (64bit) die Anwender sehen nur das erste Video, nach Vor/Zur�ck-Bl�ttern
'           kommt nur ein senkrechter wei�er Strich und der Ton ist noch zu h�ren.
'           Jeder Versuch das msdxm.ocx in die IDE (Entwicklungsumgebung) einzubinden ist mi�lungen. Das wird durch Recherchen im Internet
'           best�tigt. Deshalb habe ich entfernt in Form1 mediaplayer2. Es gibt nur noch die Auswahl zwischen Videos abspielen mit
'           -interner Mediaplayer 7 oder h�her
'           -externer Mediaplayer
'26.10.2011 13.3.21 �nderung alle Versionen:
'           Ab sofort entf�llt in frmSprache die Auswahlm�glichkeit
'           Wie sollen die Verkn�pfungen angelegt werden
'           Die Installation von GERBING Fotoalbum 13 muss in jedem Benutzerkonto ausgef�hrt werden, das damit arbeiten will,
'           weil sonst unsicher ist ob jedes Benutzerkonto seine fotos.ini bekommt
'26.10.2011 13.3.21 �nderung alle Versionen:
'           Ab sofort entf�llt das Erzeugen von Audio-Kommentaren �ber die KommentarForm. Begr�ndung: Audio-Kommentare k�nnen vielf�ltig angelegt
'           werden, oft sogar mit einer digitalen Kamera, f�r den Fall mit Mikrofon gibt es gen�gend Freeware. Au�erdem gibt es keine
'           R�ckmeldungen, da� diese Funktion je benutzt worden ist.
'04.11.2011 13.3.21 �nderung alle Versionen:
'           Ich habe den Winkelmann-Fehler im Windows 7 gefunden. Bei Dr�cken der Taste F5 kommt ein leeres Grid.
'           und beim �ffnen der Query-Form kommt Fehler-Nr.: -2147467262
'           Ein nackiges Windows 7 ohne Microsoft Office bringt diesen Fehler. Die Installation einer beliebigen Office Komponente
'           ab Office 2003 (probiert mit Word) beseitigt den Fehler. Er tritt auch dann nicht mehr auf, wenn Office wieder deinstalliert
'           wird.
'           Ich muss in frmSprache zu Beginn ermitteln in welchem Betriebssystem ich arbeite.
'           Bei XP und Vista geht es weiter mit der Sprachauswahl.
'           Bei Windows7 und h�her, muss ich fragen ob Office 2003 oder h�her installiert ist, wenn ja geht es weiter mit der Sprachauswahl.
'           Wenn nein, kommt eine MsgBox mit dem Hinweis, da� erst Office 2003 oder h�her installiert werden muss. Dann endet das Programm.
'05.11.2011 13.3.21 �nderung alle Versionen:
'           Ab sofort wird CommonDialog ohne comdlg32.ocx realisiert, weil es Registrierungsprobleme gab, nachdem ich abwechselnd mal in der
'           Entwicklungsumgebung unter win7 und mal unter XP gearbeitet habe.
'           Fehlernr '339' comdlg32.ocx ...nicht registriert
'07.11.2011 13.3.21 Verbesserung alle Versionen:
'           Verbesserung f�r Multi-Nutzer-Umgebung. Vermeidung von overhead, der entsteht bei Benutzung einer fotos.exe vom fremden PC.
'           Jeder PC hat seine lokale fotos.exe und w�hlt aus, mit welcher fotos.mdb aus einem fremden Ordner oder fremden PC er arbeiten will.
'           Dazu muss der Nutzer beim Start der lokalen fotos.exe die Shift-Taste festhalten. Daraufhin geht ein CommonDialog (ohne ocx) auf zur
'           Auswahl der fotos.mdb
'           Der Ordnername der fotos.mdb steht in gstrFotosMdbLocation.
'           Wenn gstFotosMdbLocation leer ist, wird AppPath benutzt. Wenn gstrFotosMdbLocation <> "" ist, werden die Tools FotosMdb und Renammdb
'           mit Aufrufparameter gstrFotosMdbLocation gestartet.
'15.11.2011 13.3.21 Fehlerkorrektur alle Versionen:
'           Der externe Windows Mediaplayer l�uft kontinuierlich weiter, wenn man F8 gedr�ckt hatte. Das darf nicht sein.
'15.11.2011 13.3.21 Fehlerkorrektur Professional-Version:
'           Im Win7 passiert es, dass die Professional Version sich nicht herstellen l��t. Sie behauptet, sie w�re Shareware-Version.
'           Das kommt von RegProfi.exe, dies bildet sich ein, es schreibt die Datei msprivs.log nach GetSystemDirectoryA (C:\Windows\system32)
'           schreibt aber in Wirklichkeit nach C:\users\vm\AppData\VirtualStore\Windows\System32
'           Das liegt daran, dass RegProfi.exe eigentlich mit Manifest arbeiten m��te, Aber dann kommt Installer-Fehler 1721.
'           Darum schreibe ich die Datei msprivs.log ab sofort in den Pfad von fotos.ini (gstrFotosIniAnwendungsOrdner)
'----------------------------------------------------------------------------------------------------------
'23.11.2011 13.3.22 Verbesserung alle Versionen:
'           Generelles Entfernen von CommonDialog comdlg32.ocx  zuletzt aus Diashow.exe stattdessen Benutzung von standarddialoge.bas
'           Generelles Entfernen von sysinfo.ocx aus fotos.exe/fotosmdb.exe/renammdb.exe. Es wurde benutzt um Win98 zu erkennen.
'24.11.2011 13.3.22 Verbesserung alle Versionen:
'           Windows7 Drag&Drop aus einem Explorer-Fenster nach FotosMdb.exe oder Diashow.exe geht nicht, weil es nicht m�glich ist, ein Explorer-
'           Fenster in seinen Rechten genauso hoch anzuheben wie FotosMdb.exe oder Diashow.exe gehoben sind durch das Manifest mit
'           requireAdministrator und ich kann nicht verlangen, dass alle Nutzer den Total Commander benutzen.
'           L�sung:
'           Ich ersetze im Manifest requireAdministrator durch asInvoker und gebe den Rat, das GERBING Fotoalbum nicht nach C:\Programme zu
'           installieren, sondern nach C:\Fotoalbum
'25.11.2011 13.3.22 Verbesserung alle Versionen:
'           Der Zwang beim Start von fotos.exe als Administrator zu arbeiten ist im Windows7 nur vorhanden, wenn GERBING Fotoalbum
'           nach C:\Programme oder C: installiert wird. Ich erstelle in Zukunft ein MSI-Paket immer mit
'           Installationsordner = Pers�nliche Daten des Benutzers (C:\Dokumente und Einstellungen\gottfried\Eigene Dateien\GERBING Fotoalbum)
'           Der erste Start von fotos.exe mit Administratorrechten bleibt aber weiterhin n�tig, damit msdmo.log nach
'           C:\windows\SysWOW64\msdmo.log installiert werden kann, mit Inhalt 'start-end'
'04.12.2011 13.3.22 Verbesserung alle Versionen:
'           Ursache f�r Laufzeitfehler '13' Typen unvertr�glich gefunden
'           Die fotos.ini wird nicht mehr ausgeliefert, sondern im Ordner
'           gstrFotosIniAnwendungsOrdner = GetSpecialFolder(CSIDL_APPDATA)  & "\GERBING Fotoalbum 13" selbst angelegt
'           im XP          x:\Dokumente und Einstellungen\user\Anwendungsdaten
'           im Windows7    C:\Users\gottfried\AppData\Roaming
'28.12.2011 13.3.22 Verbesserung alle Versionen:
'           Es gibt Nutzer, bei denen weder der interne windows mediaplayer noch der externe windows mediaplayer funktioniert
'           Solche Nutzer erhalten die Gelegenheit, sich selbst eine mediaplayer nach eigenem Geschmack auszuw�hlen
'           In WertxForm gibt es daf�r die Option optOtherExternalPlayer
'           Geeignete mediaplayer sind VLC player oder Irfan View, aber man muss dort ein H�kchen setzen bei
'           Werkzeuge -> Einstellungen -> 'Nur eine Instanz erlauben'
'----------------------------------------------------------------------------------------------------------
'29.12.2011 13.4.0 Neue Funktion alle Versionen:
'           Bereitstellung einer SQL-Server-Version
'           Diese Version gibt es nicht kostenlos. Jede Lizenz(jeder Nutzer) kostet 9,95 Euro
'
'           Es gibt ein Installationspaket f�r den Server das muss f�r jeden Kunde neu geschn�rt werden wegen dem LicenseCode
'               sql_fotos.mdf       mit drei Datens�tzen in Tabelle Fotos
'               $sql_fotos.mdf      mit drei Datens�tzen in Tabelle Fotos
'               2005 Fotos          mit drei Fotos
'               EnterNewUsers.exe
'           Im Hinweis-Fenster muss stehen, da� der Installations-Standort notiert werden muss. Der Client muss wissen, wo die Fotos
'           stehen. Der SQL-Server-Administrator muss die mdf-Dateien attachen.
'
'           Das Installationspaket f�r den Client ist dasselbe wie bei der Access-Version, aber es fehlt die Fotos.mdb und Ordner 2005
'               alle Programmdateien
'           F�r den ersten Start werden die SQL-Server-Connect-Parameter gebraucht, dann werden sie in die fotos.ini geschrieben.
'           Genauso der Standort der Fotos/Videos.
'           Wenn noch keine Nutzer eingetippt sind, m�ssen mit EnterNewUsers.exe welche erzeugt werden, bis zur Maximalanzahl
'
'           fotos.exe fotosmdb.exe renammdb.exe m�ssen aus der fotos.ini entnehmen, wo die Fotos/Videos stehen und
'           die SQL-Server-Connect-Parameter. Wenn Parameter mit CommandLine �bergeben werden, erscheint kein Connect-Fenster
'           fotos.exe fotosmdb.exe renammdb.exe machen zwar Connect zur Datenbank, aber kein Login
'
'           Es gibt eine neue Tabelle LicenseCode mit einer Spalte LicenseCode (string max 60)
'           Dort ist der bisherige license code erweitert um eine vorneran gestellte Kolonne von 5 Bytes
'           SQL01 f�r eine Lizenz
'           SQL99 f�r die unbegrenzte Anzahl Lizenzen (verschl�sselt mit dem Name), solche Benutzer brauchen kein Login zu machen
'
'           Die Tabelle loggedinusers enth�lt die Spalten username LoggedIn Management
'               Die Tabelle dient der �berpr�fung der gekauften Lizenzen. Es k�nnen sich maximal soviele user anmelden, wie Lizenzen gekauft wurden.
'               Alle user m�ssen sich Einloggen. Beim Programmende erfolgt das Ausloggen
'               Der Zeitpunkt f�r das Login ist nach dem erfolgreichen Connect.
'               K�ufer von SQL01 bis SQL98 Lizenzen m�ssen vor dem Einloggen des ersten Benutzers 01 bis 98 usernames anlegen.
'               Wenn ein regul�res Ausloggen verpa�t wurde, zB wegen Stromausfall oder Laufzeitfehler muss der betroffene user zum SQL-Administrator
'               gehen und sich zur�cksetzen lassen. Es wird in Spalte Management 'OUT&datum&Uhrzeit' eingetragen und veschl�sselt mit den username
'               Auch ein gleichzeitiges Reset aller usernames geht. Da sollten alle Nutzer vorher Ausloggen, sonst werden sie gewaltsam ausgeloggt.
'               In der Spalte LoggedIn steht unverschl�sselt ob der username eingeloggt ist oder nicht
'               In der Spalte Management steht 'IN ' & Datum&Uhrzeit oder 'OUT' & Datum&Uhrzeit und verschl�sselt mit dem username
'               Beim Anlegen neuer Nutzer und beim Reset wird eingetragen 'OUT&Datum&Uhrzeit' und verschl�sselt mit dem username
'               Wenn fotos.exe bei der regul�ren Arbeit merkt, dass in der Spalte Management nicht 'IN ...' steht, oder es ist ein
'               verpfuschtes Datum, dann wird das Programm gewaltsam beendet
'               Editieren ist nur in Spalte username erlaubt und LoggedIn und Management verboten
'
'           Die Spalte Dateiname wird ab sofort zum Prim�rschl�ssel, daf�r gibt es kein Pr�fenD mehr in fotosmdb.exe
'           Die Spalte Jahr beim SQL Server ist ab sofort nvarchar(4)
'
'           ExportForm/ImportForm: Die Frage, ob es bei einer SQL-Server-Datenbank Import per Drag&Drop geben soll, wird vertagt
'               vorl�ufig wird beim SQL-Server die Funktion unsichtbar gemacht
'
'           Es gibt bei Access f�nf Tabellen: Fotos FET SpaltenBreite Temp_Haken ErsterStart
'           Es gibt bei sql server sieben Tabellen: LicenseCode LoggedInUsers Fotos FET SpaltenBreite Temp_Haken ErsterStart
'           Die Tabellen: Temp_Haken ErsterStart werden nur in Fotosmdb gebraucht
'
'           M�llentfernung: Pr�fenD, Pr�fen4 und Pr�fen5 entf�llt
'               Das Umnennen der Tabellen Fotos <-> EFotos ist v�llig �berfl�ssig, es reicht die Spalten umzunennen
'               tempor�r scharf/unscharf entf�llt
'               tempor�r Helligkeit �ndern entf�llt
'
'           Anmelden von einem anderen PC an die SQL-Server-Datenbank mit Windows Authentication
'           geht nur dann, wenn auf allen Rechnern gleichlautende Benutzernamen und Password angelegt sind
'           Es muss ein Nutzername mit Administratorrechten sein
'
'           Nur bei der Access-Shareware-Version ist es n�tig, da� beim ersten Start von fotos.exe Language = "9" ist
'           nur dann wird msdmo.log erzeugt
'           mit Hilfe des Alters von msdmo.log nerve ich die Shareware-Nutzer mit Einblendung des Shareware-Hinweises
'           Das Datum 30.12.2011 ist das Datum der Fotos.mdb im Auslieferungszustand
'           Die wird ins Installationspaket aufgenommen als fotos.mdeutsch.Auslieferung.mdb
'
'           MDF und LDF sind ab sofort keine erlaubten Dateitypen
'
'29.12.2011 13.4.0 Korrektur Professional Version:
'           ab sofort kann man den txtSQLGespeicherteAbfrage immer editieren ohne erst ein H�kchen setzen zu m�ssen
'29.12.2011 13.4.0 Uralt-Fehler korrigiert:
'           ab sofort ist Query.CheckWeitereFilterAktiv.Enabled = False und Query.CheckNutzerdefinierteFelder.Enabled = False
'           sonst kam bei Query.CheckDifferenzen H�kchen rausnehmen immer wieder angeblich aktivierte
'           Query.CheckWeitereFilterAktiv.Value = 1 und Query.CheckNutzerdefinierteFelder.Value = 1
'29.12.2011 13.4.0 Seltsam Seit Version 13.3.22
'           ist die Installation von Office nicht mehr Voraussetzung. Vermutlich hatte ich �bersehen, da� ein nichtbenutzter
'           Verweis auf Microsoft Access xx.x Object Library(msacc.olb) im Projekt enthalten war
'08.02.2012 13.4.0 Verbesserung Fotosmdb kann ab jetzt wenn gew�nscht in der fotos.mdb nur alle *.jpg Fotos l�schen
'10.02.2012 13.4.0 Verbesserung m�glichst alle Formulare sollen das G-icon anzeigen (BorderStyle=1)
'14.02.2012 13.4.0 Verbesserung Im Windows 7 sieht die Standardschrift Ms Sans Serif schei�e aus. Ersetzt durch "Arial" aus Form1.txtFont.FontName
'           ebenso DbGridForm.DBGridNeu.FontName=Arial
'----------------------------------------------------------------------------------------------------------
'15.02.2012 13.4.1 Fehlerkorrektur alle Versionen:
'           Suche nach NULL hat nicht funktioniert falsche SQL-Anweisung es fehlte Klammer zu
'15.02.2012 13.4.1 nicht l�sbares Problem
'           KommentarForm.RichTextBox1.Font.Name="Arial" hat keine Auswirkung
'           Das w�rde erst wirksam nach RichTextBox1.TextRTF = DbGridForm.Adodc1.Recordset("Kommentar")
'           Folglich - Das Formatieren muss der Nutzer selber machen
'03.03.2012 13.4.1 Verbesserung alle Versionen:
'           Mit Tastenkombination Num+Strg+N einschalten der Anzeige eine Zeile Bildbeschreibung ganz oben bei jedem Foto/Video
'           Mit Tastenkombination Num+Strg+M ausschalten der Anzeige eine Zeile Bildbeschreibung ganz oben bei jedem Foto/Video
'04.03.2012 13.4.1 Fehlerkorrektur:
'           Man kann mich austricksen und aus einer Shareware-Version eine SQL-Server-Version machen bei fotos.mdbnichtda
'           frmConnectSQL darf bei gblnProversion=False nicht erscheinen
'           Der Programmabschnitt zur Bestimmung welche Version vorliegt muss vor Call SpracheFestlegen verschoben werden
'----------------------------------------------------------------------------------------------------------
'04.03.2012 13.4.2 Verbesserung:
'           nur bei #if Proversion gibt es ein Formular frmConnectSql, sonst wird es bei der Compilierung weggelassen
'05.03.2012 in den Eigenschaften der .exe soll erkennbar sein, ob Proversion=0 oder =-1
'           ich trage ein bei Projekteigenschaften -> Erstellen -> Copyright -> GERBING Software Chemnitz -1 oder 0
'05.03.2012 in Version 13.4.1 war Pr�fung auf Datum 30.12.2011 unwirksam
'----------------------------------------------------------------------------------------------------------
'11.03.2012 13.4.3 Verbesserung zum 03.03.2012
'           Form1.txtBildBeschreibung.Visible = False gleich in der Entwicklungsumgebung eintragen
'22.03.2012 13.4.3 Formale Korrektur:
'           In ExportForm waren 2 Buttons nach oben verschoben btnAbbrechen btnHilfe
'----------------------------------------------------------------------------------------------------------
'29.03.2012 13.5.0 Verbesserung
'           Bildzeichnen und Zoom erfolgt mit GDIPlus antialiasing, GDI+ ist Bestandteil des Betriebssystems seit XP
'           2 neue native Dateitypen PNG TIF, aber CUR gestrichen
'           Benutzung einer frmBildMitGDIPlus f�r die Bilder
'           frmBildMitGDIPlus wird entladen und neu geladen bei jedem Bildwechsel, Verkleinern, Vergr��ern, Rechtecklupe, Bildverschieben
'           Das Entladen der Form frmBildMitGDIPlus mit Unload Me und anschlie�ende Neuladen aus Form1 heraus ist n�tig,
'           weil ich keinen anderen Weg gefunden habe die �berreste eines gezeichneten Bildes zu l�schen bevor ein neues Bild
'           gezeichnet wird. Wenn ich das nicht mache, �bermalen neue Bilder schon gezeichnete Bilder.
'           Die Videos werden auf Form1 angezeigt wie bisher
'           Form1 muss schwarzen Hintergrund erhalten und so gro� sein wie frmBildMitGDIPlus
'           Wenn das Schlie�kreuz der MDIForm (oder ein Aufruf 'Unload Formx' aus einer anderen Form)
'           unwirksam ist, steht in der Form 'cancel = True'
'           Damit nicht 2 Forms in der Taskbar stehen, setze ich in der Entwicklungsumgebung keine in die Taskbar
'           Form1.Borderstyle = 5 �nderbares Werkzeugfenster, Form erscheint nicht in der Taskbar
'           frmBildMitGDIPlus.Borderstyle = 5 �nderbares Werkzeugfenster, Form erscheint nicht in der Taskbar
'           jetzt brauche ich aber eine Ersatzl�sung, damit bei Bedarf doch die frmBildMitGDIPlus in der Taskbar gezeigt wird
'           das ist Query.FormInTaskbar
'           �berraschung im Windows8 Query.FormInTaskbar scheint garnicht zu reagieren siehe 09.05.2012
'
'           DbGridForm gibt nur F1 F2 F3 F4 F8 weiter
'           F5MehrereZeilen gibt keine Tastendr�cke weiter an Form1.Form_KeyDown - MarkiertenTextInZwischenAblageStellen wird nicht gebraucht
'           DbGridForm zeigt keine Hilfebox
'           KommentarForm zeigt keine Hilfebox
'           ImportForm und ExportForm werden ab sofort modal geladen
'           WertxForm wird ab sofort modal geladen
'           F5MehrereZeilen wird ab sofort modal geladen
'           anstelle des bisherigen Timer1 f�r automatische Diashow gibt es jetzt Timer1Ersatz mit API Methoden
'
'           Eigenartiges Fehlerbild gekl�rt: Warum springt es ohne Grund aus einer Prozedur heraus und arbeitet bei der Prozedur weiter,
'           aus der der Aufruf gekommen ist.
'           Beispiel 'If PublicPlayVideosWith <> 10 Then' Ursache ist weil das erste ist String das zweite ist Long
'           richtig ist 'If PublicPlayVideosWith <> "10" Then'
'
'           am 21.04.2012
'           GDI+ ThumbnailAnzeigen funktioniert manchmal nicht. Reproduzierbar mit SQL Server Version aus der Entwicklungsumgebung heraus,
'           genauso wie mit der Access-Version, wenn
'           die Fotos und die fotos.exe auf verschiedenen Laufwerken stehen
'           GdipLoadImageFromFile liefert rc = 3
'           reproduzierbar mit XP Win7 Win8
'           L�sung: f�r die Thumbnailanzeige in DbGridForm wird die herk�mmliche Methode benutzt, PNG und TIF sind nicht m�glich
'06.05.2012 13.5.0 kosmetische Korrekturen
'           Wenn die Checkbox Query.CheckDifferenzen angeklickt wird muss die Checkbox Query.CheckUseAudioComments ausgehen
'           Tooltip zu Query.CheckUseAudioComments muss korrigiert werden
'               zu einer Foto-Datei kann eine gleichnamige Audio-Datei aufgenommen oder abgespielt werden 'aufgenommen' wird gestrichen
'09.05.2012 13.5.0 Gewissensfrage was sieht besser aus, wenn mit Fenstergr��e �nderbar gearbeitet wird
'           schwarze R�nder an beiden Seiten, weil die 150 Pixel gro�e Taskbar vom Windows7 ber�cksichtigt wird
'           oder vom unteren Rand des Bildes fehlt ein kleines St�ck, da muss man das Bild nach oben schieben
'09.05.2012 13.5.0 Gewissensfrage was sieht besser aus, wenn mit Fenstergr��e �nderbar gearbeitet wird
'           Fenstergr��e Form1 halbieren oder fast bis an den Rand reichen lassen
'09.05.2012 13.5.0
'           nicht reproduzierbar im Windows7
'           Bei Fenstergr��e �nderbar bekommt man manchmal kein Icon in der Taskleiste zu sehen.
'           Bei Fenstergr��e nicht �nderbar muss man ohnehin erst die Windows-Taste dr�cken
'           In beiden F�llen findet man aber das Fotoalbum-Icon durch Bl�ttern mit der Taste Alt+Tab
'           Im Windows8 scheint Query.FormInTaskbar garnicht zu reagieren, die Taskleiste verschwindet �berhaupt nicht
'               alle .Borderstyle = auskommentiert, ohnehin nur read-only at runtime
'               Im Windows8 scheint trotz Borderstyle = 5 das Fenster trotzdem angezeigt zu werden
'           Damit beim Bildwechsel der schwarze Hintergrund erhalten bleibt, darf Form1 nie verschwinden
'           Wenn mit Drag&Drop gearbeitet werden soll, wird unbedingt ein Icon in der Taskleiste gebraucht
'               Query.FormInTaskbar wird nach Module1 verlagert
'               FormInTaskbar muss mit dem window handle aufgerufen werden, das sofort nach Form_Load abgespeichert wurde
'           Im Windows8 geht Import mit Drag&Drop nur, wenn 'als Administrator ausf�hren' benutzt wird
'----------------------------------------------------------------------------------------------------------
'21.05.2012 13.5.1 Verbesserung
'           Eine frmSchwarz mit schwarzem Hintergrund wird als erste Form geladen und bleibt geladen
'           weil sonst beim Bildwechsel das Hintergrungbild durchflackert
'           frmSchwarz wird in den Aufruf von FormInTaskbar einbezogen
'           FormInTaskbar enth�lt LockWindowUpdate das darf im Windows8 nicht aufgerufen werden
'           verhindern da� bei externem Videoplayer DbGridform durchflackert
'           Vor jedem Foto/Video wird anhand von Query.chkFensterGr��e�nderbar.Value eingestellt, ob
'           in Form1.Caption etwas steht oder nicht
'30.05.2012 13.5.1 Fehlerkorrektur
'           Bisher wird PublicZoomToFullscreen verz�gert gesetzt. Erst beim N�chsten �ffnen von WertxForm
'30.05.2012 13.5.1 Fehlerkorrektur
'           betrifft meine private fotos.mdb
'           Nach Kommentar-�nderung (viel Text mehr als 2000 Bytes) kommt manchmal runtime error (80004005) Fehler beim Auswerten der
'           CHECK-Beschr�nkung. Ich habe eine G�ltigkeitsregel im Feld Kommentar, wegen �bernahme in IPTC
'           ich muss Fehler abfragen -> Msgbox bringen -> weiterarbeiten
'06.06.2012 13.5.1 Fehlerkorrektur
'           Das passiert nur bei externer Videoplayer oder privater externer Videoplayer.
'           Ich will verhindern, dass bei F12 schwarzer Hintergrund kommt. Das aktuelle Bild verschwindet.
'07.06.2012 13.5.1 Fehlerkorrektur
'           Im Windows8 flackert es weniger durch diese Korrektur
'           ----------------
'           nur in Version 13.5.1
'           Bei einem H�kchen in Fenstergr��e �nderbar fehlt ein Taskbar-Icon. Damit wird Drag&Drop erschwert. Man muss die richtigen
'           Fenster anfassen und selber so ziehen das sie unterscheidbar sind. Fenster ziehen nach rechts geht mit Anfassen an
'           oberer linker Ecke.
'==========================================================================================================
'           Version 13.5.1 ist die letzte, die XP unterst�tzt,
'           wegen der Watze, da� nach etwa 100 mal F3 klicken ein schwarzer Screen kommt, wo dann nur noch verkleinerte Bilder angezeigt
'           werden
'==========================================================================================================
'10.06.2012 13.5.2 Verbesserung
'           Es ist mir gelungen, mit dauerhaft geladener Form frmBildMitGDIPlus zu arbeiten.
'           Anstatt auf die Form wird mit Picture1 gezeichnet.
'           Damit ist der Hauptgrund f�r das Flackern beim Bildwechsel beseitigt.
'           Jetzt gibt es ein neues Problem: Sobald eine andere Form �ber dem GDIPlus Bild liegt, wird dieser
'           Teil vom GDIPlus Bild unsichtbar(schwarz) und muss neu gezeichnet werden. Zu diesem Zweck wird in frmBildMitGDIPlus
'           ein neuer Timer TimerRefresh eingef�hrt mit Timer intervall = 100 Millisekunden.
'           FormInTaskbar wird nur noch f�r Query benutzt.
'           frmSchwarz entf�llt.
'           XYPos hat Quatsch angezeigt bei der Position des Mauszeigers.
'           Alt + Pfeil nach links war wirkungslos
'           Die Arbeit mit dem Timer funktioniert nicht im XP. Im XP legt sich bei eingeschaltetem Timer das GDIPlus Bild �ber alle anderen
'           Fenster, auch die Fenster anderer Anwendungen.
'----------------------------------------------------------------------------------------------------------
'12.06.2012 13.5.3 Verbesserung
'           frmBildMitGDIPlus ist nicht n�tig, alles in Form1 machen. Den TimerRefresh brauche ich trotzdem.
'           Form1.Borderstyle = 2 ist n�tig, damit die Form in der Taskleiste erscheint.
'           ----------------
'           seit Version 13.5.3
'           Bei Videos mit internem Mediaplayer auf Form1 fehlen die Bedienelemente
'----------------------------------------------------------------------------------------------------------
'16.06.2012 13.5.4 Fehlerkorrektur
'           Videos nicht auf Form1 abspielen sondern frmVideo, da sind die Bedienelemente wieder da
'           frmVideo.Form_Resize und form1.Form_Resize wird gestrichen. Bei Bedarf soll der Nutzer selber seine Fenster ziehen.
'           gblnComeFromVideo wird gebraucht, weil nach Tasten dr�cken auf frmVideo die frmVideo wieder gezeigt werden muss
'16.06.2012 13.5.4 Verbesserung
'           Wenn die Form Query gezeigt wird, soll ein Icon in der Taskleiste erscheinen. Das geschieht bisher nur nach Programmstart.
'           Ab jetzt auch nach Taste F8
'03.07.2012 13.5.4 Fehlerkorrektur
'           damit korrigiere ich den Fehler, dass beim ersten Video nur ein senkrechter wei�er Strich zu sehen ist
'           Die Breite/H�he eines Videos wird nicht mehr von frmvideo.wmp.currentMedia.imageSourceWidth bzw
'           frmvideo.wmp.currentMedia.imageSourceHeight entnommen. Problematisch, weil das Video erst im playstate=3 playing sein muss,
'           sondern wie schon bisher bei Benutzung eines externen mediaplayers �ber
'           MM.extractDefaultMovieSize(wancho, walto)
'04.07.2012 13.5.4 Fehlerkorrektur
'           Bei Query.mnuEinstellungen_Click kam nach Schlie�en der WertxForm bisher sofort das erste Bild.
'           Abhilfe - man muss den Form1.TimerRefresh disablen
'26.07.2012 13.5.4 Fehlerkorrektur
'           Es funktionierte nicht - Weiterselektieren nur die mit Merkerspalte markierten Dateien anzeigen
'11.08.2012 13.5.4 Fehlerkorrektur
'           Das Bild sitzt bisher stets unterhalb von txtBildBeschreibung, auch wenn gar keine Bildbeschreibung gew�nscht ist
'11.08.2012 13.5.4 Fehlerkorrektur
'           Bisher lief die Automatik, nachdem sie einmal mit zuf�llig ausgew�hlt war, auch bei Widerholungen �ber F12 (wenn dort steht
'           aufsteigend) trotzdem mit zuf�llig weiter
'11.08.2012 13.5.4 Fehlerkorrektur
'           Form1.Picture1.BackColor muss Schwarz sein, sonst sieht man einen grauen Fleck, wenn Form Query verschoben wird
'11.08.2012 13.5.4 Fehlerkorrektur
'           in Prozedur Timer1Ersatz kam manchmal runtime error 402 modales Fenster kann nicht .... wenn .....
'11.08.2012 13.5.4 Fehlerkorrektur
'           Wenn die Hilfebx oben ist, funktionieren zwar alle Bedienfunktionen mit der Maus, aber nicht alle mit der Entsprechenden Taste
'11.08.2012 13.5.4 Fehlerkorrektur
'           In Form Query war nach Version 13.5.1 MinButton=True verschwunden
'27.08.2012 13.5.4 Fehlerkorrektur
'           Kooperationsfehler zwischen Fotos.exe und Renammdb.exe
'           Wenn aus Fotos.exe heraus Renammdb.exe aufgerufen wird und der in Fotos.exe gerade aktuelle Dateiname ge�ndert oder gel�scht werden soll
'           kommt errornumber = 75 Fehler beim Zugriff auf Pfad/Datei beim �ndern
'           kommt errornumber = 70 Zugriff verweigert beim L�schen
'           Die einzige fehlerfreie L�sung die ich gefunden habe, besteht darin nach dem Aufruf von RenamMdb.exe die Fotos.exe zu beenden
'           und es geht nur, wenn die Entwicklungsumgebung von fotos.exe beendet worden ist
'04.09.2012 13.5.4 Verbesserung
'           Ich kann vermeiden, f�r XP die Version 13.5.1 als letzte unterst�tzte Version auszuliefern, das Problem kommt
'           wegen der Wirkungsweise von Form1.TimerRefresh.
'           Im XP legt sich bei eingeschaltetem Timer das GDIPlus Bild �ber alle anderen Fenster, auch die Fenster anderer Anwendungen.
'           Wenn ich im XP arbeite, soll grunds�tzlich chkFensterGr��e�nderbar.Value = 1 sein
'           und ich will erzwingen, da� Form1.Controlbox = True ist und es einen MinButton und einen MaxButton gibt, dann kann der Nutzer selber
'           die fotos.exe minimieren und sein gew�nschtes Fenster wieder aktivieren.
'           Dazu dient die Function ShowTitleBar, die wahlweise f�r Fotos oder Videos aufgerufen wird.
'           Es war n�tig da� in frmVideo Controlbox = True ist und es einen MinButton und einen MaxButton gibt und BorderStyle=2
'           FormInTaskBar wird ganz entfernt
'           Wenn die Taskleiste automatisch ausgeblendet wird, soll das Bild bis ganz unten hin gehen. Das muss bei jedem
'           BildAnzeigen gepr�ft werden.
'04.09.2012 13.5.4 Fehlerkorrektur
'           Beim Import mit Drag&Drop kam Fehler 3265 'element in dieser Auflistung nicht gefunden' wenn $fotos.mdb und fotos.mdb
'           nicht in derselben Sprache sind
'05.09.2012 13.5.4 Fehlerkorrektur
'           bisher kam Laufzeitfehler '6' �berlauf wenn man die Taste F4 zum Vergr��ern festh�lt. Das wird ab sofort ignoriert.
'17.09.2012 13.5.4 Fehlerkorrektur
'           Im Kommentarfenster konnte bisher kein H�kchen gesetzt werden bei 'Kommentar soll editiert werden, weil Form1.TimerRefresh zu schnell
'           dem Kommentarfenster den Focus entzieht. Wenn der Nutzer sehr schnell mit der Maus ist, geht es
'19.09.2012 13.5.4 Fehlerkorrektur
'           Bisher hat das Anzeigen der Aboutform gehunzt, wenn der Nutzer nach Dr�cken von F8 �ber die Men�leiste '�ber...' ausgew�hlt hat.
'           Es wurde das zuletzt aktuell gewesene Bild gezeigt.
'19.09.2012 13.5.4 Fehlerkorrektur
'           Bisher wurde das erste Video nicht gezeigt, wenn in der Datenbank nur Videos sind
'27.09.2012 13.5.4 Fehlerkorrektur
'           Bisher wurde bei 'Mit diesen Such-Kriterien wurde kein einziger Datensatz gefunden' das zuletzt aktiv gewesene Bild angezeigt
'           Bei F5MehrereZeilen wurde die BackColor ge�ndert
'04.10.2012 13.5.4 Fehlerkorrektur
'           Der Dateiname '1960-Wandertag Dieter Knopf, Irmscher, Ullrich Krausse, Guenter Jacob(v. l.).jpg'
'           wird als 5-stellige Dateinamen-Erweiterung erkannt, die vom Programm als Link-Dateityp behandelt wird
'23.10.2012 13.5.4 Fehlerkorrektur
'           Wenn ein Bild einen Kommentar hatte, dann blieb dieser im Kommentarfenster, auch wenn das Bild keinen Kommentar enthielt
'           Das Editieren im KommentarFenster war schwierig, wenn das KommentarFenster von selber aufging, weil KommentarFensterEinblenden = True
'           Die neue L�sung schlie�t die KommentarForm grunds�tzlich, wenn sie nicht mehr gebraucht wird.
'           Es ist nicht mehr m�glich, da� das Kommentarfenster von selber aufgeht, wenn es einen Kommentar gibt.
'27.10.2012 13.5.4 Fehlerkorrektur
'           Bisher wird bei Num + F5 im Formular F5MehrereZeilen der Datensatz angezeigt, der im Grid aktuell ist. Das kann ein ganz anderer
'           Datensatz sein, als der, von dem das aktuelle Bild gezeigt wird. Bisher konnte der Nutzer sich nur so helfen, da� er vor der
'           Benutzung von Num + F5 die aktuelle Zeile im Grid doppelgeklickt hat.
'           Neue L�sung:
'           Die Prozedur Form1.FelderAusf�llenF5MehrereZeilen reduziert sich auf F5MehrereZeilen.chkIptc_Click
'           Das aktuell gezeigte Bild hat Priorit�t. Die aktuelle Zeile im Grid wird ignoriert.
'           Ich muss DbGridForm.rsDataGrid mit Find durchsuchen um das aktuell gezeigte Bild zu finden.
'27.10.2012 13.5.4 Fehlerkorrektur
'           Bisher wird bei F10 im Formular KommentarForm der Datensatz angezeigt, der im Grid aktuell ist. Das kann ein ganz anderer
'           Datensatz sein, als der, von dem das aktuelle Bild gezeigt wird. Bisher konnte der Nutzer sich nur so helfen, da� er vor der
'           Benutzung von F10 die aktuelle Zeile im Grid doppelgeklickt hat.
'           Neue L�sung:
'           Das aktuell gezeigte Bild hat Priorit�t. Die aktuelle Zeile im Grid wird ignoriert.
'           Ich muss DbGridForm.rsDataGrid mit Find durchsuchen um das aktuell gezeigte Bild zu finden.
'06.11.2012 13.5.4 Fehlerkorrektur
'           Es gab eine Situation da bekam der Nutzer einen schwarzen Screen zu sehen.
'           F5 dr�cken -> �ffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist. Dann die markierte Datei in ein anderes Explorer-Fenster
'           verschieben. Dann auf das Icon vom Fotoalbum klicken.
'08.11.2012 13.5.4 Verbesserung
'           Kolossale Beschleunigung beim Start wenn ich die Inhalte der Comboboxen
'           TSituation TOrt TLand TPersonen
'           an die Auswahl eines Jahres oder Jahresbereiches anpasse. Das ist gleichzeitig viel logischer.
'           Als Vorgabe stelle ich das h�chste in der Datenbank vorkommende Jahr in die TextBox TJahr
'           Immer wenn der Nutzer ein anderes Jahr w�hlt, muss sich der Inhalt der ComboBoxen �ndern
'           Immer wenn der Nutzer einen Wert der ComboBoxen TSituation TOrt TLand TPersonen ausw�hlt, muss sich der Inhalt der
'           anderen ComboBoxen anpassen
'           SendKeyStroke vbKeyEnd, -1      'das muss ich machen weil sonst der Cursor links von der Eingabe in der ComboBox steht,
'           ich will ihn aber rechts
'15.11.2012 13.5.4 Fehlerkorrektur
'           Manchmal kam ein nichtreproduzierbarer Fehler '401' nicht-modales Formular kann nicht gezeigt werden, w�hrend ein modales Formular
'           gezeigt wird, wenn die erste Datei ein Video war und F5 gedr�ckt wurde
'           Daraufhin habe ich in Form1 ge�ndert von Query.Show 1 in Query.Show
'           Dann habe ich alles was nach Query.Show 1 stand ans Ende von Query_btnOK geschoben
'           Dann habe ich in DbGridForm.Form_Paint am Anfang eingef�gt: On Error Resume Next
'15.11.2012 13.5.4 Fehlerkorrektur
'           Ab jetzt mache ich Unload frmVideo zu Beginn von BildAnzeigen, weil der Nutzer sonst beide Formen auf der Taskbar sieht und
'           anklicken kann, wenn ich starte mit H�kchen bei 'Fenstergr��e �nderbar'
'----------------------------------------------------------------------------------------------------------
'19.11.2012 13.5.5 Verbesserung SQL Server Version
'           Zum P�fen der ersten Kolonne des LicenseCode wird nicht mehr der Name, sondern die mittlere Kolonne benutzt
'21.11.2012 13.5.5 Verbesserung SQL Server Version
'           Die bisherige Verschl�sselung der Zahl der Lizenzen ist zu leicht zu knacken durch Probieren
'           Ich verschl�ssele jetzt die Zahl an zwei Positionen
'           bisher SQL99
'           jetzt  99S99 und in der Mitte bleibt ein S stehen
'26.11.2012 13.5.5 Verbesserung
'           frmVideo.WMP.stretchToFit = True und gleichzeitiges Maximieren des WMP-Fensters f�hrt dazu, da� alle Videos, die kleiner sind als das
'           WMP-Fenster unter Beibehaltung von aspect radio ans Maximum angepa�t werden.
'           Videos die gr��er sind als das WMP-Fenster, werden ebenfalls angepa�t und demzufolge etwas kleiner.
'           Umschalten auf fullscreen mit Mausbedienung geht trotzdem noch, nach Ende des Video wird fullsceen aber wieder ausgeschaltet.
'           Die Video Controls liegen etwas h�her als die Taskleiste. Durch Rechtsklick auf den Zwischenraum zwischen video controls
'           und task bar bekommt man das Hilfe-Fenster
'26.11.2012 13.5.5 Verbesserung
'           Ab sofort wird Rechtsklick auf DbGridForm.Picture1 und DbGridForm.DbGridNeu zugelassen
'           Wenn n�mlich die Videos die ganze frmVideo ausf�llen, funktioniert kein Rechtsklick
'26.11.2012 13.5.5 Verbesserung
'           Ab sofort wird Rechtsklick auf frmVideo.WMP zugelassen
'           Wenn n�mlich die Videos die ganze frmVideo ausf�llen, funktioniert kein Rechtklick. Ab jetzt kommt beim ersten Rechtsklick
'           die Hilfe, beim zweiten Rechtsklick das Rechtsklick-Men� des Mediaplayers
'27.11.2012 13.5.5 Verbesserung
'           Bisher blieb der Zoom-Mauszeiger eingeschaltet, auch wenn zwischendurch eine Fx-Taste gedr�ckt wurde.
'           Ich will sperren, da� beim gleichen gstrFRODN wiederholt die Rechtecklupe benutzt werden kann. Da klappt n�mlich die Anzeige nicht.
'           Jeder Bildwechsel setzt die Sperre zur�ck
'27.11.2012 13.5.5 Fehlerkorrektur
'           beim Sprache wechseln blieb er h�ngen
'30.11.2012 13.5.5 Verbesserung
'           Ich will da� bei Tastendruck F5 die aktuelle Zeile komplett schwarz markiert wird
'04.12.2012 13.5.5 Verbesserung
'           Gro�er Sieg - Wenn Form1.TimerRefresh dran war, wird er als erstes ausgeschaltet, so vermeide ich st�ndiges Bildflackern.
'           Ganz verzichten kann ich auf Form1.TimerRefresh nicht. Wenn ich beispielsweise die Form Hilfebx schlie�e und in der Unload-Prozedur
'           das Bild neu zeichne, wird zwar das Bild korrekt neu gezeichnet, aber die Form Hilfebx verschwindet erst mit Exit Sub und
'           hinterl��t eine schwarze Fl�che.
'           Form1.TimerRefresh.Interval wird auf 1 gesetzt
'31.12.2012 13.5.5 Verbesserung
'           Gro�er Sieg - Ich habe gefunden, warum im XP bei etwa 100 Bildern nacheinander anzeigen Feierabend ist,
'           und warum nicht funktioniert: L�schen der mit Merkerspalte markierten Dateien aus Datenbank und Ordner
'           Version 13.5.1 ist die letzte, die XP unterst�tzt,
'           wegen der Watze, da� nach etwa 100 mal F3 klicken ein schwarzer Screen kommt, wo dann nur noch verkleinerte Bilder angezeigt
'           werden
'           Es liegt daran da� alle angeguckten Fotos als ge�ffnete Datei stehenbleiben. Das k�nnen Hunderte sein.
'           L�sung: Ich mu� GdipDisposeImage und GdipDeleteGraphics und GdipCreateFromHDC in der richtigen Reihenfolge benutzen
'31.12.2012 13.5.5 Verbesserung
'           Wenn ein Video l�uft und das Programm soll �ber Schlie�kreuz beendet werden, muss man bisher zweimal klicken
'           L�sung: in FrmVideo Form_Unload abschreiben von Form1.Form_Unload
'15.01.2013 13.5.5 Verbesserung
'           Bisher konnte die Form nicht minimiert werden, wenn gerade ein Video lief
'           Bisher hat ein Klick auf das G-Icon in der Taskleiste manchmal das Bild gezeigt, manchmal ein schwarzes Fenster
'15.01.2013 13.5.5 Verbesserung
'           Ab jetzt vermeide ich das Meckern �ber nichtkorrekte Jahreszahl f�r das Auff�llen der ComboBoxen
'18.01.2013 13.5.5 Fehlerkorrektur
'           Bisher gab es einen Fehler bei Fehlen des Abschnitts [Mediaplayer] in der fotos.ini
'           Es wurde angenommen Einstellungen Videos 'play videos with other external video player'
'           und beim versuchten Video Abspielen passierte garnichts
'04.02.2013 13.5.5 Fehlerkorrektur
'           frmVideos geht nicht bis zum rechten Rand
'           Ein Video mit 1920 x 1080 Pixel hatte auch bei 'Fenstergr��e unver�nderbar' rechts und links einen schwarzen Rand
'05.02.2013 13.5.5 Fehlerkorrektur
'           Es kam Laufzeitfehler '5' bei der Auswahl von mehreren Personen
'09.02.2013 13.5.5 Verbesserung Professional Version
'           Ich will parallel zu 'erster Treffer pro Jahr' eine Option 'ein Zufallstreffer pro Jahr' erfinden
'           Bei der Option 'ein Zufallstreffer pro Jahr' kann man die Suche beliebig oft wiederholen und bekommt immer ein neues
'           zuf�lliges Bild pro Jahr, vorausgesetzt es gibt pro Jahr mehr als ein Bild
'           Wenn diese Funktion auf alle S�tze der Datenbank angewendet wird, dauert es sehr lange.
'           Das mache ich sichtbar in Query.txtArbeitsfortschritt
'           Man sollte also eine Person oder eine Situation oder einen Ort ausw�hlen
'11.02.2013 Problem seit Version 13.5.5 im XP (ist zwar nur bis Version 13.5.1 im XP vorgesehen)
'           Bei Videos mit dem internen mediaplayer kontinuierlich abspielen mit Strg+F6 st�rzt vb6.exe ab
'           Die Anweisung ...... verweist auf Speicher in ...... Der Vorgang konnte nicht auf dem Speicher durchgef�hrt werden.
'           Das letzte was geht ist WMP_PlayStateChange bis zum Ende der Prozedur
'           Videos kontinuierlich mit dem externen Mediaplayer abspielen funktioniert jedoch
'           L�sung:
'           Prozedur WMP_PlayStateChange wird gestrichen
'           TimerVideoDurationTimer wird jetzt f�r externen und internen Mediaplayer benutzt
'15.02.2013 13.5.5 Fehlerkorrektur
'           Wenn Hilfebx Form sichtbar war, hat bisher ein weiterer Rechtsklick (meist versehentlich) manchmal die
'           Hilfebx Form wieder verschwinden lassen.
'           Ab sofort muss es unbedingt ein Linksklick auf einen Label der Hilfebx Form sein
'16.02.2013 13.5.5 Fehlerkorrektur zum 09.02.2013
'           Wenn ein Feld ge�ndert wurde und die Korrektur wurde nachgef�hrt, dann wurde daraufhin eine neue Auswahl mit
'           'ein Zufallstreffer pro Jahr' bereitgestellt
'           L�sung: keine neue Suche ausf�hren, sondern erneut die vorliegenden Tabellen Fotos mit FET Inner Join verkn�pfen
'20.02.2013 13.5.5 Fehlerkorrektur
'           Die Aufhebung der Rechtecklupe soll wieder das Bild zeigen genauso wie beim ersten Laden
'04.03.2013 13.5.5 Fehlerkorrektur
'           Ab jetzt wieder richtige Reaktion bei KommentarFensterEinblenden = True
'04.03.2013 13.5.5 Fehlerkorrektur
'           Wenn ich mit Pfeiltasten im DBGridNeu auf/abbl�ttere soll die aktuelle Zeile schwarz werden
'==========================================================================================================
'04.03.2013 Neue Funktion 14.0.0
'           �berraschung: Das DataGrid msdatgrd.ocx ist unicode f�hig, Ms Access vermutlich schon lange
'           Unicode-Unterst�tzung durch die Timosoft Controls und durch FSO
'               geht nicht im XP: Diashow.exe st�rzt ab und auch IDE st�rzt ab beim Schlie�en des Programms, vermutlich weil
'               bei Diashow.Form_Unload set fso=Nothing und das Unload f�r alle Forms gefehlt hat
'               Geht im Win7
'           Zum �ndern von FontSize muss die Eigenschaft des Timosoft Controls UseSystemFont = False sein
'           Viele Events bei den Timosoft Controls sind standardm��ig disabled. Man muss im gezeichneten Control element rechtsklicken ->
'           Properties -> H�kchen rausnehmen
'           Drag&Drop als Target geht erst wenn man RegisterForOLEDragDrop = True setzt
'           Das Debuggen spinnt in DiashowForm bei eingeschaltetem Subclassing wenn man _OLEDragDrop debuggen will ->
'           probiere ob es mit Alt+F4 weitergeht
'           Die Form.Caption mit unicode sieht man nicht in der IDE sondern erst wenn man die exe startet
'           Alle Datei read/write Operationen f�r Text-Dateien sollte man mit FSO machen. Da muss man vorher testen ob der Dateiname
'               auch nur ein unicode Zeichen enth�lt oder alles ANSI ist.
'               Man muss daraufhin die Datei mit FSO entweder als unicode oder ANSI Datei �ffnen.
'               Alles auswechseln was wie 'Open Path For Binary Access Read As #Handle' aussieht.
'               Achtung bei FSO Gefunden bei Microsoft http://support.microsoft.com/kb/189751/en-us
'               Reads only ASCII data - while the FileSystemObject can create an ASCII or Unicode text file, the FileSystemObject can only
'               read ASCII text files.
'           Die scrrun.dll muss mit ausgeliefert werden Sie ist zust�ndig f�r FSO
'           Chronologie:
'           init_global bei Start des Programms
'           Form Query - alle Controls, die Abfragenamen oder Stichworte anzeigen k�nnten, werden ausgetauscht
'           Form MP - alle Controls, die Stichworte anzeigen k�nnten, werden ausgetauscht
'           Form ND - alle Controls die Feldnamen oder Stichworte  anzeigen k�nnten, werden ausgetauscht
'           Form F5MehrereZeilen - alle Text Controls die Feldnamen oder Stichworte  anzeigen k�nnten, werden ausgetauscht
'               RichTextBox ersetzen durch Timosoft Text Control - Konsequenz alle Kommentare mit Formatzeichen m�ssen editiert werden
'               iptcinfo.dll entf�llt das mach ich jetzt selber
'           Alle FileDateTime unicode f�hig machen durch FSO
'           Alle FileCopy esetzen durch file_copy
'           Alle Dir( ersetzen durch file_path_exist
'           Alle MkDir ersetzen durch CreateDirectoryW
'           LoadPicture ersetzen durch LoadPictureW, au�er die MouseIcons
'               F�r die MouseIcons wird die .res-Datei benutzt.
'               Hinein kommen sie mit dem VB-Ressourcen-Editor.
'               Heraus kommen sie mit zB Me.MouseIcon = LoadResPicture(105, 1) 105-Das ist Squarezoom.ico
'               Dadurch brauche ich die .ico-Dateien nicht mehr auszuliefern
'           INI file wird unicode f�hig durch schreiben mit FSO und Benutzen von GetPrivateProfileStringW und
'               WritePrivateProfileStringW
'           F�r command gibt es ein VBA replacement (overwrites VBAs Command$) to get unicode support in UnivbzGlobal.bas
'           F�r Kill gibt es ein VBA replacement for "Kill(PathName)" with UNICODE support in UnivbzGlobal.bas
'               besser file_delete                                                                                      'Gerbing 04.09.2013
'           F�r SetAttr gibt es ein VBA replacement for SetAttr, supports unicode and network in UnivbzGlobal.bas
'           Alle MsgBox wo file names vorkommen ersetzen durch MessageBoxW
'           GERBING Fotoalbum 13 ersetzen durch GERBING Fotoalbum 14
'           App.Path ersetzen durch getCurrentDir
'           App.Major App.Minor App.Revision ersetzen durch GetFileVersionInfo
'           chm-files lassen sich in unicode Pfad nicht �ffnen, das hat Microsoft nicht vorgesehen
'           ShellExecute ersetzen durch ShellExecuteW (RunShellExecute)
'           Der SQL Server verlangt f�r die Suche nach unicode ein N vor dem Suchbegriff
'               beispielsweise .Source = "select * from loggedinusers where (username = N'" & gstrLoggedInName & "')"
'               aber die fotos.mdb (d.h. Microsoft Access) versteht kein N'
'           Beim MSI-Paket zu beachten:
'               Die 3 ocxe von Timosoft binde ich ein als COM-Objekte - cblctlsu.ocx editctlsu.ocx ExLvwU.ocx
'               Die CLSID habe ich aus der vbp-Datei genommen
'           Damit im VM - Win7 - Benutzer mit Benutzerkonto - nicht der Fehler kommt 'DataGrid Control Cannot initialize bindings'
'               packe ich MSBIND.MSM und MSSTDFMT.MSM ins projekt
'
'
'16.04.2013 14.0.0 Professional Version
'           Gewaltige Beschleunigung bei 'Ein Zufallstreffer pro Jahr'
'29.04.2013 14.0.0 Fehlerkorrektur
'           Bisher habe ich rechts oben das mittlere Symbol gedr�ckt zum Maximieren des Fensters, diese Einstellung wurde
'           immer beim n�chsten Bild wieder r�ckg�ngig gemacht, ab jetzt nicht mehr
'29.04.2013 14.0.0 Fehlerkorrektur
'           Bisher hat das erste Bild viermal geflackert bis es endlich ruhig war. Ich habe jedes unn�tze Form1.Show abgeschaltet
'29.04.2013 14.0.0 Verbesserung
'           'Beim Bildladen vergr��ern auf Vollbild' neue Option 'immer'
'01.05.2013 14.0.0 Verbesserung
'           Es gibt Videos, bei denen FotosMdb nicht in der Lage ist, Video size/duration zu bestimmen. Die entsprechenden
'           Datenbankfelder bleiben dann leer. Ich kann jedoch in fotos.exe w�hrend des Abspielens mit dem internen Mediaplayer
'           diese Werte feststellen und in die Datenbank eintragen.
'           frmVideo.WMP_PlayStateChange wird wieder eingef�hrt
'06.05.2013 14.0.0 Fehlerkorrektur
'           Im Windows8 stelle ich 3 Sekunden R�delei zwischen zwei avi-Dateien mit dem internen mediaplayer abzuspielen fest.
'           Ich ermittle Videoduration nur, wenn kontinuierlich mit dem externen mediaplayer abgespielt werden soll
'07.05.2013 14.0.0 Fehlerkorrektur
'           Ab jetzt realisiere ich, dass w�hrend des Abspielens eines Video mit dem internen mediaplayer gesagt werden kann,
'           weiter Abspielen kontinuierlich Tasten Strg+F6
'07.05.2013 14.0.0 Fehlerkorrektur
'           Es gibt Funktionen, die bei Videos unwirksam sein sollen, wenn sie vom Rechsklick-Men� kommen
'11.05.2013 14.0.0 Verbesserung
'           Anstelle von msprivs.log benutze ich ab jetzt gerbingsoft.log - msplugin.log bleibt
'22.05.2013 14.0.0 Fehlerkorrektur
'           Bisher war die Tastenkombination Alt + F4 ins Foto oder laufende Video unwirksam, aber wirksam �ber Rechtsklick
'30.05.2013 14.0.0 Verbesserung
'           Weil die Kommentarform beim schlie�en entladen wird,
'           merke ich mir ab jetzt die Fensterpositionen/Breite/H�he der Kommentarform f�r die Dauer einer Sitzung
'           von fotos.exe
'31.05.2013 14.0.0 Fehlerkorrektur
'           Der Fehler trat auf bei Fotos/Videos angucken -> F8 -> Einstellungen -> Einstellungen schlie�en -> es wurde
'           das zuletzt aktiv gewesene Bild gezeigt, bei Video war es jedoch richtig
'03.06.2013 14.0.0 Fehlerkorrektur
'           Es gibt Bilder, die angeblich 18768 Felder EXIF-INformationen enthalten. Wenn die alle gezeigt werden sollen, l�uft sich
'           der Klassenmodul clsEXIF den Wolf.
'           Das kann nicht sein und wird deshalb auf maximal 200 EXIF-Felder begrenzt.
'04.06.2013 14.0.0 Verbesserung
'           �berarbeitung der EXIF-Felder, es werden jetzt GPS-Felder erkannt
'           und man kann in nutzerdefinierten Feldern danach suchen
'           Es ist n�tig bei Text-Feldern nicht nur die Vergleichsoperanden = und <> zuzulassen, sondern alle Vergleichsoperanden
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
'----------------------------------------------------------------------------------------------------------
'25.06.2013 14.0.1 Fehlerkorrektur/Verbesserung alle Versionen
'           Die Kombination
'           DbGridForm.Adodc1.Recordset.MoveLast        kann entfallen
'           DbGridForm.Adodc1.Recordset.MoveFirst       kann entfallen
'           Query.RecordCount = DbGridForm.Adodc1.Recordset.RecordCount
'           um den Recordcount zu ermitteln war ganz am Anfang von Access-Zeiten n�tig, kann jetzt entfallen, f�hrt nur unn�tig zum
'           Aufruf von DbGridForm.DbGridNeu_RowColChange
'25.06.2013 14.0.1 Verbesserung alle Versionen
'           CheckWeitereFilterAktiv und CheckNutzerdefinierteFelder werden wegen verbesserter Lesbarkeit erg�nzt um um die rotgef�rbten
'           Label lblWeitereFilterAktiv und lblNutzerdefinierteFelder
'05.08.2013 14.0.1 Verbesserung alle Versionen
'           Ich habe zu bem�ngeln, dass bei Videos bei F2/F3 kurz der Desktop oder das aktive Fenster durchkommt. Allerdings nicht von
'           Anfang an, sondern erst nach F8. Das liegt daran, dass nach F8 Form1.Hide drankommt.
'           Ich habe zu bem�ngeln, dass der Nutzer bei Videos immer noch beide Formen auf der Taskbar sieht und
'           anklicken kann, wenn ich starte mit H�kchen bei 'Fenstergr��e �nderbar'
'           Ich habe zu bem�ngeln, dass frmVideo ohne Icon und ohne Caption in der Taskbar steht.
'           Ich habe zu bem�ngeln, dass bei Fotos kontinuierlich immer erst kurz die Kommentarform durchkommt.
'           Ich habe zu bem�ngeln, da� Videos kontinuierlich abspielen mit dem internen player stehenbleibt beim Wechsel vom ersten auf
'           das zweite Video. Siehe 11.02.2013 es gibt jetzt wieder getrennte Timer f�r externen player und internen player
'07.08.2013 14.0.1 Verbesserung alle Versionen
'           Gel�stes Problem im Windows7
'           Shellexecute "print" geht nicht - genauso wie im Windows Explorer Rechtsklick auf eine Bilddatei -> Drucken nicht geht
'           Windows-Fotoanzeige reagiert nicht
'           L�sung siehe e:\Faq&Lehrmaterial\3 Mein Win7 PC Macken und L�sungen\Windows 7 Windows-Fotoanzeige druckt nicht.txt
'           Besere L�sung: Ich biete jetzt mit rundll32.exe einen Dialog an,
'           wo der Nutzer sein gew�nschtes Programm zum Drucken des Fotos aussuchen kann
'16.08.2013 14.0.1 Verbesserung alle Versionen
'           App.Major App.Minor App.Revision ersetzen durch GetFileVersionInfo
'           sonst kommt in einem unicode Pfad Laufzeitfehler 326 bei Ermitteln der Version der exe
'           Resource with identifier 'VERSION' not found (Error 326)
'04.09.2013 14.0.1 Verbesserung alle Versionen
'           Manchmal bleibt ein Foto blockiert f�r FotosMdb -> Iptc... Felder eintragen
'           Das passiert unter folgenden Bedingungen:
'           FotosMdb wird �ber Men� Tools aus fotos.exe heraus gestartet
'           Vorher ist schon mal mindestens ein Bild angeguckt worden und danach wurde Taste F8 gedr�ckt
'           Wenn das zuletzt angeguckte Bild mit IPTC-Feldern r�ckgeschrieben werden soll, dann ist es blockiert
'04.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           im unicode Pfad lie� sich die Sprache auf Englisch wechseln aber nicht zur�ck -> Laufzeitfehler Feld 'Ort' nicht gefunden
'           das trat nur bei meiner privaten fotos.mdb auf
'05.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Kill ersetzen durch file_delete
'           Name ... as ersetzen durch NameAs
'           war vergessen worden bei: Sprache �ndern, bei DBEngine.CompactDatabase
'10.09.2013 Fehlerkorrektur 14.0.1
'           Im unicode Pfad hunzt die Function SetFileTime bei
'           r�cksetzen fotos.mdb auf das Datum 30.12.2011 und bei Datei msdmo.log auf das Datum von heute - 100
'           es wird immer das aktuelle Datum eingetragen (Zeitpunkt des Programmablaufes)
'           L�sung: anstelle CreateFile verwenden CreateFileW
'26.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Zustand: wenn ein Video abgespielt wird und mehrere Formen sind offen und es wird das Schlie�kreuz geklickt,
'           wird nur frmVideo geschlossen, die anderen Formen bleiben offen
'           L�sung: gblnComeFromBildanzeigen = False setzen bei Prozedur VideoAbspielen
'26.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Zustand: bei Videos l��t sich kein Kommentar eintragen mit F10
'           L�sung: in KommentarForm wird die Prozedur Form_KeyDown auskommentiert
'26.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Wenn das Listenfenster oben ist, geht Taste F10 nicht
'           L�sung: auch F10 wird an Form1.Form_KeyDown(KeyCode, Shift) weitergereicht
'27.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Zustand: bei Videos in der Form F5MehrereZeilen kann ein Klick auf chkExif oder chkIPTC zur Schleife f�hren
'           L�sung: nur bei Dateinamenerweiterung 'JPG' wird ein Klick auf Exif oder IPTC akzeptiert
'30.09.2013 14.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Bei Programmstart ist neben der Form Query auch die Form Form1 sichbar
'           L�sung: vor Query.Show mache ich Form1.Hide
'----------------------------------------------------------------------------------------------------------
'10.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Bei Video abspielen mit internem videoplayer ist der Anzeiger f�r L�nge des Videos so weit rechts,
'           da� man ihn nicht erkennen kann. Das st�rt wenn ich im Kommentar eine Zeitangabe gemacht habe und das Video
'           dorthin positionieren will.
'23.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: Bei Shareware-Programmstart ist neben der Form Copy auch die Form Form1 sichbar
'           L�sung: vor Copy.Show 1 mache ich Form1.Hide
'23.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: Bei Programmstart und gleich Klick auf Einstellungen ist neben der Form WertxForm auch die Form Form1 sichbar
'           L�sung: vor WertxForm.Show 1 mache ich Form1.Hide
'23.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: Wenn ein Video spielt und ich klicke auf Einstellungen ist neben der Form WertxForm auch die Form Form1 sichbar
'           L�sung: In WertxForm wird nicht mehr aufgerufen Form1.MediaPlayerStop
'23.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: Bei Benutzung eines externen videoplayers ist neben frmVideo auch Form1 sichtbar
'           L�sung: Form1.Hide
'24.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: am Ende des Abspielens eines Videos mit dem internen mediaplayer wird mediaplayer control schwarz
'           falsche L�sung: man muss jedes video in einer Schleife abspielen (WMP.settings.setmode "loop",True)bis der Nutzer eingreift
'                           da blockiert es nachdem ich etwa 10 mal F3 gedr�ckt habe
'           richtige L�sung: ich warte in WMP_PlayStateChange bis NewState = 1(stopped) und wiederhole WMP.Controls.play
'                           dadurch wird jedes Video wiederholt bis der Nutzer eingreift
'25.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: seit 24.10.2013 passiert es, da� beim Wechsel von einem Video auf ein Bild das Video immer weiter spielt
'           L�sung: Bei MediaPlayerStop muss frmVideo.WMP.URL = "" eingestellt werden
'25.10.2013 14.0.2 Fehlerkorrektur/Verbesserung alle Versionen
'           Zustand: Nutzerdefinierte Felder beim SQL-Server sind die Datentypen durch andere Zahlen gekennzeichnet als bei Access
'           L�sung SQL-String entsprechend anpassen
'29.10.2013 Nachbesserung zum 30.09.2013
'25.11.2013 14.0.2 Verbesserung alle Versionen
'           Bei Screenwidth >= 1920 werden die Felder zum Anzeigen der EXIF/IPTC-Informationen maximal verbreitert
'01.01.2014 14.0.2 Verbesserung alle Versionen
'           Zustand: Beim Editieren in DbGridForm kann ich bisher nur eine Zelle kopieren und in eine andere Zelle eintragen
'           L�sung: Multiselect: Der mit Ctrl+C kopierte Wert einer Zelle wird in alle anderen Zellen der gleichen Spalte eingetragen, die
'           durch Klicken auf den Zeilenmarkierer mit Festhalten der Tasten Shift+Ctrl markiert(gebookmarked) werden
'           neuer Modul: module2
'           �nderungen in den verbotenen Spalten werden ignoriert
'           DbGrid.Row = x liefert den Laufzeitfehler '6148' Ung�ltige Zeilennummer, wenn die row unsichtbar ist
'           (au�erhalb des sichtbaren Bereichs). Man muss die gew�nschte Zeile zuerst in den sichtbaren Teil verschieben mit
'           DbGrid.FirstRow
'18.01.2014 14.0.2 Verbesserung alle Versionen
'           "DOC" und "DOCX" files sind ab sofort erlaubt.
'           Daraufhin hat Windows8 zwar die docx files ge�ffnet aber nicht die doc files (doc files nur wenn als Administrator gestartet)
'           Ich musste den Aufruf RunShellExecute ver�ndern
'----------------------------------------------------------------------------------------------------------
'04.02.2014 14.0.3 Fehlerkorrektur alle Versionen
'           Nachbesserung zu 01.01.2014
'           Zustand: seit 01.01.2014 geht Strg+C nicht mehr wie gewohnt, wenn ich einen Teil des Feldinhalts kopieren will,
'           kopiert es den gesamten Feldinhalt
'           L�sung: nicht Strg+C zum Kopieren benutzen, sondern Strg+(Minuszeichen auf dem Ziffernblock)
'12.02.2014 L�sung: nicht Strg+C zum Kopieren benutzen, sondern Strg+(Multiplikationszeichen auf dem Ziffernblock)
'           Strg+C bekommt wieder seine herk�mmliche Bedeutung
'           Zustand: die bisherige L�sung kann nicht mit Unicode Clipboard arbeiten
'           L�sung: Benutzung des Moduls modCopyUnicodeText.bas
'15.02.2014 14.0.3 Verbesserung alle Versionen
'           Zustand: Bisher kommt bei einem nichtvorhandenen Foto/Video ein schwarzes Fenster und eine MsgBox. Nach klicken auf OK
'           bleibt ein deprimierendes schwarzes Fenster �brig.
'           L�sung: Anstelle des schwarzen Fensters will ich einen blaugr�nen Farbverlauf und nutze die API function GradientFill
'15.02.2014 14.0.3 Verbesserung alle Versionen
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
'----------------------------------------------------------------------------------------------------------
'27.03.2014 14.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Es kann den Nutzer verwirren, wenn die Zeile in der er herumklickt, nicht mit dem aktuell gezeigten Bild
'           �bereinstimmt.
'           L�sung: Ab sofort wird beim Klicken in eine Zeile auch das dazugeh�rige Bild gezeigt. Aber nicht wenn
'           -aktuelle Zeile und aktuelles Bild bereits �bereinstimmt
'           -gerade Kopieren mit multiselect l�uft
'           -Form1.F6Continous eingeschaltet ist
'           Damit wird der Doppelklick auf eine Zeile �berfl�ssig.
'           gblnComeFromF2F3 dient zum Kennzeichnen, dass F2 oder F3 gedr�ckt wurde
'28.03.2014 14.0.4 Verbesserung alle Versionen
'           Wie Taste F1 soll auch Tastenkombination Strg+'-' wirken
'           Wie Taste F4 soll auch Tastenkombination Strg+'+' wirken
'28.03.2014 14.0.4 Fehlerkorrektur alle Versionen
'           Die Kontrolle, ob eine zweite Instanz von GERBING Fotoalbum ge�ffnet ist, hat bisher versagt
'           L�sung: bisher wurde abgefragt lstFensterTitel.ListCount < 2
'           ich muss abfragen lstFensterTitel.ListCount < 3
'13.04.2014 14.0.4 Fehlerkorrektur alle Versionen
'           bei Kopieren mit multiselect trat ein Fehler auf 'falscher Spaltenindex' wenn in lngGew�hlteSpalte der Wert = -1 steht
'19.04.2014 14.0.4 Nachbesserung zum 27.03.2014
'           Zustand: Beim Herumklicken im Listenfenster soll das Listenfenster nicht flackern
'           L�sung: Der Aufruf von MediaPlayerStop darf nur bei Videos drankommen
'22.04.2014 14.0.4 Verbesserung alle Versionen
'           nicht dokumentiert
'           Beim Dr�cken von Alt + F5 geschieht beinahe dasselbe wie beim Dr�cken von F5, aber
'           alle Buttons im oberen Teil der DbGridForm werden nicht angezeigt und das Grid DbGridNeu wandert in der Form
'           an den oberen Rand. Damit kann ich dem Nutzer mit noch weniger Platzbedarf erm�glichen, dass er das Listenfenster
'           dauerhaft angezeigt bekommt und die Beschreibung des aktuellen Bildes nur eine Zeile Platz braucht. Mit den Pfeil-nach-oben
'           und Pfeil-nach-unten-Tasten kann der Nutzer jeweils ein Bild weiterbl�ttern.
'           !!!wenn ein Video l�uft, ist Alt + F5 unwirksam
'25.04.2014 14.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Auff�llig oft endet fotos.exe beim Klicken auf das Schlie�-Kreuz mit runtime-Absturz
'           ich habe herausgefunden, da� Refresh_MyDrawImage zuletzt dran war, dann Fehler in GDIPLUS.DLL
'           m�glicherweise nur gefunden weil ich eingestellt habe Projekteigenschaften -> Compilieren -> P-Code anstelle native code
'           L�sung: in der Prozedur 'Beenden' einf�gen 'Form1.TimerRefresh.Enabled = False'
'15.05.2014 14.0.4 Verbesserung alle Versionen
'           Shareware user bekommen einen Hinweis auf Professional Version
'----------------------------------------------------------------------------------------------------------
'24.06.2014 14.0.5 Fehlerkorrektur
'           ist nur aufgetreten in BatchHistogramCorrection Diashow FotosMdb WallPaperChanger
'           Fehler: Bei der Funktion Rekursive ist ein Dateiname von >130 Bytes L�nge bisher ignoriert worden.
'           Das ist aufgetreten seit Version 14.0.0
'           L�sung: Die function FindFirstFileW und FindNextFileW in Module1 sind falsch deklariert, in UnivbzGlobal richtig
'           Falsch ist      Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'           richtig ist     Public Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFFData As Long) As Long
'           falsch ist      hSearch = FindFirstFileW(StrPtr(Path & "*"), wfd)
'           richtig ist     hSearch = FindFirstFileW(StrPtr(Path & "*"), VarPtr(wfd))
'           falsch ist      DirName = StripNulls(StrConv(wfd.cFileName, vbFromUnicode))
'           richtig ist     DirName = RemoveNulls((wfd.cFileName))
'02.07.2014 14.0.5 Fehlerkorrektur alle Versionen
'           Zustand: Es gibt Fotos, bei denen werden die EXIF-Informationen unvollst�ndig angezeigt
'           bei YCbCrCoefficients ist Schlu�
'           L�sung: in clsEFIF.cls wird abgefragt ob IFD(i).Length = 0 / Bei IFD(i).Length = 0 gab es einen unbehandelten Fehler
'           und die Prozedur wurde vorzeitig beendet.
'23.08.2014 14.0.5 Komfort-Erh�hung alle Versionen
'           Zustand: Bisher kann man die Form F5MehrereZeilen nur durch das Schlie�kreuz verlassen
'           Verbesserung: ab sofort sind neben dem Schlie�kreuz auch die Tasten F1 F2 F3 F4 erlaubt
'----------------------------------------------------------------------------------------------------------
'10.09.2014 14.0.6 Fehlerkorrektur alle Versionen
'           siehe 22.04.2014 Tastenkombination Alt + F5
'           Zustand: bei wiederholter Benutzung von Alt +F5 soll die oberste schwarze Zeile weiterhin die oberste schwarze Zeile bleiben
'                   Das tut sie nicht
'           L�sung: DbGridForm.DBGridNeu.Row korrigieren
'11.09.2014 14.0.6 Fehlerkorrektur alle Versionen
'           wenn das Kommentarfenster zu sehen ist, habe ich vorgesehen, da� die Tasten F1 F2 F3 F4 F11 erkannt werden und zum Schlie�en
'           des Kommentarfensters f�hren.
'           Das passiert manchmal erst, wenn ich ein zweites mal F10 dr�cke
'           Ich finde keinen Weg das zu verbessern.
'21.09.2014 14.0.6 Fehlerkorrektur alle Versionen
'           Wenn das Listenfenster(DbGridForm) ge�ffnet ist soll ab sofort die Tastenkombination Strg + F5 akzeptiert werden
'26.09.2014 14.0.6 Komfort-Erh�hung alle Versionen
'           siehe 29.07.2006
'           Es gibt keine Beschr�nkung mehr auf maximal 100 gefundene Fotos bei 'beim Bildwechsel individuell vergr��erten/verkleinerten
'           Bildausschnitt beibehalten'. Heutigen Tages sollte es genug memory geben um alle Foto-Informationen in ein array zu speichern.
'08.10.2014 14.0.6 Fehlerkorrektur alle Versionen
'           Zustand: Obwohl das gehen m��te, kann man das aktuelle Foto mittels Explorer oder mittels RenamMdb nicht umnennen oder l�schen
'           L�sung: ich darf keine Aufrufe 'retcode = GdipDrawImageRect...' machen, sondern mu� 'Call MyDrawImage(gstrFRODN, glngZoomProzent)'
'                   benutzen, weil nur so richtig
'                   If lngPointer Then
'                       retcode = GdipDisposeImage(lngPointer)
'                   End If
'                   ausgef�hrt wird
'                   F�r das Anzeigen der Hilfebox den Schalter blnHilfeboxStehenLassen benutzen
'08.10.2014 siehe 27.08.2012
'           ich mache die �nderung r�ckg�ngig '�ffne RenamMdb f�r die aktuelle Datei und beende das Programm
'           ab sofort mu� nicht mehr fotos.exe beendet werden
'09.10.2014 14.0.6 Fehlerkorrektur alle Versionen
'           Zustand: Kooperation mit RenamMdb geht falsch, wenn der erste Satz der Datenbank �bergeben werden soll,
'                   ist gstrRowColChangeName leer
'           L�sung: abfragen ob gstrRowColChangeName = ""
'14.10.2014 14.0.6 Fehlerkorrektur alle Versionen
'           Zustand: Rechtecklupen-Ausschnitt flackert kurz auf und verschwindet wieder
'           L�sung:  Form1.Picture1_Click mu� sofort wieder beendet werden, wenn das soeben gezeigte Bild mit Rechtecklupe erzeugt wurde
'16.10.2014 14.0.6 private Verbesserung nur in meiner privaten fotos.mdb
'           privat hei�t: nur wenn es Vollversion ist und nur wenn Spalte ExifDateTimeOriginal vorhanden ist
'           Zustand: Von mir festgelegte G�ltigkeitsregeln sind zwar wirksam, aber sie wirken lautlos. Wenn ein Ort l�nger ist als 32 Bytes,
'                   dann wird ohne Kommentar der alte Wert beibehalten
'           L�sung: Ich frage ab was in ValidationRule und ValidationText zu einem Datenbank-Feld definiert ist.
'                   In frmGridAndThumb kommt bei �berschreitung des Maximums eine MsgBox und es wird der alte Wert beibehalten.
'                   In KommentarForm wird auf die L�nge des Maximums abgeschnitten.
'                   Das erfordert ein vorgeschriebenes Format der G�ltigkeitsregel zB 'L�nge([Ort])<33'
'                   Das Programm sucht nach dem Zeichen < dahinter steht die L�ngenbegrenzung
'17.10.2014 14.0.6 Fehlerkorrektur alle Versionen
'           Zustand: Bei 'Fenstergr��e �nderbar' nicht ausgew�hlt kommt bei jedem Bildwechsel kurz erst ein Rahmen, dann kommt das Bild
'                   ohne Rahmen
'           L�sung: Ich muss in Prozedur Form1.Bildanzeigen �ndern von Form1.WindowState = 0 in Form1.WindowState = 2
'----------------------------------------------------------------------------------------------------------
'23.10.2014 14.0.7  Fehlerkorrektur alle Versionen
'           Zustand: scheinbar nicht l�sbares Problem in XP und Windows7 seit Version 13.5.4
'           Nach dem Verschieben der Formen AboutForm Copy DbGridForm F5MehrereZeilen ExportForm Hilfebx ImportForm KommentarForm
'           WertxFormXYPos ZielverzeichnisForm bleiben schwarze Fl�chen. Damit muss der Nutzer leben.
'           Ich habe versucht mit Subclassing FormMove gegenzusteuern, aber das sieht immer noch Schei�e aus. Hab es sein lassen.
'           Das Problem tritt nicht auf im Windows8
'           ebenfalls nicht bei Diashow.exe im Windows7
'
'           L�sung: so was von einfach.
'           falsch ist  Form1.Picture1.AutoRedraw = False. Da h�tte ich das Event Form1.Picture1.Paint benutzen m�ssen. Das geht, aber
'           sieht unsauber aus, man sieht immer noch schwarze R�nder.
'           richtig ist Form1.Picture1.AutoRedraw = True
'           Zitat aus der vb6-Hilfe: Setting AutoRedraw to True automatically redraws the output in a Form object
'           or PictureBox control when, for example, the object is resized or redisplayed after being hidden by another object.
'           Die einzige Schwierigkeit besteht darin, da� bei Form1.Picture1.AutoRedraw = True kein schraffiertes Rechteck gezeichnet
'           werden kann. F�r diesen Fall mu� ich das Bild mit Form1.Picture1.AutoRedraw = False zeichnen und darauf das Rechteck.
'           Wenn Mousepointer=Rechtecklupe sichtbar ist, lasse ich keine andere Aktion zu als das Rechteck zu zeichnen.
'           Form1.TimerRefresh entf�llt
'22.11.2014 14.0.7 Fehlerkorrektur alle Versionen
'           Zustand: Wenn 'weitere Filter sind aktiv' angezeigt wird und man �ndert danach die Jahreszahl, dann h�tte
'           'weitere Filter sind aktiv' ausgeschaltet werden m�ssen
'
'           L�sung:  Wenn das Jahr ver�ndert wird, muss ich alles zur�cksetzen, was in Form MP ver�ndert sein k�nnte und Label
'           'weitere Filter sind aktiv' unsichtbar machen
'03.12.2014 14.0.7  Kosmetische Korrektur
'           In Form ExportForm waren bisher die Optionen
'           -F�r den Export das Zielverzeichnis benutzen oder
'           -Mit Drag&&Drop sofort in andere ge�ffnete Anwendung  (Datenbank fotos.mdb) exportieren
'           nicht korrekt angeordnet. Bisher gab es einen unn�tzen FrameDragDrop
'10.01.2015 14.0.7  Verbesserung alle Versionen
'           Es gibt im Form Query eine neue Menu-Auswahl 'ResetAll'
'           Damit wird auch das Jahr auf den Wert '*' zur�ckgesetzt
'----------------------------------------------------------------------------------------------------------
'26.01.2015 14.0.8  Fehlerkorrektur alle Versionen
'           Zustand: Wenn ich im Listenfenster auf den Refresh-Button klicke und ich stehe nicht auf dem ersten Bild, dann flackert kurz das
'           erste Bild auf, dann erst das aktuelle.
'           L�sung: Ich muss verhindern da� bei DbGridNeu_RowColChange ein Call Form1.BildAnzeigen gemacht wird.
'           Dazu wird beim klicken auf den Refresh-Button der Schalter ExportForm.blnExportGestartet = True eingeschaltet
'18.03.2015 14.0.8 Fehlerkorrektur alle Versionen
'           Zustand: Wenn eine andere Software ein IPTC-Feld bearbeitet hat, kann es passieren dass GERBING Software
'                   keine IPTC-Felder anzeigt (gar keine = leer)
'           L�sung: modIPTC wird korrigiert und fragt nach dem ersten IPTC-Header wenn nicht "Photoshop 3.0" gefunden wird
'                   ob es weiter hinten noch einen IPTC-Header mit "Photoshop 3.0" gibt
'----------------------------------------------------------------------------------------------------------
'29.03.2015 14.1.1  neue Funktion alle Versionen
'           Nach jahrelanger Verweigerung bin ich jetzt bereit, die gefundenen Bilder als Thumbnails darzustellen. Ein Doppel-Klick auf einen
'           Thumbnail �ffnet das Bild in Vollbildansicht. Es gibt eine neue Form frmGridAndThumb. Das ist die bisherige Form DbGridForm
'           erg�nzt um eine Thumbnail-Ansicht. Es gibt zwei Panele das obere f�r DbGridNeu, das untere f�r die Thumbnails. Die Panele sind
'           resizable. Die Thumnails und die zugeh�rige Zeile im DbGridNeu sind synchronisiert. Bei ge�nderter Sortierfolge im DbGridNeu
'           werden auch die Thumbnails in dieselbe Sortierfolge gebracht.
'           Im oberen Panel gibt es eine Scrollbar, im unteren Panel nicht. Ins untere Panel kommen immer soviel Thumbnails, wie hineinpassen.
'           F�r Videos und Nicht-Bild-Dateien gibt es keine Thumbnails. Obwohl es m�glich w�re f�r Videos Thumbnails zu erzeugen.
'           Ein VB6-Beispiel steht unter d:\VISUALBA.SIC\VB6BeispielCode\Multimedia\Video\Video Thumbnails\
'           Wegen Unicode muss ich �ber den OptionButton der ein Thumbnail aufnimmt noch ein Unicode-Label legen.
'           Statusbar und Tooltip sind ebenfalls unicode-f�hig.
'           Es gibt keine Maximalzahl erlaubte Thumbnails aber das Synchronisieren(Neuanordnen) mit frmGridAndThumb.ChangePicFrameSize
'           verz�gert sich bei jedem Klicken ins Grid um so mehr je mehr Thumbnails es gibt. Sp�rbar ab etwa 500 Thumbnails.
'           Der Grund ist der Abschnitt ChangePicFrameSize2 wo zuerst alle Thumbnails unsichtbar gemacht werden. Das habe ich verbessert,
'           ich mache nur die Thumbnails unsichtbar, die ich zuvor angezeigt habe.
'
'           Zustand: clsToolTip und ttToolTip muss ich w�hrend des Testens auskommentieren, sonst ist Debuggen nahezu unm�glich
'           L�sung:  mit 'isIDE' abfragen, ob ich in der IDE-Umgebung arbeite
'
'           Beachten bei der Auslieferung: Die Spaltengr��en in der Datenbank-Tabelle SpaltenBreite sind jetzt in Pixel vorher in Twips
'               angegeben. Scheinbar nicht sichtbare Feldbegrenzungen zum kleiner ziehen bekommt man zu sehen,
'               wenn man an den rechten Rand der �berschrift geht dann minimal nach links. Der Mauszeiger
'               verwandelt sich in zwei senkrechte Striche mit einem Pfeil nach links.
'19.04.2015 L�sung: Es gibt einen neuen Men�-Punkt 'Spaltenbreite' mit dem k�nnen alle Spalten auf 100 Pixel Breite eingestellt werden.
'               Zus�tzlich wird beim ersten Start (vor der Spracheinstellung) eine SpaltenBreiteKontrolle gemacht.
'               Wenn die SummeSpaltenbreite h�her als 15000 ist vermute ich, da� bisher mit Twips gerechnet wurde
'               und �ndere die Standardspaltenbreite auf 100 Pixel
'
'           Beachten bei der Auslieferung: Wieder umstellen auf 'Kompilieren zu Systemcode (Native Code)' ohne Debug.Informationen.
'               Ich habe w�hrend der Entwicklung
'               'Kompilieren zu Systemcode (Native Code)' mit Debug.Informationen eingestellt weil ich hoffe den Fehler
'               in der fertigen exe besser zu finden, dass VB6 abst�rzt
'               manchmal nach Schlie�en des Programms und wenn der Nutzer ohne Admin-Rechte arbeitet
'               siehe "D:\P4Disks\disks\VB Komponenten und OCX\VC6SP6\Kompilierte Visual-Basic-Projekte debuggen.pdf"
'               es gibt eine Datei fotos.pdb.  Sie enth�lt die Quellcode-Informationen, die f�r den VC++-Debugger notwendig sind
'22.04.2015 Bei Systemen mit Bildschirmeinstellung DPI 96 ist der blaue Rand zur Markierung kaum zu erkennen, ich habe ihn verbreitert
'----------------------------------------------------------------------------------------------------------
'29.03.2015 14.1.1 gel�stes Problem mit der Benutzung von wmp.dll
'           Zustand: Die Funktionstasten Shift+F5 sowie F8 wird vom Mediaplayer abgefangen, darum waren sie bisher auf der Form frmVideo wirkungslos
'                   Genauso Tastenkombinationen Umsch+Strg+N und Umsch+Strg+M
'                   Umgehungsl�sung: Rechtsklick-Men� benutzen
'                   Es gibt auch Funktionen, die bei Videos unwirksam sein sollen, wenn sie vom Rechsklick-Men� kommen
'           L�sung: API Function GetAsyncKeyState und Timer TimerKeyboardHook ist in der Lage die Tasten F8  und Shift+F5 zu erkennen
'13.04.2015 14.1.1 gel�stes Problem
'           Zustand: Wenn kein H�kchen gesetzt ist in 'Fenstergr��e �nderbar' soll FullScreen kommen, aber die Taskbar bleibt stehen
'                   Beim ersten Start verschwindet nur die Titelzeile, beim zweiten Start verschwindet auch die Taskbar
'           L�sung: ShowTitleBar False, True                        'taskbar unvisible, Foto
'----------------------------------------------------------------------------------------------------------
'04.05.2015 14.1.1 gel�stes Problem
'           Zustand:manchmal bekomme ich eine schwarze Form1, nur ganz links oben ist ein St�ck von Picture1
'               Reproduzieren durch:
'               1.Fotoalbum mit Thumbnails
'               2.einen anderen Thumbnail anklicken
'               3.Form1 minimieren
'               4.Form1 wieder anzeigen -> schwarze Form1
'           L�sung: Wenn Form1 minimiert wird, sollen alle anderen offenen Formen auch minimiert werden
'----------------------------------------------------------------------------------------------------------
'04.05.2015 14.1.1 gel�stes Problem
'           Zustand: Wenn F10 gedr�ckt war (Die Kommentarform gezeigt werden soll, falls im Kommentarfeld etwas steht),
'                   dann geht es beim Klick aufs Grid richtig, aber beim Klick auf einen Thumbnail kommt der Kommentar des
'                   vorher aktiv gewesenen Thumbnails
'           L�sung: Schalter gblnWasOptThumbClick und neue Prozedur KommentarNachBildAnzeigen
'----------------------------------------------------------------------------------------------------------
'22.05.2015 14.1.1 Fehlerkorrektur alle Versionen Folgeerscheinung vom 18.03.2015
'           Zustand: Die Fotos von Ralph haben ewig lange gebraucht um mit Fotosmdb/Pr�fen3 in die Datenbank aufgenommen zu werden
'                   oder mit Diashow oder Fotoalbum oder WallpaperChanger anzezeigt zu werden.
'                   Bei pos = InStr(1, strImageString, IPTCHeader, vbTextCompare) braucht die Programmausf�hrung ewig lange
'                   Scheinbar wird ab einer bestimmten L�nge eines Strings die Function InStr arschlangsam.
'           L�sung: ich muss schreiben pos = InStrB(1, strImageString, IPTCHeader, vbTextCompare)
'                   InStrB geht blitzschnell
'                   aber beim anschlie�enden Vergleich muss die pos korrigiert werden
'31.05.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 22.05.2015
'01.06.2015 14.1.1 Fehlerkorrektur Folgeerscheinung vom 22.05.2015
'----------------------------------------------------------------------------------------------------------
'06.06.2015 14.1.1 Fehlerkorrektur alle Versionen
'           Zustand: Wenn die erste gefundene Datei docx ist (vermutlich genauso bei allen nicht nativ Typen), dann bleibt Form1 unsichtbar
'           L�sung: Form1 sichtbar machen bzw gstrRowColChangeName = "" setzen
'----------------------------------------------------------------------------------------------------------
'09.06.2015 lange Zeit nicht l�sbares Problem seit Version 14.1.1
'           Zustand: tritt im Windows8 und Windows10 auf, mit Thumbnails
'                   tritt nicht in der IDE auf, nur in der exe
'                   beim Klick aufs Schlie�kreuz von Form1 oder auch bei der F8-Taste kommt die Meldung
'                   Foto/Video Datenbank funktioniert nicht mehr
'                   -> Programm schlie�en
'                   -> Programm debuggen
'                   Der Absturz kommt am Ende von frmGridAndThumb.Form_Unload
'           scheinbare L�sung: 08.06.2015 Schuld war scheinbar clsToolTip von Timosoft
'                   ich habe andere Klasse clsToolTip von Dana Seaman eingebaut
'                   siehe http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_22951434.html
'           Zustand: sogar bei der L�sung ohne unicode -optThumb(Index).ToolTipText = Ulabel(Index).Tag- kommt der Fehler
'           keine L�sung: ich habe alle Varianten der Kompilierung zu P-Code oder native code durchprobiert -> keine L�sung
'           L�sung: ich ersetze die VB-Anweisung 'End' durch einen API-Aufruf im Modul modTerminateExe
'                   TerminateEXE "fotos.EXE"
'                   und frage ab, ob mit Thumbnails gearbeitet wurde. Wenn ja, f�hrt Form1.Form_Unload zu TerminateEXE "fotos.EXE" und
'                   wenn ja f�hrt das Dr�cken von F8 zu TerminateEXE "fotos.EXE"
'
'                   Beachten: Wenn mit Thumbnails gearbeitet wurde, f�hrt Dr�cken von F8 zum Beenden des Programms
'
'24.06.2015 14.1.1 Fehlerkorrektur Shareware-Version
'           Zustand: Der Mecker-Hinweis 'Sie benutzen die Shareware-Version' soll auch beim Wechsel von Thumbnails
'                   und bei Wechsel der Gridzeile kommen
'           L�sung: Einbau des Mecker-Hinweises
'25.06.2015 14.1.1 Fehlerkorrektur Shareware-Version
'           Zustand: Der Mecker-Hinweis 'Sie benutzen die Shareware-Version' kommt auch wenn ich an der Spaltenbreite ziehe und dabei
'                   mit der Maus ins Grid abrutsche. Es ist m�hsam nach dieser Situation weiterzuarbeiten.
'                   Man mu� erst auf die Taskleiste klicken.
'           L�sung: Ich ersetzen den Aufruf Copy.Show 1 durch eine MsgBox
'----------------------------------------------------------------------------------------------------------
'27.07.2015 14.1.2 Fehlerkorrektur alle Versionen
'           In frmGridAndThumb ist ab sofort das Grau der Buttons wieder besser vom Background unterscheidbar
'----------------------------------------------------------------------------------------------------------
'27.08.2015 14.1.2 neue Funktion
'           Geocoding: Falls in den EXIF-Felder des Fotos GPS-Positionen vorkommen, sollen sie in einer Landkarte wie bei Google maps
'           angezeigt werden.
'           Zusammenarbeit mit Picasa ist m�glich. Picasa kann GPS-Koordinaten ins Bild einf�gen.
'           Unter Aktion w�hlen und ausf�hren... kommt eine neue Aktion 'Zeige Geo-Position'
'           Bei Klick auf 'Zeige Geo-Position' mu� ich untersuchen, ob in den EXIF-Felder des Fotos GPS-Positionen vorkommen.
'           Wenn ja,
'           muss ich die GPS-Koordinaten aus der sexagesimalen Darstellung in die Dezimaldarstellung umrechnen, weil clsEXIF generell die
'           sexagesimale Darstellung liefert. Dann �ffnet sich die Form frmGEOPosition.
'           Dort macht das user control ucGMap den Hauptteil der Arbeit
'           ucGMap nutzt die http://maps.googleapis.com/maps/api
'           Es hei�t, der Aufruf der API ist (limited to 1000 requests per User and Day)
'           Wenn ucGMap breiter/h�her als 640 Pixel ist, wird die Schrift unscharf
'           ich habe stundenlang gek�mpft mit der tan() function, schlie�lich habe ich sie weggelassen und durch 1 ersetzt
'           ich habe stundenlang gek�mpft mit Mouse_Move auf der ucGMap. Die Positionsangaben der Maus stimmen nicht, Schlie�lich habe
'           ich alles Mouse_Move weggelassen
'----------------------------------------------------------------------------------------------------------
'15.09.2015 14.1.2 Fehlerkorrektur alle Versionen
'           Zustand: siehe 09.06.2015 TerminateEXE "fotos.EXE" wirkt nur bei einer exe. Soll in der IDE genauso wirken
'           L�sung: die n�chste Anweisung nach TerminateEXE "fotos.EXE" muss eine End-Anweisung sein
'----------------------------------------------------------------------------------------------------------
'15.09.2015 14.1.2 Fehlerkorrektur alle Versionen
'           Zustand: scheinbar nicht l�sbares Problem seit Version 14.1.1
'                   Klicken auf den ersten Thumbnail bewirkt nichts
'           L�sung: In Form frmGridAndThumb d�rfen Ulabel() und optThumb() nicht au�erhalb von picFrame angeordnet sein, sondern
'                   sie m�ssen innerhalb von picFrame angeordnet sein
'----------------------------------------------------------------------------------------------------------
'22.10.2015 14.1.2 Fehlerkorrektur alle Versionen
'           Zustand: In der Kommentarform sind die Scrollbalken kaum zu sehen
'           L�sung: Beim Form_Resize mehr Platz f�r die Scrollbalken lassen.
'----------------------------------------------------------------------------------------------------------
'25.10.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Wenn bei der Suche nichts gefunden wird, kann die erste MsgBox entfallen
'           L�sung: auskommentiert
'----------------------------------------------------------------------------------------------------------
'07.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: im Windows 10 und Windows 8.1 und vom Standard abweichender DPI-Einstellung zeigt mein Programm verschwommene Schrift
'                   Das kann der Nutzer korrigieren, indem er die exe markiert -> Eigenschaften -> Kompatibilit�t ->
'                   DPI-Skalierung nicht anwenden
'           L�sung: Ein Programm erkl�rt sich selbst als DPI-kompatibel. Das geht durch sein Manifest
'                   siehe d:\VISUALBA.SIC\Foto\Manifest Einf�gen mit DPI kompatibel\ManifestEinf�gen.exe
'----------------------------------------------------------------------------------------------------------
'08.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Wenn ein Nutzer ohne Administratorrechte im Windows 10 zum ersten mal nach der Sprachauswahl fotos.exe startet,
'                   kann eine MsgBox kommen 'kein einziger Datensatz gefunden'
'           L�sung: zus�tzlich kommt errornumber und errortext
'----------------------------------------------------------------------------------------------------------
'12.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Auch mov videos k�nnen EXIF-Felder besitzen. Die werden vom FotosMdb.exe angezeigt, aber nicht von fotos.exe
'           L�sung: Ich hebe die Beschr�nkung auf, dass nur JPG files EXIF-Felder haben k�nnen.
'----------------------------------------------------------------------------------------------------------
'13.11.2015 14.1.2 Verbesserung alle Versionen
'           Zustand: Ich habe nicht ber�cksichtigt, da� es auch EXIF-Felder mit unicode Inhalt geben kann die werden als ?????
'                   dargestellt
'           L�sung: In Form F5MehrereZeilen habe ich txtEXIFInfo durch ein unicode f�higes control ersetzt
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
'           Ursprung        Autoren         -> EXIF-XPAuthor
'
'           Ich stelle die Funktion IPTC... generell um auf Benutzung von Exiftool.exe. Alle zu schreibenden tags werden ab jetzt mit
'           exiftool.exe geschrieben.
'           Phil Harvey (exiftool-Autor) schreibt die EXIF-XP... Felder generell als unicode, aber die IPTC-Felder nur, wenn der Parameter
'           -charset IPTC=UTF8 benutzt wird
'----------------------------------------------------------------------------------------------------------
'16.11.2015 14.2.0 Verbesserung alle Versionen
'           Zustand: Wenn ein IPTC-Feld UTF-8 Code enth�lt (erzeugt durch exiftool mit einem unicode Feld), dann wird Mist angezeigt
'                   Ich sehe zweistellige Zeichen im UTF-8 Code.
'           L�sung: modIPTC.bas function VorhandeneEinzelsegmenteSuchen wird ver�ndert.
'                   Eventuell vorhandene UTF-8 Zeichen werden konvertiert. Es kommt FromUTF8String dran.
'----------------------------------------------------------------------------------------------------------
'21.12.2015 14.2.1 Verbesserung meine private Datenbank
'           Zustand: ExifDateTimeOriginal wird in ND.ListNutzerdefinierteFelder nicht als nutzerdefiniertes Datenfeld angeboten
'           L�sung: ich hatte 'gefunden = False' vergessen
'----------------------------------------------------------------------------------------------------------
'29.12.2015 14.2.1 Verbesserung meine private Datenbank
'           Ich m�chte EXIFDateTimeOriginal aktualisieren k�nnen f�r das aktuelle Suchergebnis
'           Dazu dient der Button btnEXIFDateTimeOriginalaktualisieren in der Form frmGridAndThumb. Dieser ist nur sichtbar, wenn
'           If Gefundenexifdatetimeoriginal = True And gblnVollversion = True And Sprache = 0
'           Zuvor k�nnte ich suchen in nutzerdefinierten Datenfeldern
'               EXIFDateTimeOriginal = null oder
'               EXIFDateTimeOriginal < 0
'           F�r 1000 Dateien braucht das Programm 30 Sekunden
'11.01.2016 Nachbesserung zu 29.12.2015
'           Aufgetauchte Unklarheit: Warum gibt es viele Fotos zb aus dem Jahr 2004 wo in EXIFDateTime das Jahr 2000 oder 2001 steht?
'           Die stammen aus den Anfangsjahren der digitalen Kameras. M�glicherweise war das Datum falsch oder gar nicht eingestellt
'----------------------------------------------------------------------------------------------------------
'07.03.2016 14.2.1 Fehlerkorrektur alle Versionen
'           Zustand: beim Anzeigen der Geoposition kommt Laufzeitfehler '5'
'           L�sung: Es gen�gt nicht nach GPSLatitudeRef: zu suchen - Abfragen ob GPSLatitude: gefunden wurde
'----------------------------------------------------------------------------------------------------------
'07.03.2016 14.2.1 Fehlerkorrektur alle Versionen
'           Zustand: Fenster 'Angaben zum aktuellen Bild' zeigt nicht immer das aktuelle Bild sondern manchmal auch die Angaben vom
'                   zuletzt benutzten Bild wo ich Geopositionen anzeigen lassen habe
'           L�sung: Vor dem Aufruf von F5MehrereZeilen.Show 1 muss ich ausf�hren Unload F5MehrereZeilen
'----------------------------------------------------------------------------------------------------------
'08.03.2016 14.2.1 Fehlerkorrektur meine private Datenbank
'           Zustand: Nur bei meiner privaten Datenbank(nicht SQL) sollen die G�ltigkeitspr�fung f�r Ort und Kommentar stattfinden
'                   sonst kommt Laufzeitfehler '91'
'           L�sung: if Gefundenexifdatetimeoriginal = True And gblnVollversion = True And Sprache = 0 And gblnSQLServerVersion = False
'----------------------------------------------------------------------------------------------------------
'11.03.2016 14.2.1 Fehlerkorrektur alle Versionen
'           Zustand: in Form F5MehrereZeilen kann ich aus EXIF-Feldern Texte kopieren, aber aus IPTC-Feldern nicht
'           L�sung: ich entferne die Listbox LstU und ersetze sie durch die Textbox txtIPTCInfo
'----------------------------------------------------------------------------------------------------------
'25.03.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: Es st�rt mich, wenn als erste Datei zB eine DOC-Datei gefunden wird, wo ich doch JPG-Fotos erwarte.
'                   Bisher kann ich direkt nach zB DOC-Dateien suchen, wenn ich bei 'Suche Begriff in jedem Feld' eingebe '.DOC'
'                   Ich will wenigstens teilweise eine Automatisierung.
'           L�sung: in der Suchmaske von Form Query gibt es eine neue Combobox TFileType
'                   Das Feld TFileType setzen, das ist eine Combobox mit Style = 0, dadurch kann der Nutzer selbst etwas eintippen
'                   Wenn der erste Satz der Datenbank die Dateinamen-Erweiterung 'JPG' hat, trage 'JPG' als ersten item ein und '*' als zweiten
'                   Wenn der erste Satz der Datenbank die Dateinamen-Erweiterung 'AVI' hat, trage 'AVI' als ersten item ein und '*' als zweiten
'                   sonst trage '*' als ersten item ein
'                   Wenn TFileType.Text = '*' dann verl�uft die Suche wie bisher
'                   Wenn TFileType.Text <> '*' dann m�ssen alle SQL-Strings erweitert werden um AND Dateiname not like "*." & TFileType.Text
'----------------------------------------------------------------------------------------------------------
'28.03.2016 14.2.1 Verbesserung SQL-Server-Version
'           Zustand: Wenn ich in frmConnectSQL auf die Spalte username klicke zum Aufsteigend/absteigend Sortieren, kommt Laufzeitfehler
'                   '91' Objektvariable oder With-Blockvariable nicht festgelegt
'           L�sung: rstsql benutzen
'----------------------------------------------------------------------------------------------------------
'10.04.2016 14.2.1 Verbesserung alle Versionen
'           am 05.07.2010
'           Es gab die Idee, �nderungen an den Datenbank-Feldern gleichzeitig in die IPTC-Felder zu kopieren.
'           Diese Idee scheitert daran, da� ich den Nutzer nicht bei der Bearbeitung jedes Fotos fragen will, in welche
'           IPTC-Felder er den Inhalt der Datenbank-Felder kopieren will.
'           Als L�sung biete ich in Fotosmdb.exe an, die Synchronisierung r�ckw�rts in die JPG-Fotos mit allen Fotos auf einen Aufwasch
'           zu machen.
'           Diese L�sung kann ich noch verbessern.
'           L�sung: Jedes Editieren im Listenfenster(frmGridThumb) oder in der KommentarForm hat zur Folge, da� im Feld IPTCPresent = 0
'                   eingetragen wird, wenn es ein JPG-Foto ist.
'                   Damit kann ich sehr einfach in Fotosmdb.exe bei der Funktion 'EXIF/IPTC...' die dritte Option ausw�hlen
'                   'F�r IPTCPresent=False' um die �nderungen in die Fotos zu synchronisieren
'----------------------------------------------------------------------------------------------------------
'11.04.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: Beim Klicken von frmGridAndThumb.btnRefresh tritt ein Laufzeitfehler auf, wenn ich durch Editieren aller bei der Suche
'                   gefundenen Datens�tze daf�r gesorgt habe, da� durch Refresh kein einziger Datensatz mehr gefunden wird.
'           L�sung: Ich muss RecordCount = 0 abfragen, dann eine MsgBox bringen, dann die Prozedur verlassen
'----------------------------------------------------------------------------------------------------------
'12.06.2016 14.2.1 Verbesserung alle Versionen
'           Video-Filter
'           werden im Fenster MP hinzugef�gt. Der Nutzer kann jetzt Fotos/Videos filtern nach Breite, H�he, Dauer ausw�hlen
'----------------------------------------------------------------------------------------------------------
'12.07.2016 14.2.1 Fehlerkorrektur alle Versionen
'           Zustand: Bei Programm starten, dann 'Suche Begriff in jedem Feld' leer dann Klicken auf 'Fotos finden' kommt ein Programmabsturz
'           L�sung: Ich muss abfragen, ob 'Suche Begriff in jedem Feld' = leer ist
'----------------------------------------------------------------------------------------------------------
'14.07.2016 14.2.1 Verbesserung alle Versionen
'           Erweiterung zum 25.03.2016
'           Ich will dem Anwender zeigen welche verschiedenen Dateinamenerweiterungen es gibt und ihn '*' oder eine bestimmte ausw�hlen lassen
'----------------------------------------------------------------------------------------------------------
'29.08.2016 14.2.1 Verbesserung alle Versionen
'           Zustand: Im XP bei Streubs kommt ein Fehler beim Beenden von fotos.exe
'                   Fehler beim L�schen der Datei oder des Ordners' Datei kann nicht gel�scht werden......
'                   Im XP meckert 'file_delete' wenn die Datei nicht existiert ab win7 wird nicht mehr gemeckert
'           L�sung: vorher file_path_exist ausf�hren
'----------------------------------------------------------------------------------------------------------
'02.09.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Die Geo-Position in der Grad-Minuten-Sekunden-Darstellung wird manchmal richtig angezeigt und manchmal falsch.
'           L�sung: im Modul clsExif: ich vermeide das Umrechnen in Grad-Minuten-Sekunden sondern benutze gleich eine Dezimalzahl,
'                   so wie sie in frmGEOPosition gebraucht wird
'                   bei Suche nach GPSLatitude/GPSLongitude muss ich das Komma stehen lassen, anstelle es in Punkt zu verwandeln
'----------------------------------------------------------------------------------------------------------
'05.09.2016 15.0.0 Verbesserung Professional Version
'           Das ist eine neue Funktion.
'           Suche nach Fotos mit GEO-Positionen innerhalb eines Rechtecks auf der Landkarte.
'           F�r das Zeichnen des Rechtecks gibt es die neue Form frmGEOFinden.
'           Es wird vorausgesetzt, dass der Nutzer die nutzerdefinierten Felder GPSLatitude und GPSLongitude(Text) selbst erzeugt hat.
'           Nur wenn es die nutzerdefinierten Felder GPSLatitude und GPSLongitude gibt �ffnet sich frmGEOFinden, sonst
'           MsgBox "F�r die Suche mit GEO-Daten muss die Datenbank die nutzerdefinierten Felder GPSLatitude und GPSLongitude enthalten."
'           Die Informationen aus dem zuletzt gezeichneten Rechteck werden gespeichert in Tabelle ErsterStart
'           Feld LetzterGEOPunkt
'           Feld ZoomListIndex
'           In einer Datenbank fotos.mdb wo die Felder LetzterGEOPunkt und ZoomListIndex nicht vorkommen, werden sie von Fotosmdb.exe
'           oder fotos.exe je nachdem welches zuerst ausgef�hrt wird, vom Nutzer unbemerkt erzeugt, wenn es die Professional Version ist.
'           Die Eckpunkte des gezeichneten Rechtecks werden in den SQL-String eingebaut
'----------------------------------------------------------------------------------------------------------
'22.09.2016 15.0.0 Fehlerkorrektur alle Versionen
'           Zustand: F�r die Merkerspalte wirkt kein Multiselect
'           L�sung: es gab eine Abfrage, ob die gew�hlte Spalte die Spalte Null ist, dann wurde nichts gemacht. Diese Abfrage muss entfallen.
'----------------------------------------------------------------------------------------------------------
'27.09.2016 15.0.0 Fehlerkorrektur alle Versionen
'           Zustand: Aktion w�hlen und ausf�hren... -> Hyperlink wird verbessert
'           L�sung: Aktion w�hlen und ausf�hren... -> Hyperlink ist in jeder Version sichtbar, aber bei der Shareware-Version kommt
'                   Msgbox 'F�r diese Funktion ben�tigen Sie die Professional Version'
'                   Bei der Professional Version wird kein #hyperlink# gebraucht das ist nur f�r hyperlinks auf der eigenen Festplatte n�tig
'                   # kann also weggelassen werden.
'                   Es wird kein Feld mit dem Feldtype Hyperlink ben�tigt, den gibt es beim SQL-Server ohnehin nicht.
'                   jedes beliebige Textfeld kann einen Hyperlink enthalten.
'                   Alle Untersuchungen ob HyperlinkField k�nnen entfallen.
'----------------------------------------------------------------------------------------------------------
'28.09.2016 15.0.0 Verbesserung Professional Version
'           Das ist eine neue Funktion zur Verbesserung des Umgangs mit
'           GPSLatitude
'           GPSLongitude
'           EXIFDateTimeOriginal
'           Die erw�hnten Felder k�nnen f�r die aktuelle Datei-Auswahl erneut importiert werden
'           Dazu wird das Men� 'Aktion w�hlen und ausf�hren...' erweitert um 'Feld-Aktualisierung durch Import-Wiederholung'
'           bei der Shareware-Version kommt die Msgbox 'F�r diese Funktion ben�tigen Sie die Professional Version'
'           Es gibt die neue Form frmFeldAktualisierung
'----------------------------------------------------------------------------------------------------------
'03.10.2016 15.0.0 Fehlerkorrektur alle Versionen
'           Zustand: Ich habe bei manchen Buttons vergessen, sie mittels Tastatur steuerbar zu machen
'           L�sung: Das wird jetzt nachgeholt durch das & Zeichen
'                   frmGridAndThumb 3 Buttons
'                   frmGridAndThumb alle Aktionen...  diese m�ssen ohne Alt-Taste gedr�ckt werden
'                   ND 1 Button
'----------------------------------------------------------------------------------------------------------
'03.10.2016 15.0.0 Verbesserung alle Versionen
'           Ich will bei Dr�cken der Tasten Ctrl+G sofort die GEO-Position anzeigen
'           und ich will dass mit F1 F2 F3 F4 die Form entladen wird
'----------------------------------------------------------------------------------------------------------
'03.10.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Wenn ich Geo-Positionen �bertrage aus EXIF in die Felder GPSLatitude und GPSLongitude sind diese bis zu 13 Stellen
'                   hinter dem Komma lang
'           L�sung: clsEXIF Ich begrenze die Anzahl Stellen hinter dem Komma auf 8
'----------------------------------------------------------------------------------------------------------
'09.10.2016 15.0.0 Verbesserung SQL Server Version
'           Zustand: In AboutForm - SQL Server Version sind 4 Textboxen editierbar, das verwirrt den user
'           L�sung: 4 mal Enabled = False in der IDE eintragen
'----------------------------------------------------------------------------------------------------------
'12.10.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Bei der Anzeige der IPTC-Felder in F5MehrereZeilen kommt beispielsweise in 10 Zeilen die Vorsilbe 'Caption'
'           L�sung : Ich entferne in modIPTC.LeseIPTC alle Loop-Konstruktionen
'                   zB Caption wird jetzt auf eine Zeile geschrieben
'----------------------------------------------------------------------------------------------------------
'12.10.2016 15.0.0 Fehlerkorrektur alle Versionen zu 05.09.2016
'           Zustand: Ich habe bisher nicht ber�cksichtigt, dass es negative GPSLatitude und GPSLongitude geben kann
'                   S�dhalbkugel und westliche Hemisph�re
'                   Damit die Vergleiche im SQL-String richtig ablaufen, muss der Datentyp von GPSLatitude und GPSLongitude Double sein
'           L�sung: In Form frmFeldAktualisierung. Wenn 'GPSLatitudeRef: S' gefunden wird, dann Minus vor GPSLatitude
'                   Wenn 'GPSLongitudeRef: W' gefunden wird, dann Minus vor GPSLongitude
'                   Konvertieren von String in Double mit CDbl(...) beim Lesen aus den EXIF-Feldern
'----------------------------------------------------------------------------------------------------------
'26.10.2016 15.0.0 Fehlerkorrektur SQL Server Version
'           Zustand: Im SQL Server Connect Fenster ist 'Windows Authentication' Standard. Aber es werden auch user name und password gezeigt
'           L�sung: Wenn PublicWindowsAuthentication = "1" dann werden diese Felder unsichtbar gemacht
'----------------------------------------------------------------------------------------------------------
'10.11.2016 15.0.0 Verbesserung alle Versionen
'           Ich speichere Thumbnails im Ordner ...\GerbingThumbs\...
'           Bei 'Mit Thumbnails' untersuche ich den Ordner ...\GerbingThumbs\... Alles was es bisher nicht als Thumbnail gibt, wird erzeugt.
'           Bei 'Mit Thumbnails' werden schon existierende gleichnamige Thumbnails nicht neu erzeugt.
'           Wenn jemand ein anderes Bild mit dem gleichen Namen ins Fotoalbum stellt, sieht er einen falschen Thumbnail.
'           und bei schon existierenden Thumbnails wird gleich der fertige Thumbnail geladen.
'           Ich f�ge eine Scrollbar hinzu, damit der Nutzer durch alle Thumbnails bl�ttern kann.
'           Die VB6-eigene VScrollbar bringt manchmal Laufzeitfehler '380' bei vsbSlide.Value = Wert und manchmal nicht
'           Darum lasse ich bei Fehlern einfach weiterarbeiten (On Error Resume Next)
'24.11.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Bisher werden f�r andere Dateien als 'JPG' keine Thumbnails angezeigt. Ich finde es besser, wenn dann wenigstens ein
'                   leeres wei�es Thumbnail-Bild angezeigt wird
'           L�sung: In die Kollektion Koll d�rfen alle file names
'                   In den Ordner ...\GerbingThumbs\... werden nur 'JPG' files aufgenommen
'24.11.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Bisher gibt es keine Video-Thumbnails
'           L�sung: Ich kopiere die Verfahrensweise aus "d:\VISUALBA.SIC\Foto\Thumbnails mit Unicode Label\"
'                   wenn es den Thumbnail in ...\GerbingThumbs\ nicht gibt wird dort zuerst einer erzeugt
'                   dann wird f�r die Thumbnails-Kollektion Koll dieser Thumbnail genommen
'                   F�r "AVI", "MPG", "MOV", "WMV", "ASF" k�nnen Thumbnails erzeugt werden
'                   F�r MP4 und ASX geht es nicht
'                   vielleicht geht es mit MP4 und ASX ja auch wenn andere Codecs installiert sind
'                   bei http://www.vbforums.com/showthread.php?761717-VB6-Shell-Video-Thumbnail-Images nachlesen
'                   Da ist erw�hnt. dass es nur funktioniert, wenn Codecs installiert sind
'27.11.2016 Bei 'Mit Thumbnails' soll das Ziehen am Scrollbalken nicht zum Flackern f�hren
'30.11.2016 Wenn ein Video l�uft soll Klick auf Foto-Thumbnail bewirken dass Form1 gezeigt wird
'30.11.2016 Ich habe Versuche mit fr�her Bindung/sp�ter Bindung von Shell32.dll gemacht, Sp�te Bindung erzeugt keine Thumbnails
'30.11.2016 Zustand: runtime error '430' bei frmGridAndThumb.Form_Load bei Shell32.dll
'           L�sung: Objekt nicht 'As New' binden sondern mit CreateObject(...)
'           Zustand: runtime error '13' type mismatch in frmGridAndThumb.Form_Load
'                   bei Set ShellObject = CreateObject(CVar("Shell.Application"))
'           ===================================================================================================
'           L�sung: ausgelieferte Versionen m�ssen im Win10 erzeugt werden und gehen nur im Win10, meine Privat-Version geht auch im Win7
'           ===================================================================================================
'                   anstelle 'Folder is Nothing' sage ich dem Nutzer dass er Windows 10 benutzen soll
'           Zustand: runtime error '91' Objectvariable nicht festgelegt nach Klick auf das Schlie�kreuz, aber nur wenn  w�hrend des Abspielens
'                   eines Videos geklickt wurde
'           L�sung: Ausf�hren von Call Form1.MediaPlayerStop
'07.12.2016 Auch andere Fotos als JPG und auch die Thumbnails von 'anderen' Dateien kommen nach ...\GerbingThumbs\...
'           Die Dateien im Ordner ...\GerbingThumbs\... hei�en zB video1.avi.jpg oder foto33.jpg.jpg
'----------------------------------------------------------------------------------------------------------
'12.12.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: 'Mit Thumbnails' spinnt, wenn ich vorher das gesamte Fotoalbum(ohne Thumbnails) nach Spalte SWF sortiert hatte.
'           L�sung: Ich hatte vergessen, bei Dr�cken der Taste F8 zu codieren  'gblnWasHeadClick = False'
'----------------------------------------------------------------------------------------------------------
'13.12.2016 15.0.0 Verbesserung alle Versionen
'           Zustand: Problem mit exiftool.exe und PSP X8. Alles was ich mit exiftool hinzugef�gt habe,
'                   (�ber fotosmdb.exe Funktion EXIF/IPTC...), wird von PSP X8 wieder rausgeschmissen.
'                   PSP X8 l��t den Abschnitt IPTC2 - aber l�scht den Abschnitt IPTC
'                   PSP X8 l�scht aus dem Abschnitt IFD0 alle Felder XPTitle XPKeywords XPAuthor XPSubjects XPComment
'                   Wenn ich eine neue Datenbank bei Null erzeuge durch Import der EXIF/IPTC-Felder, dann bekommen diese Bilder leere
'                   Datenbankfelder.
'           L�sung: Gegenma�nahme: fotosmdb.exe Funktion EXIF/IPTC... wiederholen f�r IPTCPresent = False,
'                   bevor eine neue Datenbank bei Null erzeugt wird
'                   Keine L�sung aber Milderung: ich stelle 'IPTCPresent = 0(False)'
'                   bei '�ffnen der mit 'xyz' verkn�pften Anwendung f�r die aktuelle Datei
'                   bei '�ffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist
'                   bei Fotosmdb 'Pr�fen1' wenn das aktuelle Dateidatum h�her ist als im Feld DDatum
'----------------------------------------------------------------------------------------------------------
'22.01.2017 15.0.0 15.0.0 Verbesserung alle Versionen
'           Zustand: Nach 'Thumbnails abbrechen' geht die Eieruhr nicht aus
'           L�sung: Screen.MousePointer = vbDefault
'10.02.2017 15.0.0 15.0.0 Verbesserung alle Versionen
'           Zustand: Bei Notebooks mit 1366x768 Pixel kann man in der Form frmGridAndThumb die Unterkante der Form nicht sehen und nicht
'                   anfassen und nicht verschieben
'           L�sung: Ich lege Me.Height fest auf Me.Height = 718 * Screen.TwipsPerPixelY
'=========================================================================================================
'11.03.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Wenn 4K-Monitore benutzt werden, muss es m�glich sein, die Schriftgr�esse besser als bisher anzupassen
'           L�sung1: Es gibt die Schriftgr��en
'                   klein=1
'                   mittel=2
'                   gross=3
'                   Die Einstellung wird gespeichert in der ini-Datei   [Adjustments]
'                                                               CheckForDPI 1 oder 2 oder 3
'           L�sung2: oder es gen�gt die Bildschirmaufl�sung auf zB 200 DPI einzustellen (Windows 10 kann noch weit h�her als 200 DPI)
'----------------------------------------------------------------------------------------------------------
'11.04.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Es gibt viele nichtssagende Dateinamen wie IMG_0432 Juni028.jpg Chemnitz23.jpg Ostern014.jpg
'                   Diese will ich ersetzen durch NameAlt & Ort & Situation & Personen
'                   Wenn alle Feldinhalte Ort & Situation & Personen IsNull dann bleibt der alte Name erhalten
'           L�sung: 1.Suchlauf �ber gespeicherte Abfrage zB Dateiname k�rzer als xx Bytes
'                   2.Korrekturlauf �ber Aktionen nach F5 ausw�hlen mit 'NamenErsetzen'
'                   3.Ung�ltige Zeichen im Dateiname(wie :?"\) werden ersetzt durch '-'
'----------------------------------------------------------------------------------------------------------
'19.04.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Ich bekomme trotz aller Gegenma�nahmen hin und wieder einen vb6-Absturz
'                   mit der Meldung
'                   Foto/Video Datenbank funktioniert nicht mehr
'                   -> Programm schlie�en
'                   -> Programm debuggen
'           L�sung: Endlich habe ich die Ursache gefunden
'                   Jetzt kommt kein Absturz mehr bei Beenden von fotos.exe mit Schlie�kreuz
'                   Jetzt kommt kein Absturz mehr bei Taste F8
'                   Jetzt kann ich auf TerminateEXE verzichten
'                   Ich muss nach GdipDisposeImage lngPointer = 0 setzen
'----------------------------------------------------------------------------------------------------------
'05.06.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: siehe 24.11.2016 Thumbnails bei Videos
'                   es ging bisher nicht f�r mp4 videos und es ging nur im Win10
'           L�sung: jetzt geht es doch und auch im Win7. Ursache unbekannt, ich bin darauf gesto�en, weil es im FotosMdb.exe funktioniert hat
'                   Bisher leere wei�e Thumbnails muss ich l�schen und neu erzeugen lassen
'----------------------------------------------------------------------------------------------------------
'10.08.2017 15.0.1 Verbesserung Professional Version
'           Ich benutze ab sofort InnoSetup f�r die Erstellung von Installationspaketen
'           deshalb muss regprofi.exe die Datei gerbingsoft.log in AppPath erzeugen
'           und fotos.exe muss sie dort lesen
'           Das ist auch Voraussetzung f�r eine portable Version
'----------------------------------------------------------------------------------------------------------
'11.08.2017 15.0.1 Verbesserung alle Versionen
'           Benutzung der Tastenkombination Strg+C kopiert das aktuelle Bild
'           wahlweise in die Zwischenablage oder in einen Ordner (dort Einf�gen mit Strg+V)
'----------------------------------------------------------------------------------------------------------
'13.08.2017 15.0.1 Verbesserung alle Versionen
'           Ab sofort kann im Email-Fenster die aktuelle Datei als Anhang �bernommen werden
'           Voraussetzung ist ein einstalliertes Outlook
'----------------------------------------------------------------------------------------------------------
'24.08.2017 15.0.1 Fehlerkorrektur alle Versionen
'           Zustand: Fehler-Msgbox 'Die Datei C:\GERBING FotoAlbum 15\Fotos.mdb  ist nicht vorhanden'
'                   kommt bei Installation Installation nach C:\GERBING Fotoalbum 15 (das macht mein InnoSetup als Standard)
'                   Ursache ist dass Newfotos.mdb nicht in fotos.mdb umgenannt werden kann, aber fotos.mdb ist bereits gel�scht
'                   und anschlie�end wird Newfotos .mdb gel�scht. Dann ist keine der beiden .mdb mehr da
'           L�sung: Nach DBEngine.CompactDatabase in Form1.Form_Load
'                   kein Umnennen machen, sondern erst bei 'Query.Beenden'
'           Jetzt funktioniert aber CompactDatabase h�ufig nicht
'           Ausweg: einmal Sprache �ndern und wieder zur�ck
'----------------------------------------------------------------------------------------------------------
'27.08.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Bei Auswahl einer fotos.mdb auf einem anderen Rechner gibt es keinen Hinweis, dass das Programm arbeitet,
'                   nicht einmal die Sanduhr
'           L�sung: frmFotoAlbumWirdGeladen topmost anzeigen
'                   Kein DBEngine.CompactDatabase bei Auswahl einer fotos.mdb auf einem fremden Rechner
'----------------------------------------------------------------------------------------------------------
'01.09.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Das Men� 'Aktion w�hlen und ausf�hren...' ist umst�ndlich zu erreichen
'           L�sung: Bei abwechselndem Dr�cken der Alt-Taste erscheint und verschwindet in Form1 oder in frmVideo ein Men�,
'                   Datei Version Hilfe
'                   wo unter Datei alles enthalten ist was bisher in WeiterAnShellExecute stand und zus�tzlich
'                   Strg+C
'                   Strg+K
'                   Strg+I
'                   Das was der Men�-Editor erzeugt hat, habe ich direkt aus Form1.frm nach frmVideo.frm kopiert
'----------------------------------------------------------------------------------------------------------
'15.09.2017 15.0.1 Verbesserung Shareware-Version
'           Zustand: Bisher gibt es keine portable Shareware-Version von GERBING Fotoalbum 15
'           L�sung: Ich k�nnte die Shareware-Version ab sofort als Portable Version anbieten als zip-file, aber ich tus nicht. F�r wen denn?
'                   neben der Version mit InnoSetup
'                   das ist ein leicht modifizierter Ordner 'publish FotoalbumPortable Shareware'
'                   F�r die Vollversion und Professional Version mache ich das nicht
'                   Eine portable Version muss
'                       ohne registry laufen
'                       muss ohne Administrator-Rechte laufen (das gelingt mir nicht ganz)
'                       darf keine Spuren hinterlassen (Das gelingt mir nicht ganz manche ocx'e registrieren sich selbst)
'
'                   Ich habe im Internet UMMM zum erzeugen eines Manifestes gefunden
'                   Dieses Manifest braucht man, wenn man dem System mitteilen will, dass die ocx'e im eigenen Ordner benutzt werden sollen
'                   msdatgrd.ocx
'                   MSSTDFmt.dll
'                   MSBind.dll
'                   Mit '...Manifest Einf�gen mit DPI regfree\ManifestEinf�gen.exe' f�ge ich das Manifest in folgende exe files ein
'                   aber nur f�r die portable Shareware-Version von GERBING Fotoalbum 15, f�r die anderen Versionen nicht
'                   fotos.exe
'                   BatchHistogramCorrection.exe
'                   Diashow.exe
'                   Fotosmdb.exe
'                   Renammdb.exe
'                   Die portable Version geht sogar als Multiuserversion und kann von verschiedenen usern benutzt werden. Jeder user
'                   muss fotos.exe und fotos.ini im gleichen pers�nlichen Ordner haben. Das Fotoalbum steht irgendwo und wird �ber
'                   Fotos.exe mit Shift-Taste aufgerufen
'----------------------------------------------------------------------------------------------------------
'27.09.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: es nervt mich, dass ich gleich beim ersten angezeigten Bild auf oben rechts 'Fenster maximieren' klicken muss
'                   damit die fotos den Bildschirm ausf�llen
'           L�sung: Form1.WindowState = 2
'----------------------------------------------------------------------------------------------------------
'02.10.2017 15.0.1 Verbesserung alle Versionen
'           Zustand: Bei der MsgBox 'Mit diesen Such-Kriterien wurde kein einziger Datensatz gefunden' kann bei wiederholtem Klicken auf
'                   'Fotos finden' die MsgBox immer l�nger werden
'           L�sung: Ich muss 'msg = ""' ausf�hren
'----------------------------------------------------------------------------------------------------------
'03.10.2017 15.0.1 Problem CompactDatabase
'           Zustand: Seit 27.08.2017 funktioniert CompactDatabase h�ufig nicht
'                   Beim Debuggen kommt error 3356 (You attempted to open a database that is already opened exclusively
'                   by user 'x' on machine 'y')
'                   Der Task-Manager findet eine aktive fotos.exe. Irgendwas ist h�ngengeblieben.
'           L�sung: einmal Sprache �ndern und wieder zur�ck da wird entweder CompactDatabase ausgef�hrt oder es kommt error 3356,
'                   dann mit Task-Manager die h�ngegebliebene fotos.exe schlie�en, dann wiederholen
'----------------------------------------------------------------------------------------------------------
'18.10.2017 15.0.1 Problem CompactDatabase
'           Zustand: Bei Installation nach C:\GERBING Fotoalbum 15\fotos.mdb kann Newfotos.mdb nicht umgenannt werden in fotos.mdb
'           L�sung: anstelle von 'rename altername, neuername' benutze ich'file_copy(Quellname, Zielname)'
'                   warum rename nicht funktioniert aber file_copy funktioniert, weis ich nicht
'----------------------------------------------------------------------------------------------------------
'22.11.2017 15.0.1 Fehlerkorrektur Professional Version
'           Zustand: Im Fenster Num + F5 (F5MehrereZeilen) f�hrt Klicken auf btnNutzerdef (Nutzerdefinierte Felder einstellen)
'                   zu Laufzeitfehler 3265:
'                   Ein Objekt, das dem angeforderten Namen oder dem Ordinalverweis entspricht, kann nicht gefunden werden.
'           L�sung: On Error resume Next
'=========================================================================================================
'23.11.2017 15.0.2 Problem mit unicode filename wenn zB GGCnopt\fotos.mdb  Access Datenbank
'           kein Problem mit SQL-Server-Version
'           Zustand: Es kommt 'Kein zul�ssiger Dateiname' fr�her ging das schon mal
'                   Vermutlich hat Microsoft daran herumgedreht.
'                   Die Datenbank l�sst sich aber mit MS Access �ffnen.
'                   Ich komme mit DBsql.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;... dar�ber hinweg
'                   aber dann spinnen andere Stellen im Code, die mit DAO programmiert sind
'           L�sung: DAO Code durch ADO Code ersetzen
'                   Verweise... -> C:\Program Files (x86)\Common Files\System\ado\msjro.dll#Microsoft Jet and Replication Objects 2.6 Library
'                   DBsql.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0
'                   Pr�fung ob die Datenbank schreibgesch�tzt ist mit SQL = "UPDATE FET SET FN = 'test'"
'                   es gibt nur noch DBado f�r beide Versionen Access oder SQL-Server
'                   anstelle von DBEngine.CompactDatabase -> CompactDB,
'                   CompactDatabase geht nicht mit Unicode filename, wenn Provider=Microsoft.Jet.OLEDB.4.0 angegeben wird
'                   Reference auf Microsoft DAO 3.6 Object Library wird nicht mehr gebraucht. dao360.dll wird nicht mehr gebraucht
'           Zitat aus https://stackoverflow.com/questions/14401729/difference-between-microsoft-jet-oledb-and-microsoft-ace-oledb
'           With version 2007 onwards, Access includes an Office-specific version of Jet, initially called the
'           Office Access Connectivity Engine (ACE), but which is now called the AccessDatabaseEngine. This engine is fully
'           backward-compatible with previous versions of the Jet engine, so it reads and writes (.mdb) files from earlier Access versions.
'           It introduces a new default file format, (.accdb), that brings several improvements to Access, including complex data types
'           such as multivalue fields, the attachment data type and history tracking in memo fields. It also brings security and encryption
'           improvements and enables integration with Microsoft Windows SharePoint Services 3.0 and Microsoft Office Outlook 2007
'           In addition, ACE provides a 64-bit driver, so can be used on 64-bit machines, whereas JET cannot.
'           The driver is not part of the Windows operating system, but is available as a redistributable(AccessDatabaseEngine.exe).
'           Previously the Jet Database Engine was only 32-bit and did not run natively under 64-bit versions of Windows.
'           download the ACE components separately. I got them from the link Microsoft Access Database Engine 2010 Redistributable.
'           This is likely because I had installed a 32-bit version of Office under 64-bit Windows; in any case,
'           the necessary files are easy to obtain from Microsoft.
'16.12.2017 Weglassen einer MsgBox wenn CompactDB nicht m�glich ist.
'           Bei der Auslieferung f�r die Access-Version muss AccessDatabaseEngine.exe mit ausgeliefert werden und vom
'           Setup-Paket aufgerufen werden
'----------------------------------------------------------------------------------------------------------
'10.12.2017 15.0.2 Verbesserung alle Versionen
'           Zustand: Video-Datei-Typ "MKV" und "FLV" wird bisher nicht akzeptiert
'           L�sung: Ab sofort wird "MKV" und "FLV" akzeptiert
'                   In MKV und FLV videos kann der Explorer keine Eigenschaften reinschreiben
'                   bei "MKV" und "FLV" gibt es keine Vorschaubilder
'----------------------------------------------------------------------------------------------------------
'07.01.2018 15.0.2 Verbesserung SQL-Server-Version
'           Zustand: Bei frmConnectSQL geht bei Klick auf btnConnect die Sanduhr an, aber nicht wieder aus
'           L�sung: Am Ende von btnConnect_Click wird die Sanduhr wieder ausgeschaltet
'----------------------------------------------------------------------------------------------------------
'07.01.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: Die Formulare �ffnen an unterschiedlichen Positionen, meist mit StartUpPosition=3=Windows-Standard.
'                   Ich will alle in Fenstermitte
'           L�sung: StartUpPosition=1=Fenstermitte
'21.01.2018         Ausnahmen sind frmGridAndThumb und KommentarForm (StartUpPosition=0=Manuell)
'----------------------------------------------------------------------------------------------------------
'08.01.2018 15.0.2 Nachbesserung zum 01.09.2017 alle Versionen
'           Zustand: Das Men� in Form1 und frmVideo ist nur Deutsch
'           L�sung: mnuxxx.Caption mit LoadResString laden
'----------------------------------------------------------------------------------------------------------
'09.01.2018 15.0.2 Fehlerkorrektur Nachbesserung zum 23.11.2017 alle Versionen
'           Zustand: DAO Querydefs gibt es nicht mehr bei ADO
'                   Ich war der Meinung ich kann die L�sung vom wie beim SQL Server benutzen, aber da kommen Laufzeitfehler
'           L�sung: DAO Querydefs m�ssen durch OpenSchema und ADODB.command ersetzt werden.
'                   Mann hat das gedauert.
'                   Hilfe gab es bei https://www.labath.org/docs/sys/mssql2000/mdacxml/htm/wpmigratingschema.htm
'                   Defining and Retrieving a Database�s Schema
'----------------------------------------------------------------------------------------------------------
'18.01.2018 15.0.2 Nachbesserung zum 11.03.2017 alle Versionen
'           Zustand: Ver�ndern der Schriftgr��e wirkt erst nach einem Neustart
'           L�sung: Durch Aufruf von 'Call AnpassenNutzerWunsch(frm)' wirkt das Ver�ndern der Schriftgr��e sofort
'----------------------------------------------------------------------------------------------------------
'20.01.2018 15.0.2 Fehlerkorrektur Nachbesserung zum 01.09.2017
'           Zustand: Wenn das H�kchen aus der Checkbox 'Fenstergr��e �nderbar' entfernt wird und man Schaltet �ber Rechtsklick
'                   die Men�zeile ein/aus -> dann kommen abwechselnd 2 Men�zeilen/1 Men�zeile.
'                   Das passiert nur in der exe, nicht in der IDE.
'           L�sung: es hat ein DoEvents gefehlt
'----------------------------------------------------------------------------------------------------------
'21.01.2018 15.0.2 Fehlerkorrektur Nachbesserung zum 04.03.2013
'           Zustand: MkDir funktioniert nicht im Unicode-Pfad wenn ein Ordner 'GerbingThumbs' erzeugt werden soll
'           L�sung: F�r MkDir gibt es ein Unicode-�quivalent CreateDirectoryW
'----------------------------------------------------------------------------------------------------------
'23.01.2018 15.0.2 Fehlerkorrektur Nachbesserung zum 10.04.2016
'           Zustand: Wenn der Dateiname Hochkomma enth�lt,
'                   dann kommt Laufzeitfehler sowohl bei ..Find wie bei Open Recordset mit Suche nach Dateiname
'           L�sung: Schlechte L�sung:
'                   In Fotosmdb.exe werden ab Version 15.0.2 Hochkommas im Dateinamen durch - ersetzt.
'                   Was passiert aber mit Dateinamen wo schon Hochkommas enthalten sind:
'                   1.fotos.exe ignoriert derartige Dateinamen zb in F5MehrereZeilen oder in KommentarForm oder in frmGridAndThumb beim
'                       synchronisieren Thumbnail-Ansicht mit Listen-Ansicht
'                   2.Renammdb.exe l��t gemeinsames L�schen im Ordner und der Datenbank zu, danach wird das Programm beendet
'                   3.Man muss alle Dateinamen mit Hochkomma im Fotoalbum-Ordner finden(zB mit Everything) -> in einen Retteordner kopieren->
'                       im Fotoalbum-Ordner l�schen ->
'                       Pr�fen1 ausf�hren -> nicht gefundene Datens�tze l�schen -> im Retteordner umnennen ohne Hochkomma -> umgenannte in den
'                       Fotoalbum-Ordner zur�ckkopieren -> Pr�fen3 ausf�hren
'           Schlechte L�sung wird zur�ckgedreht
'           Gute L�sung: Wo im Dateiname ein Hochkomma vorkommt, wird beim Aufbau des SQL-Strings nach 2 Hochkommas gesucht
'----------------------------------------------------------------------------------------------------------
'27.01.2018 15.0.2 Verbesserung alle Versionen
'           Zustand: Das Tool Diashow.exe kann nicht �ber die Men�leiste von fotos.exe erreicht werden
'           L�sung: Neuer Men�-Eintrag in Form Query: unter Tools -> Diashow starten
'----------------------------------------------------------------------------------------------------------
'12.02.2018 15.0.2 Fehlerkorrektur alle Versionen
'           Zustand: Es kommt Laufzeitfehler '3705' Der Vorgang ist f�r ein ge�ffnetes Objekt nicht zugelassen
'                   Das passiert wenn ich �bers Men� zuerst ausw�hle 'Namen Ersetzen...' dann nichts mache (Antwort nein)
'                   Dann �bers Men� ausw�hle 'L�schen markierte...' -> Laufzeitfehler '3705'
'           L�sung: on error resume next
'=========================================================================================================
'02.03.2018 15.0.3 Verbesserung alle Versionen
'           Zustand: Der GERBING Fotoalbum user k�nnte bei unaufmerksamem Arbeiten seine Datenbank-Datei fotos.mdb einb��en.
'                   Das k�nnte passieren beim UnInstall, wenn keine Rettekopie von fotos.mdb existiert.
'           L�sung: Ich lasse beim UnInstall das Programm backupdatabase.exe ausf�hren.
'                   Dort wird der user gefragt, ob er eine Rettekopie von fotos.mdb anlegen will.
'                   Wenn er schon eine hat kann er mit 'Nein' antworten.
'                   Es w�re auch gegangen nach dem Beenden mit Compact in Form Query die Datei 'Newfotos.mdb' existieren zu lassen,
'                       Aber auch da kann bei unaufmerksamem Arbeiten der user seine Datenbank-Datei fotos.mdb durchs UnInstall einb��en.
'----------------------------------------------------------------------------------------------------------
'25.03.2018 15.0.3 Fehlerkorrektur Professional Version
'           Zustand: Wenn mit Audio-Datei, dann klappt der Bildwechsel nicht bei Klick auf den Thumbnail
'                   er klappt jedoch bei Klick auf die Zeile im Grid
'           L�sung: es war falsch und �berfl�ssig 'Call FRODateiname auszuf�hren
'----------------------------------------------------------------------------------------------------------
'25.03.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Wenn mit dem aktuell gezeigten Bild nach RenamMdb verzweigt wird, kann dieses Bild weder gel�scht noch umgenannt
'                   werden. Es wird angemeckert, es w�re noch ge�ffnet.
'           L�sung: Das aktuell gezeigte Bild muss entladen werden, bis zum GdiplusShutdown m_lngInstance
'                   In Form1.MyDrawImage muss ausgef�hrt werden GdiplusStartup(m_lngInstance, udtData, 0)
'                   oh Wunder das Bild bleibt trotzdem angezeigt, obwohl
'                   lngPointer m_lngGraphics m_lngInstance alle gel�scht sind
'----------------------------------------------------------------------------------------------------------
'26.03.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Beim NamenErsetzen (siehe 11.04.2017) wurden .wav und .mp3 Dateien nicht ber�cksichtigt
'                   Diese Dateien verschwinden sogar aus dem GERBING Fotoalbum, bei Ausf�hrung von FotosMdb -> Pr�fenS
'                   Das ist ganz b�se, wenn man keine Rettekopien dieser Sound-Dateien hat.
'           L�sung: Prozedur AudioDateiMitUmnennen ausf�hren
'----------------------------------------------------------------------------------------------------------
'02.04.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Bei Videos mit Video-Ordner mit Unicode-Zeichen kommen in Zeile Bildbeschreibung ????? anstelle unicode
'           L�sung: Ich hatte vergessen in frmVideo eine unicode Textbox txtBildbeschreibung zu benutzen
'                   kopiert aus Form1
'----------------------------------------------------------------------------------------------------------
'02.04.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Bei Videos mit Video-Ordner mit Unicode-Zeichen kommen in der Titelleiste ????? anstelle unicode
'           L�sung: Es war ein Fehler in 'ShowTitleBar'
'                   falsch ist  frmVideo.Caption = Form1.FotoAlbumTitle
'                   richtig ist formCaption frmVideo.hWnd, Form1.FotoAlbumTitle
'----------------------------------------------------------------------------------------------------------
'25.04.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Nicht zu fassen, da hat doch der Programmierer 4 Jahre lang nicht bemerkt, dass die Rechteck-Lupe kein Rechteck zeichnet.
'                   Oder es nicht bemerken wollen, weil er keine L�sung hatte siehe 23.10.2014
'           L�sung: So was von einfach. Bisher war das Shape1 Control auf der Form1 angeordnet.
'                   Es muss aber innerhalb von Picture1 angeordnet werden.
'                   gblnMouseIconSquare wird wieder entfernt.
'----------------------------------------------------------------------------------------------------------
'29.04.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Wiederholbarer Fehler, wenn nach Bildern > 3000 Pixel gesucht wird und die Taste F3 festgehalten wird
'                   Es kommt kein runtime error, sondern VB6 st�rzt ab. Das passiert bei einem gdiplus-Aufruf. Das sagt der Visual C++
'                   Debugger, nachdem ich eine Datei Fotos.pdb habe erzeugen lassen. Es ist der 5. Aufruf in der Prozedur MyDrawImage
'                   In Version 15.0.2 passiert das nicht
'           L�sung: ich habe einen Abschnitt auskommentiert und bei jedem fehlerhaften gdiplus-Aufruf Msgbox Fehlermeldungen eingef�gt
'----------------------------------------------------------------------------------------------------------
'04.05.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Ich habe eine zus�tzliche Idee, wie ich verhindern kann, dass bei unaufmerksamer Neuinstallation die
'           bereits bearbeitete fotos.mdb verschwindet
'           L�sung: Die Beispiel-Datenbank mit 3 Bildern wird ausgeliefert als fotosStart.mdb
'                   Daduch wird bei der Deinstallation versucht die fotosStart.mdb zu deinstallieren die fotos.mdb bleibt
'                   Bei der ersten Benutzung fragen, ob es fotos.mdb gibt.
'                   Bei 'Nein' -> fotosStart.mdb umnennen in fotos.mdb
'                   Bei 'Ja' -> fotos.mdb bleibt unangetastet
'                   Fehler, wenn es weder fotos.mdb noch fotosStart.mdb gibt
'18.05.2018 Nacharbeiten n�tig, weil in der SQL Server Version kein Fehler vorliegt, wenn es weder fotos.mdb noch fotosStart.mdb gibt
'           L�sung: Ob Vollversion=SQL Server Version vorliegt erkenne ich an Datei-Eigenschaften -> Details -> Copyright wenn dort
'                   ganz rechts "-1" steht
'                   Ich benutze den Modul modFileInfo
'----------------------------------------------------------------------------------------------------------
'17.05.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Problem: F�r das Men� 'Start Diashow gibt es keine zweisprachige Version
'           L�sung: mnuDiashow.Caption = LoadResString(3193 + Sprache)    'Diashow starten
'----------------------------------------------------------------------------------------------------------
'29.09.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: GPS die Google Maps API gehen seit Juni 2018 nur noch mit einem API key zu benutzen, heute noch kostenlos,
'                   aber wer wei� wie lange?
'                   Ich habe einen key, den benutze ich in ucGMap in frmGEOFinden
'           L�sung: Ich steige um auf OpenStreetMap, da gibts auch ein API(Frameworks/Overpass API)
'                   neue Form: frmMap
'                   zoom=14 -> zoom=16
'                   MinButton = False
'           in frmMap Form_KeyDown zu erlauben hat im Windows10 zum kommentarlosen Programmende nach 2-3 Fotos gef�hrt
'           keine L�sung: zuerst Me.Hide dann weiterleiten Form_KeyDown dann Unload me
'           L�sung: frmMap Form_KeyDown auskommentiert
'                   gestrichene Form: frmGEOPosition
'----------------------------------------------------------------------------------------------------------
'16.10.2018 15.0.3 Nachbesserung zum 29.09.2018
'           Zustand: In VM Win7 geht die L�sung mit frmMap und OpenStreetMap nicht
'                   Es kommt 'In dem Skript auf dieser Seite ist ein Fehler aufgetreten'
'                   .
'                   URL: https://www.openstreetmap.org/assets/application-.......js
'           L�sung1: Ich brauche eine neue Form frmStrgG, wenn Strg + G gedr�ckt wird
'                   Der Nutzer soll ausw�hlen k�nnen ob mit frmMap oder mit ShellExecute
'           L�sung2: Die VM muss Microsoft Internet Explorer 11 installieren oder wenn ShellExecute benutzt wird, mu� Google Chrome installiert sein
'
'----------------------------------------------------------------------------------------------------------
'29.11.2018 15.0.3 Verbesserung Professional Version
'           Zustand: Es gibt gegenw�rtig keine Software, die die GPS-Angaben eines Smartphone MP4-Videos auf einer Landkarte anzeigen kann.
'                   Ebenso gibt es gegenw�rtig keine Software, die die GPS-Daten eines Smartphone MP4-Videos unangetastet l��t.
'                   beim Editieren oder Schneiden oder Zusammenf�gen gehen die GPS-Daten eines Smartphone MP4-Videos verloren.
'                   Die GPS-Daten verschwinden sogar beim Zurechtschneiden auf dem Smartphone.
'           L�sung: Ich muss mit dem Smartphone viele kurze Clips herstellen und diese weder editieren, schneiden noch zusammenf�gen
'                   1. Smartphone-Videos auf den Computer kopieren
'                   2. Bleibende Namen vergeben, solche wie sie im Foto/Video-Album stehen sollen, das geht leider nicht mit Diashow.exe
'                   3. Beim Aufnehmen von mp4 files mit MediaInfo.DLL nach dem Feld "xyz" suchen, das ist das GPS-Feld. Was dort steht,
'                      wandert in die Datenbank-Felder GPSLatitude und GPSLongitude.
'                      Es passiert nichts, wenn in der Datenbank die Felder GPSLatitude und GPSLongitude fehlen.
'                      Ich ignoriere die Felder Exif-GPSLatitude und Exif-GPSLongitude
'                   4. Ab jetzt ist editieren, schneiden und zusammenf�gen mit Smartphone MP4-Videos m�glich
'                   5. Wenn ein Name ver�ndert werden soll, dann nur mit RenamMdb
'                   6. In fotos.exe kann in der Professional Version bei Dr�cken von Strg+G bei einem mp4 video eine Landkarte gezeigt werden
'           Organisatorische L�sung: Man muss neben dem Smartphone-Video stets auch einige Smartphone-Fotos machen.
'                   Diese k�nne die GPS-Daten dauerhaft speichern.
'----------------------------------------------------------------------------------------------------------
'30.11.2018 15.0.3 Fehlerkorrektur alle Versionen
'           Zustand: Die fotos.mdb ist verschwunden. Ganz b�ser Fehler. Aber Newfotos.mdb ist noch da.
'                   Zuletzt ist ein fehlerhaftes Video dran gewesen.
'                   Dann wurde das Programm beendet. Beim Beenden sollte eine Datenbank-Komprimierung ausgef�hrt werden.
'                   Dann kam run time error in frmVideo.WMP_PlayStateChange
'           L�sung: wenn Schalter gblnComeFromBeenden ein ist, dann frmVideo.WMP_PlayStateChange sofort wieder verlassen
'=========================================================================================================
'24.03.2019 15.0.3 ich habe vermutlich eine falsche 'fotos.mdeutsch.Auslieferung.mdb' ausgeliefert
'           Es kommt Fehler
'           Fehler Nr.: -2147467259
'           Unrecognized database format 'C:\GERBING_FotoAlbum_15\fotos.mdb'. Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\GERBING_FotoAlbum_15\fotos.mdb;
'           es ist n�tig eine mit dem Stand Provider=Microsoft.ACE.OLEDB.12.0 auszuliefern
'=========================================================================================================
'08.04.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Obwohl in einem Foto GPS-Daten eingetragen sind(Kontrolle mit Windows-Explorer -> Eigenschaften -> Details -> �berschrift GPS)
'                   kommt eine MsgBox 'Geo positions not available'
'                   Das betraf die Fotos von PahrenHellau
'           Ursache: Die GEO-Positionen sind im XMP-Abschnitt des Fotos eingetragen.
'                   Der XMP-Abschnit beginnt mit '<?xpacket begin' und endet mit '<?xpacket end'
'                   Das macht zB Geosetter(mit Hilfe von Exiftool),
'                   ich suche sie aber nur im EXIF-Abschnitt.
'                   Andere Software findet diese GEO-Positionen zB ExifToolGUI, PSP 2019, Irfan View, Fotos App, XnViewMP.
'                   Fotos App(Win10) korrigiert sogar selbst�ndig aus dem XMP-Abschnitt in den EXIF-Abschnitt
'           L�sung: Da ich in clsEXIF sowieso jedes JPG-Foto durchsuche, um den EXIF-Abschnitt zu finden, kann ich dort ebenso nach den
'                   XMP-GEO-Positionen suchen
'                   Ich suche nach exif:GPSLatitude und exif:GPSLongitude mit InstrB, weil das rasend schnell geht
'                   Die gefundenen Werte gstrLatXMP und gstrLongXMP muss ich dann noch in ein Format verwandeln, das OpenStreetMap versteht
'                   zB gstrLatXMP 50,38.7309456N -> 50.64551575
'                   zB gstrLongXMP 11,53.9826786E -> 11.89971130
'                   Genauso arbeite ich in frmFeldAktualisierung beim nachtr�glichen Aktualisieren
'----------------------------------------------------------------------------------------------------------
'29.04.2019 15.0.4 Fehlerkorrektur Vollversion Portable
'           Ursache: eine portable Vollversion meldet sich als Shareware-Version, weil msplugin.log und gerbingsoft.log fehlen
'           L�sung: siehe 18.05.2018 Ob Vollversion vorliegt erkenne ich an Datei-Eigenschaften -> Details -> Copyright wenn dort
'                   ganz rechts "-1" steht
'                   Ich benutze den Modul modFileInfo
'                   Das geht aber nicht in der IDE, weil die Zeile Copyright in der exe abgefragt wird
'----------------------------------------------------------------------------------------------------------
'29.04.2019 15.0.4 Fehlerkorrektur Vollversion Portable
'           Ursache: Wenn fotos.exe umbenannt wird in FotosPortable.exe dann bringt die AboutForm Fehler
'           L�sung: Mit MsgBox darauf hinweisen, dass fotos.exe nicht da ist
'----------------------------------------------------------------------------------------------------------
'08.05.2019 15.0.4 Fehlerkorrektur Professional Version
'           Zustand: Bei Suche in nutzerdefinierten Feldern vom Typ Text gibt es bisher als Vergleichsoperand
'                   = > < >= <= <>
'                   Es fehlt like
'           L�sung: Wenn like, dann auch %
'----------------------------------------------------------------------------------------------------------
'10.05.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Bei sehr gro�en JPG-Fotos ist das Drehen mit dem Explorer ohne Auswirkung.
'                   Im Explorer-Icon kann man die erfolgte Drehung sehen.
'                   In fotos.exe erfolgt keine Drehung, aber jede andere Software macht es.
'                   Die vom Explorer durchgef�hrte Drehung wird in EXIF.Orientation eingetragen, aber von fotos.exe nicht ausgewertet
'                   Es gibt 8 m�gliche Werte siehe 'Readme EXIF PropertyTagOrientation.docx'
'           Notl�sung: Das Foto verkleinern, dann wird es von fotos.exe korrekt angezeigt
'           L�sung: Das nichtssagende 'mirrored or turned' bzw 'horizontal normal' wird erweitert um die vorangesetzte EXIF.Orientation clsEXIF.cls
'                   und gleichzeitig wird der Wert von EXIF.Orientation gespeichert in gstrEXIFOrientation.
'                   Abh�ngig von gstrEXIFOrientation wird in frmBildMitGDIPlus.MyDrawImage das Foto gedreht mit GdipImageRotateFlip
'                   Durch das Drehen haben sich Width und Height des Bildes ver�ndert. Mit GdipGetImageDimension muss ich diese neu ermitteln.
'                   Bei Bildern im Hoch-Format muss ich einen dblKorrekturFaktor ausrechnen, weil sonst der untere Bildteil nicht mit gezeigt wird.
'                   Dann kann das Bild wie bisher gezeichnet werden.
'                   Bei der  Rechteck-Lupe brauche ich SaveMyZoomPercent
'           Mangel: oder Nicht-Mangel: das Vorschaubild Thumbnail wird nicht gedreht dargestellt
'----------------------------------------------------------------------------------------------------------
'28.05.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: fotos.mdb war verschwunden, aber keine Ursache zu finden
'                   M�glicherweise gab es einen Fehler, dabei kam Query.Beenden dran
'                   dort sollte Datenbank-Komprimierung gemacht werden, dabei ist fotos.mdb verschwunden
'           versuchte L�sung: bei Query.Beenden keine Datenbank-Komprimierung machen - auskommentiert
'                   r�ckg�ngig gemacht
'           L�sung: Bei Programmstart
'                   1.L�schen fotos_copy.mdb
'                   2.fotos.mdb kopieren in fotos_copy.mdb - Das funktioniert auch bei schreibgesch�tzter Datenbank auf CD, dort
'                       kommt lediglich der Hinweis 'Schreibgesch�tzt' - Wollen Sie weiterarbeiten?
'----------------------------------------------------------------------------------------------------------
'30.05.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: das Video '1901\ZJanuarRabensteinSchneeResteHoppberg.mp4' hat irgendeinenFehler, der zu st�ndigem Beenden/Neustarten f�hrt
'                   frmVideo_PlayStateChange NewState= 3 8 9 1 ... in st�ndiger Schleife
'                   andere Video-Abspiel-Software:
'                   Windows-Media-Player spielt es nicht
'                   andere Microsoft-Software spielt es nicht
'                   Irfan View spielt es nicht
'                   VLC-Media-Player spielt es
'           Not-L�sung: Man muss das Programm abw�rgen
'           L�sung: ich kontrolliere, ob kurz nach frmVideo_PlayStateChange NewState=3=playing
'                   frmVideo_PlayStateChange NewState=8=MediaEnded kommt, das melde ich als Fehler
'----------------------------------------------------------------------------------------------------------
'04.07.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Nacharbeiten zum 08.04.2019
'                   Die GPS-Daten in den EXIF-Feldern werden richtig angezeigt, aber sind fehlerhaft in den Datenbank-Feldern
'                   Ursache ist ein Fehler in Form1.GEOKoordinatenUmrechnenXMP
'                   Beispiel: Minuten = 0.287564
'                   Nachkomma = MinutenDouble / 60 'liefert Ergebnis=0
'           L�sung: Wenn Komma als Dezimaltrennzeichen verwendet wird, muss der Punkt im String Minuten in Komma verwandelt werden
'                   sonst kommt bei MinutenDouble / 60 Ergebnis=0
'----------------------------------------------------------------------------------------------------------
'05.07.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Checkbox 'Gespeicherte Abfragen benutzen' liefert ein Ergebnis in deutschsprachiger exe
'                   Checkbox 'use saved queries' liefert kein Ergebnis bei englischsprachiger exe
'           L�sung: Zus�tzliche Untersuchung von DBado.OpenSchema(adSchemaProcedures) auf PROCEDURE_NAME
'----------------------------------------------------------------------------------------------------------
'08.07.2019 15.0.4 Fehlerkorrektur alle Versionen
'           Zustand: Nacharbeiten zum 29.09.2018
'                   am 29.09.2018 konnte ich die Form frmMap nicht auf Eingabe von F-Tasten reagieren lassen -> Programmabbruch
'           L�sung: Heute geht es, aber nur in der IDE. Als Exe geht es nicht. Ursache ungekl�rt
'----------------------------------------------------------------------------------------------------------
'30.09.2019 15.0.4 Fehlerkorrektur Nacharbeiten zum 10.05.2019
'           Zustand: Bei Hochkantfotos aus Katjas Handy oder bei Panoramafotos aus Jens Handy
'                   werden diese nur in halber H�he oder halber Breite dargestellt
'           L�sung: zuerst wird gstrEXIFOrientation ermittelt
'                   dann kommt sofort GdipImageRotateFlip
'                   wegen GdipImageRotateFlip �ndert sich Width und Height des Bildes
'                   erst danach X und Y ausrechnen
'=========================================================================================================
'02.10.2019 15.0.5 Verbesserung Professional Version Neue Funktion
'           Zustand: Bisher muss ich Fremd-Software zu Hilfe nehmen um die GEO-Position zu einem Foto nachtr�glich
'                   festzulegen. F�r Videos gibt es noch keine brauchbare Fremd-Software.
'                   Bei JPG-Fotos tragen zB Picasa oder GeoSetter die GEO-Position in den EXIF-Abschnitt ein.
'                   Anschlie�end kann ich mit Men� Datei.. -> Feldaktualisierung durch Import-Wiederholung
'                   die Datenbank-Felder GPSLatitude und GPSLongitude auff�llen, das bleibt auch so.
'           L�sung: Es gibt die neue Form frmGPSInDatenbankEintragen.
'                   Jetzt kann jede Datei in der Datenbank mit den Feldern GPSLatitude und GPSLongitude versehen werden.
'                   Diese Form wird aufgerufen, wenn der Nutzer Ctrl+G dr�ckt und 'Keine Geo-Position vorhanden' als Antwort erh�lt
'                   und wenn er 'Wollen Sie eine Geo-Position eintragen?' mit JA beantwortet
'                   oder sie wird aufgerufen �ber den neuen Menupunkt Datei... -> Einf�gen Geo-Position.
'                   Dort wird ein WebBrowser Control mit OpenStreetMap gef�ttert.
'                   Die Start-Position f�r OpenStreetMap kommt aus der Tabelle ErsterStart Feld LetzterGeoPunkt und ZoomListIndex
'                   Wenn Feld LetzterGeoPunkt und ZoomListIndex leer sind, wird OpenStreetMap mit der Weltkarte gestartet.
'                   Das ist https://www.openstreetmap.org/#map=2/0/0
'                   Die vom Nutzer ins Feld Geo-Position kopierte Geo-Position wird
'                   in die Tabelle Fotos - Datenbankfelder GPSLatitude und GPSLongitude eingetragen
'                   und in Tabelle ErsterStart - Datenbankfelder LetzterGeoPunkt und ZoomListIndex
'
'           Zustand: Die Google Maps API macht �rger. J�hrlich muss der Quellcode mit einem neuen key codiert werden.
'                   Das ist Schei�e und unzumutbar.
'                   F�r die Suche in einem virtuellen Rechteck auf der Landkarte weiche ich aus auf OpenStreetMap.
'           L�sung: Es gibt jetzt die neue Form frmGPSRechteck.
'                   Diese wird aufgerufen, wenn auf Suche nach nutzerdefinierten Feldern geklickt wird.
'                   Dort wird ein WebBrowser Control mit OpenStreetMap gef�ttert.
'                   Die Start-Position f�r OpenStreetMap kommt aus der Tabelle ErsterStart Feld LetzterGeoPunkt und ZoomListIndex
'                   Wenn Feld LetzterGeoPunkt und ZoomListIndex leer sind, wird OpenStreetMap mit der Weltkarte gestartet.
'                   Das ist https://www.openstreetmap.org/#map=2/0/0
'                   Der Nutzer muss auf der Karte ein virtuelles Rechteck definieren von links oben nach rechts unten
'                   Bei korrekter Definition werden die Strings gstrGEOStartPunkt und gstrGEOEndPunkt gebildet,
'                   mit den dann der Suchstring verkettet wird.
'           Problem: Ich kann manuell keine Werte im DataGrid aus den Spalten GPSLatitude und GPSLongitude in eine andere Zeile kopieren
'                   aus einer Zahl mit Komma wie zB 50,4367 wird eine Zahl ohne Komma gemacht wie zB 504367
'                   deshalb verbiete ich das manuelle Ver�ndern dieser Spalten
'                   in �lteren Versionen ging das, Ursache unbekannt
'----------------------------------------------------------------------------------------------------------
'04.10.2019 15.0.5 Fehlerkorrektur
'           Zustand: Nur wenn vorher F12 gedr�ckt wurde,
'                   und anschlie�end wird Taste Strg gedr�ckt, da geht frmVideo auf (viel schwarz)
'           L�sung: in Form1.Form_Keydown bei Case vbKeyF12
'                   auskommentiert
                    '            If gblnComefromVideo = False Then                                               'Gerbing 23.10.2013 04.10.2019
                    '                frmVideo.lblLeereForm.Visible = True
                    '            End If
'----------------------------------------------------------------------------------------------------------
'12.10.2019 21.11.2019 15.0.5 Neuerung
'           Zustand: Ich will ab sofort keine Shareware-Version und keine Professional Version mehr pflegen. Schade um den Aufwand.
'                   Elke verkauft mehr Leseknochen als ich je Software verkauft habe.
'                   Es soll nur noch eine Freeware Vollversion geben. Die SQL-Server-Version wird nicht kostenlos.
'                   W�re aber kostenlos m�glich mit einer 99-Lizenz.
'           L�sung: �nderungen in frmGeschichteDieserSoftware
'                   �nderungen in der Website
'                   �nderungen in der Hilfe
'                   fotos.mdb mit den Feldern GPSLatitude und GPSLongitude und ExifDateTimeOriginal und VideoDuration ausliefern.
'                             Tabelle UserDefined mit Feld-Zuordnungen GPSLatitude und GPSLongitude und ExifDateTimeOriginal ausliefern
'                   gblnProversion = true und gblnVollVersion = true gleich zu Programmbeginn einschalten und nie wieder aus
'                   Dateien gerbingsoft.log, msplugin.log, msdmo.log werden nicht mehr erzeugt und nicht abgefragt
'----------------------------------------------------------------------------------------------------------
'20.10.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Wenn der Nutzer sehen will, ob das Eintragen einer GEO-Position erfolgreich war,
'                   muss er den Refresh-Button in frmGridAndThumb clicken
'           L�sung: Das Programm l�st selber ein frmGridAndThumb.btnRefresh_Click aus
'----------------------------------------------------------------------------------------------------------
'23.10.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Bisher verlange ich vom Nutzer, dass er die Felder GPSLatitude/GPSLongitude selber erzeugt mit MS Access
'                   Das hat in anderen �hnlichen F�llen das Programm selbst gemacht.
'           L�sung: Falls GPSLatitude/GPSLongitude nicht angelegt sind, legt das Programm sie an sowohl bei der Access-Version wie beim SQL Server
'----------------------------------------------------------------------------------------------------------
'29.10.2019 Zustand: Leute wie Streubs brauchen auch im Listenfenster(frmGridAndThumb) Hilfe durch die Form Hilfebox
'           L�sung: An m�glichst vielen Stellen in frmGridAndThumb auf rechten Maus-Klick reagieren
'----------------------------------------------------------------------------------------------------------
'14.11.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Bisher verlange ich vom Nutzer, dass er das Feld EXIFDateTimeOriginal selber erzeugt mit MS Access
'                   Das hat in anderen �hnlichen F�llen das Programm selbst gemacht.
'           L�sung: Falls EXIFDateTimeOriginal nicht angelegt ist, legt das Programm es an sowohl bei der Access-Version wie beim SQL Server
'----------------------------------------------------------------------------------------------------------
'14.11.2019 15.0.5 Nachbesserung zum 02.10.2019
'           Zustand: Bisher verlange ich vom Nutzer, dass er das Feld VideoDuration selber erzeugt mit MS Access
'                   Das hat in anderen �hnlichen F�llen das Programm selbst gemacht.
'           L�sung: Falls VideoDuration nicht angelegt ist, legt das Programm es an sowohl bei der Access-Version wie beim SQL Server
'----------------------------------------------------------------------------------------------------------
'15.11.2019 15.0.5 Verbesserung
'           Zustand: Ich kann bisher im Listenfenster frmGridAndThumb.DbGridNeu nicht mit dem Mausrad scrollen
'           L�sung: aus dem Internet abgeschrieben
'                   DBGridNeu.MarqueeStyle = dbgHighlightCell
'----------------------------------------------------------------------------------------------------------
'18.11.2019 15.0.5 Verbesserung
'           Zustand: Es gibt gegenw�rtig nur eine mir bekannte Software, die die GPS-Angaben eines Smartphone MP4-Videos oder MOV-Videos
'                   von der digitalen Kamera nach dem Editieren unangetastet l��t. Ebenso das Feld 'Encoded Date'.
'                   Das ist im Windows 10 die Fotos App von Microsoft.
'                   Andere Software macht folgendes:
'                   beim Editieren oder Schneiden oder Zusammenf�gen gehen die GPS-Daten eines Smartphone MP4-Videos verloren.
'                   Die GPS-Daten verschwinden sogar beim Zurechtschneiden auf dem Smartphone.
'           L�sung: 1. Ich muss zum Editieren von mp4 oder mov videos im Windows 10 die Fotos App von Microsoft benutzen.
'                   2. Bei Import-Wiederholung von mp4 oder mov files mit MediaInfo.DLL(must be i386 version, getestet mit Version 18.8.1.0)
'                      nach dem Feld "xyz" suchen,
'                      das ist das GPS-Feld. Was dort steht, wandert in die Datenbank-Felder GPSLatitude und GPSLongitude.
'                      Ich ignoriere die Felder Exif-GPSLatitude und Exif-GPSLongitude, weil dort eh nichts steht.
'                   3. Nach dem Feld 'Encoded Date' suchen. Was dort steht, wandert in das Datenbank-Feld ExifDateTimeOriginal.
'           Notl�sung: Manuell die Geo-Position nachtragen mit Strg+G
'=========================================================================================================
'08.03.2020 16.0.0 Fehlerkorrektur
'           Zustand: B�ser Fehler
'                   Beim Start von fotos.exe passiert scheinbar garnichts.
'                   fotos.exe startet kurz(das merkt aber keiner) und beendet sich selbst ohne jegliche Mitteilung.
'                   Die Ursache ist dass JRO.CompactDatabase nicht ausgef�hrt wird und sich selbst ohne jegliche Mitteilung beendet.
'                   Der Fehler verschwindet, wenn man AccessDatabaseEngine.exe erneut ausf�hrt.
'                   Office 2019 Updates macht das zunichte was AccessDatabaseEngine.exe aufbaut.
'                   das ist passiert nach der Installation eines Office 2019 Updates.
'           L�sung: Ich brauche eine Kontrolle, ob CompactDatabase ausgef�hrt worden ist
'                   Zur Kontrolle, ob CompactDB ausgef�hrt werden konnte wird vorher CheckCompactDatabase.exe gestartet
'                   Wenn CheckCompactDatabase.exe nach 10 Sekunden feststellt, dass in der fotos.ini
'                   CompactDatabaseEnded <> 1 ist, dann muss CheckCompactDatabase.exe die Meldung bringen,
'                   dass AccessDatabaseEngine.exe wiederholt werden muss und vom Nutzer auch gleich gestartet werden kann.
'                   Die fotos.ini erh�lt zwei neue Felder und einen neuen Abschnitt
'                   [CheckCompactDatabase]
'                   CompactDatabaseStarted= 0 oder 1
'                   CompactDatabaseEnded= 0 oder 1
'----------------------------------------------------------------------------------------------------------
'15.04.2020 16.0.0 Verbesserung
'           Zustand: Wenn GPS-Daten vorhanden sind werden sie mit OpenStreetmap angezeigt.
'                   Das geht manchmal sehr schleppend.
'                   Das hat den Nachteil dass nicht auf Satellitenansicht umgeschaltet werden kann, und auch nicht auf Google Street View.
'           L�sung: Als dritte Variante wird Google Maps Anzeigen im Browser angeboten
'----------------------------------------------------------------------------------------------------------
'20.04.2020 16.0.0 Fehlerkorrektur
'           Zustand: Eingabe Feld Dateiname= "E" - Es kommt eine MsgBox '�nderungen in diesem Feld sind verboten'
'                   Der Dateiname wird aber trotzdem ge�ndert, und das zieht massenhaft weitere Fehler nach sich(siehe n�chster).
'                   Das ist so seit ich DBGridNeu mit dem Mausrad scrollbar gemacht habe.
'           L�sung: Die Prozedur DBGridNeu.Change darf nicht zur �berpr�fung benutzt werden, sondern die Prozedur DBGridNeu.BeforeColUpdate
'----------------------------------------------------------------------------------------------------------
'20.04.2020 16.0.0 Fehlerkorrektur
'           Zustand: Laufzeitfehler '5' beim Programmstart
'                   Es passiert beim F�llen der Combobox TFileType.
'                   Wenn hier ein Dateiname auftaucht, der keinen Punkt besitzt, dann kommt der Fehler.
'           L�sung: Abfragen ob Pos = 0
'----------------------------------------------------------------------------------------------------------
'03.05.2020 16.0.0 Fehlerkorrektur Nachbesserung zum 08.03.2020
'           Zustand: Die SQL Server Version erwartet die CheckCompactDatabase.exe in AppPath. Die wird nicht gebraucht.
'           L�sung: Call SpracheFestlegen muss fr�her drankommen
'----------------------------------------------------------------------------------------------------------
'08.05.2020 16.0.0
'           Zustand: Die Trefferauswahl ist unsichtbar
'           L�sung: Neu compilieren. Es ist keine Ursache erkennbar, warum es mit der neuen exe pl�tzlich geht
'----------------------------------------------------------------------------------------------------------
'08.05.2020 16.0.0 Fehlerkorrektur
'           Zustand: Ich dr�cke die Num-Taste beim Start von fotos.exe weil ich eine fremde mdb ausw�hlen will.
'                   Scheinbar tut sich nichts. Ich sehe die Sanduhr. Es blinkt auch nichts auf der Task-Bar.
'                   Der ShowOpen Dialog ist hinter allen anderen offenen Fenstern versteckt. Ich muss die anderen Fenster minimieren.
'           L�sung: Nur wenn mit fremder mdb gearbeitet werden soll, mache ich
'                       AppActivate Me.Caption
'                   AppActivate erfordert aber, dass es eine Form geben muss in der der angegebenen Titel vorkommt
'                   Das mache ich unmittelbar vorher mit
'                       Me.Caption = "GERBING Fotoalbum"
'                       Me.Show
'                   Jetzt kommt entweder sofort der Showopen Dialog oder es blinkt auf der Task-Bar
'----------------------------------------------------------------------------------------------------------
'02.07.2020 16.0.0 Fehlerkorrektur
'           Zustand: Bei Dr�cken der Tasten Strg+G(Geo-Position) landet manchmal der Buchstabe 'g' in der Zelle, die in frmGridAndThumb
'                   gerade aktiv ist. Das kann die Spalte 'Merker' sein oder 'Ort'(wenn dieser bisher leer war)
'                   In der IDE ist das nur reproduzierbar, wenn die IDE nicht als Administrator gestartet wird.
'           L�sung: In frmGridAndThumb.DBGridNeu_KeyDown Keycode = 0 setzen
'----------------------------------------------------------------------------------------------------------
'11.07.2020 16.0.1 kosmetische Verbesserung
'           Zustand: In der Form MP(Weitere Filter) wird das Wort Filter nicht benutzt
'           L�sung: das Wort Filter benutzen
'----------------------------------------------------------------------------------------------------------
'28.08.2020 16.0.1 kosmetische Verbesserung
'           Zustand: Num+Strg+N Zeile Bildbeschreibung einschalten soll sofort wirksam sein, nicht erst beim n�chsten Bild
'                    Num+Strg+M Zeile Bildbeschreibung ausschalten soll sofort wirksam sein, nicht erst beim n�chsten Bild
'           L�sung: nicht nur gblnBildBeschreibung = True oder False einschalten sondern anschlie�end noch
'                   If gblnComefromVideo = True Then                                            'Gerbing 28.08.2020
'                       Call VideoAbspielen                                                     'Gerbing 28.08.2020
'                   Else
'                       Call MyDrawImage(gstrFRODN, glngZoomProzent)                            'Gerbing 28.08.2020
'                   End If
'           Zustand: Num+Strg+N wirkt nicht bei Videos
'           keine L�sung:
'                   STRG+UMSCHALT+N bedeutet beim Windows Media Player Wiedergabe mit Normalgeschwindigkeit
'           Umgehungsl�sung:
'                   �ber die rechte Maustaste und Hilfebox gehen
'----------------------------------------------------------------------------------------------------------
'12.09.2020 16.0.1 Fehlerkorrektur
'           Zustand: Nach dem Dr�cken des Refresh-Buttons oder programmatischem Ausf�hren des Refresh-Buttons
'                   kann es sein, dass das aktuelle Bild oder andere Datei nicht mehr den Such-Kriterien entspricht.
'                   Beispielsweise wurde mit den Suchkriterien nach Person = Elke gesucht
'                   dann wird Elke in Helene ge�ndert und der Datensatz gesichert
'                   dann wird der Refresh-Button bet�tigt
'                   dann soll mit Num+F5 die Form F5MehrereZeilen ge�ffnet werden
'                   Jetzt fehlt ein Fehler-Hinweis, dass das aktuelle Bild nicht mehr den Such-Kriterien entspricht
'           L�sung: Fehlerhinweis ausgeben









'**********************************************************************************************************
'           Diverse Probleme
'           seit XP SP3 werden vom MSI-Installer 3 tlb-Dateien installiert
'           amcompat.tlb
'           msado25.tlb
'           nscompat.tlb
'           Wenn diese 3 Dateien fehlen, kommt beim Start von fotos.exe Laufzeitfehler'13' Typen unvertr�glich
'           Problem tritt nicht mehr auf ab Version 13.5.0
'----------------------------------------------------------------------------------------------------------
'           In Vista und Windows 7 festgestellte Probleme
'           Mit aktivierter UAC = User Access Control = Benutzerkonstensteuerung werden �nderungen an fotos.ini
'           in C:\Benutzer\Maloi\AppData\Local\VirtualStore\fotos.ini gemacht
'           Ohne UAC werden die �nderungen an fotos.ini im Installationsordner(C:\Programme...) gemacht
'           Mit UAC hat jeder Nutzer einen eigenen Ordner ...\AppData\Local\VirtualStore\
'           wo s�mtliche Dateien stehen die �nderbar sind. Auch die Datenbankdatei fotos.mdb. Damit w�re eine
'           gemeinsame Nutzung der Datenbank durch mehrere Nutzer unm�glich.
'           Oder man mu� die Benutzung von
'           ...\VirtualStore\ ausschalten durch folgende Mittel:
'               siehe http://www.jondavis.net/techblog/post/2008/02/Beyond-Disabling-UAC-Disable-Virtual-Store.aspx
'
'               oder man muss einen anderen Installations-Ordner w�hlen als den
'               Standard-Installationsordner C:\Programme\...(das habe ich �berpr�ft und es stimmt),
'
'               oder man muss nach der Installation durch Kopieren die 3-Einigkeit
'               von Programm-Datenbank-Fotos au�erhalb von C:\Programme\... herstellen (das habe ich �berpr�ft und es stimmt),
'
'               oder man benutzt das Microsoft Application Compatibility Toolkit 5.5 um ein einzelnes Programm ohne UAC zu starten.
'
'               Oder ich als Programmierer f�ge ein Manifest in die fotos.exe ein mit requestedExecutionLevel. Bei Benutzung eines Manifestes
'               mit beliebigen Parametern wird ...\VirtualStore\ nicht verwendet.
'
'           Ergebnisse des Tests mit Manifest und Installationsordner(C:\Programme...):
'           1. Als Nutzer mit Administratorrechten
'                                       Bei Desktop-Icon-Click                                                          wird VirtualStore erzeugt
'               asInvoker               verlangt Admin-Rechte, aber im Rechtsklick-Men� gibts keine Admin-Rechte        nein
'                                       Ausweg: Starten fotos.exe �ber das Explorer-Fenster als Admin
'               highestAvailable        es kommt ein Zusatzfenster mit Auswahlm�glichkeit                               nein
'                                       Abbrechen oder
'                                       Zulassen
'               requireAdministrator    es kommt ein Zusatzfenster mit Auswahlm�glichkeit                               nein
'                                       Abbrechen oder
'                                       Zulassen
'               ohne Manifest           startet sofort                                                                  ja
'
'           2. Als Nutzer ohne Administratorrechte
'                                       Bei Desktop-Icon-Click                                                          wird VirtualStore erzeugt
'               asInvoker               verlangt Admin-Rechte, aber im Rechtsklick-Men� gibts keine Admin-Rechte        nein
'                                       Ausweg: Starten fotos.exe �ber das Explorer-Fenster als Admin
'               highestAvailable        verlangt Admin-Rechte, aber im Rechtsklick-Men� gibts keine Admin-Rechte        nein
'                                       Ausweg: Starten fotos.exe �ber das Explorer-Fenster als Admin
'               requireAdministrator    es kommt ein Zusatzfenster mit Auswahlm�glichkeit                               nein
'                                       Abbrechen oder
'                                       Zulassen nach Pa�worteingabe f�r ein Administratorkonto
'               ohne Manifest           verlangt Admin-Rechte, aber im Rechtsklick-Men� gibts keine Admin-Rechte        ja
'                                       Ausweg: Starten fotos.exe �ber das Explorer-Fenster als Admin
'
'           Ergebnisse des Tests mit Manifest, aber mit ausgeschalteter UAC (Benutzerkontensteuerung):
'           1. Als Nutzer mit Administratorrechten
'               es startet sofort
'           2. Als Nutzer ohne Administratorrechte
'               Die Arbeit im Lesemodus wird angeboten
'
'
'           Man mu� das Manifest in alle exe-Dateien des Installationspaketes einf�gen, damit keines der Programme \VirtualStore\ verwendet
'----------------------------------------------------------------------------------------------------------
'           In XP TNL max2 Halloween festgestellte Probleme
'           Es startet nicht, wenn Windows Media Player 9 oder h�her nicht installiert ist(Windows Media Player 9 oder h�her ist in Halloween nicht enthalten)
'           Nach Start von fotos.exe h�rt man ein paar Festplattenger�usche, dann passiert garnichts, auch
'           keine Fehlernachricht.
'           Aber daf�r gibt es in XP TNL max2 Halloween keine Probleme damit, da� XP SP3 zuerst installiert wird und
'           erst anschlie�end GERBING Fotoalbum 13.
'----------------------------------------------------------------------------------------------------------
'siehe auch 02.05.2006
'           manchmal nach l�ngerer Benutzung ohne Neustart und variierten Stichworten ist die Merker-Spalte
'           des ersten Datensatzes gel�scht(weder 0 noch 1). Als Folgefehler werden alle �nderungen im ersten Datensatz
'           nicht geupdatet.
'           Manchmal beim Wechsel von Suchkriterien kommt danach die MsgBox
'           'Sie d�rfen in die Spalte Merker nur 0 oder 1 eintragen'
'----------------------------------------------------------------------------------------------------------
'am 10.02.2007
'           1. fotos.exe hat gerade irgend etwas gemacht mit 'Fenstergr��e �nderbar'
'           2. fotos.exe soll fotosmdb.exe starten, aber fotosmdb.exe l��t sich nicht starten ->
'           fotosmdb.exe - Fehler in der Anwendung
'           "unknown software exception"......
'           Nach etwa 2 Minuten geht alles wieder.
'----------------------------------------------------------------------------------------------------------
'am 21.09.2008
'           im Windows 98 und Windows 2000 - letzte Version 13.3.8 funktioniert bei Fotosmdb Pr�fen1 das
'           Neuberechnen von Pixelhoehe und PixelBreite nicht.
'           Abhilfe: man muss nachtr�glich Gdiplus.dll (1.736 KB) in den Programmordner stellen.
'           Meine L�sung: Version 13.3.8 ist die letzte bereitgestellte Version f�r Windows 98 und Windows 2000
'----------------------------------------------------------------------------------------------------------
'am 10.10.2008
'           -2147221164 Klasse nicht registriert
'           Dieser Fehler tritt auf, wenn Sie GERBING Fotoalbum 13 sp�ter installiert haben als ein Update mit
'           XP SP3 (Service Pack 3),
'           und tritt nicht auf, wenn GERBING Fotoalbum 13 bereits installiert war, als Sie das Update mit
'           XP SP3 ausgef�hrt haben.
'           Solange Microsoft keine L�sung dieses Fehlers bietet, haben Sie nur die M�glichkeit, die
'           erforderliche Reihenfolge einzuhalten.
'           You get this error if you have first installed XP SP3 (service pack 3) and then GERBING Fotoalbum 13,
'           and will not get this error if GERBING Fotoalbum 13 was already installed before you installed XP SP3.
'           As long as microsoft does not deliver a solution you must follow the required installation sequence.
'
'           M�glicherweise ist nur bei mir der Fehler aufgetreten, weil meine Installations-DVD zwar SP3 enthalten hat,
'           aber kein SP2. Und es hei�t in der Literatur da� beide drauf sein m�ssen, weil SP3 auf SP2 aufbaut.
'           Mit XP Halloween tritt der Fehler nicht auf.
'----------------------------------------------------------------------------------------------------------
'am 20.12.2008
'           Wenn fotos.exe mit meinen aktuell etwa 16000 Fotos l�uft, dann zeigt der task manager etwa 70%
'           CPU Verbrauch an. Andere Programme werden aber nicht wirklich in ihrer Geschwindigkeit behindert.
'           Irgendein Microsoft MI oder MP (ein zertifizierter Microsoft Spezialist) hat im Internet beschrieben,
'           dass MS Access mehrere threads startet, die monitoring Aufgaben mit der access datenbank ausf�hren,
'           diese arbeiten aber mit so niedriger priorit�t, dass andere tasks sofort bedient werden, wenn sie
'           mehr Zeit brauchen.
'----------------------------------------------------------------------------------------------------------
'am 09.07.2011 nicht wiederholbar
'           Der btnStart der ImportForm ist manchmal wirkungslos.
'           Wenn er korrekt funktioniert muss er das Textfeld txtDragDropDatenbank sichtbar machen, welches nach Form_Load
'           unsichtbar ist.
'           Einmal habe ich s�mtliche Suchkriterien zur�ckgesetzt, danach ging es wieder.
'           Wiederholbarkeit mit Suchkriterien zB Jahr 2011 war nicht gegeben.
'----------------------------------------------------------------------------------------------------------
'am 05.10.2011
'           Wie kann man in der Datenbank nach EXIF:DateTimeOriginal suchen
'           oder wie kann ich erreichen, dass im Feld DDatum das Datum steht an dem das Foto gemacht wurde(und nicht das der letzten Bearbeitung)
'
'           L�sung1 Professional Version ab Version 14.0.2:
'           Man muss ein nutzerdefiniertes Text-Feld anlegen zB ExifDateTimeOriginal und bei der Aufnahme in die Datenbank
'           ausw�hlen, dass dieses Feld aus dem Feld EXIF:DateTimeOriginal aufgef�llt wird. Da bekommt man Datum und Uhrzeit
'           im Format 2010:12:31 12:01:05
'           Mit Pr�fen3 kann ich das �bernehmen von EXIF:DateTimeOriginal nach Spalte ExifDateTimeOriginal starten
'
'           L�sung2 auch f�r Shareware-Version:
'           Man benutzt die Software ExifToolGUI.exe
'           ExifToolGUI.exe kann EXIF:DateTimeOriginal auswerten und als Datei-Datum 'Ge�ndert am' (Date modified) eintragen.
'           Das geht auch stapelweise. Man muss ausf�hren Men� -> Various -> File:Date modified = Exif:DateTimeOriginal
'           Danach muss man FotosMdb.exe Pr�fen1 wiederholen f�r alle Dateien, dadurch wird die Spalte DDatum neu aufgef�llt,
'           allerdings ohne die Uhrzeit.
'----------------------------------------------------------------------------------------------------------
'seit 21.05.2012 Version 13.5.1
'           Im Windows8 bekommt man den externen Windows Media Player obendrauf, wenn man ein H�kchen setzt bei
'           'Aktuelle Wiedergabe' immer oben anzeigen
'----------------------------------------------------------------------------------------------------------
'nur in 21.05.2012 Version 13.5.1
'           Bei Benutzung eines externen Videoplayers muss man manuell nach Video-Wechsel mittels F2 oder F3 dem externen Videoplayer
'           den Focus geben.
'           Ausweg: kontinuierlich die Videos abspielen lassen, da beh�lt der externe Videoplayer den Focus
'           tritt bei Version 13.5.4 nicht mehr auf
'----------------------------------------------------------------------------------------------------------
'nur in Version 13.5.1
'           Bei einem H�kchen in Fenstergr��e �nderbar fehlt ein Taskbar-Icon. Damit wird Drag&Drop erschwert. Man muss die richtigen
'           Fenster anfassen und selber so ziehen das sie unterscheidbar sind. Fenster ziehen nach rechts geht mit Anfassen an
'           oberer linker Ecke.
'----------------------------------------------------------------------------------------------------------
'           L�sbares Problem mit der IDE
'           You might notice after successfully installing VB6 on Win7/Win8 that working in the IDE is a bit, well, sluggish.
'           For example, resizing objects on a form is a real pain.
'           After installing VB6, you must change the compatibility settings for the IDE executable.
'           1.  Using Windows Explorer, browse the location where you installed VB6.
'           By default, the path is C:\Program Files\Microsoft Visual Studio\VB98\
'           2.  Right click the VB6.exe program file, and select properties from the context menu.
'           3.  Click on the Compatibility tab.
'           4.  Place a check in each of these checkboxes:
'           o   Run this program in compatibility mode for Windows XP (Service Pack 3)
'           o   Disable Visual Themes
'           o   Disable Desktop Composition
'           o   Disable display scaling on high DPI settings
'----------------------------------------------------------------------------------------------------------
'           L�sbares Problem mit der IDE
'           Normalerweise wird im Add-In-Manager der VB-Entwicklungsumgebung u.a. auch der VB6 Ressourcen-Editor angezeigt.
'           Ist der Ressourcen-Editor bei Ihnen nicht aufgef�hrt, k�nnen Sie diesen durch Registrieren der Datei RESEDIT.DLL
'           wieder aktivieren. Beispiel ist anzupassen. Man muss als Administrator arbeiten.
'           REGSVR32 "C:\Programme\Microsoft Visual Studio\VB98\Wizards\RESEDIT.DLL"
'           REGSVR32 "C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RESEDIT.DLL"
'           und man mu� die vb6 IDE als Administrator starten
'           Bessere L�sung: Eine Ressourcen-Datei (Language.res) entweder mit 'Microsoft (R) Developer Studio' oder
'                           'Resource Editor by Anders Melander' bearbeiten
'----------------------------------------------------------------------------------------------------------
'           L�sbares Problem mit der IDE
'           Zustand: Ich w�nsche mir schon lange, da� ich in der IDE mit dem Mausrad duch ein Formular scrollen kann.
'           L�sung: Wenn ich die Scroll-Funktion mit Mouse Rad brauche, starte ich VB6ScrollwheelFix.exe
'                   Wenn ich sie nicht mehr brauche, klicke ich auf das rote mouse icon in der task bar und w�hle Quit
'----------------------------------------------------------------------------------------------------------
'           nicht l�sbares Problem mit externen mediaplayern
'           Es gibt keine R�ckkopplung ob der externe Player das Abspielen pausiert hat.
'           Wenn kontinuierliches Abspielen ausgew�hlt ist und der externe mediaplayer wird pausiert, dann wird nach Ablauf von
'           VideoDuration trotzdem das n�chste Video gestartet.
'           Ebensowenig kann ich den externen player stoppen, wenn mit F2/F3 ein Foto ausgew�hlt wird.
'----------------------------------------------------------------------------------------------------------
'           l�sbares Problem mit meiner privaten englischsprachiger Datenbank
'           Wenn Fotosmdb S�tze �ndern oder l�schen soll kommt
'           Laufzeitfehler '-2147467259 (80004005)' Feld 'Ort' wurde nicht gefunden
'           Beim �ndern von Feldinhalten in der Datenbank fotos.mdb kommt - Feld 'Ort' wurde nicht gefunden (Fehler 3799)
'           Es sind die G�ltigkeitsregeln schuld, die nur in meiner Datenbank vorkommen zu Ort und Kommentar
'           Wenn ich die Datenbank englischsprachig mache, bleiben doch die deutschsprachigen G�ltigkeitsregeln erhalten
'----------------------------------------------------------------------------------------------------------
'           l�sbares Problem seit Version 14.0.0
'           bei Installation in einen unicode-Pfad kommt zB error 1904 Module C:\cnopt\msvbvm50.dll failed to register.
'           Man muss hier die Installation fortsetzen und nicht abbrechen.
'           und es kommt beim Starten �ber das Desktop-Icon von GERBING Fotoalbum
'           C:\users\MeinName\Desktop\fotos.mdb ist nicht vorhanden
'           Ausweg: einfach das Desktop-Icon von GERBING Fotoalbum l�schen und selber neu erzeugen
'----------------------------------------------------------------------------------------------------------
'           nicht l�sbares Problem Mit dem Kommentarfenster
'           wenn das Kommentarfenster zu sehen ist, habe ich vorgesehen, da� die Tasten F1 F2 F3 F4 F11 erkannt werden und zum Schlie�en
'           des Kommentarfensters f�hren.
'           Das passiert manchmal erst, wenn ich ein zweites mal F10 dr�cke
'           Ich finde keinen Weg das zu verbessern.
'----------------------------------------------------------------------------------------------------------
'           l�sbares Problem nur begrenzt auf mein pers�nliches Win7
'           '�ffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist' �ffnet nur den Ordner aber macht keine Markierung
'           Das tritt nicht auf im Win8 und nicht im Win10 und nicht in virtuellen Maschinen
'           offenbar bin ich selbst schuld. Ich hatte eingestellt Systemsteuerung -> programs and features -> turn windows features
'           on or off und ausgeschaltet Indizierung und 'windows search'
'           Daraufhin waren verschwunden Systemsteuerung -> Indizierungsoptionen und der Dienst 'Windows Search'
'           Aber die korrekte L�sung war ein neues Konto Gottfried (ohne Administratorrechte) anzulegen
'           das alte Konto hei�t GottfriedAlt
'----------------------------------------------------------------------------------------------------------
'           nicht l�sbares Problem seit Version 14.0.7
'           Wenn die Tasten Strg+Z gedr�ckt werden flackert das Bild schwarz auf
'           Das liegt an der Prozedur MyDrawImage dort wird gemacht Picture1.Picture = LoadPicture("")
'           vielleicht doch l�sbar, aber ich kann damit leben
'----------------------------------------------------------------------------------------------------------
'           nicht l�sbares Problem seit Version 14.1.1
'           wenn ich im Thumbnail-Bereich mit der -> oder <- Taste schnell hintereinander klicke, bleiben mehrere Thumbnails
'           blau umrandet und werden erst nach und nach grau.
'           Damit kann man leben.
'----------------------------------------------------------------------------------------------------------
'09.11.2015 geplante Verbesserung nach Version 14.1.2
'           auch f�r Videos sollen Metadaten angezeigt werden k�nnen.
'           analog zu IPTC... soll es eine Funktion geben, die Datenbankfelder in die Videos schreibt
'           12.11.2015 diese Idee wird wieder fallengelassen
'           ich zweifle am Nutzen. Wer au�er mir wird diese Funktion nutzen. Es gen�gt ein gut gew�hlter Dateiname.
'           Lesen avi oder mp4 w�rde gehen mit mediainfo.dll siehe
'               d:\VISUALBA.SIC\Video\Video Mit Exif und Metadaten\GMF Get Media Info\gmf\Get Media Info.vbp
'               mediainfo.dll liest:                              und Explorer kann schreiben bei mp4:
'               Titel
'               Kommentare
'               Mitwirkende Interpreten
'               Genre
'               Komponisten
'           1.schreiben w�rde gehen mit Aufruf von Exiftool als Command line
'               aktuell am 09.11.2015 unterst�tzt exiftool noch nicht das Schreiben in avi files, aber in mp4 files, und nur in die XMP-Felder
'               der Explorer schreibt aber nicht in XMP-Felder sondern in Quicktime-Felder bzw Microsoft-Felder
'           2.schreiben in avi oder mp4 w�rde gehen mit ffmpeg command line tool
'               man muss es mit direct stream copy machen
'               Die folgenden 5 tags erscheinen auch im Windows Explorer
'               Titel = title
'               Kommentare = comment
'               Mitwirkende Interpreten = artist
'               Genre = genre
'               Komponisten = composer
'               aber ffmpeg entfernt beim schreiben dieser tags die evtl vorhandenen microsoft tags (Untertitel und Markierungen u.a.)
'               das macht mp4tags nicht
'           3.schreiben in mp4 w�rde gehen mit mp4tags command line tool
'               Die folgenden 5 tags erscheinen auch im Windows Explorer
'               Titel = song
'               Kommentare = comment
'               Mitwirkende Interpreten = artist
'               Genre = genre
'               Komponisten = writer
'           4.schreiben in avi oder mp4 w�rde gehen mit VirtualDub im men� file -> set text information...
'               Die folgenden 4 tags erscheinen auch im Windows Explorer, aber k�nnen bei avi mit dem Windows Explorer nicht geschrieben werden.
'               Titel = virtualdub name
'               Mitwirkende Interpreten = virtualdub artist
'               Genre = virtualdub genre
'               Copyright = virtualdub copyright
'               Jedoch s�mtliche mit virtualdub geschriebenen tags (set text information...) k�nnen mit mediainfo.dll gelesen werden
'
'           Bei mp4 videos kann man mit dem Explorer Eigenschafts-Fenster Metadaten eintragen, auch f�r mehrere gleichzeitig
'           L�sung k�nnte so aussehen:
'           Schwerpunkt von avi weg auf mp4 files legen.
'           1. ich benutze VirtualDub2 zum Beschneiden und schreiben 4 tags - und Avidemux zum Append
'           2. optional kann ich dann mit dem Windows Explorer tags eintragen
'           3. Ausf�hren von Fotosmdb zur Aufnahme in die Datenbank(Lesen mit mediainfo.dll)
'           4. bei IPTC... mp4tags oder exiftool benutzen um in mp4 videos zu schreiben
'           nicht geeignet ist die Idee mit NTFS Streams zu arbeiten (siehe dsofile.dll) weil diese Daten der Explorer nicht sehen kann
'----------------------------------------------------------------------------------------------------------
'05.12.2015 verworfene Verbesserungs-Idee
'           Ich wollte einen allgemeing�ltigen Resize-Algorithmus erfinden. keine von Form1 ge�ffnete Form darf gr��er sein als Form1
'           oder �ber die Begrenzungen von Form1 hinausragen.
'           Das Hinausragen kann ich nicht verhindern, weil es keine R�ckmeldung gibt, wenn ein ge�ffnetes Fenster auf dem Screen
'           verschoben wird.
'           Vielleicht gibt es diese R�ckmeldung mit API.
'           Angeblich geht Subclassing siehe d:\VISUALBA.SIC\Foto\Check if any form has moved\
'           L�sung: siehe 07.01.2018
'                   Alle Formulare zentriert �ffnen (StartUpPosition=1=Fenstermitte)
'----------------------------------------------------------------------------------------------------------
'01.05.2016 scheinbarer Fehler
'           Zustand: Wenn ich im Listenfenster zum Spaltenverbreitern oder Spaltenverkleinern in der Grid-Headline einen senkrechten Strich
'                   mit der Maus anfasse und ziehe, ver�ndert dieser sich in einen senkrechten Strich mit Pfeil nach links und Pfeil
'                   nach rechts. Wenn ich die linke Maustaste loslasse, springt die Anzeige auf ein weiter oben liegendes Foto.
'           L�sung: Da bin ich selber schuld. Ich habe ein Ereignis RowColChange ausgel�st.
'                   Wenn ich nach dem Verschieben der Spaltenbreite aus der �berschrift herausrutsche und erst dann die linke Maustaste
'                   loslasse, wenn ich bereits in einer Tabellenzeile bin, dann wird ein Ereignis RowColChange ausgel�st. Logisch.
'----------------------------------------------------------------------------------------------------------
'02.05.2016 Fehler mit Umgehungsl�sung
'           Zustand: tritt nur auf wenn mit Thumbnails gew�hlt ist.
'                   Tritt nur auf, wenn mehrfach zwischen 2 Fotos hin- und hergewechselt wird.
'                   Tritt nicht auf, wenn zum Hin- und Herwechseln nur in das Grid geklickt wird.
'                   Tritt auf, wenn zum Bildwechsel abwechselnd Klick ins Grid und Klick auf Thumbnail genutzt wird.
'                   Sobald mit Klick ins Grid zur�ckgewechselt wird, wird die Umrandung des Thumbnail dunkelblau, das ist die Situation,
'                   wo ein Klick zur�ck mit Thumbnail-Klick wirkungslos ist. Mit anderen Worten, ein Ereignis optThumb_Click findet
'                   nicht statt. Da kann ich nichts dagegen tun.
'           L�sung: Man muss einmal irgendwo anders hinklicken als auf den Thumbnail der nicht reagiert, dann reagiert er wieder.
'----------------------------------------------------------------------------------------------------------
'10.11.2016.2016 scheinbarer Fehler
'           Zustand: tritt nur auf wenn mit Thumbnails gew�hlt ist.
'                   Wenn ich im Grid in ein Feld der Spalten�berschrift klicke, werden auch die Thumbnails in neue Reihenfolge
'                   gebracht.
'                   Wenn ich das Ende der Thumbnail-Anordnung nicht abwarte und in andere Felder in der Spalten�berschrift klicke,
'                   konnte ich beobachten da� die Thumbnail-Anodnung 4-mal wiederholt wurde.
'           L�sung: Da bin ich selber schuld. Das Programm macht das was ihm gesagt wird. Ich muss einfach das Ende der
'                   Thumbnail-Anordnung abwarten.
'----------------------------------------------------------------------------------------------------------
'27.11.2016 nicht l�sbares Problem
'           Zustand: Beim Wechsel von einem Video auf ein anderes Video sehe ich kurz ein schwarzes Fenster. Das passiert immer bei
'                   'frmVideo.WMP.URL = ""' oder 'frmVideo.WMP.URL = videoname', aber ohne 'frmVideo.WMP.URL = videoname' spielt es nicht
'----------------------------------------------------------------------------------------------------------
'13.12.2016 Problem mit exiftool.exe und PSP X8 oder h�her
'           Zustand: Alles was ich mit exiftool hinzugef�gt habe, (�ber fotosmdb.exe Funktion EXIF/IPTC...),
'                   wird von PSP X8 oder PSP 2019 wieder rausgeschmissen.
'                   Das passiert, wenn ich nach der Aufnahme ins Fotoalbum Fotos mit PSP X8 nochmal bearbeite
'                   PSP X8 l�scht den Abschnitt IPTC
'                       aber erzeugt stattdessen den Abschnitt IPTC2
'                   PSP X8 l�scht aus dem Abschnitt IFD0 alle Felder XPTitle XPKeywords XPAuthor XPSubjects XPComment
'                       und schreibt sie stattdessen in den Abschnitt XMP-photoshop
'                   Wenn ich nachfolgend eine neue Datenbank bei Null erzeuge durch Import der EXIF/IPTC-Felder, dann bekommen diese Bilder
'                   leere Datenbankfelder.
'           L�sung: Gegenma�nahme: entweder damit leben und leere Felder manuell auff�llen oder
'                   fotosmdb.exe ausf�hren
'                   1. Pr�fen1 (ohne Zusatzberechnung) setzt das Feld IPTCPresent = 0 f�r die Dateien, die ein aktuelleres Datum haben als
'                       in der Datenbank.
'                   2. Funktion EXIF/IPTC... ausf�hren, f�r die Dateien, wo das Feld IPTCPresent = 0, bevor eine neue Datenbank bei Null
'                       erzeugt wird.
'                   �ffnen der mit 'jpg' verkn�pften Anwendung f�r die aktuelle Datei setzt IPTCPresent = 0
'                   �ffne ein Explorer-Fenster, wo die aktuelle Datei markiert ist setzt IPTCPresent = 0
'                   exiftool kann alle IPTC-Felder l�schen mit '-iptc:all='
'                   exiftool kann den Abschnitt XMP-photoshop l�schen mit '-Xmp-photoshop='
'           Versuchte L�sung: Zum Lesen von EXIF/IPTC exiftool.exe benutzen, scheitert am Schreiben in IPTC-Felder mit unicode
'                   exiftool.exe kann nicht lesen was es selbst geschrieben hat, wenn es IPTC-Felder mit unicode sind
'                   aber ich kann das korrigieren
'                   siehe d:\VISUALBA.SIC\VB6BeispielCode\Multimedia\Exif IPTC Info.OCX\2A Benutzung von exiftool.exe CreatePipe\Projekt1.vbp
'----------------------------------------------------------------------------------------------------------
'10.08.2017 Verbesserungsidee realisiert
'           Zustand: Das portable InnoSetup-Paket ist etwa 26 MB gro�. Davon entfallen auf WMP.dll schon 10 MB
'           L�sung: Fotoalbum Portable kann die WMP.dll benutzen, ohne dass ich sie ausliefere
'----------------------------------------------------------------------------------------------------------
'22.01.2018 gel�stes Problem
'           Zustand: Von der Datenbankdefinition festgelegten G�ltigkeitsregeln sind zwar wirksam, aber sie wirken lautlos.
'                   Wenn zB Situation l�nger ist als 255 Bytes,
'                   dann wird ohne Kommentar der alte Wert beibehalten
'----------------------------------------------------------------------------------------------------------
'15.06.2018 Problem mit Umgehungs-L�sung
'           Zustand: Fehlermeldung: Die zum Aktualisieren angegebene Zeile wurde nicht gefunden.
'                                   Einige Werte wurden seit dem letzten Lesen ggf. ge�ndert
'           Situation: Ich habe nach Dr�cken von F5 die Spalte Kommentar editiert. Dann diesen Kommentar mit Dr�cken von F10 nochmal
'                   �berarbeitet.
'           Gegenma�nahme: Nach Dr�cken von F5 auf den Refresh-Button klicken, dann erst den Kommentar editieren
'----------------------------------------------------------------------------------------------------------
'10.01.2019 Ich habe ausprobiert, ob ein neues InnoSetup wegen Benutzung von Office 16 gebraucht wird -> wird nicht
'           Nach Benutzung von Office 2016 habe ich mit 'InnoToolbarVB6 Advanced for VB6' ein neues InnoSetup erzeugen lassen.
'           Hier sind 17 dll und eine olb mehr drin als vor der Benutzung von Office 2016
'           Ich wollte wissen, ob jetzt bei Nichtinstalliertem Outlook trotzdem eine Email erzeugt werden kann -> kann nicht
'           und ob ohne installierten Internet Explorer 11 eine GPS-Position angezeigt werden kann -> kann nicht
'----------------------------------------------------------------------------------------------------------
'11.01.2019 scheinbares Problem mit Anzeige GEO-Position, GEO-Position zeigt leeren Inhalt(PGSLongitude=0 GPSLatitude=0) das ist im Ozean
'           das passiert bei mir bei Salt Lake City
'           L�sung:
'           Ich muss mehrmals das Bild verkleinern (Minus-Button)
'----------------------------------------------------------------------------------------------------------
'13.01.2019 ungel�stes Problem
'           Zustand: Die Funktion Strg+C(Kopieren in Zwischenablage oder Ordner) funktioniert nicht mit Unicode Ordnern zB 'Video AlbumCnopt'
'                   es funktioniert jedoch mit Ansi-Ordnern und Unicode filename
'           L�sung: offen, es geht auch nicht mit commandline und powershell
'           Umgehungs-L�sung: Der Nutzer darf unicode filenames benutzen, aber keine unicode pathnames
'----------------------------------------------------------------------------------------------------------
'08.05.2019 ungel�stes Problem
'           Zustand: Nach 3 Sekunden beendet sich die IDE von selbst
'                   unter folgenden Bedingungen
'                   1. wenn 'Fotos finden' nicht drangekommen ist
'                   und
'                   2. die IDE nicht als Administrator gestartet wurde
'           Not-L�sung: Die IDE als Administrator starten
'           Neues Problem wegen der Notl�sung:
'                   Wenn die IDE als Administrator startet, kann es passieren dass ich einen Fehler nicht entdecke,
'                   der aufgetreten w�re bei Start als Nichtadministrator,
'                   dann erzeuge ich die exe
'                   und falle auf die Nase bei Ausf�hrung in einem fremden PC, der als Nichtadministrator l�uft
'                   das ist so passiert bei CheckCompactDatabase.exe als diese AccessDatabaseEngine.exe starten sollte
'                   da kam Laufzeitfehler '5'
'----------------------------------------------------------------------------------------------------------
'20.01.2020 scheinbares Problem siehe auch 09.05.2012
'           Zustand: Auf einem 1920x1200 Pixel Monitor soll ein 1920x1200 Pixel Foto gezeigt werden. Am unteren Rand fehlt ein Streifen oder
'                   Auf einem 1920x1080 Pixel Monitor soll ein 1920x1080 Pixel Foto gezeigt werden. Am unteren Rand fehlt ein Streifen
'           L�sung: Man muss Fotos.exe starten mit entferntem H�kchen bei 'Fenstergr��e �nderbar'
'----------------------------------------------------------------------------------------------------------
'30.06.2020 ungel�stes Problem
'           Zustand: B�ses Problem
'                   Beim Start von fotos.exe in der IDE passiert scheinbar garnichts.
'                   Nach einigen Sekunden beendet sich vb6.
'           L�sung: AccessDatabaseEngine.exe neu starten
'----------------------------------------------------------------------------------------------------------
'31.08.2020 Problem mit ungenauen GPS-Positionen
'           Zustand: Wenn mit einem Smartphone und eingeschalteter Standortbestimung Fotos gemacht werden, dann ist unmittelbar nach
'                   Antippen des Kamera-Symbols der Standort noch ungenau. Der ungenaue Standort wird in die EXIF-Daten des Fotos eingetragen.
'                   Bei entsprechender Einstellung werden die EXIF-Daten unver�ndert in die Datenbank �bernommen.
'                   Solche vom Smartphone �bernommenen Felder GPSLatitude und GPSLongitude sind gekennzeichnet durch viel Stellen
'                   nach dem Komma.
'           L�sung: Ich kann nach Feldern GPSLatitude und GPSLongitude mit vielen Nachkommastellen suchen, wenn ich trickreich arbeite.
'                   Ich muss ein H�kchen setzen bei 'SQL �berarbeiten' dann eine von mir vorbereitete SQL-Anweisung eingeben
'                   zB SELECT Fotos.* FROM Fotos WHERE len(Fotos.GPSLongitude)>15
'                   oder
'                   SELECT Fotos.* FROM Fotos WHERE len(Fotos.GPSLatitude)>15



Private Sub Form_Load()

End Sub
