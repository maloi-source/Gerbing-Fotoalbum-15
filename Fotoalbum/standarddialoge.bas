Attribute VB_Name = "Module2"
Option Explicit
' ###########################################
'
'                      MODUL "CommonDialog"
'
'                     (c) Ingo Steinhaus 2000
'                     ingo.steinhaus@gmx.de
'
'  Funktionen zur Anzeige der Windows-Standarddialoge "�ffnen"
'  und "Speichern"
'
'  Dieses Modul ist urheberrechtlich gesch�tzte Freeware.
'  Die originale Copyright-Meldung darf nicht entfernt oder ver-
'  �ndert werden. Der Quelltext darf nicht ver�ndert werden-
'
'############################################

Public Const MAX_PATH = 260

'**********************************************************
Rem Die Datenstruktur "OpenFilename" dient der Konfiguration
Rem des Dialogs.

Private Type OpenFilename
    lStructSize As Long
    'Gr��e der Datenstruktur. Kann mit Len() bestimmt werden.
    hWndOwner As Long
    'Handle des Besitzers (mit GetActiveWindow() abfragen).
    hInstance As Long
    'Handle der Dialogfeldvorlage, wenn OFN_ENABLETEMPLATEHANDLE
    'in Flags gesetzt ist. Wenn OFN_EXPLORER gesetzt ist, wird der
    'Dialog vom Standardialog des Explorers abgeleitet. Andernfalls
    'wird ein Dialog im Windows-3.x-Stil erzeugt.
    lpstrFilter As String
    'Ein VB-String mit paarweise angeordneten nullterminierten Strings.
    'Der letzte nullterminierte String mu� mitr einem weiteren NULL-Zeichen
    'abgeschlossen werden.
    'Ein Filter besteht aus zwei nullterminierten Strings. Der erste enth�lt
    'die Zeichenkette, die im Kombifeld "DateiTyp" angezeigt wird, der zweite
    'die zugeh�rtigen Dateimasken wie z. B. "*.doc".
    'Beispiel: "Word-Dokumente" + Chr$(0) + "*.doc" + Chr$(0)
    'Sie k�nnen mehrere Dateimasken durch Semikola abtrennen.
    'Beispiel: "Grafiken" + Chr$(0) + "*.bmp;*.jpg;*.gif" + Chr$(0)
    lpstrCustomFilter As String
    'Ein VB-String mit dem im Kombifeld "DateiTyp" ausgew�hlten Filter.
    nMaxCustFilter As Long
    'Die Gr��e von lpstrCustomFilter.
    nFilterIndex As Long
    'Der 1-basierte Index des im Kombifeld "DateiTyp" ausgew�hlten Filters.
    lpstrFile As String
    'Ein VB-String mit dem ausgew�hlten Dateinamen inkl. Laufwerk und Pfad.
    'Der String mu� vorher in der entsprechenden Gr��e erzeugt werden.
    'Er kann vor dem Aufruf des Dialogs mit dem Namen einer existierenden
    'Datei belegt werden.
    nMaxFile As Long
    'Die Gr��e von lpstrFile.
    lpstrFileTitle As String
    'Ein VB-String mit dem ausgew�hlten Dateinamen ohne Laufwerk und Pfad.
    nMaxFileTitle As Long
    'Die Gr��e von lpstrFileTitle.
    lpstrInitialDir As String
    'Ein VB-String mit dem Pfadnamen des Ordners, dessen Inhalt der Dialog
    'beim Anzeigen darstellen soll.
    lpstrTitle As String
    'Ein VB-String mit Titel des Dialogfeldes.
    Flags As Long
    'Flags, die die Anzeigeoptionen des Dialogfeldes bestimmen
    nFileOffset As Integer
    'Index zum Beginn des ersten Dateinamens in lpstrFile.
    nFileExtension As Integer
    'Index zum Beginn der Dateierweiterung in lpstrFile.
    lpstrDefExt As String
    'Die Standarderweiterung, die an einen Dateinamen vergeben wird, wenn
    'er keine Erweiterung besitzt.
    lCustData As Long
    'Ein Zeiger auf anwendungsspezifiasche Daten, f�r die R�ckruffunktion.
    lpfnHook As Long
    'Adresse einer R�ckruffunktion, die in der Anwendung definiert wird.
    'Sie k�nnen hier NULL eintragen.
    lpTemplateName As String
    'Der Name der Dialogfeldvorlage (siehe hInstance)
End Type

'**********************************************************
Rem Die folgenden Konstanten sind die erlaubten Werte f�r
Rem OpenFilename->Flags.

Public Const OFN_ALLOWMULTISELECT = &H200
'Zeigt ein Dialogfeld mit der M�glichkeit, mehrere Dateien auszuw�hlen.
'In diesem Fall enth�lt lpstrFile den Pfad und anschlie�end alle Dateinamen.
'nFileOffset zeigt auf den Index des ersten Dateinamens nach der Pfadangabe.
'lpstrFile enth�lt alle Dateinamen durch Chr$(0) getrennt. Em Ende folgt ein
'zweites Chr$(0). Bei alten Win-3.x-Dialoge) sind die Dateinamen durch
'Leerzeichen getrennt. Diese Variante kennt keine langen Dateinamen.

Public Const OFN_CREATEPROMPT = &H2000
'Zeigt eine Meldung, wenn die Datei nicht existiert und fragt den Anwender, ob
'sie erzeugt werden soll.

Public Const OFN_ENABLEHOOK = &H20
'Aktiviert die R�ckruffunktion lpfnHook.

Public Const OFN_ENABLETEMPLATE = &H40
'Aktiviert die Dialogfeldvorlage.

Public Const OFN_ENABLETEMPLATEHANDLE = &H80
'Aktiviert die Dialogfeldvorlage.

Public Const OFN_EXPLORER = &H80000
'Nutzt Explorer-Dialoge. Diese Einstellung ist die Vorgabe, selbst wenn Sie
'dieses Flag nicht angeben. F�r alte Win-3.x-Dialoge m�ssen Sie das Flag
'l�schen.
'Sie m�ssen es in den folgenden F�llen setzen:
'- bei OFN_ALLOWMULTISELECT.
'- wenn Sie Dialogfeldvorlagen und R�ckruffunktionen benutzen.

Public Const OFN_EXTENSIONDIFFERENT = &H400&
'Gibt an, dass der Anwender einen Dateinamen mit einer anderen Erweiterung
'als lpstrDefExt eingeben kann.

Public Const OFN_FILEMUSTEXIST = &H1000
'Gibt an, dass der Anwender nur die Namen von existierenden Dateien eingeben
'kann. Andernfalls wird eine Warnmeldung ausgegeben.
'OFN_PATHMUSTEXIST mu� ebenfalls gesetzt werden.

Public Const OFN_HIDEREADONLY = &H4&
'Versteckt das Kontrollk�stchen "Nur lesen".

Public Const OFN_LONGNAMES = &H200000
'Aktiviert die Unterst�tzung von langen Dateinamen in den alten Win-3.x-Dialogen.

Public Const OFN_NOCHANGEDIR = &H8&
'Stellt das urspr�ngliche Verzeichnis bei Ende des Dialoges wieder her, wenn
'der Anwender anderes Verzeichnis eingestellt hat.

Public Const OFN_NODEREFERENCELINKS = &H100000
'Weist das Dialogfeld an, bei einer markierten Verkn�pfung Namen und Pfad der
'Verkn�pungsdatei zur�ckzugeben, snstatt Namen und Pfad der Datei, auf die die
'Verkn�pfung verweist.

Public Const OFN_NOLONGNAMES = &H40000
'Deaktiviert die Unterst�tzung von langen Dateinamen in den alten
'Win-3.x-Dialogen.

Public Const OFN_NONETWORKBUTTON = &H20000
'Versteckt die Schaltfl�che "Netzwerk".

Public Const OFN_NOTESTFILECREATE = &H10000
'Gibt an, dass keine Testdatei erzeugt wird, bevor der Dialog endet. In diesem
'Fall �berpr�ft das Dialogfeld nicht auf Schreibschutz, Platzmangel auf dem
'Datentr�ger oder korrekten Netzwerkzugriff.

Public Const OFN_OVERWRITEPROMPT = &H2&
'Gibt im Dialog "Speichern" eine Warnmeldung aus, wenn die Datei bereits
'existiert und durch das Speichern �berschrieben wird.

Public Const OFN_PATHMUSTEXIST = &H800
'Gibt an, dass der Anwender nur die Namen von existierenden Verzeichnissen
'eingeben kann. Andernfalls wird eine Warnmeldung ausgegeben.

Public Const OFN_READONLY = &H1
'Gibt an, das das Kontrollk�stchen "Nur Lesen" angekreuzt ist, wenn der Dialog
'angezeigt wird.

Public Const OFN_SHAREAWARE = &H4000
'Gibt an, dass die Funktion fehlschl�gt, wenn ein Netzwerkfehler auftritt.

Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1

Public Const OFN_SHOWHELP = &H10
'Zeigt im Dialogfeld den Hilfe-Schalter an. hwndOwner mu� auf ein Fenster zeigen,
'das die Hilfe anzeigen kann. ExplorerDialoge senden die Nachricht CDN_HELP
'an die R�ckruffunktion.

Public Const OFS_MAXPATHNAME = 128

'**********************************************************
Rem *** GetSaveFileName ***
Rem Funktion zum Anzeigen des Dialogs "Speichern"

Private Declare Function GetSaveFileName Lib "comdlg32" Alias _
    "GetSaveFileNameA" (lpOpenfilename As OpenFilename) As Long

'**********************************************************
Rem *** GetOpenFileName ***
Rem Funktion zum Anzeigen des Dialogs "�ffnen"

Private Declare Function GetOpenFileName Lib "comdlg32" Alias _
    "GetOpenFileNameA" _
    (lpOpenfilename As OpenFilename) As Long

'**********************************************************
Rem *** CommDlgExtendedError ***
Rem Funktion zum Ermitteln der Fehlernummer

Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Integer

'**********************************************************
Rem *** GetActiveWindow ***
Rem Eine Funktion zum Ermitteln des Fenster-Handles.

Private Declare Function GetActiveWindow Lib "user32" () As Long

'**********************************************************
Rem *** PrepareFilter ***
Rem Eine Funktion zum Aufbereiten des Filters
Rem Beispielfilter: "Word Dokument (.doc)|*.doc|Word Vorlage (.dot)|*.dot|"
Rem Die Funktion ersetzt "|" durch Chr$(0) und f�gt das abschlie�ende
Rem Chr$(0) ein.

Public Function PrepareFilter(Flt$) As String
Const O$ = "|"
Dim Temp$
Dim i As Integer
    'Mit einer Kopie arbeiten
    Temp$ = Flt$
    'Beim ersten Zeichen beginnen
    i = 1
    '"|" gefunden?
    Do While InStr(i, Flt$, O$) <> 0
        'Alles bis zum ersten "|" kopieren und Chr$(0) anh�ngen
        PrepareFilter = PrepareFilter + _
                Mid(Temp$, i, InStr(i, Temp$, O$) - i) + vbNullChar
        'Index auf Zeichen nach "|" setzen
        i = InStr(i, Temp$, O$) + Len(O$)
    Loop
    'Evtl. Rest vom String und abschlie�endes Chr$(0) anh�ngen
    PrepareFilter = PrepareFilter + Right(Temp$, Len(Temp$) - i + 1) + vbNullChar
End Function

'**********************************************************
Rem *** GetSaveName ***
Rem Eine VB/VBA-Funktion als einfach zu nutzender Mantel f�r den
Rem Aufruf des Dialogs "Speichern".

Public Function GetSaveName(ByVal Filter$, ByVal DefExt$, ByVal InitialDir$) As String
Dim OFN As OpenFilename
Dim Temp$
Dim n As Integer

    'Bestimmen der Optionen f�r den Dialog
    With OFN
        'Gr��e der Struktur festlegen
        .lStructSize = Len(OFN)
        'Das aktive Fenster (= Word) wird zum Besitzer des Dialogs
        .hWndOwner = GetActiveWindow()
        'Der Filtzer wird vorbereitet
        .lpstrFilter = PrepareFilter(Filter$)
        'Speicher reservieren f�r kompletten Pfad
        .lpstrFile = String$(700, vbNullChar)
        'Gr��e des reservierten Speichers angeben
        .nMaxFile = 700
        'Speicher reservieren f�r Dateinamen
        .lpstrFileTitle = String$(MAX_PATH, vbNullChar)
        'Gr��e des reservierten Speichers angeben
        .nMaxFileTitle = MAX_PATH
        'Das Vorgabeverzeichnis bestimmen
        .lpstrInitialDir = InitialDir$
        'Der Titel des Dialoges
        .lpstrTitle = "Speichern"
        'Optionen bestimmen
        .Flags = OFN_EXTENSIONDIFFERENT Or _
            OFN_NOCHANGEDIR Or _
            OFN_OVERWRITEPROMPT _
            Or OFN_HIDEREADONLY
        'Standarderweiterung f�r die Dateien bestimmen
        .lpstrDefExt = DefExt$
    End With

    If GetSaveFileName(OFN) Then
        Temp$ = OFN.lpstrFile
        'Alles nach dem NULL-Zeichen verwerfen
        n = InStr(Temp$, vbNullChar)
        If n > 1 Then
            GetSaveName = left$(Temp$, n - 1)
        Else
            GetSaveName = ""
        End If
    Else
        GetSaveName = ""
    End If
End Function

'**********************************************************
Rem *** GetOpenName ***
Rem Eine VB/VBA-Funktion als einfach zu nutzender Mantel f�r den
Rem Aufruf des Dialogs "�ffnen".

Public Function GetOpenName(ByVal Filter$, ByVal InitialDir$, ByVal Title$) As String
Dim OFN As OpenFilename
Dim Temp$
Dim n As Integer

    'Bestimmen der Optionen f�r den Dialog
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = GetActiveWindow()
        .lpstrFilter = PrepareFilter(Filter$)
        'Speicher reservieren
        .lpstrFile = String$(700, vbNullChar)
        .nMaxFile = 700
        .lpstrFileTitle = String$(MAX_PATH, vbNullChar)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = InitialDir$
        '.lpstrTitle = "�ffnen"
        .lpstrTitle = Title
        .Flags = OFN_EXTENSIONDIFFERENT Or _
            OFN_NOCHANGEDIR Or _
            OFN_OVERWRITEPROMPT _
            Or OFN_HIDEREADONLY
    End With

    If GetOpenFileName(OFN) Then
        Temp$ = OFN.lpstrFile
        'Alles nach dem NULL-Zeichen verwerfen
        n = InStr(Temp$, vbNullChar)
        If n > 1 Then
            GetOpenName = left$(Temp$, n - 1)
        Else
            GetOpenName = ""
        End If
    Else
        GetOpenName = ""
    End If
End Function



