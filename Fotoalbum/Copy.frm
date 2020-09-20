VERSION 5.00
Begin VB.Form Copy 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   Icon            =   "Copy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblHinweis 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   $"Copy.frx":038A
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Copy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                                                   'Gerbing 11.03.2017
    If Sprache = 0 Then
        'me.Caption = "Shareware-Hinweis"
        Me.Caption = Chr(83) & Chr(104) & Chr(97) & Chr(114) & Chr(101) & Chr(119) & Chr(97) & Chr(114) & Chr(101) & Chr(45)   'Shareware-
        Me.Caption = Me.Caption & Chr(72) & Chr(105) & Chr(110) & Chr(119) & Chr(101) & Chr(105) & Chr(115) 'Hinweis
    
        'lblHinweis = "Sie benutzen die Shareware-Version von GERBING Fotoalbum."
        lblHinweis = Chr(83) & Chr(105) & Chr(101) & Chr(32)                                         'Sie
        lblHinweis = lblHinweis & Chr(98) & Chr(101) & Chr(110) & Chr(117) & Chr(116) & Chr(122) & Chr(101) & Chr(110) & Chr(32) 'benutzen
        lblHinweis = lblHinweis & Chr(100) & Chr(105) & Chr(101) & Chr(32)                           'die
        lblHinweis = lblHinweis & Chr(83) & Chr(104) & Chr(97) & Chr(114) & Chr(101) & Chr(119) & Chr(97) & Chr(114) & Chr(101) & Chr(45)   'Shareware-
        lblHinweis = lblHinweis & Chr(86) & Chr(101) & Chr(114) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(32)  'Version
        lblHinweis = lblHinweis & Chr(118) & Chr(111) & Chr(110) & Chr(32)                           'von
        lblHinweis = lblHinweis & Chr(71) & Chr(69) & Chr(82) & Chr(66) & Chr(73) & Chr(78) & Chr(71) & Chr(32) 'GERBING
        lblHinweis = lblHinweis & Chr(70) & Chr(111) & Chr(116) & Chr(111) & Chr(97) & Chr(108) & Chr(98) & Chr(117) & Chr(109) & Chr(46) & Chr(32)  'Fotoalbum.
        
        'lblHinweis = lblHinweis & " Lesen Sie in der Hilfe wie Sie eine Professional-Version erhalten können."
        lblHinweis = lblHinweis & Chr(32) & Chr(76) & Chr(101) & Chr(115) & Chr(101) & Chr(110) & Chr(32)   'Lesen
        lblHinweis = lblHinweis & Chr(83) & Chr(105) & Chr(101) & Chr(32)                           'Sie
        lblHinweis = lblHinweis & Chr(105) & Chr(110) & Chr(32)                                     'in
        lblHinweis = lblHinweis & Chr(100) & Chr(101) & Chr(114) & Chr(32)                          'der
        lblHinweis = lblHinweis & Chr(72) & Chr(105) & Chr(108) & Chr(102) & Chr(101) & Chr(32)     'Hilfe
        lblHinweis = lblHinweis & Chr(119) & Chr(105) & Chr(101) & Chr(32)                          'wie
        lblHinweis = lblHinweis & Chr(83) & Chr(105) & Chr(101) & Chr(32)                           'Sie
        lblHinweis = lblHinweis & Chr(101) & Chr(105) & Chr(110) & Chr(101) & Chr(32)               'eine
        lblHinweis = lblHinweis & Chr(80) & Chr(114) & Chr(111) & Chr(102) & Chr(101) & Chr(115) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(97) & Chr(108) & Chr(45) 'Professional-
        lblHinweis = lblHinweis & Chr(86) & Chr(101) & Chr(114) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(32)   'Version
        lblHinweis = lblHinweis & Chr(101) & Chr(114) & Chr(104) & Chr(97) & Chr(108) & Chr(116) & Chr(101) & Chr(110) & Chr(32) 'erhalten
        lblHinweis = lblHinweis & Chr(107) & Chr(246) & Chr(110) & Chr(110) & Chr(101) & Chr(110) & Chr(46) 'können
    End If
    If Sprache = 3000 Then
        'me.Caption = "Shareware"
        Me.Caption = Chr(83) & Chr(104) & Chr(97) & Chr(114) & Chr(101) & Chr(119) & Chr(97) & Chr(114) & Chr(101)   'Shareware
        'lblHinweis = "You use the Shareware version of GERBING Fotoalbum."
        lblHinweis = Chr(89) & Chr(111) & Chr(117) & Chr(32)                                        'You
        lblHinweis = lblHinweis & Chr(117) & Chr(115) & Chr(101) & Chr(32)                          'use
        lblHinweis = lblHinweis & Chr(116) & Chr(104) & Chr(101) & Chr(32)                          'the
        lblHinweis = lblHinweis & Chr(83) & Chr(104) & Chr(97) & Chr(114) & Chr(101) & Chr(119) & Chr(97) & Chr(114) & Chr(101) & Chr(32)   'Shareware
        lblHinweis = lblHinweis & Chr(118) & Chr(101) & Chr(114) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(32)  'version
        lblHinweis = lblHinweis & Chr(111) & Chr(102) & Chr(32)                                     'of
        lblHinweis = lblHinweis & Chr(71) & Chr(69) & Chr(82) & Chr(66) & Chr(73) & Chr(78) & Chr(71) & Chr(32) 'GERBING
        lblHinweis = lblHinweis & Chr(70) & Chr(111) & Chr(116) & Chr(111) & Chr(97) & Chr(108) & Chr(98) & Chr(117) & Chr(109) & Chr(46) & Chr(32)  'Fotoalbum.
        
        'lblHinweis = lblHinweis & " Read in the helpfile how you can get a Professional version."
        lblHinweis = lblHinweis & Chr(82) & Chr(101) & Chr(97) & Chr(100) & Chr(32)                 'Read
        lblHinweis = lblHinweis & Chr(105) & Chr(110) & Chr(32)                                     'in
        lblHinweis = lblHinweis & Chr(116) & Chr(104) & Chr(101) & Chr(32)                          'the
        lblHinweis = lblHinweis & Chr(104) & Chr(101) & Chr(108) & Chr(112) & Chr(102) & Chr(105) & Chr(108) & Chr(101) & Chr(32)   'helpfile
        lblHinweis = lblHinweis & Chr(104) & Chr(111) & Chr(119) & Chr(32)                          'how
        lblHinweis = lblHinweis & Chr(121) & Chr(111) & Chr(117) & Chr(32)                          'you
        lblHinweis = lblHinweis & Chr(99) & Chr(97) & Chr(110) & Chr(32)                            'can
        lblHinweis = lblHinweis & Chr(103) & Chr(101) & Chr(116) & Chr(32)                          'get
        lblHinweis = lblHinweis & Chr(97) & Chr(32)                                                 'a
        lblHinweis = lblHinweis & Chr(80) & Chr(114) & Chr(111) & Chr(102) & Chr(101) & Chr(115) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(97) & Chr(108) & Chr(32) 'Professional
        lblHinweis = lblHinweis & Chr(118) & Chr(101) & Chr(114) & Chr(115) & Chr(105) & Chr(111) & Chr(110) & Chr(46) & Chr(32) 'version.
    End If
    
    Me.Top = Query.Top + Query.height \ 2 - Me.height \ 2
    Me.Left = Query.Left + Query.width \ 2 - Me.width \ 2
    
    btnOK.Caption = LoadResString(3001 + Sprache)       '&OK
End Sub
