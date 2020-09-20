VERSION 5.00
Begin VB.Form XYPos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Position des Mauszeigers"
   ClientHeight    =   1776
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3168
   Icon            =   "XYPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1776
   ScaleWidth      =   3168
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label lblBildgröße 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bildgröße in Pixel:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2412
   End
   Begin VB.Label lblYPos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Y:"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblXPos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "X:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   252
   End
End
Attribute VB_Name = "XYPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then             'Gerbing 06.12.2005
        Me.Top = Form1.Top                                      'Gerbing 06.12.2006
        Me.Left = Form1.Left
    End If

    Me.Caption = LoadResString(1100 + Sprache)        'Position des Mauszeigers
    Label2.Caption = LoadResString(1101 + Sprache)    'Bildgröße in Pixel:
    Me.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblnShowXYPos = False                           'Gerbing 04.12.2012
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Hide
    'Tastatur-Eingabe weiterreichen
    '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Form1.Form_KeyDown(KeyCode, Shift)
End Sub

