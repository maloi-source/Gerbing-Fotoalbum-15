VERSION 5.00
Begin VB.Form XYPos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Position des Mauszeigers"
   ClientHeight    =   1332
   ClientLeft      =   -12
   ClientTop       =   288
   ClientWidth     =   3936
   ControlBox      =   0   'False
   Icon            =   "XYPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1332
   ScaleWidth      =   3936
   Begin VB.Label lblBildgröße 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bildgröße in Pixel:"
      Height          =   612
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1812
   End
   Begin VB.Label lblYPos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Y:"
      Height          =   372
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   372
   End
   Begin VB.Label lblXPos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   612
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "X:"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   372
   End
End
Attribute VB_Name = "XYPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long
    
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    
    Me.Caption = LoadResString(1219 + Sprache) 'Position des Mauszeigers
    Label2.Caption = LoadResString(1220 + Sprache) 'Bildgröße in Pixel:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload Me
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Hide
    'Tastatur-Eingabe weiterreichen
    '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
End Sub

