VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDateChoose 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Datum"
   ClientHeight    =   3780
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4032
   Icon            =   "DateChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4032
   StartUpPosition =   1  'Fenstermitte
   Begin MSComCtl2.MonthView DateChoose 
      Height          =   2256
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2268
      _ExtentX        =   4001
      _ExtentY        =   3979
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   146341890
      CurrentDate     =   36445
   End
End
Attribute VB_Name = "frmDateChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DateChoose_DateClick(ByVal DateClicked As Date)
    gstrDateChoose = DateChoose.Value
    Unload frmDateChoose
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017
    If Query.chkFensterGrößeÄnderbar.Value = 1 Then             'Gerbing 06.12.2005
        Me.Top = Form1.Top                                      'Gerbing 06.12.2006
        Me.Left = Form1.Left
    End If
    
    Me.Caption = LoadResString(1035 + Sprache)        '"Datum"
    DateChoose.Value = Format(gstrDateChoose, "dd/mm/yy")
End Sub

