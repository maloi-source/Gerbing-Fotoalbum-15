VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.2#0"; "CBLCtlsU.ocx"
Begin VB.Form Test 
   Caption         =   "Teste die Methode Point"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10224
   LinkTopic       =   "Form2"
   ScaleHeight     =   10980
   ScaleWidth      =   10224
   StartUpPosition =   3  'Windows-Standard
   Begin CBLCtlsLibUCtl.ListBox LstU 
      Height          =   9492
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   9372
      _cx             =   16531
      _cy             =   16743
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
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   0   'False
      VirtualMode     =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test IPTC"
      Height          =   852
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3012
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim rc As Boolean
    
    'Call LeseIPTC(AppPath & "\unicode.jpg")
    rc = LeseIPTC(AppPath & "\" & ChrW$(&H421) & ChrW$(&H43F) & ChrW$(&H43E) & ChrW$(&H440) & ChrW$(&H442) & ".jpg", LstU, True)    'kyrillisches sport & ChrW$(&H421) & ChrW$(&H43F) & ChrW$(&H43E) & ChrW$(&H440) & ChrW$(&H442) & ".jpg"      'kyrillisches sport
End Sub

