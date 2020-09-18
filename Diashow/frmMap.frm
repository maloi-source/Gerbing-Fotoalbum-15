VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMap 
   Caption         =   "GEO Position"
   ClientHeight    =   11208
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   16896
   Icon            =   "frmMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   11208
   ScaleWidth      =   16896
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "9.1768399"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtLat 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "47.5117078"
      Top             =   120
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10572
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   16812
      ExtentX         =   29654
      ExtentY         =   18648
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label6 
      Caption         =   "Long"
      Height          =   252
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label5 
      Caption         =   "Lat"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Type ControlPositionType
        Left As Single
        Top As Single
        Width As Single
        Height As Single
        FontSize As Single
    End Type
    
    Private m_ControlPositions() As ControlPositionType
    Private m_FormWid As Single
    Private m_FormHgt As Single

Private Sub SaveSizes()
    Dim i As Integer
    Dim ctl As Control
    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                .Left = ctl.X1
                .Top = ctl.Y1
                .Width = ctl.X2 - ctl.X1
                .Height = ctl.Y2 - ctl.Y1
            Else
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl
    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                               'Gerbing 29.09.2018
    If gstrLat = "" Or gstrLong = "" Then
        MsgBox "Supply a latitude and longitude value.", "Missing Data"
    End If
    
    Dim lat As String
    Dim lon As String
    Dim queryAddress As String
    
    'queryAddress = queryAddress & "?force=tt&hl=de-AT" 'bringt nichts
    'Rabenstein 50,83136366 12,8308115
    'queryAddress = "http://www.openstreetmap.org/?mlat=" & gstrLat & "&mlon=" & gstrLong & "&zoom=14&layers=M?force=tt&hl=de-AT"
    queryAddress = "http://www.openstreetmap.org/?mlat=" & gstrLat & "&mlon=" & gstrLong & "&zoom=16&layers=M?force=tt&hl=de-AT" 'Gerbing 29.09.2018
    'Wenn die GPS-Koordinaten als Dezimal-Koordinaten vorliegen (49.4122) sind sie von openstreetmap sofort verwendbar
    'Wenn die GPS-Koordinaten als Sexagesimal-Koordinaten vorliegen (50-49-57,43), müssen sie in Dezimal-Koordinaten umgewandelt werden
    txtLat = gstrLat
    txtLong = gstrLong
    WebBrowser1.Navigate queryAddress

    SaveSizes
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

'auskommentiert Gerbing 04.10.2018
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Me.Hide
'    'Tastatur-Eingabe weiterreichen
'    '-> Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call DiashowForm.Form_KeyDown(KeyCode, Shift)
'    Unload Me
'End Sub

Private Sub ResizeControls()
    Dim i As Integer
    Dim ctl As Control
    Dim x_scale As Single
    Dim y_scale As Single
    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub
    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt
    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            Else
                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                ctl.Font.Size = y_scale * .FontSize
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl
End Sub
