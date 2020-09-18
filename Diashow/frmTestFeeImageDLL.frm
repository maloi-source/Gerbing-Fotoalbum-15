VERSION 5.00
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "EditCtlsU.ocx"
Begin VB.Form frmTestFeeImageDLL 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9348
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   9348
   StartUpPosition =   3  'Windows-Standard
   Begin EditCtlsLibUCtl.TextBox TxtU 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9132
      _cx             =   16108
      _cy             =   8700
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
         Size            =   10.2
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
      MultiLine       =   -1  'True
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   3
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "frmTestFeeImageDLL.frx":0000
      Text            =   "frmTestFeeImageDLL.frx":0020
   End
End
Attribute VB_Name = "frmTestFeeImageDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim dib As Long
    Dim bOK As Long
    Dim AlleFelder As String
    
    dib = FreeImage_Load(FIF_JPEG, gblStrAktuellGezeigtesBild, 0)
    'dib = FreeImage_Load(FIF_JPEG, App.Path & "\unicode.jpg", 0)
    'dib = FreeImage_Load(FIF_JPEG, App.Path & "\GPS Neues Rathaus in Hannover.jpg", 0)
    AlleFelder = printtagsfromdib(dib)
    TxtU.Text = AlleFelder
    ' Unload the dib
    FreeImage_Unload (dib)
End Sub

Function printtagsfromdib(dib As Long) As String
    On Error GoTo Finish
    Dim maintag() As FREE_IMAGE_TAG
    Dim exiftag() As FREE_IMAGE_TAG
    Dim gpstag() As FREE_IMAGE_TAG
    Dim result As Long
    Dim i As Integer
    Dim alltags As String

    result = FreeImage_GetAllMetadataTags(FIMD_EXIF_MAIN, dib, maintag)
    alltags = ""
    'get Main Meta Data
    alltags = alltags & "FIMD_EXIF_MAIN" & vbNewLine
    If (result > 0) Then
        For i = LBound(maintag) To UBound(maintag)
            alltags = alltags & maintag(i).Key & " - " & maintag(i).StringValue & ";" & vbNewLine
        Next i
    End If
    'Get EXIF Meta Data
    alltags = alltags & "FIMD_EXIF_EXIF" & vbNewLine
    result = FreeImage_GetAllMetadataTags(FIMD_EXIF_EXIF, dib, exiftag)
    If (result > 0) Then
        For i = LBound(exiftag) To UBound(exiftag)
            alltags = alltags & exiftag(i).Key & " - " & exiftag(i).StringValue & ";" & vbNewLine
        Next i
    End If
    'get GPS Meta Data
    alltags = alltags & "FIMD_EXIF_GPS" & vbNewLine
    result = FreeImage_GetAllMetadataTags(FIMD_EXIF_GPS, dib, gpstag)
    If (result > 0) Then
        For i = LBound(gpstag) To UBound(gpstag)
            alltags = alltags & gpstag(i).Key & " - " & gpstag(i).StringValue & ";" & vbNewLine
        Next i
    End If
    'get IPTC Meta Data
    alltags = alltags & "FIMD_IPTC" & vbNewLine
    result = FreeImage_GetAllMetadataTags(FIMD_IPTC, dib, gpstag)
    If (result > 0) Then
        For i = LBound(gpstag) To UBound(gpstag)
            alltags = alltags & gpstag(i).Key & " - " & gpstag(i).StringValue & ";" & vbNewLine
        Next i
    End If
    'get XMP Meta Data
    alltags = alltags & "FIMD_XMP" & vbNewLine
    result = FreeImage_GetAllMetadataTags(FIMD_XMP, dib, gpstag)
    If (result > 0) Then
        For i = LBound(gpstag) To UBound(gpstag)
            alltags = alltags & gpstag(i).Key & " - " & gpstag(i).StringValue & ";" & vbNewLine
        Next i
    End If
    'MsgBox alltags
    printtagsfromdib = alltags
Finish:
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
