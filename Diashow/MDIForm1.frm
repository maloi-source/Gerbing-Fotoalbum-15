VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "GERBING Diashow"
   ClientHeight    =   8916
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11868
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Height          =   1092
      Left            =   0
      ScaleHeight     =   1044
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11868
      Begin VB.TextBox txtFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5280
         TabIndex        =   1
         Text            =   "txtFont"
         Top             =   240
         Visible         =   0   'False
         Width           =   972
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Initialize()
    init_global
    'InitCommonControls                                         'Gerbing 04.03.2013
    Set IniFso = New FileSystemObject                           'Gerbing 11.03.2017
End Sub

Private Sub MDIForm_Load()
    
    PublicCheckForDPI = 2                                       'Gerbing 11.03.2017
    Call AnpassenNutzerWunsch(Me)                               'Gerbing 11.03.2017

    'Me.WindowState = 2
    
'    Load XYPos
'    Load WertxForm
'    Load ListBoxForm
'    Load HilfeBoxForm
    DiashowForm.Show
'    If DiashowForm.chkFensterGrößeÄnderbar.Value = 0 Then       'Gerbing 24.04.2012                'Gerbing 30.01.2018
'        Call DiashowForm.MDIFormOhneTitle
'    End If
    XYPos.Hide              'nach DiashowForm, sonst war Sprachauswahl noch nicht dran
    WertxForm.Hide
    ListBoxForm.Hide
    HilfeBoxForm.Hide

    If DiashowForm.CommandLine <> "" Then
        frmBildMitGDIPlus.WindowState = 2
        DiashowForm.ZOrder 1
     End If
End Sub

Private Sub MDIForm_Resize()
    'MsgBox "MDIForm_Resize"
    On Error Resume Next
    Me.Width = Screen.Width / 2                                 'Gerbing 27.01.2018
    Me.Height = Screen.Height / 2                               'Gerbing 27.01.2018
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'MsgBox "MDIForm_Unload"
    Set DiashowForm.EXF = Nothing
    Set IniFso = Nothing                                        'Gerbing 11.03.2017
    Call DiashowForm.MDIFormMitTitle
End Sub

'Public Sub CaptionW(Optional ByRef NewCaption As String)
''When working with MDI forms things get a bit more complicated if you happen to support maximized child forms: VB updates the MDI form's
''caption automatically when a child form is maximized or a maximized state of a child is removed. Thus you need to do a custom function
''that updates the MDI caption whenever you want it to be updated.
''The code looks for the current active child form, gets it's actual caption and then shows it alongside the MDI form's own caption that
''is being stored in the Tag property of the MDI form. If the child form is not maximized, the MDI caption will not contain the child
''form's caption.
'
'    ' update to new caption if a non vbNullString was passed
'    If StrPtr(NewCaption) Then Me.Tag = NewCaption
'    ' must have active form
'    If Not Me.ActiveForm Is Nothing Then
'        ' see if window state is maximized
'        If Me.ActiveForm.WindowState = vbMaximized Then
'            ' show both child caption and MDI caption
'            UniCaption(Me) = UniCaption(Me.ActiveForm) & " - " & Me.Tag
'        Else
'            ' show only MDI caption
'            UniCaption(Me) = Me.Tag
'        End If
'    Else
'        ' show only MDI caption
'        UniCaption(Me) = Me.Tag
'    End If
'End Sub

