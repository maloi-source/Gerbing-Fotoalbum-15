VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmStartSoundAutomatisch 
   Caption         =   "frmStartSoundAutomatisch"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   492
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   2952
   ScaleWidth      =   5880
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer1 
      Height          =   732
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   4692
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8276
      _cy             =   1291
   End
End
Attribute VB_Name = "frmStartSoundAutomatisch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

