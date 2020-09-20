VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovieModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Programmer: Jason Hensley

'Email: elitecobra@hotmail.com

'Website: www16.brinkster.com/codesource

'Description: Easily create a movie player to play
'avi, mpeg, mpeg2, quicktime, divx(need divx codec) and more
'Designed to be easy to use and small in size. Please email
'me if you have any feedback or problems.

'Problems/Bugs: If you use the mpegvideo device type
'and it skips the video just don't use the getPositionInFrames
'function


Option Explicit

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long 'Get the error message of the mcidevice if any
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long 'Send command strings to the mci device

Private Data As String * 128 ' Used to store our return data
Public Error As Long ' Used to store our error message
Public Filename  As String ' Used to store our file

Public Function stepFrames(Value As Long)
    'Step ahead a specified amount of frames
    'Ex. If the movie was on frame 20. And if you stepped
    '10 frames the movie would skip ahead 10 frames and
    'would be on frame 30.
    Error = mciSendString("step movie by " & Value, 0, 0, 0)
End Function
Public Function restoreSizeDefault()
    'This function will restore the movie to its original
    'size. Not if you use a child window
    Error = mciSendString("put movie window", 0, 0, 0)
End Function
Public Function openMovie()
    'Open a movie in the default window style(Popup)
    Dim a As Long
    Filename = Chr$(34) & Filename & Chr$(34)
    Error = mciSendString("close movie", 0, 0, 0)
    'Decide which way you want the mci device to work below
    
    'Specify the mpegvideo driver to play the movies
    Error = mciSendString("open " & Filename & " type mpegvideo alias movie", 0, 0, 0)
    
    'Let the mci device decide which driver to use
    'Error = mciSendString("open " & Filename & " alias movie", 0, 0, 0)
End Function
Public Function openMovieWindow(hWnd As Long, WindowStyle As String)
    'Style types = popup , child or overlapped
    'Child window would be a .hwnd window of your choice.
    'Ex. A picturebox control or a frame control would be
    'a child window
    Filename = Chr$(34) & Filename & Chr$(34)
    Error = mciSendString("close movie", 0, 0, 0)
    'Decide which way you want the mci device to work below
    
    'use the command below to play divx movies. Must have the Divx codec installed
    Error = mciSendString("open " & Filename & " type mpegvideo alias movie parent " & hWnd & " style " & WindowStyle & " ", 0, 0, 0)
    
    'Let the mci device decide which driver to use
    'Error = mciSendString("open " & Filename & " alias movie parent " & hWnd & " style " & WindowStyle & " ", 0, 0, 0)
End Function
Public Function minimizeMovie()
    'Minimize the movie window
    Error = mciSendString("window movie state minimized", 0, 0, 0)
End Function
Public Function playMovie()
    'Play the movie after you open it
    Error = mciSendString("play movie", 0, 0, 0)
End Function
Public Function hideMovie()
    'Hides the movie window
    Error = mciSendString("window movie state hide", 0, 0, 0)
End Function
Public Function showMovie()
    'Will show the window if it was hidden with the
    'hideMovie function
    Error = mciSendString("window movie state show", 0, 0, 0)
End Function
Public Function restoreMovie()
    'Will restore the window to its original state
    Error = mciSendString("window movie state restore", 0, 0, 0)
End Function
Public Function stopMovie()
    'Stops the playing of the movie
    Error = mciSendString("stop movie", 0, 0, 0)
End Function
Public Function extractCurrentMovieSize(wLeft As Long, wTop As Long, wWidth As Long, wHeight As Long)
    'Returns the size parameters of the movie
    On Error Resume Next
    Dim a As String
    Dim b As String
    Dim c As String
    Dim f As String
    Dim g As String
    a = getCurrentSize
    b = InStr(1, a, " ")
    c = InStr(b + 1, a, " ")
    f = Mid(a, c + 1)
    g = InStr(1, f, " ")
    wWidth = Val(left(f, g)) 'width
    wHeight = Val(Mid(f, g)) 'height
End Function
Public Function extractDefaultMovieSize(wWidth As Long, wHeight As Long)
    'Returns the default size of the movie even if the size
    'of the movie has been changed
    On Error Resume Next
    Dim a As String
    Dim b As String
    Dim c As String
    Dim f As String
    Dim g As String
    a = getDefaultSize
    b = InStr(1, a, " ") '2
    c = InStr(b + 1, a, " ") '4
    f = Mid(a, c + 1) '9
    g = InStr(1, f, " ")
    wWidth = Val(left(f, g)) 'width
    wHeight = Val(Mid(f, g)) 'height
End Function
Public Function getBitsPerPixel()
    'Will get the movie bitsperpixel
    'Works with avi movies only
    Error = mciSendString("status movie bitsperpel", Data, 128, 0)
    getBitsPerPixel = Val(Data)
End Function
Public Function getMovieInput() As String
    'Returns the current input source
    Error = mciSendString("status movie monitor input", Data, 128, 0)
    getMovieInput = Data
End Function
Public Function getMovieOutput() As String
    'Returns the current output source
    Error = mciSendString("status movie monitor output", Data, 128, 0)
    getMovieOutput = Data
End Function
Public Function getAudioStatus() As String
    'Check to see if the audio is on or off
    Error = mciSendString("status movie audio", Data, 128, 0)
    getAudioStatus = Data
End Function
Public Function sizeLocateMovie(left As Long, top As Long, Width As Long, Height As Long)
    'Change the size of the movie and the location of
    'the movie in Pixels
    Error = mciSendString("put movie window at " & left & " " & top & " " & Width & " " & Height, 0, 0, 0)
End Function
Public Function isMoviePlaying() As Boolean
    'Checks the status of the movie whether it is playing
    'or not
    Dim isPlaying As String
    Error = mciSendString("status movie mode", Data, 128, 0)
    isPlaying = left(Data, 7)
    If isPlaying = "playing" Then
        isMoviePlaying = True
    Else
        isMoviePlaying = False
    End If
End Function
Public Function checkError() As String
    'A very useful function for getting any errors
    'associated with the mci device
    checkError = Space$(255)
    mciGetErrorString Error, checkError, Len(checkError)
End Function
Public Function getDeviceName() As String
    'Returns the current device name in use
    Error = mciSendString("info movie product", Data, 128, 0)
    getDeviceName = Data
End Function
Public Function getDeviceVersion() As String
    'Returns the current version of the mci device in use
    Error = mciSendString("info movie version", Data, 128, 0)
    getDeviceVersion = Data
End Function
Public Function getNominalFrameRate() As Long
    'Returns the nominal frame rate of the movie file
    Error = mciSendString("status movie nominal frame rate wait", Data, 128, 0)
    getNominalFrameRate = Val(Data)
End Function
Public Function getFramePerSecRate() As Long
    'Returns the Frames Per Second of the movie file
    'avi and mpeg movies
    Error = mciSendString("status movie frame rate", Data, 128, 0)
    getFramePerSecRate = Val(Data) \ 1000
End Function
Public Function getCurrentSize() As String
    'Returns the current width, height of the movie
    Error = mciSendString("where movie destination max", Data, 128, 0)
    getCurrentSize = Data
End Function
Public Function getDefaultSize() As String
    'Returns the default width, height the movie
    Error = mciSendString("where movie source", Data, 128, 0)
    getDefaultSize = Data
End Function
Public Function getLengthInFrames() As Long
    'Get the length of the movie in frames
    Error = mciSendString("set movie time format frames", 0, 0, 0)
    Error = mciSendString("status movie length", Data, 128, 0)
    getLengthInFrames = Val(Data)
End Function
Public Function getLengthInMS() As Long
    'Get the length of the movie in milliseconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("status movie length", Data, 128, 0)
    getLengthInMS = Val(Data)
End Function
Public Function playFullScreen()
    'Play the movie in full screen mode
    Error = mciSendString("play movie fullscreen", 0, 0, 0)
End Function
Public Function getLengthInSec() As Long
    'Get the length of the movie in seconds
    getLengthInSec = getLengthInMS \ 1000
End Function
Public Function setVideoOff()
    'Set the video device off
    Error = mciSendString("set all video off", 0, 0, 0)
End Function
Public Function setVideoOn()
    'Set the video device on
    Error = mciSendString("set all video on", 0, 0, 0)
End Function
Public Function pauseMovie()
    'Pause the movie
    Error = mciSendString("pause movie", 0, 0, 0)
End Function
Public Function resumeMovie()
    'Resumes the movie
    Error = mciSendString("resume movie", 0, 0, 0)
End Function
Public Function getPositionInMS() As Long
    'Get the position of the movie in milliseconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("status movie position wait", Data, 128, 0)
    getPositionInMS = Val(Data)
End Function
Public Function getRate() As Long
    'Get the current speed of the movie
    Error = mciSendString("status movie speed", Data, 128, 0)
    getRate = Val(Data)
End Function
Public Function getPositionInFrames() As Long
    'Get the position of the movie in frames
    Error = mciSendString("set movie time format frames wait", 0, 0, 0)
    Error = mciSendString("status movie position", Data, 128, 0)
    getPositionInFrames = Val(Data)
End Function
Public Function getStatus() As String
    'Get the current mode of the movie
    'Playing, Stopped, Paused, Not Ready
    Error = mciSendString("status movie mode", Data, 128, 0)
    getStatus = StrConv(Data, vbProperCase)
End Function
Public Function closeMovie()
    'Close the mci device
    Error = mciSendString("close all", 0, 0, 0)
End Function
Public Function getFormatPosition() As String
    'Get the position in a userfriendly time format
    getFormatPosition = getThisTime(getPositionInMS)
End Function
Public Function getFormatLength() As String
    'Get the length in a userfriendly time format
    getFormatLength = getThisTime(getLengthInMS)
End Function

Private Function getThisTime(ByVal timein As Long) As String
    'Used to format the position and duration
    On Error GoTo TrapIt
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    If conH > 0 Then
        strRetTime = Trim(Str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "0"
    End If
    strRetTime = strRetTime & ":"
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    getThisTime = strRetTime
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
Public Function getVolume() As Long
    'Get the current volume level
    Error = mciSendString("status movie volume", Data, 128, 0)
    getVolume = Val(Data)
End Function
Public Function getVideoStatus() As String
    'Get the status of the video. Returns on or off
    Error = mciSendString("status movie video", Data, 128, 0)
    getVideoStatus = Data
End Function
Public Function getTimeFormat() As String
    'Returns the current time format. Frames or Millisecond
    Error = mciSendString("status movie time format", Data, 128, 0)
    getTimeFormat = Data
End Function
Public Function getLeftVolume() As Long
    'Returns the volume value of the left channel
    Error = mciSendString("status movie left volume", Data, 128, 0)
    getLeftVolume = Val(Data)
End Function
Public Function getPositionInSec() As Long
    'Get the position of the movie in seconds
    getPositionInSec = getPositionInMS \ 1000
End Function
Public Function getRightVolume() As Long
    'Get the volume value of the right channel
    Error = mciSendString("status movie right volume", Data, 128, 0)
    getRightVolume = Data
End Function
Public Function setAudioOff()
    'Turns of the audio device
    Error = mciSendString("set movie audio all off", 0, 0, 0)
End Function
Public Function setAudioOn()
    'turns on the audio device
    Error = mciSendString("set movie audio all on", 0, 0, 0)
End Function
Public Function setLeftOff()
    'Turns of the left channel
    Error = mciSendString("set movie audio left off", 0, 0, 0)
End Function
Public Function setRightOff()
    'Turns of the right channel
    Error = mciSendString("set movie audio right off", 0, 0, 0)
End Function
Public Function setLeftOn()
    'Turns on the left channel
    Error = mciSendString("set movie audio left on", 0, 0, 0)
End Function
Public Function setRightOn()
    'Truns on the right channel
    Error = mciSendString("set movie audio right on", 0, 0, 0)
End Function
Public Function setDoorOpen()
    'Open the cdrom door
    Error = mciSendString("set cdaudio door open", 0, 0, 0)
End Function
Public Function setDoorClosed()
    'Close the cdrom door
    Error = mciSendString("set cdaudio door closed", 0, 0, 0)
End Function
Public Function setVolume(Value As Long)
    'Raise or lower the volume for both channels
    '1000 max - 0 min
    Error = mciSendString("setaudio movie volume to " & Value, 0, 0, 0)
End Function
Public Function setPositionTo(Second As Long)
    'Sets the position of the movie to play at
    On Error Resume Next
    Second = Second * 1000
    If isMoviePlaying = True Then
    mciSendString "play movie from " & Second, 0, 0, 0
    ElseIf isMoviePlaying = False Then
    mciSendString "seek movie to " & Second, 0, 0, 0
    End If
    On Error GoTo 0
    End Function
Public Function restartMovie()
    'Sets the movie to the beginning and call the playMovie
    'function to start playing from the beginning
    Error = mciSendString("seek movie to start", 0, 0, 0)
    playMovie
End Function
Public Function rewindByMS(numMS As Long)
    'Rewind the movie a specified number of milliseconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInMS - numMS, 0, 0, 0)
End Function
Public Function rewindByFrames(numFrames As Long)
    'Rewind the movie by a specified number of frames
    Error = mciSendString("set movie time format frames", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInFrames - numFrames, 0, 0, 0)
End Function
Public Function rewindBySeconds(numSec As Long)
    'Rewind the movie by a specified number of seconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInMS - 1000 * numSec, 0, 0, 0)
End Function
Public Function forwardByFrames(numFrames As Long)
    'Forward the movie a specified number of frames
    Error = mciSendString("set movie time format frames", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInFrames + numFrames, 0, 0, 0)
End Function
Public Function forwardByMS(numMS As Long)
    'Forward the movie a specified number of milliseconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInMS + numMS, 0, 0, 0)
End Function
Public Function forwardBySeconds(numSec As Long)
    'Forward the movie a specified number of seconds
    Error = mciSendString("set movie time format ms", 0, 0, 0)
    Error = mciSendString("play movie from " & getPositionInMS + 1000 * numSec, 0, 0, 0)
End Function
Public Function checkDeviceReady() As String
    'Returns true or false depending if the mci device
    'is ready or not
    Error = mciSendString("status movie ready", Data, 128, 0)
    checkDeviceReady = Data
End Function
Public Function setSpeed(Value As Long)
    'Set the current playing spped of the movie
    '0 = as fast as possible without losing frames
    'Values 1 - 2000 - 2000 being fastest
    Error = mciSendString("set movie speed " & Value, 0, 0, 0)
End Function
Public Function setLeftVolume(Value As Long)
    'Set the value of the left volume
    Error = mciSendString("setaudio movie left volume to " & Value, 0, 0, 0)
End Function
Public Function setRightVolume(Value As Long)
    'Set the value of the right volume
    Error = mciSendString("setaudio movie right volume to " & Value, 0, 0, 0)
End Function
Sub timeOut(duration)
    'Pauses for a specified amount of milliseconds
    Dim StartTime As Long
    Dim x As Long
    StartTime = Timer
    Do While Timer - StartTime < duration
        x = DoEvents()
    Loop
    Exit Sub
End Sub
