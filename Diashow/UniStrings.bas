Attribute VB_Name = "UniStrings"
'*************************************************************************************************
'* UniStrings - Unicode string support for native controls & forms
'* ---------------------------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* Notes: in Windows 2000 you probably get the infamous question marks with UniCaption.
'*        However, in the taskbar you will see the correct Unicode caption.
'*************************************************************************************************
Option Explicit

' UniCaption
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

' three times faster than Split: only BinaryCompare is supported
Public Sub QuickSplit(Expression As String, ResultSplit() As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByRef IgnoreDelimiterWithin As String = vbNullString)
    Dim lngA As Long, lngB As Long, lngCount As Long, lngDelLen As Long, lngExpLen As Long, lngExpPtr As Long, lngIgnLen As Long, lngResults() As Long
    lngExpLen = LenB(Expression)
    lngDelLen = LenB(Delimiter)
    If lngExpLen > 0 And lngDelLen > 0 And (Limit > 0 Or Limit = -1&) Then
        lngIgnLen = LenB(IgnoreDelimiterWithin)
        If lngIgnLen Then
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            lngB = InStrB(1, Expression, IgnoreDelimiterWithin, Compare)
            Do Until (lngB And 1) Or (lngB = 0)
                lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
            Loop
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        Do Until (lngA And 1) Or (lngA = 0)
                            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                        Loop
                        lngCount = lngCount + 1
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                        Loop
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            Do Until (lngA And 1) Or (lngA = 0)
                                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                            Loop
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                                Do Until (lngB And 1) Or (lngB = 0)
                                    lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                                Loop
                            End If
                        End If
                    End If
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        Do Until (lngA And 1) Or (lngA = 0)
                            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                        Loop
                        lngCount = lngCount + 1
                        If lngCount = Limit Then Exit Do
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                        Loop
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            Do Until (lngA And 1) Or (lngA = 0)
                                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                            Loop
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                                Do Until (lngB And 1) Or (lngB = 0)
                                    lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                                Loop
                            End If
                        End If
                    End If
                Loop
            End If
        Else
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                    Loop
                    lngCount = lngCount + 1
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0 And lngCount < Limit
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                    Loop
                    lngCount = lngCount + 1
                Loop
            End If
        End If
        ReDim Preserve ResultSplit(0 To lngCount)
        If lngCount = 0 Then
            ResultSplit(0) = Expression
        Else
            lngExpPtr = StrPtr(Expression)
            ResultSplit(0) = LeftB$(Expression, lngResults(0) - 1)
            For lngCount = 0 To lngCount - 2
                ResultSplit(lngCount + 1) = MidB$(Expression, lngResults(lngCount) + lngDelLen, lngResults(lngCount + 1) - lngResults(lngCount) - lngDelLen)
            Next lngCount
            ResultSplit(lngCount + 1) = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
        End If
    Else
        ResultSplit = VBA.Split(vbNullString)
    End If
End Sub
' byte version of QuickSplit
Public Sub QuickSplitB(Expression As String, ResultSplit() As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByRef IgnoreDelimiterWithin As String = vbNullString)
    Dim lngA As Long, lngB As Long, lngCount As Long, lngDelLen As Long, lngExpLen As Long, lngExpPtr As Long, lngIgnLen As Long, lngResults() As Long
    lngExpLen = LenB(Expression)
    lngDelLen = LenB(Delimiter)
    If lngExpLen > 0 And lngDelLen > 0 And (Limit > 0 Or Limit = -1&) Then
        lngIgnLen = LenB(IgnoreDelimiterWithin)
        If lngIgnLen Then
            lngA = InStrB(1, Expression, Delimiter, Compare)
            lngB = InStrB(1, Expression, IgnoreDelimiterWithin, Compare)
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        lngCount = lngCount + 1
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                            End If
                        End If
                    End If
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0
                    If lngA + lngDelLen <= lngB Or lngB = 0 Then
                        lngResults(lngCount) = lngA
                        lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                        lngCount = lngCount + 1
                        If lngCount = Limit Then Exit Do
                    Else
                        lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                        If lngB Then
                            lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                            If lngA Then
                                lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                            End If
                        End If
                    End If
                Loop
            End If
        Else
            lngA = InStrB(1, Expression, Delimiter, Compare)
            If Limit = -1& Then
                ReDim lngResults(0 To (lngExpLen \ lngDelLen))
                Do While lngA > 0
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    lngCount = lngCount + 1
                Loop
            Else
                ReDim lngResults(0 To Limit - 1)
                Do While lngA > 0 And lngCount < Limit
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    lngCount = lngCount + 1
                Loop
            End If
        End If
        ReDim Preserve ResultSplit(0 To lngCount)
        If lngCount = 0 Then
            ResultSplit(0) = Expression
        Else
            lngExpPtr = StrPtr(Expression)
            ResultSplit(0) = LeftB$(Expression, lngResults(0) - 1)
            For lngCount = 0 To lngCount - 2
                ResultSplit(lngCount + 1) = MidB$(Expression, lngResults(lngCount) + lngDelLen, lngResults(lngCount + 1) - lngResults(lngCount) - lngDelLen)
            Next lngCount
            ResultSplit(lngCount + 1) = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
        End If
    Else
        ResultSplit = VBA.Split(vbNullString)
    End If
End Sub
' Split directly to multiple string variables, documentation: http://www.vbforums.com/showthread.php?t=538612
Public Sub SplitToVar(Expression As String, ByVal Delimiter As String, IgnoreDelimiterWithin As String, ParamArray Results())
    Dim lngA As Long, lngB As Long, lngCount As Long, lngDelLen As Long, lngExpLen As Long, lngExpPtr As Long, lngIgnLen As Long, lngResults() As Long, Compare As VbCompareMethod, Limit As Long
    If LenB(Delimiter) = 0 Then Delimiter = " "
    lngExpLen = LenB(Expression)
    lngDelLen = LenB(Delimiter)
    Compare = vbBinaryCompare
    For Limit = 0 To UBound(Results)
        Results(Limit) = vbNullString
    Next Limit
    If lngExpLen > 0 And lngDelLen > 0 And Limit > 0 Then
        lngIgnLen = LenB(IgnoreDelimiterWithin)
        If lngIgnLen Then
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            lngB = InStrB(1, Expression, IgnoreDelimiterWithin, Compare)
            Do Until (lngB And 1) Or (lngB = 0)
                lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
            Loop
            ReDim lngResults(0 To Limit - 1)
            Do While lngA > 0
                If lngA + lngDelLen <= lngB Or lngB = 0 Then
                    lngResults(lngCount) = lngA
                    lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                    Loop
                    lngCount = lngCount + 1
                    If lngCount = Limit Then Exit Do
                Else
                    lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                    Do Until (lngB And 1) Or (lngB = 0)
                        lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                    Loop
                    If lngB Then
                        lngA = InStrB(lngB + lngIgnLen, Expression, Delimiter, Compare)
                        Do Until (lngA And 1) Or (lngA = 0)
                            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                        Loop
                        If lngA Then
                            lngB = InStrB(lngB + lngIgnLen, Expression, IgnoreDelimiterWithin, Compare)
                            Do Until (lngB And 1) Or (lngB = 0)
                                lngB = InStrB(lngB + 1, Expression, IgnoreDelimiterWithin, Compare)
                            Loop
                        End If
                    End If
                End If
            Loop
        Else
            lngA = InStrB(1, Expression, Delimiter, Compare)
            Do Until (lngA And 1) Or (lngA = 0)
                lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
            Loop
            ReDim lngResults(0 To Limit - 1)
            Do While lngA > 0 And lngCount < Limit
                lngResults(lngCount) = lngA
                lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                Do Until (lngA And 1) Or (lngA = 0)
                    lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                Loop
                lngCount = lngCount + 1
            Loop
        End If
        If lngCount = 0 Then
            Results(0) = Expression
            Expression = vbNullString
        Else
            lngExpPtr = StrPtr(Expression)
            Results(0) = LeftB$(Expression, lngResults(0) - 1)
            For lngCount = 0 To lngCount - 2
                Results(lngCount + 1) = MidB$(Expression, lngResults(lngCount) + lngDelLen, lngResults(lngCount + 1) - lngResults(lngCount) - lngDelLen)
            Next lngCount
            If UBound(Results) > lngCount Then
                Results(lngCount + 1) = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
                Expression = vbNullString
            Else
                Expression = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
            End If
        End If
    End If
End Sub
Public Property Get UniCaption(ByRef Control As Object) As String
    Dim lngLen As Long, lngPtr As Long
    ' validate supported control
    If Not Control Is Nothing Then
        If _
            (TypeOf Control Is CheckBox) _
        Or _
            (TypeOf Control Is CommandButton) _
        Or _
            (TypeOf Control Is Form) _
        Or _
            (TypeOf Control Is Frame) _
        Or _
            (TypeOf Control Is MDIForm) _
        Or _
            (TypeOf Control Is OptionButton) _
        Then
            ' get length of text
            lngLen = DefWindowProcW(Control.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
            ' must have length
            If lngLen Then
                ' create a BSTR of that length
                lngPtr = SysAllocStringLen(0, lngLen)
                ' make the property return the BSTR
                PutMem4 ByVal VarPtr(UniCaption), ByVal lngPtr
                ' call the default Unicode window procedure to fill the BSTR
                DefWindowProcW Control.hWnd, WM_GETTEXT, lngLen + 1, ByVal lngPtr
            End If
        Else
            ' go ahead and try the default property
            On Error Resume Next
            UniCaption = Control
        End If
    End If
End Property
Public Property Let UniCaption(ByRef Control As Object, ByRef NewValue As String)
    ' validate supported control
    If Not Control Is Nothing Then
        If _
            (TypeOf Control Is CheckBox) _
        Or _
            (TypeOf Control Is CommandButton) _
        Or _
            (TypeOf Control Is Form) _
        Or _
            (TypeOf Control Is Frame) _
        Or _
            (TypeOf Control Is MDIForm) _
        Or _
            (TypeOf Control Is OptionButton) _
        Then
            ' call the default Unicode window procedure and pass the BSTR pointer
            DefWindowProcW Control.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
        Else
            ' go ahead and try the default property
            On Error Resume Next
            Control = NewValue
        End If
    End If
End Property
