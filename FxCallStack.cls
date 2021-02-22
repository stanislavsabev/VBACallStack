VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FxCallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Static class to manage callstack"

Option Explicit

Private Type TThis
    Frames As Collection
End Type
Private this As TThis

Public Sub Push(Module As String, Member As String, _
    ParamArray Params() As Variant)
Attribute Push.VB_Description = "Pushes a new stack frame onto the call stack."
    Dim Frame As New FxStackFrame
    If UBound(Params) <> -1 Then
        Dim Pars() As Variant
        Pars = Params
        this.Frames.Add Frame.Create(Module, Member, Pars)
    Else
        this.Frames.Add Frame.Create(Module, Member)
    End If
End Sub

Public Function Pop() As FxStackFrame
Attribute Pop.VB_Description = "Removes the last stack frame from the to of the stack."
    If Count = 0 Then Exit Function
    Set Pop = Peek
    this.Frames.Remove this.Frames.Count
End Function

Public Function Peek() As FxStackFrame
Attribute Peek.VB_Description = "Returns the top-most stack frame"
    If Count = 0 Then Exit Function
    Set Peek = this.Frames(this.Frames.Count)
End Function

Public Property Get ToCollection() As Collection
    Set ToCollection = this.Frames
End Property

Public Property Get Count() As Long
    Count = this.Frames.Count
End Property

Public Sub Sync(ByVal Module As String, ByVal Member As String, _
        ParamArray Params() As Variant)
    Dim Frame As New FxStackFrame

    If Count = 0 Then Exit Sub

    If UBound(Params) <> -1 Then
        Dim Pars() As Variant
        Pars = Params
        Frame.Create Module, Member, Pars
    Else
        Frame.Create Module, Member
    End If

    Do Until Peek().ToString() = Frame.ToString()
        Pop
    Loop
End Sub

Public Sub Clear()
    Set this.Frames = Nothing
    Set this.Frames = New Collection
    Err.Clear
End Sub

Public Function ToString(Optional Limit As Variant) As String
Attribute ToString.VB_Description = "Returns a String containing the stack trace."
    Dim Result  As String
    Dim i       As Long
    Dim j       As Long
    Dim Lim     As Long

    On Error GoTo Failed
    If IsMissing(Limit) Then
        Lim = this.Frames.Count - 1
    Else
        Lim = CLng(Limit)
        If Lim < 1 Or Lim > Count Then
            Err.Raise 5, , "Error in FxCallStack.ToString(): Invalid stack trace limit: " & Lim
        End If
    End If

    j = (this.Frames.Count - Lim)
    For i = this.Frames.Count To j Step -1
        Result = "> " & this.Frames(i).ToString() & vbNewLine & Result
    Next i

    ToString = Result
Exit Function
Failed:
    Debug.Print "Error "; Err.Number; vbNewLine; Err.Description
End Function

Public Function TracebackToString(Optional Limit As Variant) As String
    TracebackToString = Join(Array( _
            WorksheetFunction.Rept("-", 34), _
            "Traceback (most recent call last):", _
            Me.ToString() _
        ), vbNewLine)
End Function

Public Sub PrintExcInfo(ErrNumber As Long, ErrDescription As String)
    Debug.Print ExcInfo(ErrNumber, ErrDescription)
End Sub

Public Sub PrintExcInfoAndClear(ErrNumber As Long, ErrDescription As String)
    Call PrintExcInfo(ErrNumber, ErrDescription)
    Me.Clear
End Sub

Public Property Get ExcInfo(ErrNumber As Long, ErrDescription As String) As String
    If ErrNumber = 0 Then Exit Property
    ExcInfo = "Runtime error " & ErrNumber & ": " & ErrDescription _
        & vbNewLine & Me.TracebackToString()
End Property

Public Sub ShowExcInfo(ErrNumber As Long, ErrDescription As String)
    MsgBox ExcInfo(ErrNumber, ErrDescription), vbCritical, "Error"
End Sub

Public Sub ShowExcInfoAndClear(ErrNumber As Long, ErrDescription As String)
    Call ShowExcInfo(ErrNumber, ErrDescription)
    Me.Clear
End Sub

Private Sub Class_Initialize()
    Set this.Frames = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Frames = Nothing
End Sub
