VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FxStackFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type TThis
    ModuleName  As String
    MemberName  As String
    Params      As Collection
End Type
Private this As TThis

''' Creates a new stack frame object,
''' representing procedure call and it's arguments
Public Function Create(ByVal ModuleName As String, MemberName As String, _
        Optional Params As Variant) As FxStackFrame
Attribute Create.VB_Description = "Creates a new stack frame object, representing procedure call and it's arguments"
    With this
        .ModuleName = ModuleName
        .MemberName = MemberName

        If Not IsMissing(Params) Then
            Dim i As Long
            Set .Params = New Collection
            For i = LBound(Params) To UBound(Params)
                .Params.Add Params(i)
            Next i
        End If
    End With
    Set Create = Self
End Function

Public Property Get Self() As FxStackFrame
    Set Self = Me
End Property

Public Property Get ModuleName() As String
    ModuleName = this.ModuleName
End Property

Public Property Get MemberName() As String
    MemberName = this.MemberName
End Property

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a String containing the stack frame."
    ToString = this.ModuleName & "." & this.MemberName & "(" & ParamsToString() & ")"
End Function

Public Function ParamsToString() As String
Attribute ParamsToString.VB_Description = "Returns a String containing the param types and their values."
    Dim Result      As String
    Dim i           As Long
    Dim Param       As Variant

    If this.Params Is Nothing Then Exit Function

    For Each Param In this.Params
        i = i + 1
        Result = Result & TypeName(Param) & ":"
        If IsObject(Param) Then
            Result = Result & ObjPtr(Param)
        ElseIf IsArray(Param) Then
            Result = Result & "[" & LBound(Param) & "-" & UBound(Param) & "]"
        ElseIf VarType(Param) = vbString Then
            Result = Result & Chr$(34) & Param & Chr$(34)
        Else
            Result = Result & CStr(Param)
        End If

        Result = Result & (IIf(i = this.Params.Count, vbNullString, ","))
    Next Param

    ParamsToString = Result
End Function

