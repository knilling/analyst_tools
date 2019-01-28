VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' https://stackoverflow.com/a/4878884
Dim pStack As Collection
Public Function Pop() As Variant
    With pStack
        If .Count > 0 Then
            Pop = .Item(.Count)
            .Remove .Count
        End If
    End With
End Function
Public Function Push(newItem As Variant) As Variant
    With pStack
        .Add newItem
        Push = .Item(.Count)
    End With

End Function
Public Sub init()
    Set pStack = New Collection
End Sub
