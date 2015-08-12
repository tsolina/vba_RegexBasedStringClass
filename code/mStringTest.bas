Attribute VB_Name = "mStringTest"
Option Explicit

Public F As clsFactory

Private Sub doInitialization()
    Set F = New clsFactory
End Sub
Private Sub doTermination()
    Set F = Nothing
End Sub



Private Sub regexTest()
    Debug.Print "regexTest:", F.regexCreate("/\d/g").Replace("H3e3l7l2o1 Wor64ld!", vbNullString)
End Sub
Private Sub stringtest()
    With F.stringCreate("H e. l.l o.").TrimLeft.ReplaceS("/\.?\s?/g", vbNullString).Add(" World.........").Slice(1, -8).Add("!")
        Debug.Print "stringTest:", .Value, .Length
    End With
End Sub



Public Sub CATMain()
    doInitialization

    regexTest
    stringtest

    doTermination
End Sub
