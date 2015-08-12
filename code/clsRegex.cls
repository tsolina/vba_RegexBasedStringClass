VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' - Author: Tomislav Solina -
' - eMail : tsolina@gmail.com -

Private mRegex As VBScript_RegExp_55.RegExp
Private mPattern As String
Private mFlags As String


Private Sub Class_Initialize()
    Set mRegex = New VBScript_RegExp_55.RegExp
    With mRegex
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = vbNullString
    End With
End Sub

Private Sub Class_Terminate()
    Set mRegex = Nothing
End Sub

Private Function IsRegex(iString As String) As Boolean
    Dim regex As VBScript_RegExp_55.RegExp
    Set regex = New VBScript_RegExp_55.RegExp
    
    regex.Global = True
    regex.Pattern = "^\/.*\/[gmi]{0,3}$"
    
    IsRegex = regex.Test(iString)
    
    Set regex = Nothing
End Function

Private Sub Parse(ByVal iPattern As String, Optional ByVal iFlags As String = vbNullString)
    mRegex.Global = False
    mRegex.IgnoreCase = False
    mRegex.MultiLine = False
    mRegex.Pattern = iPattern
    
    If iPattern = vbNullString Then Exit Sub

    Dim regex As VBScript_RegExp_55.RegExp
    Set regex = New VBScript_RegExp_55.RegExp
    
    regex.Pattern = "^\/(.*)\/([gmi]{0,3})$"
    
    If regex.Test(iPattern) Then ' - extract regex-
        Dim Matches As VBScript_RegExp_55.MatchCollection
        Set Matches = regex.Execute(iPattern)
        With Matches(0)
            mRegex.Pattern = .SubMatches(0)
            If iFlags = vbNullString Then iFlags = .SubMatches(1)
        End With
        Set Matches = Nothing
    Else ' - replace special regex chars in string -
        regex.Global = True
        regex.MultiLine = True
        regex.Pattern = "([.$^{[(|)*+?\\])"
        mRegex.Pattern = regex.Replace(iPattern, "\$1")
    End If
    
    If InStr(iFlags, "g") Then mRegex.Global = True
    If InStr(iFlags, "m") Then mRegex.MultiLine = True
    If InStr(iFlags, "i") Then mRegex.IgnoreCase = True
    
    Set regex = Nothing
End Sub

Public Property Get getPattern() As String
    getPattern = mRegex.Pattern
End Property


Public Sub Pattern(iPattern As String, Optional iFlags As String = vbNullString)
    Parse iPattern, iFlags
End Sub

Public Function Execute(iSourceString As String) As VBScript_RegExp_55.MatchCollection
    Set Execute = mRegex.Execute(iSourceString)
End Function

Public Function Match(iString As String) As Collection
    Set Match = New Collection
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = Me.Execute(iString)
    
    Dim i As Long
    For i = 1 To Matches.Count
        Match.Add Matches(i - 1).FirstIndex + 1 ', Matches(i - 1).Value
    Next
End Function

Public Function Replace(iSourceString As String, iReplaceVar As String) As String
    Replace = mRegex.Replace(iSourceString, iReplaceVar)
End Function

Public Function Test(iSourceString As String) As Boolean
    Test = mRegex.Test(iSourceString)
End Function
