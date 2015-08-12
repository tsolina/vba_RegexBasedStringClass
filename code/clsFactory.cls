VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function regexCreate(iPattern As String, Optional iFlags As String = vbNullString) As clsRegex
    Set regexCreate = New clsRegex
    regexCreate.Pattern iPattern, iFlags
End Function
Public Function stringCreate(iStr As String) As clsString
    Set stringCreate = New clsString
    stringCreate.Create iStr
End Function
