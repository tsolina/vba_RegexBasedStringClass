VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' - Author: Tomislav Solina -
' - eMail : tsolina@gmail.com -

Private Const mChunk As Long = 16
Private mString As String
Private mAllocated As Long
Private mStringLen As Long

Private Sub Class_Initialize()
    mAllocated = 0
    mStringLen = 0
    mString = vbNullString
End Sub

Private Sub Class_Terminate()
End Sub


Public Property Get Value() As String
    Value = Mid(mString, 1, mStringLen)
End Property
Public Property Get Length() As Long
    Length = mStringLen
End Property




Private Function setAllocationLength(iLen As Long) As Long
    setAllocationLength = (Int(iLen / mChunk) + 1) * mChunk
End Function

Private Function regexCreate(iPattern As String, Optional iFlags As String = vbNullString) As clsRegex
    Set regexCreate = New clsRegex
    regexCreate.Pattern iPattern, iFlags
End Function
Private Function stringCreate(iStr As String) As clsString
    Set stringCreate = New clsString
    stringCreate.Create iStr
End Function

Public Function Add(iString As String) As clsString
    Set Add = Me
    
    Dim sLen As Long
    sLen = Len(iString)
    
    If sLen = 0 Then Exit Function
    
    If Me.Length = 0 Then
        Me.Create iString
        Exit Function
    End If
    
    Dim bckString As String, bckLength As Long
    bckString = Me.Value
    bckLength = Me.Length
    
    mStringLen = sLen + Me.Length
    If mStringLen > mAllocated Then
        mAllocated = setAllocationLength(mStringLen)
        mString = Space(mAllocated)
        Mid$(mString, 1) = bckString
    End If
    
    Mid$(mString, bckLength + 1) = iString
End Function
Public Function AddCopy(iString As String) As clsString
    Set AddCopy = stringCreate(Me.Value)
    
    If Len(iString) = 0 Then Exit Function
    
    AddCopy.Add iString
End Function



Public Function Create(iString As String) As clsString
    Set Create = Me
    
    mStringLen = Len(iString)
    If mStringLen = 0 Then Exit Function
    
    If mStringLen > mAllocated Then
        mAllocated = setAllocationLength(mStringLen)
        mString = Space(mAllocated)
    End If
   
    Mid$(mString, 1) = iString
End Function



Public Function CharAt(Optional iPos As Long = 1, Optional iCompare As VbCompareMethod = vbBinaryCompare) As String
    CharAt = Mid$(Me.Value, iPos, 1)
End Function

Public Function Contains(iPattern As String) As Boolean
    Contains = regexCreate(iPattern).Test(Me.Value)
End Function

Public Function EndsWith(iPattern As String, Optional iFlags As String = vbNullString) As Boolean
    EndsWith = regexCreate("/" & regexCreate(iPattern).getPattern & "$/", iFlags).Test(Me.Value)
End Function

Public Function IndexOf(iSearch As String, Optional iStart As Long = 1, Optional iCompare As VbCompareMethod = vbBinaryCompare) As Long
    IndexOf = InStr(iStart, Me.Value, iSearch, iCompare)
End Function

Public Function LastIndexOf(iSearch As String, Optional iStart As Long = -1) As Long
    LastIndexOf = InStrRev(Me.Value, iSearch, iStart)
End Function

Public Function LeftS(iCount As Long) As clsString
    Set LeftS = stringCreate(Mid$(Me.Value, 1, iCount))
End Function

Public Function MatchS(iPattern As String) As Collection
    Set MatchS = regexCreate(iPattern).Match(Me.Value)
End Function

Public Function MatchAsArray(iPattern As String, Optional iDelimiter As String = " ") As Variant()
    Dim mColl As Collection, vMatch() As Variant, i As Long
    Set mColl = Me.MatchS(iPattern)
    
    If mColl.Count = 0 Then mColl.Add CLng(0)

    ReDim vMatch(mColl.Count - 1)
    For i = 1 To mColl.Count
        vMatch(i - 1) = mColl(i)
    Next
    
    MatchAsArray = vMatch
    Erase vMatch
End Function

Public Function MatchAsString(iPattern As String, Optional iDelimiter As String = " ") As clsString
    Set MatchAsString = stringCreate(Join(Me.MatchAsArray(iPattern, iDelimiter)))
End Function

Public Function RightS(iCount As Long) As clsString
    Set RightS = stringCreate(Mid$(Me.Value, Me.Length - iCount + 1))
End Function

Public Function ReplaceNative(iFind As String, iReplace As String, Optional iStart As Long = 1, Optional iCount As Long = -1, Optional iCompare As VbCompareMethod = vbBinaryCompare) As clsString
    Set ReplaceNative = stringCreate(Replace(Me.Value, iFind, iReplace, iStart, iCount, iCompare))
End Function
Public Function ReplaceS(iPattern As String, iReplace As String) As clsString
    Set ReplaceS = stringCreate(regexCreate(iPattern).Replace(Me.Value, iReplace))
End Function

Public Function Slice(iStart As Long, Optional iEnd As Long = 0) As clsString
    Select Case iStart
        Case Is < 0
            iStart = Me.Length + iStart + 1
            If iStart < 1 Then iStart = 1
        Case 0, Is > Me.Length
            Set Slice = stringCreate(vbNullString)
            Exit Function
        Case Else
            
    End Select
    
    Select Case iEnd
        Case Is < 0
            iEnd = Me.Length + iEnd
            If iEnd < 1 Then iEnd = 1
        Case 0, Is > Me.Length + 1
            iEnd = Me.Length + 1
        Case Else
    End Select
    If iEnd > iStart Then iEnd = iEnd - iStart Else iEnd = 0

    Set Slice = stringCreate(Mid$(Me.Value, iStart, iEnd))
End Function

Public Function SplitS(Optional iDelimiter As String = " ") As Variant()
    Dim vArr() As Variant
    
    Dim mColl As VBScript_RegExp_55.MatchCollection, i As Long, begin As Long, fIndex As Long
    Set mColl = regexCreate(iDelimiter).Execute(Me.Value)
    
    ReDim vArr(mColl.Count)
    begin = 1
    For i = 1 To mColl.Count
        fIndex = mColl(i - 1).FirstIndex + 1
        vArr(i - 1) = Me.SubString(begin, fIndex - begin).Value
        begin = fIndex + Len(mColl(i - 1).Value)
    Next
    vArr(UBound(vArr)) = Me.SubString(begin).Value
    
    SplitS = vArr
    Erase vArr
    Set mColl = Nothing
End Function

Public Function StartsWith(iPattern As String, Optional iFlags As String = vbNullString) As Boolean
    StartsWith = regexCreate("/^" & regexCreate(iPattern).getPattern & "/", iFlags).Test(Me.Value)
End Function

Public Function SubString(iStart As Long, Optional iLength As Long = -1) As clsString
    If iLength = -1 Then iLength = mStringLen - iStart + 1
    
    Set SubString = stringCreate(Mid$(Me.Value, iStart, iLength))
End Function

Public Function ToLowerCase() As clsString
    Set ToLowerCase = stringCreate(LCase(Me.Value))
End Function

Public Function ToUpperCase() As clsString
    Set ToUpperCase = stringCreate(UCase(Me.Value))
End Function

Public Function TrimS() As clsString
    Set TrimS = stringCreate(Trim(Me.Value))
End Function

Public Function TrimLeft() As clsString
    Set TrimLeft = stringCreate(LTrim(Me.Value))
End Function

Public Function TrimRight() As clsString
    Set TrimRight = stringCreate(RTrim(Me.Value))
End Function
