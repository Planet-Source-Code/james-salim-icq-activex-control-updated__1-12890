Attribute VB_Name = "modFunction"
Option Explicit

'-----------------------------------------------------------------------'
' Peek / Poke Memory function - Byte, Boolean, Integer, Long and String '
'-----------------------------------------------------------------------'
Public Function Peek(ByVal Address As Long, ByVal VariableType As VbVarType, Optional varLength As Integer) As Variant
  If varLength = 0 Then
    Select Case VariableType
      Case vbByte
        Dim TempByte As Byte
        CopyMemory VarPtr(TempByte), Address, 1
        Peek = TempByte
      Case vbInteger
        Dim TempInt As Integer
        CopyMemory VarPtr(TempInt), Address, 2
        Peek = TempInt
      Case vbBoolean
        Dim TempBool As Boolean
        CopyMemory VarPtr(TempBool), Address, 2
        Peek = TempBool
      Case vbLong
        Dim TempLong As Long
        CopyMemory VarPtr(TempLong), Address, 4
        Peek = TempLong
    End Select
  Else
    Select Case VariableType
      Case vbByte
        ReDim TempByteArray(varLength - 1) As Byte
        CopyMemory VarPtr(TempByteArray(0)), Address, 1 * varLength
        Peek = TempByteArray
      Case vbInteger
        ReDim TempIntArray(varLength - 1) As Byte
        CopyMemory VarPtr(TempIntArray(0)), Address, 2 * varLength
        Peek = TempIntArray
      Case vbBoolean
        ReDim TempBoolArray(varLength - 1) As Boolean
        CopyMemory VarPtr(TempBoolArray(0)), Address, 2 * varLength
        Peek = TempBoolArray
      Case vbLong
        ReDim TempLongArray(varLength - 1) As Long
        CopyMemory VarPtr(TempLongArray(0)), Address, 4 * varLength
        Peek = TempLongArray
      Case vbString
        Dim TempString As String
        TempString = String(varLength, vbNullChar)
        CopyMemory StrPtr(TempString), Address, 2 * varLength
        Peek = TempString
    End Select
  End If
End Function

Public Function Poke(ByVal Address As Long, ByVal Value As Variant, Optional VariableType As VbVarType = vbVariant) As Variant
  If VariableType = vbVariant Then VariableType = varType(Value)
  Select Case VariableType
    Case vbByte
      Dim TempByte As Byte
      TempByte = Value
      CopyMemory Address, VarPtr(TempByte), 1
    Case vbInteger
      Dim TempInt As Integer
      TempInt = Value
      CopyMemory Address, VarPtr(TempInt), 2
    Case vbBoolean
      Dim TempBool As Boolean
      TempBool = Value
      CopyMemory Address, VarPtr(TempByte), 2
    Case vbLong
      Dim TempLong As Long
      TempLong = Value
      CopyMemory Address, VarPtr(TempLong), 4
    Case vbString
      Dim TempStr As String
      TempStr = Value
      CopyMemory Address, StrPtr(TempStr), LenB(TempStr)
  End Select

End Function

'-------------------------------------------'
' Conversion functions -- Cvx - String to X '
'                      -- Mkx - X to String '
'-------------------------------------------'
Function Cvi(ByVal X As String) As Integer
    X = Left(X & String(2, Chr(0)), 2)
    CopyVariable Cvi, ByVal X, 2&
End Function

Function Cvl(ByVal X As String) As Long
    X = Left(X & String(4, Chr(0)), 4)
    CopyVariable Cvl, ByVal X, 4&
End Function

Function Mki(ByVal X As Integer) As String
   Dim Temp As String * 2
   CopyVariable ByVal Temp, X, 2&
   Mki = Temp
End Function

Function Mkl(ByVal X As Long) As String
   Dim Temp As String * 4
   CopyVariable ByVal Temp, X, 4&
   Mkl = Temp
End Function

' -- IP to String Conversion (and Vice versa) --
Function MkIP(ByVal strHostIP As String) As String
  Dim Output As Long
  Dim TempIP As Variant
  Dim i As Integer
  Dim Value As Integer
  
  TempIP = Split(strHostIP, ".")
  ReDim Preserve TempIP(3)
  
  For i = 0 To 3
    Value = Val(TempIP(i))
    If Value > &HFF Then Value = 0
    Poke VarPtr(Output) + i, Value, vbByte
  Next i
  
  MkIP = Mkl(Output)
End Function

Function CvIP(ByVal strIPDump As String) As String
  Dim TempIP As Variant
  Dim i As Integer
  
  strIPDump = StrConv(strIPDump, vbFromUnicode)
  TempIP = Peek(StrPtr(strIPDump), vbByte, 4)
  
  For i = 0 To 2
    CvIP = CvIP + Trim(Str$(TempIP(i))) + "."
  Next i
  CvIP = CvIP + Trim(Str$(TempIP(3)))
End Function

'-----------------------'
' String Function - Cut '
'-----------------------'
Function lCut(strText As String, intBytesCount As Integer) As String
  Dim Temp As String
  Dim OutLen As Integer
  
  OutLen = Len(strText) - intBytesCount
  
  lCut = Left$(strText, intBytesCount)
  If OutLen < 0 Then strText = "": Exit Function
  
  Temp = String$(OutLen, vbNullChar)
  CopyMemory StrPtr(Temp), StrPtr(strText) + (intBytesCount * 2), OutLen * 2
  strText = Temp
End Function

Function rCut(strText As String, intBytesCount As Integer) As String
  Dim Temp As String
  Dim OutLen As Integer
  
  OutLen = Len(strText) - intBytesCount
  
  rCut = Right$(strText, intBytesCount)
  If OutLen < 0 Then strText = "": Exit Function
  
  Temp = String$(OutLen, vbNullChar)
  CopyMemory StrPtr(Temp), StrPtr(strText), OutLen * 2
  strText = Temp
End Function

'-----------------'
' Packet Function '
'-----------------'
Function StrAppend(ParamArray strText()) As String
  Dim i As Integer
  
  For i = LBound(strText) To UBound(strText)
    StrAppend = StrAppend + _
      Mki(Len(strText(i)) + 1) + _
      strText(i) + vbNullChar
  Next i
End Function

Function PacketRead(strText As String, ByVal varType As VbVarType, Optional varLength As Integer = -1, Optional bTrimString As Boolean = True)
  Dim TempStr As String
  Dim TempLen As Integer
  
  Select Case varType
    ' Byte Handling
    Case vbByte
      If varLength = -1 Then
        PacketRead = Asc(lCut(strText, 1))
      ElseIf varLength >= 0 Then
        TempStr = lCut(strText, varLength)
        TempStr = StrConv(TempStr, vbFromUnicode)
        PacketRead = Peek(StrPtr(TempStr), vbByte, varLength)
      End If
      
    ' Integer Handling
    Case vbInteger
      If varLength = -1 Then
        PacketRead = Cvi(lCut(strText, 2))
      ElseIf varLength >= 0 Then
        TempStr = lCut(strText, 2 * varLength)
        TempStr = StrConv(TempStr, vbFromUnicode)
        PacketRead = Peek(StrPtr(TempStr), vbInteger, varLength)
      End If
      
    ' Boolean Handling
    Case vbBoolean
      If varLength = -1 Then
        PacketRead = CBool(Cvi(lCut(strText, 2)))
      ElseIf varLength >= 0 Then
        TempStr = lCut(strText, 2 * varLength)
        TempStr = StrConv(TempStr, vbFromUnicode)
        PacketRead = Peek(StrPtr(TempStr), vbBoolean, varLength)
      End If

    ' Long Handling
    Case vbLong
      If varLength = -1 Then
        PacketRead = Cvl(lCut(strText, 4))
      ElseIf varLength >= 0 Then
        TempStr = lCut(strText, 4 * varLength)
        TempStr = StrConv(TempStr, vbFromUnicode)
        PacketRead = Peek(StrPtr(TempStr), vbLong, varLength)
      End If
      
    ' String Handling
    Case vbString
      If varLength = -1 Then
        Dim TempOutput As String
        TempLen = PacketRead(strText, vbInteger)
        If bTrimString Then
          PacketRead = lCut(strText, TempLen - 1)
          lCut strText, 1
        Else
          PacketRead = lCut(strText, TempLen)
        End If
      ElseIf varLength >= 0 Then
        PacketRead = lCut(strText, varLength)
      End If
  End Select
End Function

Function strJoinFE(ParamArray strText()) As String
  Dim i As Integer
  Dim Delimiter As String
  
  If UBound(strText) = -1 Then Exit Function
  
  Delimiter = Chr$(&HFE)
  For i = LBound(strText) To UBound(strText) - 1
    strJoinFE = strJoinFE + strText(i) + Delimiter
  Next i
  strJoinFE = strJoinFE + strText(UBound(strText))
End Function

Function JoinArray(ByVal ArgList As Variant) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim tempvar As Variant
  Dim bFoundArray  As Boolean
  ReDim Output(0) As Variant
    
  k = 0
  
  For i = 0 To UBound(ArgList)
    tempvar = ArgList(i)
    
    If IsArray(tempvar) = True Then
      For j = 0 To UBound(tempvar)
        If IsArray(tempvar(j)) Then bFoundArray = True
        
        ReDim Preserve Output(k) As Variant
        Output(k) = tempvar(j)
        k = k + 1
      Next j
    Else
      ReDim Preserve Output(k) As Variant
      Output(k) = tempvar
      k = k + 1
    End If
    
  Next i
  
  If bFoundArray = True Then Output = JoinArray(Output)
  JoinArray = Output
End Function

