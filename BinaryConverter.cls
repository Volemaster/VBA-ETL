VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private mHexMap(0 To &HFF) As String * 2
Private mBytes(0 To 13) As Byte
Private mVariantType As Integer
Private mPointer As LongPtr
Private mSize As Long
Private mIndex As Long
Private mOffset As Long
Private mPrefixSize As Long

Implements IBinaryConverter

Private Sub Class_Initialize()
  Dim i As Long
  For i = 0 To &HFF
    mHexMap(i) = Right$("0" & Hex(i), 2)
  Next
End Sub

Public Function GetBinary(value As Variant, Optional prefix As String = "0x", Optional Profile As BinaryProfile = Nothing) As String
  Call Me.ToBinary(value, GetBinary, prefix)
End Function

Public Sub ToBinary(bin As Variant, ByRef binaryOut As String, Optional prefix As String = "0x", Optional Profile As BinaryProfile = Nothing)
  If IsNull(bin) Then
    binaryOut = "NULL"
    Exit Sub
  End If
  If Profile Is Nothing Then
    Set Profile = GetProfile(bin)
  ElseIf Profile.size = 0 Then
    Set Profile = GetProfile(bin)
  End If
  
  With Profile
  
    mSize = .size
    If mSize = 0 Then
      Set Profile = Nothing
      binaryOut = prefix & "00"
      Exit Sub
    End If
    
    mVariantType = .VarType
    mPrefixSize = Len(prefix)
    
    If (mVariantType And VarTypeEnum.VT_ARRAY) Then
      If mSize = 1 Then
        mOffset = LBound(bin)
        mSize = UBound(bin) - mOffset + 1
        binaryOut = prefix & String((mSize) * 2, "0")
        For mIndex = 0 To mSize - 1
          Mid$(binaryOut, 1 + mPrefixSize + 2 * mIndex, 2) = mHexMap(CByte(bin(mIndex + mOffset)))
        Next
        Exit Sub
      Else
        Err.Raise vbObjectError + CustomErrorEnum.erNotImplementedException, "BinaryConverter.AppendBinary", "No binary conversion specified for " & typeName(bin) & " arrays."
        Exit Sub
      End If
    End If
    
  ' Decimal values! What fun!
    
    If .IsPointer Then
    ' Dereference the pointer at +0x08 bytes. _
      This really shouldn't happen, since I would expect it _
      to be handled in the byte array up above.
      Debug.Assert False
      CopyMemory mPointer, ByVal VarPtr(bin) + .PointerOffset, 8
    Else
    ' Should only be decimals, I think?
      Debug.Assert .PointerOffset = 2
      mPointer = VarPtr(bin) + .PointerOffset
    End If
    
    binaryOut = prefix & String((mSize) * 2, "0")
    CopyMem ByVal VarPtr(mBytes(0)), ByVal mPointer, mSize
    For mIndex = mSize - 1 To 0 Step -1
      Mid$(binaryOut, 1 + mPrefixSize + (mSize - mIndex - 1) * 2, 2) = mHexMap(mBytes(mIndex))
    Next
  End With

End Sub

Private Function GetProfile(value) As BinaryProfile
  If IsNull(value) Then
    Exit Function
  End If
'
'  ' Yes, we've declared everything at module level so we can _
'    avoid the tiny overhead of allocating space on the stack _
'    every time we run this. This is generally a pretty horrible _
'    idea, but we're trying to save as much time as we can here.
  
  Set GetProfile = New BinaryProfile
  CopyMemory mVariantType, ByVal VarPtr(value), 2
  
  With GetProfile
    .VarType = mVariantType
    If mVariantType And VarTypeEnum.VT_BYREF Then
    ' Dereference the pointer at +0x08 bytes
      .IsPointer = True
      .PointerOffset = 8
    ElseIf (mVariantType And VarTypeEnum.VT_DECIMAL) = VarTypeEnum.VT_DECIMAL Then
    ' The decimal is stored within the variant itself, _
      starting at +0x02
      .PointerOffset = 2
    Else
    ' The is stored within the variant itself at +0x08
      .PointerOffset = 8
    End If
    Select Case (mVariantType And &H3F)
      Case VT_EMPTY, VT_NULL
        .size = 0
      Case VT_I1, VT_UI1
        .size = 1
      Case VT_I2, VT_UI2, VT_BOOL
        .size = 2
      Case VT_INT, VT_UINT, VT_UI4, VT_I4, VT_R4, VT_ERROR
        .size = 4
      Case VT_R8, VT_CY, VT_DATE, VT_I8, VT_UI8, VT_PTR, VT_INT_PTR, VT_UINT_PTR
        .size = 8
      Case VT_DECIMAL
        .size = 14
      Case VT_BSTR, VT_DISPATCH, VT_VARIANT, VT_UNKNOWN, VT_VOID, VT_HRESULT, VT_SAFEARRAY, VT_CARRAY, VT_USERDEFINED, VT_LPSTR, VT_LPWSTR, VT_RECORD, VT_ARRAY, VT_BYREF                  'vbString, vbObject, vbError, vbVariant, vbDataObject, vbUserDefinedType
        ' Nope. Write your own.
        Err.Raise vbObjectError + CustomErrorEnum.erNotImplementedException, "BinaryConverter.GetProfile", "No binary conversion specified for " & typeName(value) & "."
        Exit Function
      Case Else
        ' Nope. Write your own.
        Err.Raise vbObjectError + CustomErrorEnum.erNotImplementedException, "BinaryConverter.GetProfile", "No binary conversion specified for " & typeName(value) & "."
        Exit Function
    End Select
  
  End With
End Function

Private Sub IBinaryConverter_AppendBinary(ByRef binaryOut As IStringBuilder, value As Variant, Optional prefix As String = "0x", Optional Profile As BinaryProfile = Nothing)
  Dim result As String
  Call IBinaryConverter_ToBinary(result, value, prefix, Profile)
  binaryOut.Add result
End Sub

Private Function IBinaryConverter_GetBinary(bin As Variant, Optional prefix As String = "0x", Optional Profile As BinaryProfile = Nothing) As String
  IBinaryConverter_GetBinary = Me.GetBinary(bin, prefix, Profile:=Profile)
End Function

Private Sub IBinaryConverter_ToBinary(ByRef binaryOut As String, value As Variant, Optional prefix As String = "0x", Optional Profile As BinaryProfile = Nothing)
  Me.ToBinary value, binaryOut, prefix, Profile
End Sub
