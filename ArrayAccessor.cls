VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mOriginalArray As Variant
Private mRows As Long
Private mRowStart As Long
Private mColumns As Long
Private mColumnStart As Long
Private mActiveArray As Variant
Private mRowIndex As Long
Private mAddressPointer As LongPtr
Private mOriginalPointerTarget As LongPtr
Private mPointerSize As Long
Private mAddressOffset As Long
Private mVariantTypeSize As Long
Private mSafeArrayTypeSize As Long
Private mRowSize As LongPtr
Private mArrayBaseAddress As LongPtr

Private Sub Class_Initialize()
  Dim t As SAFEARRAY
  Dim vt As VARIANT_TYPE
  mVariantTypeSize = LenB(vt)
  mSafeArrayTypeSize = LenB(t)
  mAddressOffset = CLng(VarPtr(t.pvData) - VarPtr(t))
  mPointerSize = LenB(mAddressPointer)
End Sub

Private Sub Class_Terminate()
  If (mAddressPointer <> 0) Then
    CopyMem mAddressPointer, VarPtr(mOriginalPointerTarget), mPointerSize
  End If
End Sub

Public Sub MoveFirst()
  mRowIndex = mRowStart - 1
End Sub

Public Sub MoveNext()
  mRowIndex = mRowIndex + 1
End Sub

Public Property Get EOF() As Boolean
  EOF = Not (mRowIndex < UpperBound)
End Property

Public Property Get LowerBound() As Long
  LowerBound = mRowStart
End Property
  
Public Property Get UpperBound() As Long
  UpperBound = mRowStart + mRows - 1
End Property

Public Property Get Count() As Long
  Count = mRows
End Property

Private Sub InitializeCopy()
  Dim arrayInfo As SAFEARRAY
  Dim variantInfo As VARIANT_TYPE
  Dim arrayPointerBase As LongPtr
  ReDim mActiveArray(0 To mColumns - 1)
  
  variantInfo = GetVariantType(mActiveArray)
  
  arrayInfo = GetVariantSafeArray(variantInfo)
  mOriginalPointerTarget = arrayInfo.pvData
  arrayPointerBase = FollowPointer(VarPtr(variantInfo.data(0)))
  mAddressPointer = arrayPointerBase + mAddressOffset
  
  mOriginalPointerTarget = FollowPointer(mAddressPointer)
End Sub

Private Sub ChangeArrayTargetAddress(v As Variant, targetPointer As LongPtr)
  Dim vt As VARIANT_TYPE
  vt = GetVariantType(v)
End Sub

Private Function GetVariantType(v As Variant) As VARIANT_TYPE
  CopyMemory GetVariantType, ByVal VarPtr(v), LenB(GetVariantType)
End Function

Private Function GetVariantSafeArray(headerVariantInfo As VARIANT_TYPE) As SAFEARRAY
  If headerVariantInfo.VarType <> VarTypeEnum.VT_VariantArray Then
    Err.Raise vbObjectError + 5, "ArrayAccessor.GetSafeArray", "The value passed to the GetSafeArray() procedure must be a variant pointing to a variant array."
    Exit Function
  End If
  Dim ptr2 As LongPtr
  
  CopyMem VarPtr(GetVariantSafeArray), FollowPointer(VarPtr(headerVariantInfo.data(0))), LenB(GetVariantSafeArray)
  If GetVariantSafeArray.fFeatures <> (SafeArrayFlagsEnum.FADF_HAVEVARTYPE Or SafeArrayFlagsEnum.FADF_VARIANT) Then
    Err.Raise vbObjectError + 5, "ArrayAccessor.GetVariantSafeArray", "The value passed to the GetVariantSafeArray() procedure must be a variant pointing to a variant array."
    Exit Function
  End If

End Function

Private Function GetSafeArray(v As Variant) As SAFEARRAY
  Dim headerVariantInfo As VARIANT_TYPE
  headerVariantInfo = GetVariantType(v)
  GetSafeArray = GetVariantSafeArray(headerVariantInfo)
End Function

Public Sub Create(twoDimensionalArray As Variant)
  Dim headerVariantInfo As VARIANT_TYPE
  Dim arrayInfo As SAFEARRAY
  Dim PointerOffset As LongPtr
  mOriginalArray = twoDimensionalArray
  
  mRowStart = LBound(twoDimensionalArray, 2)
  mColumnStart = LBound(twoDimensionalArray, 1)
  
  mRows = 1 + UBound(twoDimensionalArray, 2) - mRowStart
  mColumns = 1 + UBound(twoDimensionalArray, 1) - mColumnStart
  
  InitializeCopy
  arrayInfo = GetSafeArray(twoDimensionalArray)
  
  If mRows > 1 Then
    mRowSize = VarPtr(twoDimensionalArray(0, mRowStart + 1)) - VarPtr(twoDimensionalArray(0, mRowStart))
  End If
  mRowIndex = mRowStart - 1
  
  mArrayBaseAddress = arrayInfo.pvData
  PointerOffset = VarPtr(arrayInfo.pvData) - VarPtr(arrayInfo)
  
End Sub

Public Function NextRow() As Variant
  If mRowIndex >= (mRowStart + mRows - 1) Then
    Err.Raise vbObjectError + 5, "ArrayAccessor.NextRow()", "Already at last row."
    Exit Function
  Else
    mRowIndex = mRowIndex + 1
  End If
  Dim ai1 As SAFEARRAY, ai2 As SAFEARRAY, vi1 As VARIANT_TYPE, vi2 As VARIANT_TYPE
  Dim ptr1 As LongPtr, ptr2 As LongPtr
  ptr1 = mArrayBaseAddress + ((mRowIndex - mRowStart) * mRowSize)
  CopyMem mAddressPointer, VarPtr(ptr1), mPointerSize
  
  NextRow = mActiveArray
End Function

Private Function FollowPointer(ByVal address As LongPtr) As LongPtr
  CopyMem VarPtr(FollowPointer), address, mPointerSize
End Function

