Attribute VB_Name = "DataTypes"
Option Explicit

Public Enum CustomErrorEnum
  erInvalidNumberFormat = 513
  erNotImplementedException
  erInvalidAdoDataTypeException
  erNotProperlyInitializedException
End Enum

Public Enum FormatRuleEnum
  frNone
  frNullOnly
  frText
  frDate
  frBinary
  frBoolean
  frDbTime
  frDbDateTimeOffset
End Enum

Public Type DBTime2
  hour As Integer
  minute As Integer
  second As Integer
  padding As Integer
  fraction As Long
End Type

Public Type DBTIMESTAMPOFFSET
  year As Integer
  month As Integer
  day As Integer
  hour As Integer
  minute As Integer
  second As Integer
  fraction As Long
  timezone_hour As Integer
  timezone_minute As Integer
End Type

Public Type SAFEARRAYBOUNDS
  elements As Long  'actually a ulong
  LowerBound As Long
End Type

Public Type SAFEARRAY
  cDims As Integer      ' Unsigned
  fFeatures As Integer  ' actually FADF flags
  cbElements As Long    ' Unsigned
  cLocks As Long        ' Unsigned
  pvData As LongPtr     ' Pointer to data
  firstSafeArrayBound As SAFEARRAYBOUNDS ' followed by cDims - 1 additional SafeArrayBound structures
End Type

' We don't actually use these... but in case you were wondering.
Public Type VARIANT_TYPE
  VarType As Integer
  reserved1 As Integer
  reserved2 As Integer
  reserved3 As Integer
  data(0 To 7) As Byte  '16
  padding As LongLong   '24
End Type

Public Type VARIANT_DECIMAL
  VarType As Integer
  Scale As Byte
  Sign As Byte
  Hi32 As Long
  Lo64 As LongLong ' or "Data" for a non-decimal variant
End Type

Public Enum VarTypeEnum
  VT_EMPTY = &H0&
  VT_NULL = &H1&
  VT_I2 = &H2&
  VT_I4 = &H3&
  VT_R4 = &H4&
  VT_R8 = &H5&
  VT_CY = &H6&
  VT_DATE = &H7&
  VT_BSTR = &H8&
  VT_DISPATCH = &H9&
  VT_ERROR = &HA&
  VT_BOOL = &HB&
  VT_VARIANT = &HC&
  VT_UNKNOWN = &HD&
  VT_DECIMAL = &HE&
  VT_I1 = &H10&
  VT_UI1 = &H11&
  VT_UI2 = &H12&
  VT_UI4 = &H13&
  VT_I8 = &H14&
  VT_UI8 = &H15&
  VT_INT = &H16&
  VT_UINT = &H17&
  VT_VOID = &H18&
  VT_HRESULT = &H19&
  VT_PTR = &H1A&
  VT_SAFEARRAY = &H1B&
  VT_CARRAY = &H1C&
  VT_USERDEFINED = &H1D&
  VT_LPSTR = &H1E&
  VT_LPWSTR = &H1F&
  VT_RECORD = &H24&
  VT_INT_PTR = &H25&
  VT_UINT_PTR = &H26&
  VT_ARRAY = &H2000&
  VT_BYREF = &H4000&
  VT_VariantArray = VT_ARRAY Or VT_VARIANT
End Enum

Public Enum SafeArrayFlagsEnum
  FADF_AUTO = &H1
  FADF_STATIC = &H2
  FADF_EMBEDDED = &H4
  FADF_FIXEDSIZE = &H10
  FADF_RECORD = &H20
  FADF_HAVEIID = &H40
  FADF_HAVEVARTYPE = &H80
  FADF_BSTR = &H100
  FADF_UNKNOWN = &H200
  FADF_DISPATCH = &H400
  FADF_VARIANT = &H800
  FADF_RESERVED = &HF008
End Enum

