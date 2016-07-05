VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INullableType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Summary:
'   Obtains a boolean indicating whether this INullableType has a value.
' Returns:
'   True, if INullableType.Value is not Null; else, false.
Public Property Get HasValue() As Boolean
End Property


' Summary:
'   Gets the value of this INullableType.
' Returns:
'   Null, if INullableType.HasValue is false, else, this INullableType's
'   non-Null value. The type of which is defined by the implementing class.
Public Property Get Value() As Variant
End Property


' Summary:
'   Sets the value of this INullableType.
' Params:
'   vNewValue:
'     Null, or a valid value for this INullableType. The range of valid values
'     is defined by the implementing class.
' Errors:
'   - 6. Overflow.
'   - 13. Type mismatch.
'TODO: Add detail to errors.
Public Property Let Value(ByVal vNewValue)
End Property


' Summary:
'   Clones this INullableType to another INullableType.
' Returns:
'   A new INullableType instance.
Public Function Clone() As INullableType
End Function


' Summary:
'   Gets the value of this INullableType or a default value.
' Returns:
'   TYPE?
'   If INullableType.HasValue is true, then INullableType.Value; else, a default
'   value as defined by the implementing class.
Public Function GetValueOrDefault() As Variant
End Function


' Summary:
'   Gets the value of this INullableType or a default value.
' Params:
'   Default:
'     The default value to be returned if INullableType.HasValue is false.
' Returns:
'   TYPE?
'   If INullableType.HasValue is true, then INullableType.Value; else, the value
'   of the Default parameter.
Public Function GetValueOrDefault2(ByVal Default As Variant) As Variant
End Function
