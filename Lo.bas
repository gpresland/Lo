Attribute VB_Name = "Lo"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Lo
'
' Utility belt for Microsoft Excel inspired by Lodash.
'
' @author Greg Presland
' @date   30 Nov 2018
'

Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clamps number within the inclusive lower and upper bounds.
'     Number : The number to clamp.
'     Lower  : The lower bound.
'     Upper  : The upper bound.
' Returns the first element of array.
'
Public Function Clamp(Number As Variant, Lower As Variant, Upper As Variant) As Variant
    If Number < Lower Then
        Clamp = Lower
    ElseIf Number > Upper Then
        Clamp = Upper
    Else
        Clamp = Number
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the first element of array.
'     Arr : The array to query.
' Returns the first element of array.
'
Public Function Head(Arr As Variant) As Variant
    Head = Arr(LBound(Arr))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is in array.
'     Arr   : The array to search.
'     Value : The value to search for.
' Returns true if value is found, else false.
'
Public Function Includes(Arr As Variant, Value As Variant) As Boolean
    Includes = UBound(Filter(Arr, Value)) > -1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the index at which the first occurrence of value is found in array.
'     Arr   : The  array to search.
'     Value : The value to search for.
' Returns the index of the matched value, else -1.
'
Public Function IndexOf(Arr As Variant, Value As Variant) As Integer
    Dim i As Long
    For i = 1 To UBound(Arr, 1)
        If Arr(i) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a boolean primitive or object.
'     Value : The value to check.
' Returns true if value is a boolean, else false.
'
Public Function IsBoolean(Value As Variant) As Boolean
    IsBoolean = TypeName(Value) = "Boolean"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is null or undefined.
'     Value : The value to check.
' Returns true if value is nullish, else false.
'
Public Function IsNil(Value As Variant) As Boolean
    IsNil = IsEmpty(Value) Or IsNull(Value) Or Value Is Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a Number primitive or object.
'     Value : The value to check.
' Returns true if value is a number, else false.
'
Public Function IsNumber(Value As Variant) As Boolean
    Dim name As String: name = TypeName(Value)
    IsNumber = name = "Byte" Or _
               name = "Decimal" Or _
               name = "Double" Or _
               name = "Integer" Or _
               name = "Long" Or _
               name = "SByte" Or _
               name = "Short" Or _
               name = "Single" Or _
               name = "UInteger" Or _
               name = "Ulong" Or _
               name = "UShort"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is the language type of Object.
'     Value : The value to check.
' Returns true if value is an object, else false.
'
Public Function IsObject(Value As Variant) As Boolean
    IsObject = TypeName(Value) = "Object"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a String primitive or object.
'     Value : The value to check.
' Returns true if value is a string, else false.
'
Public Function IsString(Value As Variant) As Boolean
    IsString = TypeName(Value) = "String"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the last element of array.
'     Arr : The array to query.
' Returns the last element of array.
'
Public Function Last(Arr As Variant) As Variant
    Last = Arr(UBound(Arr))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets all but the first element of array.
'     Arr : The array to query.
' Returns the slice of array.
'
Public Function Tail(Arr As Variant) As Variant
    Dim i As Integer
    Dim iStartIndex As Integer: iStartIndex = LBound(Arr, 1)
    Dim iEndIndex As Integer: iEndIndex = UBound(Arr, 1)
    If iStartIndex >= iEndIndex Then
        Exit Function
    End If
    ReDim newArr(iStartIndex + 1 To iEndIndex) As Variant
    For i = iStartIndex + 1 To iEndIndex
        newArr(i) = Arr(i)
    Next i
    Tail = newArr
End Function
