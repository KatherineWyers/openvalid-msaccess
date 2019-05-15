Attribute VB_Name = "Validator"
Option Compare Binary
Option Explicit

'************************************************************************'
'                _   _       _ _     _       _                           '
'               | | | |     | (_)   | |     | |                          '
'               | | | | __ _| |_  __| | __ _| |_ ___  _ __               '
'               | | | |/ _` | | |/ _` |/ _` | __/ _ \| '__|              '
'               \ \_/ / (_| | | | (_| | (_| | || (_) | |                 '
'                \___/ \__,_|_|_|\__,_|\__,_|\__\___/|_|                 '
'                                                                        '
'************************************************************************'
                                   

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'NOTIFICATION MSGS
'
'Library of notification messages displayed in MsgBox
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'General
Public Const FIELD_REQUIRED As String = " field is required"

'Numeric Ranges
Public Const NUMBER_REQUIRED As String = " must be a number"
Public Const INTEGER_REQUIRED As String = " must be an integer"
Public Const DOUBLE_REQUIRED As String = " must be a double"

'Strings
Public Const INVALID_STRING_LENGTH As String = " not the required length"
Public Const CAN_ONLY_BE_LETTERS As String = " can only contain letters"
Public Const CAN_ONLY_BE_LETTERS_AND_NUMBERS As String = " can only contain letters and/or numbers"
Public Const CAN_ONLY_BE_LETTERS_AND_NUMBERS_AND_SPECIAL_CHARACTERS As String = " can only contain letters, numbers, underscores and dashes"
Public Const MUST_BE_ALL_ALLOWABLE_TEXT As String = " cannot contain carriage-returns"
Public Const IS_NOT_A_VALID_SELECTION As String = " is not a valid selection"
Public Const WHOLE_NUMBER_REQUIRED As String = " can only be a whole number (integer)"
Public Const VALUE_OUTSIDE_RANGE As String = " is outside the valid range"
Public Const MUST_BE_A_VALID_DATE As String = " must be a valid date"
Public Const MUST_BE_A_BOOLEAN As String = " must be a boolean"

'
'Configuration
'
Private Const DOUBLE_LINE_SPACING As Boolean = True
                                                                                                                                
                                                                       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STYLES                                                                  '
'                                                                        '
'Style sheet for colors etc                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Colors
Public Const TEXTBOX_BACKCOLOR_DEFAULT As Long = 16777215
Public Const TEXTBOX_BACKCOLOR_WARNING As Long = 1030655 'Darker Orange: 3381759
Public Const TEXTBOX_BACKCOLOR_ERROR As Long = 12695295 'Darker Red: 3289800
                                                                       
                                                                       
                                                                       
                                                                       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AccessQueryMgr                                                          '
'                                                                        '
'Handle usages of MS Access queries that are created using the           '
'MS Access user interface                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AccessQueryMgr_getArrayOfStringsFromAccessQuery
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'AccessQueryMgr_getArrayOfStringsFromAccessQuery
'
'Return an array of strings from the database query
'
'@param string strQueryName
'@param string strColumnName OPTIONAL. Default to first column in query
'@return variant array of strings
'
Public Function AccessQueryMgr_getArrayOfStringsFromAccessQuery(ByVal strQueryName As String, Optional ByVal strColumnName As Variant = 0) As Variant
  Dim rstData As DAO.Recordset
  Set rstData = CurrentDb.OpenRecordset(strQueryName, dbOpenDynaset, dbSeeChanges)
  Dim arrOfStrings() As Variant
  
  Do While Not rstData.EOF
    'Only set strings
    If VarType(rstData.Fields(strColumnName)) = 8 Then
      arrOfStrings = Validator.ArrayHelper_addStringTo1DArray(rstData.Fields(strColumnName), arrOfStrings)
    End If
    
    rstData.MoveNext
  Loop

  AccessQueryMgr_getArrayOfStringsFromAccessQuery = arrOfStrings
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AccessQueryMgr_getArrayOfIntegersFromAccessQuery
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'AccessQueryMgr_getArrayOfIntegersFromAccessQuery
'
'Return an array of integers from the database query
'
'@param string strQueryName
'@param string strColumnName OPTIONAL. Default to first column in query
'@return variant array of strings
'
Public Function AccessQueryMgr_getArrayOfIntegersFromAccessQuery(ByVal strQueryName As String, Optional ByVal strColumnName As Variant = 0) As Variant
  Dim rstData As DAO.Recordset
  Set rstData = CurrentDb.OpenRecordset(strQueryName, dbOpenDynaset, dbSeeChanges)
  Dim arrOfIntegers() As Variant
  
  Do While Not rstData.EOF
    'Only set bytes, integers and long integers
    If VarType(rstData.Fields(strColumnName)) = 2 Or VarType(rstData.Fields(strColumnName)) = 3 Or VarType(rstData.Fields(strColumnName)) = 17 Then
      arrOfIntegers = Validator.ArrayHelper_addIntegerTo1DArray(rstData.Fields(strColumnName), arrOfIntegers)
    End If
    
    rstData.MoveNext
  Loop

  AccessQueryMgr_getArrayOfIntegersFromAccessQuery = arrOfIntegers
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ArrayHelper                                                             '
'                                                                        '
'Library of functions to manage the                                      '
'arrays.                                                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'TABLE OF CONTENTS
'
'getLastElement
'setLastElementAsString
'getSize
'isEmptyArray
'clipLastElement
'increase1DArraySizeByOne
'addStringTo1DArray
'isInArray
'isArrayOfStrings
'numberOfArrayDimensions
'arrayIsOneDimensional
'allElementsAreVarType
'concatAllElementsIn1DArray
'allElementsAreVarTypeInArray
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'ArrayHelper_getLastElement
'
'@param variant varArray
'@return variant varLastElement
'

Public Function ArrayHelper_getLastElement(ByVal varArray As Variant) As Variant
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ArrayHelper_getLastElement = Null
    Exit Function
  End If
  
  ArrayHelper_getLastElement = varArray(UBound(varArray))
End Function

'
'ArrayHelper_setLastElementAsString
'
'@param string strVal
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_setLastElementAsString(ByVal strVal As String, ByVal varArray As Variant) As Variant
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ReDim varArray(0)
  End If
  varArray(UBound(varArray)) = strVal
  
  ArrayHelper_setLastElementAsString = varArray
End Function

'
'ArrayHelper_setLastElementAsInteger
'
'@param integer intVal
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_setLastElementAsInteger(ByVal varVal As Variant, ByVal varArray As Variant) As Variant

  If Not (VarType(varVal) = 2) And Not (VarType(varVal) = 3) And Not (VarType(varVal) = 17) Then
    Call Validator.ErrorHandler_logAndRaiseError("Validator.ArrayHelper_addIntegerTo1DArray", "Input value was not a byte, integer or long integer")
  End If
  
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ReDim varArray(0)
  End If
  varArray(UBound(varArray)) = varVal
  
  ArrayHelper_setLastElementAsInteger = varArray
End Function

'
'ArrayHelper_getSize
'
'@param variant varArray
'@return integer intSize
'
Public Function ArrayHelper_getSize(ByVal varArray As Variant) As Long
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ArrayHelper_getSize = 0
    Exit Function
  End If
  
  ArrayHelper_getSize = (UBound(varArray) - LBound(varArray)) + 1
End Function

'
'ArrayHelper_isEmptyArray
'
'@param variant varArray
'@return bool
'
Public Function ArrayHelper_isEmptyArray(ByVal varArray As Variant) As Boolean
  On Error GoTo IS_EMPTY
  If (UBound(varArray) >= 0) Then Exit Function
IS_EMPTY:
  ArrayHelper_isEmptyArray = True
End Function

'
'ArrayHelper_clipLastElement
'
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_clipLastElement(ByVal varArray As Variant) As Variant
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ArrayHelper_clipLastElement = varArray
    Exit Function
  End If
  
  'Only clip the last element if there are more than one elements
  'in the array
  If UBound(varArray) = 0 Then
    Dim varEmptyArray() As Variant
    ArrayHelper_clipLastElement = varEmptyArray
    Exit Function
  Else
    ReDim Preserve varArray(0 To UBound(varArray) - 1)
  End If

  ArrayHelper_clipLastElement = varArray
End Function

'
'ArrayHelper_increase1DArraySizeByOne
'
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_increase1DArraySizeByOne(ByVal varArray As Variant) As Variant
  If Validator.ArrayHelper_isEmptyArray(varArray) Then
    ReDim varArray(0)
  Else
    ReDim Preserve varArray(0 To UBound(varArray) + 1)
  End If
  ArrayHelper_increase1DArraySizeByOne = varArray
End Function

'
'ArrayHelper_addStringTo1DArray
'
'@param string strVal
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_addStringTo1DArray(ByVal strVal As String, ByVal varArray As Variant) As Variant

  varArray = Validator.ArrayHelper_increase1DArraySizeByOne(varArray)
  varArray = Validator.ArrayHelper_setLastElementAsString(strVal, varArray)
  
  ArrayHelper_addStringTo1DArray = varArray
End Function

'
'ArrayHelper_addIntegerTo1DArray
'
'@param integer intVal
'@param variant varArray
'@return variant varArray
'
Public Function ArrayHelper_addIntegerTo1DArray(ByVal varVal As Variant, ByVal varArray As Variant) As Variant

  If Not (VarType(varVal) = 2) And Not (VarType(varVal) = 3) And Not (VarType(varVal) = 17) Then
    Call Validator.ErrorHandler_logAndRaiseError("Validator.ArrayHelper_addIntegerTo1DArray", "Input value was not a byte, integer or long integer")
  End If

  varArray = Validator.ArrayHelper_increase1DArraySizeByOne(varArray)
  varArray = Validator.ArrayHelper_setLastElementAsInteger(varVal, varArray)
  
  ArrayHelper_addIntegerTo1DArray = varArray
End Function

'
'ArrayHelper_isInArray
'
'@param variant varToBeFound
'@param variant arr
'@return bool
'
Public Function ArrayHelper_isInArray(ByVal varToBeFound As Variant, ByVal arr As Variant) As Boolean

  If Validator.ArrayHelper_isEmptyArray(arr) Then
    ArrayHelper_isInArray = False
    Exit Function
  End If
  
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(arr) - 1)
  
    If IsNull(varToBeFound) And IsNull(arr(i)) Then
      ArrayHelper_isInArray = True
      Exit For
    End If
    
    If varToBeFound = arr(i) Then
      ArrayHelper_isInArray = True
      Exit For
    End If
  Next i
  
End Function

'
'ArrayHelper_isArrayOfStrings
'
'@param variant arr
'@return bool
'
Public Function ArrayHelper_isArrayOfStrings(ByVal arr As Variant) As Boolean
  ArrayHelper_isArrayOfStrings = True
  
  'False if array is empty
  If Validator.ArrayHelper_getSize(arr) = 0 Then
    ArrayHelper_isArrayOfStrings = False
    Exit Function
  End If
  
  'False if array is more than 1 dimension
  If Not Validator.ArrayHelper_arrayIsOneDimensional(arr) Then
    ArrayHelper_isArrayOfStrings = False
    Exit Function
  End If
  
  'Iterate through the array of control details and check that the necessary rules have been defined correctly
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(arr) - 1)
    If Not VarType(arr(i)) = vbString Then
      ArrayHelper_isArrayOfStrings = False
    End If
  Next i
  
End Function

'
'ArrayHelper_numberOfArrayDimensions
'
'@param variant arr
'@return integer intNumOfArrayDimensions
'
Public Function ArrayHelper_numberOfArrayDimensions(arr As Variant) As Long

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' NumberOfArrayDimensions
  ' This function returns the number of dimensions of an array. An unallocated dynamic array
  ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim Ndx As Long
  Dim Res As Long
  On Error Resume Next
  ' Loop, increasing the dimension index Ndx, until an error occurs.
  ' An error will occur when Ndx exceeds the number of dimension
  ' in the array. Return Ndx - 1.
  Do
      Ndx = Ndx + 1
      Res = UBound(arr, Ndx)
  Loop Until Err.Number <> 0
  
  Err.Clear
  
  ArrayHelper_numberOfArrayDimensions = Ndx - 1

End Function

'
'ArrayHelper_arrayIsOneDimensional
'
'@param variant arr
'@return bool
'
Public Function ArrayHelper_arrayIsOneDimensional(arr As Variant) As Boolean
    ArrayHelper_arrayIsOneDimensional = (Validator.ArrayHelper_numberOfArrayDimensions(arr) = 1)
End Function

'
'ArrayHelper_allElementsAreVarType
'
'@param string strVarType
'@param variant varArray
'@return bool
'
Public Function ArrayHelper_allElementsAreVarType(ByVal strVarType As String, ByVal varArray As Variant) As Boolean
  ArrayHelper_allElementsAreVarType = True
  
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varArray) - 1)
    If Not VarType(varArray(i)) = Validator.Datatypes_convertStrVarTypeToIntVarType(strVarType) Then
      ArrayHelper_allElementsAreVarType = False
    End If
  Next i
End Function

'
'ArrayHelper_concatAllElementsIn1DArray
'
'Return empty string if not 1D array
'
'@param array varArray
'@return string
'
Public Function ArrayHelper_concatAllElementsIn1DArray(ByVal varArray As Variant) As String
  ArrayHelper_concatAllElementsIn1DArray = ""
  If Not (Validator.ArrayHelper_arrayIsOneDimensional(varArray)) Then
    Exit Function
  End If
  
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varArray) - 1)
    ArrayHelper_concatAllElementsIn1DArray = concatAllElementsIn1DArray & "[" & i & "]: " & varArray(i) & "; "
  Next i
End Function

'
'ArrayHelper_allElementsAreVarTypeInArray
'
'@param variant varVarTypesAllowed
'@param variant varArray
'@return bool
'
Public Function ArrayHelper_allElementsAreVarTypeInArray(ByVal varVarTypesAllowed As Variant, ByVal varArray As Variant) As Boolean
  ArrayHelper_allElementsAreVarTypeInArray = True
  
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varArray) - 1)
    If Not Validator.ArrayHelper_isInArray(Validator.Datatypes_convertIntVarTypeToStrVarType(VarType(varArray(i))), varVarTypesAllowed) Then
      ArrayHelper_allElementsAreVarTypeInArray = False
    End If
  Next i
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BooleanValidationHelper                                                 '
'                                                                        '
'Manage Boolean input validation fields                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'BooleanValidationHelper_allBooleanInputsAreValidBooleans
'
'Check that the boolean inputs are valid booleans
'data types
'
'@param variant varValidationRulesArray
'@param object frmForm
'@param string strNotifications
'@return Boolean
'
Public Function BooleanValidationHelper_allBooleanInputsAreValidBooleans(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "Boolean"
  Dim strFunctionName As String: strFunctionName = "Validator.BooleanValidationHelper_allBooleanInputsAreValidBooleans"
  
  BooleanValidationHelper_allBooleanInputsAreValidBooleans = True

  'Iterate through the array until a Date is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
  
    'Validate input
    Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)


    'Validate boolean
    If ((varValidationRules(i, 2) = strDatatype)) Then

      'NonNullable values should never reach here as Null
      Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(varValidationRules(i, 2), frmForm.Controls(varValidationRules(i, 0)).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)

      If Not (frmForm.Controls(varValidationRules(i, 0)).Value = -1) And Not (frmForm.Controls(varValidationRules(i, 0)).Value = 0) Then
        Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
        strNotifications = strNotifications & varValidationRules(i, 0) & Validator.MUST_BE_A_BOOLEAN & vbNewLine
        BooleanValidationHelper_allBooleanInputsAreValidBooleans = False
      End If
      
    End If
    
  Next i
End Function



                                                             
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BoundaryMgr                                                             '
'                                                                        '
'Library of functions to manage the                                      '
'boundary values. These are used when a                                  '
'user enters a numeric value that must be                                '
'checked against the boundary conditions                                 '
'in the data dictionary.                                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'TABLE OF CONTENTS
'
'Get Location
' - getLocationInRangeDouble
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetLocationInRangeDouble
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'BoundaryMgr_getLocationInRangeDouble
'Take a Double as argument and determine its LocationInRange from the boundary-array
'
'@param string strTitle Title of the amount being tested
'@param Double dblValue Value of the amount being tested
'@param Variant(4) varBoundariesArr Array of boundary-values as a Variant
'@param String strWarnings String of Yes/No notifications
'@param String strOutOfBounds String with the OutOfBounds notifications
'@return integer 1 to 5 or -1 if something is wrong
'1 = Red
'2 = Orange
'3 = Green
'4 = Orange
'5 = Red
'
Public Function BoundaryMgr_getLocationInRangeDouble(ByVal strTitle As String, ByVal dblValue As Double, ByVal varBoundariesArr As Variant, ByRef strWarnings As String, ByRef strOutOfBounds As String) As Long


  Dim strFunctionName As String: strFunctionName = "Validator.BoundaryMgr_getLocationInRangeDouble"

  'Array must be one-dimensional
  If Not (Validator.ArrayHelper_arrayIsOneDimensional(varBoundariesArr) = True) Then
    Call Validator.ErrorHandler_logErrorWithoutRaisingError(strFunctionName, "varBoundariesArr is not one dimensional")
    BoundaryMgr_getLocationInRangeDouble = -1
    Exit Function
  End If

  'Array must have exactly four elements
  If Not (Validator.ArrayHelper_getSize(varBoundariesArr) = 4) Then
    Call Validator.ErrorHandler_logErrorWithoutRaisingError(strFunctionName, "varBoundariesArr is not Size 4. Size is " & Validator.ArrayHelper_getSize(varBoundariesArr))
    BoundaryMgr_getLocationInRangeDouble = -1
    Exit Function
  End If
  
  'All boundary values must be recognized numeric datatypes
  If Not (Validator.ArrayHelper_allElementsAreVarTypeInArray(Validator.Datatypes_getRecognizedNumericDataTypes, varBoundariesArr)) Then
    Call Validator.ErrorHandler_logErrorWithoutRaisingError(strFunctionName, "varBoundariesArr is not only Double or Integer")
    BoundaryMgr_getLocationInRangeDouble = -1
    Exit Function
  End If
  
  Dim strBoundaries As String

  If dblValue < varBoundariesArr(0) Then
    strBoundaries = strBoundaries & strTitle & " Ranges" & vbNewLine & "[[RED]] " & varBoundariesArr(0) & " orange " & varBoundariesArr(1) & " good " & varBoundariesArr(2) & " orange " & varBoundariesArr(3) & " red"
    strOutOfBounds = strOutOfBounds & strBoundaries & vbNewLine
    strOutOfBounds = strOutOfBounds & strTitle & " is in the RED ZONE: " & dblValue & " is too low" & vbNewLine
    strOutOfBounds = strOutOfBounds & "----------------" & vbNewLine
    BoundaryMgr_getLocationInRangeDouble = 1
    Exit Function
  End If
  
  If dblValue < varBoundariesArr(1) Then
  
    strBoundaries = strTitle & " Ranges" & vbNewLine & "red " & varBoundariesArr(0) & " [[ORANGE]] " & varBoundariesArr(1) & " good " & varBoundariesArr(2) & " orange " & varBoundariesArr(3) & " red"
    
    strWarnings = strWarnings & strBoundaries & vbNewLine
    strWarnings = strWarnings & strTitle & ": " & dblValue & " is in the ORANGE ZONE between " & varBoundariesArr(0) & " and " & varBoundariesArr(1) & vbNewLine
    strWarnings = strWarnings & "----------------" & vbNewLine
    BoundaryMgr_getLocationInRangeDouble = 2
    Exit Function
  End If
  
  If dblValue < varBoundariesArr(2) Then
    BoundaryMgr_getLocationInRangeDouble = 3
    'Do nothing. Value in the GREEN range
    Exit Function
  End If
  
  If dblValue < varBoundariesArr(3) Then
  
    strBoundaries = strTitle & " Ranges" & vbNewLine & "red " & varBoundariesArr(0) & " orange " & varBoundariesArr(1) & " good " & varBoundariesArr(2) & " [[ORANGE]] " & varBoundariesArr(3) & " red"
    
    strWarnings = strWarnings & strBoundaries & vbNewLine
    strWarnings = strWarnings & strTitle & ": " & dblValue & " is in the ORANGE ZONE between " & varBoundariesArr(2) & " and " & varBoundariesArr(3) & vbNewLine
    strWarnings = strWarnings & "----------------" & vbNewLine
    BoundaryMgr_getLocationInRangeDouble = 4
    Exit Function
  End If

  If dblValue >= varBoundariesArr(3) Then
  
    strBoundaries = strTitle & " Ranges" & vbNewLine & "red " & varBoundariesArr(0) & " orange " & varBoundariesArr(1) & " good " & varBoundariesArr(2) & " orange " & varBoundariesArr(3) & " [[RED]]"
    
    strOutOfBounds = strOutOfBounds & strBoundaries & vbNewLine
    strOutOfBounds = strOutOfBounds & strTitle & " is in the RED ZONE: " & dblValue & " is too high" & vbNewLine
    strOutOfBounds = strOutOfBounds & "----------------" & vbNewLine
    BoundaryMgr_getLocationInRangeDouble = 5
    Exit Function
  End If
  
  
  'Raise Error if the int2DLocationInRange is not in range 1 to 5
  If Not Validator.ArrayHelper_isInArray(BoundaryMgr_getLocationInRangeDouble, Array(1, 2, 3, 4, 5)) Then
    Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "Data Dictionary location in range returned invalid from BoundaryMgr_getLocationInRangeDouble.")
  End If
End Function





                                                                       

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                                               '
'DATATYPES                                                                                                      '
'                                                                                                               '                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'Datatypes_getRecognizedNumericDataTypes
'
'Get an array of the datatypes that are
'allowable integer datatypes or allowable
'decimal datatypes
'
'@return string-array or Empty array
'
Public Function Datatypes_getRecognizedNumericDataTypes() As String()
  Dim recognizedNumericDataTypes(5) As String
  recognizedNumericDataTypes(0) = "Byte"
  recognizedNumericDataTypes(1) = "Integer"
  recognizedNumericDataTypes(2) = "Long Integer"
  recognizedNumericDataTypes(3) = "Single"
  recognizedNumericDataTypes(4) = "Decimal"
  recognizedNumericDataTypes(5) = "Double"
  Datatypes_getRecognizedNumericDataTypes = recognizedNumericDataTypes
End Function

'Datatypes_getRecognizedIntegerDataTypes
'
'Get an array of the datatypes that are
'allowable integer datatypes
'
'@return string-array or Empty array
'
Public Function Datatypes_getRecognizedIntegerDataTypes() As String()
  Dim recognizedIntegerDataTypes(2) As String
  recognizedIntegerDataTypes(0) = "Byte"
  recognizedIntegerDataTypes(1) = "Integer"
  recognizedIntegerDataTypes(2) = "Long Integer"
  Datatypes_getRecognizedIntegerDataTypes = recognizedIntegerDataTypes
End Function

'Datatypes_getRecognizedDecimalDataTypes
'
'Get an array of the datatypes that are
'allowable decimal datatypes
'
'@return string-array or Empty array
'
Public Function Datatypes_getRecognizedDecimalDataTypes() As String()
  Dim recognizedDecimalDataTypes(2) As String
  recognizedDecimalDataTypes(0) = "Single"
  recognizedDecimalDataTypes(1) = "Decimal"
  recognizedDecimalDataTypes(2) = "Double"
  Datatypes_getRecognizedDecimalDataTypes = recognizedDecimalDataTypes
End Function

'
'Datatypes_getNullableDataTypes
'
'Get an array of the datatype titles that can be
'set to null and are therefore not required.
'
'@return string-array or Empty array
'
Public Function Datatypes_getNullableDataTypes() As String()
  Dim nullableDataTypes(12) As String
  nullableDataTypes(0) = "AlphaText"
  nullableDataTypes(1) = "AlphanumericText"
  nullableDataTypes(2) = "SpecialCharacterText"
  nullableDataTypes(3) = "AllAllowableText"
  nullableDataTypes(4) = "Date"
  nullableDataTypes(5) = "StringInArray"
  nullableDataTypes(6) = "IntegerInArray"
  nullableDataTypes(7) = "IntegerInRange"
  nullableDataTypes(8) = "DecimalInRange"
  nullableDataTypes(9) = "OneDimInteger"
  nullableDataTypes(10) = "OneDimDecimal"
  nullableDataTypes(11) = "TwoDimInteger"
  nullableDataTypes(12) = "TwoDimDecimal"
  Datatypes_getNullableDataTypes = nullableDataTypes
End Function

'
'Datatypes_getSanitizableDataTypes
'
'Get an array of the text datatype titles that must
'be sanitized
'
'@return string-array or Empty array
'
Public Function Datatypes_getSanitizableDataTypes() As String()
  Dim sanitizableDataTypes(3) As String
  sanitizableDataTypes(0) = "AlphaText"
  sanitizableDataTypes(1) = "AlphanumericText"
  sanitizableDataTypes(2) = "SpecialCharacterText"
  sanitizableDataTypes(3) = "AllAllowableText"
  Datatypes_getSanitizableDataTypes = sanitizableDataTypes
End Function

'
'Datatypes_getRecognizedDataTypes
'
'Get an array of the recognized datatype titles
'
'@return string-array or Empty array
'
Public Function Datatypes_getRecognizedDataTypes() As String()
    Dim validDataTypes(13) As String
    validDataTypes(0) = "StringInArray"
    validDataTypes(1) = "AlphaText"
    validDataTypes(2) = "AlphanumericText"
    validDataTypes(3) = "SpecialCharacterText"
    validDataTypes(4) = "AllAllowableText"
    validDataTypes(5) = "Date"
    validDataTypes(6) = "IntegerInArray"
    validDataTypes(7) = "IntegerInRange"
    validDataTypes(8) = "DecimalInRange"
    validDataTypes(9) = "OneDimInteger"
    validDataTypes(10) = "OneDimDecimal"
    validDataTypes(11) = "TwoDimInteger"
    validDataTypes(12) = "TwoDimDecimal"
    validDataTypes(13) = "Boolean"
  Datatypes_getRecognizedDataTypes = validDataTypes
End Function

'
'Datatypes_getDataTypesThatUseParameters
'
'Get an array of the datatypes that use parameters
'in (x, 4)
'
'@return string-array or Empty array
'
Public Function Datatypes_getDatatypesThatUseParameters() As String()
    Dim arrDatatypesThatUseParameters(7) As String
    arrDatatypesThatUseParameters(0) = "StringInArray"
    arrDatatypesThatUseParameters(1) = "AlphaText"
    arrDatatypesThatUseParameters(2) = "AlphanumericText"
    arrDatatypesThatUseParameters(3) = "SpecialCharacterText"
    arrDatatypesThatUseParameters(4) = "AllAllowableText"
    arrDatatypesThatUseParameters(5) = "IntegerInArray"
    arrDatatypesThatUseParameters(6) = "IntegerInRange"
    arrDatatypesThatUseParameters(7) = "DecimalInRange"
  Datatypes_getDatatypesThatUseParameters = arrDatatypesThatUseParameters
End Function

'
'Datatypes_getValidDatatypesForSecondDimVariable
'
'Get an array of the datatypes that can be used
'as second-dim variables
'
'@return string-array or Empty array
'
Public Function Datatypes_getValidDatatypesForSecondDimVariable() As String()
  Dim validDataTypesForSecondDimVariable(4) As String
  validDataTypesForSecondDimVariable(0) = "OneDimInteger"
  validDataTypesForSecondDimVariable(1) = "OneDimDecimal"
  validDataTypesForSecondDimVariable(2) = "TwoDimInteger"
  validDataTypesForSecondDimVariable(2) = "TwoDimDecimal"
  validDataTypesForSecondDimVariable(3) = "IntegerInRange"
  validDataTypesForSecondDimVariable(4) = "DecimalInRange"
  Datatypes_getValidDatatypesForSecondDimVariable = validDataTypesForSecondDimVariable
End Function


'
'Datatypes_getDatatypesThatHaveASecondDimVariable
'
'Get an array of the datatypes that have a
'second dim variable assigned in (x, 5)
'
'@return string-array or Empty array
'
Public Function Datatypes_getDatatypesThatHaveASecondDimVariable() As String()
  Dim arrDatatypesThatHaveASecondDimVariable(1) As String
  arrDatatypesThatHaveASecondDimVariable(0) = "TwoDimInteger"
  arrDatatypesThatHaveASecondDimVariable(1) = "TwoDimDecimal"
  Datatypes_getDatatypesThatHaveASecondDimVariable = arrDatatypesThatHaveASecondDimVariable
End Function

'
'Datatypes_getDatatypesThatAreArrayable
'
'Get an array of the datatypes that can be passed as an
'array using the multi-select combobox
'
'@return string-array or Empty array
'
Public Function Datatypes_getDatatypesThatAreArrayable() As String()
  Dim arrDatatypesThatAreArrayable(1) As String
  arrDatatypesThatAreArrayable(0) = "IntegerInArray"
  arrDatatypesThatAreArrayable(1) = "StringInArray"
  Datatypes_getDatatypesThatAreArrayable = arrDatatypesThatAreArrayable
End Function


'
'Datatypes_ifNonNullableFieldIsNullRaiseError(ByVal strDataType As String, ByVal varVal As Variant, ByVal varNullableDataTypes As Variant)
'
'@param string strDataType
'@param variant varVal
'@param variant varNullableDataTypes
'@param string strSource
'
Public Sub Datatypes_ifNonNullableFieldIsNullRaiseError(ByVal strDatatype As String, ByVal varVal As Variant, ByVal varNullableDataTypes As Variant, ByVal strSource As String)
  If Not (Validator.ArrayHelper_isInArray(strDatatype, varNullableDataTypes)) And IsNull(varVal) Then
    Call Validator.ErrorHandler_logAndRaiseError("ifNonNullableFieldIsNullRaiseError", "ControlName cannot be Null. Should be filtered before entering this function: " & strDatatype)
  End If
End Sub


'
'Datatypes_convertStrVarTypeToIntVarType
'
'If not found, return -1
'
'@param string strVarType
'@return integer intVarType or -1
'
Public Function Datatypes_convertStrVarTypeToIntVarType(ByVal strVarType As String) As Long
  Datatypes_convertStrVarTypeToIntVarType = -1
  
  If strVarType = "Empty" Then
    Datatypes_convertStrVarTypeToIntVarType = 0
  End If
  
  If strVarType = "Null" Then
    Datatypes_convertStrVarTypeToIntVarType = 1
  End If
  
  If strVarType = "Integer" Then
    Datatypes_convertStrVarTypeToIntVarType = 2
  End If
  
  If strVarType = "Long Integer" Then
    Datatypes_convertStrVarTypeToIntVarType = 3
  End If
  
  If strVarType = "Single" Then
    Datatypes_convertStrVarTypeToIntVarType = 4
  End If
  
  If strVarType = "Double" Then
    Datatypes_convertStrVarTypeToIntVarType = 5
  End If
  
  If strVarType = "Date" Then
    Datatypes_convertStrVarTypeToIntVarType = 7
  End If
  
  If strVarType = "String" Then
    Datatypes_convertStrVarTypeToIntVarType = 8
  End If
  
  If strVarType = "Boolean" Then
    Datatypes_convertStrVarTypeToIntVarType = 11
  End If
  
  If strVarType = "Variant" Then
    Datatypes_convertStrVarTypeToIntVarType = 12
  End If
  
  If strVarType = "Decimal" Then
    Datatypes_convertStrVarTypeToIntVarType = 14
  End If
  
  If strVarType = "Byte" Then
    Datatypes_convertStrVarTypeToIntVarType = 17
  End If
  
End Function


'
'Datatypes_convertIntVarTypeToStrVarType
'
'If not found, return ""
'
'@param int intVarType
'@return string strVarType or empty string
'
Public Function Datatypes_convertIntVarTypeToStrVarType(ByVal intVarType As Long) As String
  Datatypes_convertIntVarTypeToStrVarType = ""
  
  If intVarType = 0 Then
    Datatypes_convertIntVarTypeToStrVarType = "Empty"
  End If
  
  If intVarType = 1 Then
    Datatypes_convertIntVarTypeToStrVarType = "Null"
  End If
  
  If intVarType = 2 Then
    Datatypes_convertIntVarTypeToStrVarType = "Integer"
  End If
  
  If intVarType = 3 Then
    Datatypes_convertIntVarTypeToStrVarType = "Long Integer"
  End If
  
  If intVarType = 4 Then
    Datatypes_convertIntVarTypeToStrVarType = "Single"
  End If
  
  If intVarType = 5 Then
    Datatypes_convertIntVarTypeToStrVarType = "Double"
  End If
  
  If intVarType = 7 Then
    Datatypes_convertIntVarTypeToStrVarType = "Date"
  End If
  
  If intVarType = 8 Then
    Datatypes_convertIntVarTypeToStrVarType = "String"
  End If
  
  If intVarType = 11 Then
    Datatypes_convertIntVarTypeToStrVarType = "Boolean"
  End If
  
  If intVarType = 12 Then
    Datatypes_convertIntVarTypeToStrVarType = "Variant"
  End If
  
  If intVarType = 14 Then
    Datatypes_convertIntVarTypeToStrVarType = "Decimal"
  End If
  
  If intVarType = 17 Then
    Datatypes_convertIntVarTypeToStrVarType = "Byte"
  End If
  
End Function


                                                                     
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DATA VALIDATOR                                                          '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Textbox validation
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'DataValidator_isAlpha
'
'string contains only the following: a-z A-Z -' and whitespace
'
'Whitelisting
'
'@param variant userInput
'@return boolean
'
Public Function DataValidator_isAlpha(ByVal strValue As Variant) As Boolean

  'If the value is a zero-length string, return True
  If strValue = "" Then
    DataValidator_isAlpha = True
    Exit Function
  End If

  'If the value is null, return false
  If IsNull(strValue) Then
    DataValidator_isAlpha = False
    Exit Function
  End If
  
  'If the first or last characters are whitespace, return false
  If (Left(strValue, 1) = 32) Or (Right(strValue, 1) = 32) Then
    DataValidator_isAlpha = False
    Exit Function
  End If
  
  'If all characters are whitespace, return false
  Dim intPos As Long
  Dim boolAllCharsAreWhitespace As Boolean: boolAllCharsAreWhitespace = True

  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32
        'Do nothing
      Case Else
        'Found a character that is not white space
        boolAllCharsAreWhitespace = False
      Exit For
    End Select
  Next
  
  If boolAllCharsAreWhitespace Then
    DataValidator_isAlpha = False
    Exit Function
  End If
  
  'Check whether the characters a
  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32, 39, 45, 65 To 90, 97 To 122
        DataValidator_isAlpha = True
      Case Else
        DataValidator_isAlpha = False
      Exit For
    End Select
  Next
End Function

'
'DataValidator_isAlphaNumeric
'
'string contains only the following: a-z A-Z 0-9 -_/' and whitespace
'
'Whitelisting
'
'@param variant userInput
'@return boolean
'
Public Function DataValidator_isAlphaNumeric(ByVal strValue As Variant) As Boolean

  'If the value is a zero-length string, return True
  If strValue = "" Then
    DataValidator_isAlphaNumeric = True
    Exit Function
  End If

  'If the value is null, return false
  If IsNull(strValue) Then
    DataValidator_isAlphaNumeric = False
    Exit Function
  End If
  
  'If the first or last characters are whitespace, return false
  If (Left(strValue, 1) = 32) Or (Right(strValue, 1) = 32) Then
    DataValidator_isAlphaNumeric = False
    Exit Function
  End If
  
  'If all characters are whitespace, return false
  Dim intPos As Long
  Dim boolAllCharsAreWhitespace As Boolean: boolAllCharsAreWhitespace = True

  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32
        'Do nothing
      Case Else
        'Found a character that is not white space
        boolAllCharsAreWhitespace = False
      Exit For
    End Select
  Next
  
  If boolAllCharsAreWhitespace Then
    DataValidator_isAlphaNumeric = False
    Exit Function
  End If
  
  'Check whether the characters are all a-z A-Z 0-9 -_/
  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32, 39, 45, 47 To 57, 65 To 90, 95, 97 To 122
        DataValidator_isAlphaNumeric = True
      Case Else
        DataValidator_isAlphaNumeric = False
      Exit For
    End Select
  Next
End Function


'
'DataValidator_isSpecialCharacterText
'
'string contains only the following: a-z A-Z 0-9 -_/#%<>.;()+,'" and whitespace
'
'Whitelisting
'
'@param variant userInput
'@return boolean
'
Public Function DataValidator_isSpecialCharacterText(ByVal strValue As Variant) As Boolean

  'If the value is a zero-length string, return True
  If strValue = "" Then
    DataValidator_isSpecialCharacterText = True
    Exit Function
  End If

  'If the value is null, return false
  If IsNull(strValue) Then
    DataValidator_isSpecialCharacterText = False
    Exit Function
  End If
  
  'If the first or last characters are whitespace, return false
  If (Left(strValue, 1) = 32) Or (Right(strValue, 1) = 32) Then
    DataValidator_isSpecialCharacterText = False
    Exit Function
  End If
  
  'If all characters are whitespace, return false
  Dim intPos As Long
  Dim boolAllCharsAreWhitespace As Boolean: boolAllCharsAreWhitespace = True

  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32
        'Do nothing
      Case Else
        'Found a character that is not white space
        boolAllCharsAreWhitespace = False
      Exit For
    End Select
  Next
  
  If boolAllCharsAreWhitespace Then
    DataValidator_isSpecialCharacterText = False
    Exit Function
  End If
  
  'Check whether the characters are all a-z A-Z 0-9 -_/#%<>.;()+,
  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32, 34, 35, 37, 39, 40, 41, 43 To 58, 59, 60, 62, 65 To 90, 95, 97 To 122
        DataValidator_isSpecialCharacterText = True
      Case Else
        DataValidator_isSpecialCharacterText = False
      Exit For
    End Select
  Next
End Function



'
'DataValidator_isAllAllowableText
'
'string contains all text except carriage-return, Null and leading/trailing whitespace
'
'Blacklisting
'
'@param variant userInput
'@return boolean
'
Public Function DataValidator_isAllAllowableText(ByVal strValue As Variant) As Boolean

  'If the value is a zero-length string, return True
  If strValue = "" Then
    DataValidator_isAllAllowableText = True
    Exit Function
  End If

  'If the value is null, return false
  If IsNull(strValue) Then
    DataValidator_isAllAllowableText = False
    Exit Function
  End If
  
  'If the first or last characters are whitespace, return false
  If (Left(strValue, 1) = 32) Or (Right(strValue, 1) = 32) Then
    DataValidator_isAllAllowableText = False
    Exit Function
  End If
  
  'If all characters are whitespace, return false
  Dim intPos As Long
  Dim boolAllCharsAreWhitespace As Boolean: boolAllCharsAreWhitespace = True

  For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
      Case 32
        'Do nothing
      Case Else
        'Found a character that is not white space
        boolAllCharsAreWhitespace = False
      Exit For
    End Select
  Next
  
  If boolAllCharsAreWhitespace Then
    DataValidator_isAllAllowableText = False
    Exit Function
  End If
  
  'Check whether the characters are all allowable
  DataValidator_isAllAllowableText = True
  For intPos = 1 To Len(strValue)
    If Asc(Mid(strValue, intPos, 1)) = 10 Or Asc(Mid(strValue, intPos, 1)) = 13 Then
      'Found a character that is not allowed
      DataValidator_isAllAllowableText = False
    End If
  Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'String validation
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'DataValidator_stringLengthInRange
'Check whether the length of the string in within the given length range
'
'@param string str
'@param integer minLen
'@param integer maxLen
'@return Boolean
'
Public Function DataValidator_stringLengthIsInRange(ByVal str As String, ByVal minLen As Long, ByVal maxLen As Long) As Boolean
  If (Len(str) >= minLen) And (Len(str) <= maxLen) Then
    DataValidator_stringLengthIsInRange = True
    Exit Function
  End If
  DataValidator_stringLengthIsInRange = False
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Number validation
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'DataValidator_isIntegerInRange
'Check whether the length of the integer in within the given range
'
'Max value for all inputs is 2147483647
'Above these, the system will raise a Overflow error
'
'
'@param integer intVal
'@param integer minVal
'@param integer maxVal
'@return Boolean
'
Public Function DataValidator_isIntegerInRange(ByVal intVal As Long, ByVal minVal As Long, ByVal maxVal As Long) As Boolean

  DataValidator_isIntegerInRange = True
  
  If ((intVal < minVal) Or (intVal > maxVal)) Then
    DataValidator_isIntegerInRange = False
  End If
  
End Function

'
'DataValidator_isDecimalInRange
'Check whether the length of the double in within the given range
'
'Note: If the min and max values are input as the same values, you
'cannot rely on this function to EVER return true. Unlike INTEGERS,
'DOUBLE values can never be accurately equate as equal to other DOUBLE
'values. They can only ever be compared as less than or greater than.
'
'@param double dblVal
'@param integer minVal
'@param integer maxVal
'@return Boolean
'
Public Function DataValidator_isDecimalInRange(ByVal dblVal As Double, ByVal minVal As Double, ByVal maxVal As Double) As Boolean
  DataValidator_isDecimalInRange = True

  If ((dblVal < minVal) Or (dblVal > maxVal)) Then
    DataValidator_isDecimalInRange = False
  End If
  
End Function

Public Function DataValidator_isInteger(ByVal varValue As Variant) As Boolean
  DataValidator_isInteger = False
  
  If Validator.ArrayHelper_isInArray(Validator.Datatypes_convertIntVarTypeToStrVarType(VarType(varValue)), Validator.Datatypes_getRecognizedIntegerDataTypes) Then
    DataValidator_isInteger = True
  End If

End Function

'
'DataValidator_integerIsInIntegerArray
'Check whether the input integer is in the array
'
'VBA rounds down any double inputs to become integers
'
'@param integer intVal
'@param variant varIntegerArray
'@return Boolean
'
Public Function DataValidator_integerIsInIntegerArray(ByVal intVal As Long, ByRef varIntegerArray As Variant) As Boolean
  DataValidator_integerIsInIntegerArray = False
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varIntegerArray) - 1)
    If (varIntegerArray(i) = intVal) Then
      DataValidator_integerIsInIntegerArray = True
    End If
  Next i
End Function

Public Function DataValidator_isDouble(ByVal varValue As Variant) As Boolean
  DataValidator_isDouble = False
  
  If Validator.ArrayHelper_isInArray(Validator.Datatypes_convertIntVarTypeToStrVarType(VarType(varValue)), Validator.Datatypes_getRecognizedDecimalDataTypes) Then
    DataValidator_isDouble = True
  End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DATE VALIDATION HELPER                                                  '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'DateValidationHelper_allDateInputsAreValidDates
'
'Check that the date inputs are valid dates
'data types
'
'@param variant varValidationRulesArray
'@param object frmForm
'@param string strNotifications
'@return Boolean
'
Public Function DateValidationHelper_allDateInputsAreValidDates(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "Date"
  Dim strFunctionName As String: strFunctionName = "Validator.DateValidationHelper_allDateInputsAreValidDates"
  
  DateValidationHelper_allDateInputsAreValidDates = True

  'Iterate through the array until a Date is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
  
    'Validate input
    Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

    'Validate Date values
    If ((varValidationRules(i, 2) = strDatatype)) Then

      'NonNullable values should never reach here as Null
      Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(varValidationRules(i, 2), frmForm.Controls(varValidationRules(i, 0)).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)

      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Datatype is not null
        If Not IsDate(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.MUST_BE_A_VALID_DATE & vbNewLine
          DateValidationHelper_allDateInputsAreValidDates = False
        End If
      End If
      
    End If
    
  Next i
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ERROR HANDLER                                                           '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'ErrorHandler_logAndRaiseError
'
'Call the log and raise functions
'
'@param string strSource
'@param string strDescription
'
Public Sub ErrorHandler_logAndRaiseError(ByVal strSource As String, ByVal strDescription As String)
  Call Validator.Logger_logError(strSource, strDescription)
  Call Validator.ErrorHandler_raise(strSource, strDescription)
End Sub

'
'ErrorHandler_logWithoutRaisingError
'
'Call the log functions
'
'@param string strSource
'@param string strDescription
'
Public Sub ErrorHandler_logErrorWithoutRaisingError(ByVal strSource As String, ByVal strDescription As String)
  Call Validator.Logger_logError(strSource, strDescription)
End Sub


'
'ErrorHandler_raise(ByVal intErrorCode As Long, ByVal strSource As String, ByVal strDescription As String)
'
'Raise an error in the application
'
'@param integer intErrorCode
'@param string strSource
'@param string strDescription
'
Private Sub ErrorHandler_raise(ByVal strSource As String, ByVal strDescription As String)
  Call Err.raise(1, strSource, strDescription)
End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FORM CONTROL HELPER                                                     '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'FormControlHelper_getBoundableControlTypes
'
'Get an array of the control types that can be BOUND
'be sanitized
'
'@return string-array or Empty array
'
Public Function FormControlHelper_getBoundableControlTypes() As String()
  Dim arrBoundableControlTypes(4) As String
  arrBoundableControlTypes(0) = "TextBox"
  arrBoundableControlTypes(1) = "ComboBox"
  arrBoundableControlTypes(2) = "CheckBox"
  arrBoundableControlTypes(3) = "OptionGroup"
  arrBoundableControlTypes(4) = "ListBox"
  FormControlHelper_getBoundableControlTypes = arrBoundableControlTypes
End Function

Public Sub FormControlHelper_raiseErrorIfControlNameIsNullOrEmptyOrBlank(ByVal strSource As Variant, ByVal strControlName As Variant)

  'Raise Error if strSource is null or blank
  If IsNull(strSource) Or strSource = "" Or IsEmpty(strSource) Then
    strSource = "Validator.FormControlHelper_raiseErrorIfControlNameIsNullOrEmptyOrBlank"
  End If
  
  'Raise Error if strControlName is null or blank
  If IsNull(strControlName) Or strControlName = "" Or IsEmpty(strControlName) Then
    Call Validator.ErrorHandler_logAndRaiseError("raiseErrorIfControlNameIsNullOrEmptyOrBlank", "strControlName cannot be Null or Blank or Empty")
  End If
End Sub

Public Sub FormControlHelper_raiseErrorIfControlDoesNotExist(ByVal strSource As String, ByVal strControlName As Variant, ByRef frmForm As Object)

  'Raise Error is control name does not exist
  If Not Validator.FormControlHelper_controlExists(strControlName, frmForm) Then
    Call Validator.ErrorHandler_logAndRaiseError(strSource, "ControlName does not exist: " & strControlName)
  End If
End Sub

Public Function FormControlHelper_controlExists(ByVal strControlName As String, ByRef frmForm As Object) As Boolean
  Dim strTest As String
  On Error Resume Next
    strTest = frmForm(strControlName).Name
    FormControlHelper_controlExists = (Err.Number = 0)
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Getters and Setters
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'FormControlHelper_getControlValue
'
'Return the contents of the input field
'
'@param String strControlName
'@param Form frmForm
'@return variant
'
Public Function FormControlHelper_getControlValue(ByVal strControlName As Variant, ByRef frmForm As Object) As Variant

  Dim strFunctionName As String: strFunctionName = "Validator.FormControlHelper_getControlValue"
  
  'Validate input
  Call Validator.FormControlHelper_raiseErrorIfControlNameIsNullOrEmptyOrBlank(strFunctionName, strControlName)
  Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, strControlName, frmForm)

  FormControlHelper_getControlValue = frmForm.Controls(strControlName).Value
End Function

'
'FormControlHelper_setControlValue
'
'Set the contents of the input field
'
'@param String strControlName
'@param Variant varValue
'@param Form frmForm
'
Public Sub FormControlHelper_setControlValue(ByVal strControlName As Variant, ByVal varValue As Variant, ByRef frmForm As Object)

  Dim strFunctionName As String: strFunctionName = "Validator.FormControlHelper_setControlValue"
  
  'Validate inputs
  Call Validator.FormControlHelper_raiseErrorIfControlNameIsNullOrEmptyOrBlank(strFunctionName, strControlName)
  Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, strControlName, frmForm)
  
  'Set control value
  frmForm.Controls(strControlName).Value = varValue
End Sub

'FormControlHelper_isEditable
'
'Return the contents of the input field
'
'@param String strControlName
'@param Form frmForm
'@return variant
'
Public Function FormControlHelper_isEditable(ByVal strControlName As Variant, ByRef frmForm As Object) As Boolean

  FormControlHelper_isEditable = False

  Dim arrLockableControlTypes(4) As String
  arrLockableControlTypes(0) = "TextBox"
  arrLockableControlTypes(1) = "ComboBox"
  arrLockableControlTypes(2) = "CheckBox"
  arrLockableControlTypes(3) = "OptionGroup"
  arrLockableControlTypes(4) = "ListBox"

  On Error GoTo errHandler

  Dim strFormInputType As String: strFormInputType = ""
  Dim ctl As Control
  
  For Each ctl In frmForm
    If ctl.Name = strControlName Then
      'set Form Input Type
      If ctl.ControlType = acCheckBox Then strFormInputType = "CheckBox"
      If ctl.ControlType = acComboBox Then strFormInputType = "ComboBox"
      If ctl.ControlType = acCommandButton Then strFormInputType = "CommandButton"
      If ctl.ControlType = acLabel Then strFormInputType = "Label"
      If ctl.ControlType = acListBox Then strFormInputType = "ListBox"
      If ctl.ControlType = acOptionButton Then strFormInputType = "OptionButton"
      If ctl.ControlType = acOptionGroup Then strFormInputType = "OptionGroup"
      If ctl.ControlType = acSubform Then strFormInputType = "Subform"
      If ctl.ControlType = acTextBox Then strFormInputType = "TextBox"
      If ctl.ControlType = acToggleButton Then strFormInputType = "ToggleButton"
      
      If Validator.ArrayHelper_isInArray(strFormInputType, arrLockableControlTypes) Then
        'set Editable status
        If ((ctl.Locked = False) And (ctl.Enabled = True)) Then
          FormControlHelper_isEditable = True
        Else
          FormControlHelper_isEditable = False
        End If
      End If

      Exit For
    End If
  Next ctl
  
  Exit Function
  
errHandler:
  Debug.Print "Error Raised in Validator.FormControlHelper_isEditable"
End Function

'FormControlHelper_isBoundField
'
'Return the contents of the input field
'
'@param String strControlName
'@param Form frmForm
'@return boolean
'
Public Function FormControlHelper_isBoundField(ByVal strControlName As Variant, ByRef frmForm As Object) As Boolean
  FormControlHelper_isBoundField = False

  On Error GoTo errHandler
  
  Dim ctl As Control
  
  For Each ctl In frmForm
    If ctl.Name = strControlName Then
      If Validator.ArrayHelper_isInArray(Validator.FormControlHelper_getFormInputType(ctl), Validator.FormControlHelper_getBoundableControlTypes) Then
        'set isBound status
        If ((ctl.ControlSource = Null) Or (ctl.ControlSource = "")) Then
          FormControlHelper_isBoundField = False
        Else
          FormControlHelper_isBoundField = True
        End If
      End If

      Exit For
    End If
  Next ctl

  Exit Function
  
errHandler:
  Debug.Print "Error Raised in Validator.FormControlHelper_isBoundField"
End Function


'
'FormControlHelper_getFormInputType
'
'@param Control ctl
'@return strFormInputType
'
Public Function FormControlHelper_getFormInputType(ByVal ctl As Control) As String
  FormControlHelper_getFormInputType = ""
  If ctl.ControlType = acCheckBox Then FormControlHelper_getFormInputType = "CheckBox"
  If ctl.ControlType = acComboBox Then FormControlHelper_getFormInputType = "ComboBox"
  If ctl.ControlType = acCommandButton Then FormControlHelper_getFormInputType = "CommandButton"
  If ctl.ControlType = acLabel Then FormControlHelper_getFormInputType = "Label"
  If ctl.ControlType = acListBox Then FormControlHelper_getFormInputType = "ListBox"
  If ctl.ControlType = acOptionButton Then FormControlHelper_getFormInputType = "OptionButton"
  If ctl.ControlType = acOptionGroup Then FormControlHelper_getFormInputType = "OptionGroup"
  If ctl.ControlType = acSubform Then FormControlHelper_getFormInputType = "Subform"
  If ctl.ControlType = acTextBox Then FormControlHelper_getFormInputType = "TextBox"
  If ctl.ControlType = acToggleButton Then FormControlHelper_getFormInputType = "ToggleButton"
  
  'Raise an error if the input type was not recognised
  If FormControlHelper_getFormInputType = "" Then
    Call Validator.ErrorHandler_logAndRaiseError("Validator.FormControlHelper_getFormInputType", "The Control with Name " & ctl.Name & " is a " & ctl.ControlType & " control type. This is not a control type that cannot be used with the Validator. The system only accepts CheckBox, ComboBox, CommandButton, Label, ListBox, OptionButton, OptionGroup, Subform, TextBox and ToggleButton.)")
  End If
  
End Function


'FormControlHelper_isEditedField
'
'Return whether the field has been edited
'
'@param String strControlName
'@param Form frmForm
'@return boolean
'
Public Function FormControlHelper_isEditedField(ByVal strControlName As Variant, ByRef frmForm As Object) As Boolean

  FormControlHelper_isEditedField = False
  
  On Error GoTo errHandler

  Dim ctl As Control
  
  For Each ctl In frmForm
    If ctl.Name = strControlName Then
      'Found the requested ctl
      If Validator.ArrayHelper_isInArray(Validator.FormControlHelper_getFormInputType(ctl), Validator.FormControlHelper_getBoundableControlTypes) Then
        'Requested ctl is a value that has the OldValue property
        If Validator.FormControlHelper_isSameOrBothNull(ctl.OldValue, ctl.Value) Then
          FormControlHelper_isEditedField = False
        Else
          FormControlHelper_isEditedField = True
        End If
        
      End If
      
      Exit For
    End If
  Next ctl

  Exit Function
  
errHandler:
  Debug.Print "Error Raised in Validator.FormControlHelper_isEditedField"
End Function


'
'FormControlHelper_isSameOrBothNull
'
Public Function isSameOrBothNull(ByVal varInputOne As Variant, ByVal varInputTwo As Variant) As Boolean
  FormControlHelper_isSameOrBothNull = False
  
  If (varInputOne = varInputTwo) Or (IsNull(varInputOne) And IsNull(varInputTwo)) Then
    FormControlHelper_isSameOrBothNull = True
  End If
  
End Function






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'LOGGER FUNCTIONS                                                        '
'                                                                        '
'This contains all the functions that interact with the log files        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'Logger_logError
'
'Call the log with an error
'
'@param string strSource
'@param string strDescription
'
Public Sub Logger_logError(ByVal strSource As String, ByVal strDescription As String)
  'Strip out the line breaks from the strDescription. This ensures that description is written to a single line in the log file
  strDescription = Replace(strDescription, Chr(10), "")
  strDescription = Replace(strDescription, Chr(13), "")
  
  '1 is the code for LogType = Error
  Call Validator.Logger_log(1, strSource, strDescription)
End Sub

'
'Logger_logValidationNotice
'
'Call the log with a validation notice
'
'@param string strSource
'@param string strDescription
'
Public Sub Logger_logValidationNotice(ByVal strSource As String, ByVal strDescription As String)
  'Strip out the line breaks from the strDescription. This ensures that description is written to a single line in the log file
  strDescription = Replace(strDescription, Chr(10), "")
  strDescription = Replace(strDescription, Chr(13), "")
  
  '2 is the code for LogType = ValidationNotice
  Call Validator.Logger_log(2, strSource, strDescription)
End Sub

'
'Logger_logDataInputWasValid
'
'Call the log to log that the data input was valid
'
'@param string strSource
'@param string strDescription
'
Public Sub Logger_logDataInputWasValid(ByVal strSource As String, ByVal strDescription As String)
  'Strip out the line breaks from the strDescription. This ensures that description is written to a single line in the log file
  strDescription = Replace(strDescription, Chr(10), "")
  strDescription = Replace(strDescription, Chr(13), "")
  
  '3 is the code for LogType = SuccessfulSave
  Call Validator.Logger_log(3, strSource, strDescription)
End Sub

'
'Logger_logProcessCancelledByValidator
'
'Log that the process was cancelled by the validator,
'by either the user-cancelled warning, or the out-of-bounds
'notification
'
'@param string strSource
'@param string strDescription
'
Public Sub Logger_logProcessCancelledByValidator(ByVal strSource As String, ByVal strDescription As String)
  'Strip out the line breaks from the strDescription. This ensures that description is written to a single line in the log file
  strDescription = Replace(strDescription, Chr(10), "")
  strDescription = Replace(strDescription, Chr(13), "")
  
  '4 is the code for LogType = ValidationNotice
  Call Validator.Logger_log(4, strSource, strDescription)
End Sub


'
'Logger_log(ByVal intLogType As Long, ByVal strSource As String, ByVal strDescription As String)
'
'Log the message in the log file
'
'@param integer intErrorCode
'@param string strSource
'@param string strDescription
'
Private Sub Logger_log(ByVal intLogType As Long, ByVal strSource As String, ByVal strDescription As String)
  Dim fs, f, strDateTime As String
  strDateTime = Format(Now, "dddd, mmm d yyyy") & " - " & Format(Now, "hh:mm:ss AMPM")
  
  'Default directory path to current project folder if it is Null in the GlobalSettings
  Dim strDirectoryPath As String
  If (Validator_Settings.CUSTOM_DIRECTORY_PATH = "") Then
    strDirectoryPath = Application.CurrentProject.path
  Else
    strDirectoryPath = Validator_Settings.CUSTOM_DIRECTORY_PATH
  End If
  
  If Not (Validator_Settings.CUSTOM_SUBDIRECTORY_NAME = "") Then
    strDirectoryPath = strDirectoryPath & "\" & Validator_Settings.CUSTOM_SUBDIRECTORY_NAME
  End If
  
  'Set filename to error or validation-notice depending on intErrorCode
  Dim strFilename As String
  Dim strErrorMsg As String
  Select Case intLogType
    Case 1:
      strFilename = Validator_Settings.ERROR_LOG_FILENAME
      strErrorMsg = strDateTime & vbTab & "Error Source: " & strSource & vbTab & " Description: " & strDescription
    Case 2:
      strFilename = Validator_Settings.VALIDATION_NOTICE_LOG_FILENAME
      strErrorMsg = strDateTime & vbTab & "Validation Source: " & strSource & vbTab & " Description: " & strDescription
    Case 3:
      strFilename = Validator_Settings.SUCCESSFUL_SAVE_LOG_FILENAME
      strErrorMsg = strDateTime & vbTab & "Save: " & strSource & vbTab & " Description: " & strDescription
    Case 4:
      strFilename = Validator_Settings.PROCESS_CANCELLED_BY_VALIDATOR_LOG_FILENAME
      strErrorMsg = strDateTime & vbTab & "Save: " & strSource & vbTab & " Description: " & strDescription
    Case Else
      strFilename = Validator_Settings.ERROR_LOG_FILENAME
      strErrorMsg = strDateTime & vbTab & "Error Source: " & strSource & vbTab & " Description: " & strDescription
  End Select
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  
  On Error GoTo ERR_HANDLER
  
    If Validator.Logger_pathToFolderExists(strDirectoryPath) Then
      'Open the log file
      Set f = fs.OpenTextFile(strDirectoryPath & "\" & strFilename, 8, 1) 'Second arg 8 = append. Third arg 1 = make if needed
      
      'Write to the file
      f.Write strErrorMsg & vbCrLf
  
      'Close the file
      f.Close
    Else
      MsgBox ("LOG FOLDER NOT FOUND" & vbNewLine & vbNewLine & "Please create a folder titled " & Validator_Settings.CUSTOM_SUBDIRECTORY_NAME & " in the following directory: " & Application.CurrentProject.path & vbNewLine & vbNewLine & " You can customize the log folder location in the VALIDATOR_GlobalSettings module")
    End If
    
  Err.Clear
    
  Exit Sub
  
ERR_HANDLER:
  Select Case Err.Number
    Case 78 'Overflow
      MsgBox "Failed attempt to write to error log file. Please contact your system administrator. " & vbNewLine & _
      "Error Number: " & Err.Number & vbNewLine & _
      "Error Description: " & Err.Description
    Case 76 'File not found. Create the file
      Set f = fs.OpenTextFile(strDirectoryPath & "\" & strFilename, 8, 1)
    Case Else
      MsgBox "Failed attempt to write to error log file. Please contact your system administrator. " & vbNewLine & _
      "Error Number: " & Err.Number & vbNewLine & _
      "Error Description: " & Err.Description
  End Select
End Sub


Private Function Logger_pathToFolderExists(ByVal strPathToFolder As String) As Boolean
  Logger_pathToFolderExists = False
  
  Dim sFolderPath As String
  If Right(strPathToFolder, 1) <> "\" Then
      strPathToFolder = strPathToFolder & "\"
  End If
  
  If Dir(strPathToFolder, vbDirectory) <> vbNullString Then
    Logger_pathToFolderExists = True
  End If

End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'MSG BOX HELPER                                                          '
'                                                                        '
'This contains all the functions that interact with the dialog box       '
'(message box, popup box)                                                '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function MsgBoxHelper_askYesNoQuestion(ByVal strStatusBarText As String, ByVal strQuestion As String, ByVal strMsgWhenAnswerIsNo) As Boolean
    MsgBoxHelper_askYesNoQuestionaskYesNoQuestion = True
    
    Dim strYesOrNoAnswerToMessageBox As String
    strYesOrNoAnswerToMessageBox = MsgBox(strQuestion, vbYesNo, strStatusBarText)
    If strYesOrNoAnswerToMessageBox = vbNo Then
        MsgBox strMsgWhenAnswerIsNo
        MsgBoxHelper_askYesNoQuestionaskYesNoQuestion = False
    End If
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'NUMBER VALIDATION HELPER                                                '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'
'NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers
'
'Check that the input is in the array of integers
'
'@param variant varValidationRules
'@param object frmForm
'@param string strNotifications
'@return boolean
'
Public Function NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "IntegerInArray"
  Dim strFunctionName As String: strFunctionName = "Validator.NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers"
  
  NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers = True

  'Iterate through the array until a "IntegerInArray" is found
  Dim i As Long
  Dim varElem As Variant
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strDatatype) Then

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

      'NonNullable values should never reach here as Null
      Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(varValidationRules(i, 0)).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)

      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
      
      
        'If value is multi-select ComboBox, it passes an Array instead of a String
        'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
        If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 Then
          For Each varElem In frmForm.Controls(varValidationRules(i, 0)).Value

            'Validate Integers as whole numbers
            If Not Validator.DataValidator_isInteger(varElem) Then
              Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
              strNotifications = strNotifications & varValidationRules(i, 0) & Validator.WHOLE_NUMBER_REQUIRED & vbNewLine
              NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers = False
              Exit Function
            End If
      
            If (Validator.DataValidator_integerIsInIntegerArray(varElem, varValidationRules(i, 4)) = False) Then
              Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
              strNotifications = Validator.StringFormatter_addLineToBody(varValidationRules(i, 0) & Validator.IS_NOT_A_VALID_SELECTION, strNotifications)
              NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers = False
            End If

          Next varElem
          
        Else
          'Validate Integers as whole numbers
          If Not Validator.DataValidator_isInteger(frmForm.Controls(varValidationRules(i, 0)).Value) Then
            Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
            strNotifications = strNotifications & varValidationRules(i, 0) & Validator.WHOLE_NUMBER_REQUIRED & vbNewLine
            NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers = False
            Exit Function
          End If
    
          If (Validator.DataValidator_integerIsInIntegerArray(frmForm.Controls(varValidationRules(i, 0)).Value, varValidationRules(i, 4)) = False) Then
            Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
            strNotifications = Validator.StringFormatter_addLineToBody(varValidationRules(i, 0) & Validator.IS_NOT_A_VALID_SELECTION, strNotifications)
            NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers = False
          End If
        End If
        
      End If

    End If
  Next i

End Function

'
'NumberValidationHelper_allNumericInputsAreValidNumbers
'
'Check that the numeric inputs are valid numbers
'data types
'
'@param variant varValidationRulesArray
'@param object frmForm
'@param string strNotifications
'@return Boolean
'
Public Function NumberValidationHelper_allNumericInputsAreValidNumbers(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strFunctionName As String: strFunctionName = "Validator.NumberValidationHelper_allNumericInputsAreValidNumbers"
  
  NumberValidationHelper_allNumericInputsAreValidNumbers = True

  'Iterate through the array until a "Numeric" is found
  Dim i As Long
  Dim varElem As Variant
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
  
    'Validate input
    Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)
  
    'NonNullable values should never reach here as Null
    Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(varValidationRules(i, 2), frmForm.Controls(varValidationRules(i, 0)).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
    
    'If value is multi-select ComboBox, it passes an Array instead of a String
    'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
    If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 And Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getDatatypesThatAreArrayable()) Then
      Debug.Print varValidationRules(i, 0)
      Call Validator.ErrorHandler_logAndRaiseError("NumberValidationHelper.allNumericInputsAreValidNumbers", varValidationRules(i, 0) & " is an array. It must be a number. This error probably occurred because you are using a Multi-select ComboBox on a numeric value. Multi-select ComboBox can only use either IntegerInArray or StringInArray.") '
    End If


    'Validate OneDimInteger
    If (varValidationRules(i, 2) = "OneDimInteger") Then
      
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("OneDimInteger", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Validate
        If Not Validator.DataValidator_isInteger(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.INTEGER_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
    End If


    'Validate OneDimDecimal
    If (varValidationRules(i, 2) = "OneDimDecimal") Then
      
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("OneDimDecimal", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Validate
        If Not Validator.DataValidator_isDouble(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.DOUBLE_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
    End If
    

    
    'Validate TwoDimInteger
    If (varValidationRules(i, 2) = "TwoDimInteger") Then
      
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("TwoDimInteger", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Validate
        If Not Validator.DataValidator_isInteger(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.INTEGER_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
    End If
    
    
    'Validate TwoDimDecimal
    If (varValidationRules(i, 2) = "TwoDimDecimal") Then
      
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("TwoDimDecimal", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Validate
        If Not Validator.DataValidator_isDouble(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.DOUBLE_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
    End If
    
    If (varValidationRules(i, 2) = "IntegerInRange") Then
    
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("IntegerInRange", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Validate Integers as whole numbers
        If Not Validator.DataValidator_isInteger(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.WHOLE_NUMBER_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
        
        If Not Validator.DataValidator_isIntegerInRange(frmForm.Controls(varValidationRules(i, 0)).Value, varValidationRules(i, 4)(0), varValidationRules(i, 4)(1)) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.VALUE_OUTSIDE_RANGE & "Min: " & _
          varValidationRules(i, 4)(0) & "  Max: " & varValidationRules(i, 4)(1) & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
      
    End If
    
    If (varValidationRules(i, 2) = "DecimalInRange") Then
    
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray("DecimalInRange", Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
      
        'Validate Double as numeric
        If Not Validator.DataValidator_isDouble(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.NUMBER_REQUIRED & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
        
        If Not Validator.DataValidator_isDecimalInRange(frmForm.Controls(varValidationRules(i, 0)).Value, varValidationRules(i, 4)(0), varValidationRules(i, 4)(1)) Then
          Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
          strNotifications = strNotifications & varValidationRules(i, 0) & Validator.VALUE_OUTSIDE_RANGE & "Min: " & _
          varValidationRules(i, 4)(0) & "  Max: " & varValidationRules(i, 4)(1) & vbNewLine
          NumberValidationHelper_allNumericInputsAreValidNumbers = False
        End If
      End If
      
    End If
    
  Next i
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'ONE DIM VALIDATION MGR                                                  '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'OneDimValidationMgr_getDataDictionaryGrpFrom1DimensionalVariable
'1Dimensional Variable
'Set the Group to be used when the variable is a 1-dimensional variable
'
Public Function OneDimValidationMgr_getDataDictionaryGrpFrom1DimensionalVariable(ByVal strTitle As String, _
ByVal var1DimValidationRules As Variant) As Double()

  Dim strDataDictionaryTitle As String
  Dim intGrp As Long
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(var1DimValidationRules) - 1)
    If var1DimValidationRules(i, 0) = strTitle Then
      strDataDictionaryTitle = var1DimValidationRules(i, 1)
      intGrp = var1DimValidationRules(i, 2)
    End If
  Next i
  
  '
  'Check the strTitle and intGrp before passing them into the Validator_Boundaries.getGrp function
  '
  'Raise Error if the intGrpInt is Null or 0 or blank
  If IsNull(intGrp) Or (intGrp = 0) Then
    Call Validator.ErrorHandler_logAndRaiseError("getDataDictionaryGrpFrom1DimensionalVariable", "Data Dictionary group number request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
  End If
  
  'Raise Error if the mFirstDimensionName is Null or blank
  If IsNull(strTitle) Or (strTitle = "") Then
    Call Validator.ErrorHandler_logAndRaiseError("getDataDictionaryGrpFrom1DimensionalVariable", "Data Dictionary dimension ame request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
  End If
  
  OneDimValidationMgr_getDataDictionaryGrpFrom1DimensionalVariable = Validator_Boundaries.getGrp(strDataDictionaryTitle, intGrp)
End Function

'
'OneDimValidationMgr_setLocationForAll1DimensionalVariables
'
Public Sub OneDimValidationMgr_setLocationForAll1DimensionalVariables(ByVal varValidationRules As Variant, _
ByVal var1DimValidationRules As Variant, ByRef frmForm As Object, ByRef strWarnings As String, _
ByRef strOutOfBounds As String)

  Dim strFunctionName As String: strFunctionName = "Validator.OneDimValidationMgr_setLocationForAll1DimensionalVariables"

  '(x, 3) states the number of dimensions that the variable has
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = "OneDimInteger") Or (varValidationRules(i, 2) = "OneDimDecimal") Then

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Datatype is not null. Validate
        'Variable is a 1-dimensional variable
        Call Validator.UserInterfaceMgr_setControlBackColorFromLocationInRange(varValidationRules(i, 0), _
        Validator.OneDimValidationMgr_get1DLocationInRange(varValidationRules(i, 0), _
        frmForm.Controls(varValidationRules(i, 0)).Value, _
        var1DimValidationRules, _
        strWarnings, _
        strOutOfBounds), _
        frmForm)
      End If
      
    End If
  Next i
End Sub



'
'OneDimValidationMgr_get1DLocationInRange
'
'Get the location of a one-dimensional variable as an integer
'in the range of boundaries
'
'Boundary Codes
'1 = Red: Value below lowest allowable value
'2 = Orange: Value is in the low warning range
'3 = Green: Value is good
'4 = Orange: Value is in the high warning range
'5 = Red: Value above the highest allowable value
'
'Add notification to strWarnings if the value is in the ORANGE range
'Add notification to strOutOfBound if the value is in the RED range
'
'@param string title
'@param double value
'@param variant var1DimValidationRules
'@param string strWarnings
'@param string strOutOfBounds
'@return integer in range {1, 2, 3, 4, 5}
'
Public Function OneDimValidationMgr_get1DLocationInRange(ByVal mTitle As String, ByVal mValue As Double, _
ByVal var1DimValidationRules As Variant, ByRef strWarnings As String, ByRef strOutOfBounds As String) As Long

  Dim strFunctionName As String: strFunctionName = "Validator.OneDimValidationMgr_get1DLocationInRange"

  Dim arrGrp() As Double: arrGrp = Validator.OneDimValidationMgr_getDataDictionaryGrpFrom1DimensionalVariable(mTitle, var1DimValidationRules)
  
  Dim int1DLocationInRange As Long: int1DLocationInRange = Validator.BoundaryMgr_getLocationInRangeDouble(mTitle, mValue, arrGrp, strWarnings, strOutOfBounds)

  'BoundaryMgr_getLocationInRangeDouble returns a -1 when there was an invalid request
  If (int1DLocationInRange = -1) Then
    Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "Validator.BoundaryMgr_getLocationInRangeDouble returned an invalid value " & OneDimValidationMgr_get1DLocationInRange & " for mTitle: " & mTitle & " and mValue: " & mValue & " Array: " & Validator.ArrayHelper_concatAllElementsIn1DArray(arrGrp))
  End If
  
  OneDimValidationMgr_get1DLocationInRange = int1DLocationInRange
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'REPORTER                                                                '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'Reporter_run
'
Public Sub Reporter_run(ByVal frmForm As Form)
  'Only run the reporter if you are in testing mode
  'The Reporter functionality was moved from "optional" to "always print current form structure grid when running in testing mode"
  '
  '*Backwards compatibility*
  'Before this change, the implementors had the option to call Validator.Reporter_run function in their Form_frmForm code.
  'This check is made to ensure that old validator rules do not throw an error if they call Validator.Reporter_run.
  If Validator_Settings.ENVIRONMENT = "testing" Then
    Call Reporter_clearImmediateWindow
    Call Reporter_printCurrentFormStructureGrid(frmForm)
  End If
End Sub


'
'Reporter_clearImmediateWindow
'
'Clear the immediate window by printing 100 blank lines
'
Private Sub Reporter_clearImmediateWindow()
  Dim i As Long
  For i = 1 To 100
    Debug.Print "" 'Print a blank line
  Next i
End Sub

'
'Reporter_printFormLayoutGrid
'
'Print the details about the form controls to the immediate window
'
Private Sub Reporter_printCurrentFormStructureGrid(ByVal frmForm As Form)
  Debug.Print ""
  Debug.Print ""
  Debug.Print "#################################################################################################################################"
  Debug.Print "(a) FormControlName (b) FormInputType (c) FormValidation (d) Editable (e) DbColumnName"
  Debug.Print "#################################################################################################################################"
  Debug.Print "***********WARNING: THE FOLLOWING ITEMS ARE NOT IN THE ORDER THEY APPEAR ON THE SCREEN!****************"
  Debug.Print "#################################################################################################################################"
  Dim ctl As Control
  
  Dim arrPrintableControlTypes(7) As String
  arrPrintableControlTypes(0) = "CheckBox"
  arrPrintableControlTypes(1) = "ComboBox"
  arrPrintableControlTypes(2) = "CommandButton"
  arrPrintableControlTypes(3) = "ListBox"
  arrPrintableControlTypes(4) = "OptionGroup"
  arrPrintableControlTypes(5) = "Subform"
  arrPrintableControlTypes(6) = "TextBox"
  arrPrintableControlTypes(7) = "ToggleButton"
  
  Dim arrLockableControlTypes(4) As String
  arrLockableControlTypes(0) = "TextBox"
  arrLockableControlTypes(1) = "ComboBox"
  arrLockableControlTypes(2) = "CheckBox"
  arrLockableControlTypes(3) = "ListBox"
  arrLockableControlTypes(4) = "OptionGroup"
  
  Dim arrSourceableControlTypes(5) As String
  arrSourceableControlTypes(0) = "CheckBox"
  arrSourceableControlTypes(1) = "ComboBox"
  arrSourceableControlTypes(2) = "ListBox"
  arrSourceableControlTypes(3) = "OptionGroup"
  arrSourceableControlTypes(4) = "TextBox"
  arrSourceableControlTypes(5) = "ToggleButton"
  
  Dim arrMaskableControlTypes(0) As String
  arrMaskableControlTypes(0) = "TextBox"
  
  Dim arrRowSourceableControlTypes(1) As String
  arrRowSourceableControlTypes(0) = "ComboBox"
  arrRowSourceableControlTypes(1) = "ListBox"
  
  Dim strFormControlName As String
  Dim strFormInputType As String
  Dim strFormValidation As String
  Dim strEditable As String
  Dim strDbColumnName As String

  For Each ctl In frmForm

    
      strFormControlName = "N/A"
      strFormInputType = "N/A"
      strFormValidation = "N/A"
      strEditable = "N/A"
      strDbColumnName = "N/A"
    
      'set Form Control Name
      strFormControlName = ctl.Name
      
      'set Form Input Type
      'The Validator ignores subforms. To use the Validator on a
      'subform, create a seperate VBA sheet for that subform
      If ctl.ControlType = acCheckBox Then strFormInputType = "CheckBox"
      If ctl.ControlType = acComboBox Then strFormInputType = "ComboBox"
      If ctl.ControlType = acCommandButton Then strFormInputType = "CommandButton"
      If ctl.ControlType = acLabel Then strFormInputType = "Label"
      If ctl.ControlType = acListBox Then strFormInputType = "ListBox"
      If ctl.ControlType = acOptionButton Then strFormInputType = "OptionButton"
      If ctl.ControlType = acOptionGroup Then strFormInputType = "OptionGroup"
      If ctl.ControlType = acTextBox Then strFormInputType = "TextBox"
      If ctl.ControlType = acToggleButton Then strFormInputType = "ToggleButton"

      
      If Validator.Reporter_isInArray(strFormInputType, arrMaskableControlTypes) Then
        If Not (ctl.InputMask = "") Then
          'set InputMask
          strFormValidation = "Mask: " & ctl.InputMask
        End If
      End If
      
      If Validator.Reporter_isInArray(strFormInputType, arrRowSourceableControlTypes) Then
        If Not (ctl.RowSource = "") Then
          'set RowSource (set of elements populating the combo box)
          strFormValidation = "RowSource: " & ctl.RowSource
        End If
      End If
      
      On Error GoTo errHandler
  
      If Validator.Reporter_isInArray(strFormInputType, arrLockableControlTypes) Then
        'set Editable status
        If ((ctl.Locked = False) And (ctl.Enabled = True)) Then
           strEditable = "Y"
        Else
           strEditable = "N"
        End If
      End If
    
      If Validator.Reporter_isInArray(strFormInputType, arrSourceableControlTypes) Then
        If Not ctl.ControlSource = "" Then
          'set ControlSource status
          strDbColumnName = ctl.ControlSource
        Else
          strDbColumnName = "No DbColumn Assigned"
        End If
      End If
        
      'Print only the control types that are in the printable array
      If Validator.Reporter_isInArray(strFormInputType, arrPrintableControlTypes) Then
        Debug.Print "(a) " & strFormControlName & " (b) " & strFormInputType & " (c) " & strFormValidation & " (d) " & strEditable & " (e) " & strDbColumnName
      
        If Validator.DOUBLE_LINE_SPACING = True Then
          Debug.Print "---------------------------------------------------------------------------------------------------------------------------------"
        End If
    End If
    
  Next ctl
  Set ctl = Nothing
  
  Exit Sub
  
errHandler:
  Debug.Print "Error Handled: " & strFormControlName & " : " & Err.Description
  Exit Sub
End Sub


'
'Reporter_isInArray
'
'This is a copy of the ArrayHelper function.
'It is duplicated and set to Private to keep
'Reporter a self-contained application
'
'@param variant varToBeFound
'@param variant arr
'@return bool
'
Private Function Reporter_isInArray(ByVal varToBeFound As Variant, ByVal arr As Variant) As Boolean

  If Validator.Reporter_isEmptyArray(arr) Then
    Reporter_isInArray = False
    Exit Function
  End If
  
  Dim i As Long
  For i = 0 To (Validator.Reporter_getSize(arr) - 1)
  
    If IsNull(varToBeFound) And IsNull(arr(i)) Then
      Reporter_isInArray = True
      Exit For
    End If
    
    If varToBeFound = arr(i) Then
      Reporter_isInArray = True
      Exit For
    End If
  Next i
  
End Function

'
'Reporter_isEmptyArray
'
'This is a copy of the ArrayHelper function.
'It is duplicated and set to Private to keep
'Reporter a self-contained application
'
'@param variant varArray
'@return bool
'
Private Function Reporter_isEmptyArray(ByVal varArray As Variant) As Boolean
  On Error GoTo IS_EMPTY
  If (UBound(varArray) >= 0) Then Exit Function
IS_EMPTY:
  Reporter_isEmptyArray = True
End Function

'
'Reporter_getSize
'
'This is a copy of the ArrayHelper function.
'It is duplicated and set to Private to keep
'Reporter a self-contained application
'
'@param variant varArray
'@return integer intSize
'
Private Function Reporter_getSize(ByVal varArray As Variant) As Long
  If Validator.Reporter_isEmptyArray(varArray) Then
    Reporter_getSize = 0
    Exit Function
  End If
  
  Reporter_getSize = (UBound(varArray) - LBound(varArray)) + 1
End Function




                                                                       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STRING FORMATTER                                                        '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'StringFormatter_sanitize
'Remove characters not allowed in SQL and trim whitespace at the
'start and end of the string
'
'If strVal is Null, return a blank string
'
'@param variant str
'@return variant
'
Public Function StringFormatter_sanitize(ByVal strVal As Variant) As String

  If IsNull(strVal) Then
    strVal = ""
  End If

  strVal = Replace(strVal, Chr(10), "")
  strVal = Replace(strVal, Chr(13), "")
  strVal = Replace(strVal, "'", "") 'Strip single quotes
  strVal = Replace(strVal, """", "") 'Strip double quotes
  'strVal = Replace(strVal, ",", "")'Strip commas
  'strVal = Replace(strVal, "/", "-")'Strip forward slash
  'strVal = Replace(strVal, "\", "-")'Strip backslash
  
  'trim whitespace from the start and the end of the string
  strVal = Trim(strVal)
  
  StringFormatter_sanitize = strVal
End Function

'
'StringFormatter_addHeading
'
'Concatenate a string to the start of a string and
'add two line breaks
'
'@param variant strHeading
'@param variant strBody
'
Public Function StringFormatter_addHeading(ByVal strHeading As Variant, ByVal strBody As Variant) As String
  If IsNull(strHeading) Then
    strHeading = ""
  End If
  
  If IsNull(strBody) Then
    strBody = ""
  End If
  
  StringFormatter_addHeading = strHeading & vbNewLine & vbNewLine & strBody
End Function

'
'StringFormatter_addLineToBody
'
'Concatenate a string to the end of a string and
'add one line breaks
'
'@param string strLine
'@param string strBody
'
Public Function StringFormatter_addLineToBody(ByVal strLine As String, ByVal strBody As String) As String
  StringFormatter_addLineToBody = strBody & strLine & vbNewLine
End Function


'
'StringFormatter_sanitizeTextInputs
'
'Remove dangerous characters from the input strings in each
'field where the ControlDetails are marked as Plaintext
'
'@param variant varValidationRulesArray
'@param object frmForm
'
Public Sub StringFormatter_sanitizeTextInputs(ByVal varValidationRules As Variant, ByRef frmForm As Object)

  Dim strFunctionName As String: strFunctionName = "Validator.StringFormatter_sanitizeTextInputs"

  'Iterate through the array until a sanitizable datatype is found
  Dim i As Long: i = 0
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getSanitizableDataTypes)) Then

      'Validate inputs
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

      'Determine whether the input is to be sanitized. Guarding against Null inputs being passed into the string formatter
      'Also, check that the field is editable (Enabled and not locked). Without this isEditable check, Access would raise
      'an error whenever a field is not Enabled or is Locked when the sanitizer is run.
      '
      'If nullable and input is Null, do not sanitize. Else sanitize
      If (Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getNullableDataTypes)) Then
        'Is Nullable
        If IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          'Nullable and null. No sanitizing
        Else
          'Only change the value in the form-field if the field is editable
          If Validator.FormControlHelper_isEditable(varValidationRules(i, 0), frmForm) Then
            'Nullable and NOT null. Sanitize the value in the form
            frmForm.Controls(varValidationRules(i, 0)).Value = Validator.StringFormatter_sanitize(frmForm.Controls(varValidationRules(i, 0)).Value)
          End If
        End If
      Else
        'Not nullable
        If IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
          'The non-nullable Null value should not have reached here. Raise an error
          Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "ControlName cannot be Null. Should be filtered before entering this function: " & varValidationRules(i, 0))
        Else
          'Only change the value in the form-field if the field is editable
          If Validator.FormControlHelper_isEditable(varValidationRules(i, 0), frmForm) Then
            'Value is not nullable and also not null. Sanitize the value
            frmForm.Controls(varValidationRules(i, 0)).Value = Validator.StringFormatter_sanitize(frmForm.Controls(varValidationRules(i, 0)).Value)
          End If
        End If
      End If

    End If
  Next i
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'STRING VALIDATION HELPER                                                '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'
'StringValidationHelper_allStringInArrayInputsAreValidStrings
'
'Check that the input is in the array of strings
'
'@param variant varValidationRules
'@param object frmForm
'@param string strNotifications
'@return boolean
'
Public Function StringValidationHelper_allStringInArrayInputsAreValidStrings(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "StringInArray"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_allStringInArrayInputsAreValidStrings"
  
  StringValidationHelper_allStringInArrayInputsAreValidStrings = True
  
  'Iterate through the array until a strDatatype is found
  Dim i As Long
  
  
  Dim varElem As Variant
  
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strDatatype) Then

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)
      
      'NonNullable values should never reach here as Null
      Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(varValidationRules(i, 0)).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
     
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Input is nullable and not null
        
        'If value is multi-select ComboBox, it passes an Array instead of a String
        'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
        If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 Then
          For Each varElem In frmForm.Controls(varValidationRules(i, 0)).Value
            If (Validator.ArrayHelper_isInArray(varElem, varValidationRules(i, 4)) = False) Then
              'String is not in the array of valid strings
              Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
              strNotifications = Validator.StringFormatter_addLineToBody(varValidationRules(i, 0) & Validator.IS_NOT_A_VALID_SELECTION, strNotifications)
              StringValidationHelper_allStringInArrayInputsAreValidStrings = False
            End If
          Next varElem
          
        Else
          If (Validator.ArrayHelper_isInArray(frmForm.Controls(varValidationRules(i, 0)).Value, varValidationRules(i, 4)) = False) Then
            'String is not in the array of valid strings
            Call Validator.UserInterfaceMgr_setControlBackColor(varValidationRules(i, 0), "Error", frmForm)
            strNotifications = Validator.StringFormatter_addLineToBody(varValidationRules(i, 0) & Validator.IS_NOT_A_VALID_SELECTION, strNotifications)
            StringValidationHelper_allStringInArrayInputsAreValidStrings = False
          End If
        End If

      End If

    End If
  Next i

End Function



'
'StringValidationHelper_allAlphaTextInputsAreValidStrings
'
'Check that the plaintext inputs are valid strings
'data types
'
'@param string strNotifications
'@param variant varValidationRulesArray
'@param object frmForm
'@return Boolean
'
Public Function StringValidationHelper_allAlphaTextInputsAreValidStrings(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean
  
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_allAlphaTextInputsAreValidStrings"
  
  StringValidationHelper_allAlphaTextInputsAreValidStrings = True
  
  'Iterate through the array until a "AlphaText" is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = "AlphaText") Then

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)
      
        'If value is multi-select ComboBox, it passes an Array instead of a String
        'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
        If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 And Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getDatatypesThatAreArrayable()) Then
          Call Validator.ErrorHandler_logAndRaiseError("StringValidationHelper.allAlphaTextInputsAreValidStrings", varValidationRules(i, 0) & " is an array. It must be a string. This error probably occurred because you are using a Multi-select ComboBox with AlphaText Validator Constraint. Multi-select ComboBox can only use either IntegerInArray or StringInArray.")
        End If

      If Not Validator.StringValidationHelper_isValidAlphaText(varValidationRules(i, 0), varValidationRules(i, 4)(0), varValidationRules(i, 4)(1), strNotifications, frmForm) Then
        StringValidationHelper_allAlphaTextInputsAreValidStrings = False
      End If
      
    End If
  Next i

End Function

'
'StringValidationHelper_isValidAlphaText
'
'@param string control title
'@param integer minLength
'@param integer maxLength
'@param object frmForm
'
Private Function StringValidationHelper_isValidAlphaText(ByVal strTitle As String, ByVal intMinLength As Long, ByVal intMaxLength As Long, ByRef strNotifications As String, ByRef frmForm As Object) As Boolean

  Dim strDatatype As String: strDatatype = "AlphaText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_isValidAlphaText"
  
  StringValidationHelper_isValidAlphaText = True
  
  'NonNullable values should never reach here as Null
  Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(strTitle).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
  
  'If nullable and input is Null, do not validate.
  If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(strTitle).Value) Then
    Exit Function
  End If

  If (Validator.DataValidator_isAlpha(frmForm.Controls(strTitle).Value) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.CAN_ONLY_BE_LETTERS, strNotifications)
    StringValidationHelper_isValidAlphaText = False
  End If

  If (Validator.DataValidator_stringLengthIsInRange(frmForm.Controls(strTitle).Value, intMinLength, intMaxLength) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.INVALID_STRING_LENGTH & "Min: " & _
    intMinLength & "  Max: " & intMaxLength, strNotifications)
    StringValidationHelper_isValidAlphaText = False
  End If
End Function


'
'StringValidationHelper_allAlphanumericTextInputsAreValidStrings
'
'Check that the plaintext inputs are valid strings
'data types
'
'@param string strNotifications
'@param variant varValidationRulesArray
'@param object frmForm
'@return Boolean
'
Public Function StringValidationHelper_allAlphanumericTextInputsAreValidStrings(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "AlphanumericText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_allAlphanumericTextInputsAreValidStrings"
  
  StringValidationHelper_allAlphanumericTextInputsAreValidStrings = True

  'Iterate through the array until a "AlphanumericText" is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strDatatype) Then


      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

        'If value is multi-select ComboBox, it passes an Array instead of a String
        'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
        If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 And Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getDatatypesThatAreArrayable()) Then
          Call Validator.ErrorHandler_logAndRaiseError("StringValidationHelper.allAlphanumericTextInputsAreValidStrings", varValidationRules(i, 0) & " is an array. It must be a string. This error probably occurred because you are using a Multi-select ComboBox with AlphanumericText Validator Constraint. Multi-select ComboBox can only use either IntegerInArray or StringInArray.")
        End If
      
      If Not Validator.StringValidationHelper_isValidAlphanumericText(varValidationRules(i, 0), varValidationRules(i, 4)(0), varValidationRules(i, 4)(1), strNotifications, frmForm) Then
        StringValidationHelper_allAlphanumericTextInputsAreValidStrings = False
      End If

    End If
  Next i

End Function

'
'StringValidationHelper_isValidAlphanumericText
'
'@param string control title
'@param integer minLength
'@param integer maxLength
'@param object frmForm
'
Private Function StringValidationHelper_isValidAlphanumericText(ByVal strTitle As String, ByVal intMinLength As Long, ByVal intMaxLength As Long, ByRef strNotifications As String, ByRef frmForm As Object) As Boolean

  Dim strDatatype As String: strDatatype = "AlphanumericText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_isValidAlphanumericText"

  StringValidationHelper_isValidAlphanumericText = True

  'NonNullable values should never reach here as Null
  Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(strTitle).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
  
  'If nullable and input is Null, do not validate.
  If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(strTitle).Value) Then
    Exit Function
  End If
  
  If (Validator.DataValidator_isAlphaNumeric(frmForm.Controls(strTitle).Value) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.CAN_ONLY_BE_LETTERS_AND_NUMBERS, strNotifications)
    StringValidationHelper_isValidAlphanumericText = False
  End If

  If (Validator.DataValidator_stringLengthIsInRange(frmForm.Controls(strTitle).Value, intMinLength, _
  intMaxLength) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.INVALID_STRING_LENGTH & "Min: " & _
    intMinLength & "  Max: " & intMaxLength, strNotifications)
    StringValidationHelper_isValidAlphanumericText = False
  End If
End Function


'
'StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings
'
'Check that the plaintext inputs are valid strings
'data types
'
'@param string strNotifications
'@param variant varValidationRulesArray
'@param object frmForm
'@return Boolean
'
Public Function StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "SpecialCharacterText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings"
  
  StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings = True
  
  'Iterate through the array until a "AlphanumericText" is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strDatatype) Then
      

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

      'If value is multi-select ComboBox, it passes an Array instead of a String
      'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
      If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 And Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getDatatypesThatAreArrayable()) Then
        Call Validator.ErrorHandler_logAndRaiseError("StringValidationHelper.allSpecialCharacterTextInputsAreValidStrings", varValidationRules(i, 0) & " is an array. It must be a string. This error probably occurred because you are using a Multi-select ComboBox with SpecialCharacterText Validator Constraint. Multi-select ComboBox can only use either IntegerInArray or StringInArray.")
      End If

      
      If Not Validator.StringValidationHelper_isValidSpecialCharacterText(varValidationRules(i, 0), varValidationRules(i, 4)(0), varValidationRules(i, 4)(1), strNotifications, frmForm) Then
        StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings = False
      End If

    End If
  Next i

End Function

'
'StringValidationHelper_isValidSpecialCharacterText
'
'@param string control title
'@param integer minLength
'@param integer maxLength
'@param object frmForm
'
Private Function StringValidationHelper_isValidSpecialCharacterText(ByVal strTitle As String, ByVal intMinLength As Long, ByVal intMaxLength As Long, ByRef strNotifications As String, ByRef frmForm As Object) As Boolean

  Dim strDatatype As String: strDatatype = "SpecialCharacterText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_isValidSpecialCharacterText"

  StringValidationHelper_isValidSpecialCharacterText = True

  'NonNullable values should never reach here as Null
  Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(strTitle).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
  
  'If nullable and input is Null, do not validate.
  If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(strTitle).Value) Then
    Exit Function
  End If

  If (Validator.DataValidator_isSpecialCharacterText(frmForm.Controls(strTitle).Value) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.CAN_ONLY_BE_LETTERS_AND_NUMBERS_AND_SPECIAL_CHARACTERS, strNotifications)
    StringValidationHelper_isValidSpecialCharacterText = False
  End If

  If (Validator.DataValidator_stringLengthIsInRange(frmForm.Controls(strTitle).Value, intMinLength, _
  intMaxLength) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.INVALID_STRING_LENGTH & "Min: " & _
    intMinLength & "  Max: " & intMaxLength, strNotifications)
    StringValidationHelper_isValidSpecialCharacterText = False
  End If
End Function





'
'StringValidationHelper_allAllAllowableTextInputsAreValidStrings
'
'Check that the plaintext inputs are valid strings
'data types
'
'@param string strNotifications
'@param variant varValidationRulesArray
'@param object frmForm
'@return Boolean
'
Public Function StringValidationHelper_allAllAllowableTextInputsAreValidStrings(ByVal varValidationRules As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strDatatype As String: strDatatype = "AllAllowableText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_allAllAllowableTextInputsAreValidStrings"
  
  StringValidationHelper_allAllAllowableTextInputsAreValidStrings = True
  
  'Iterate through the array until a "AlphanumericText" is found
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strDatatype) Then
      

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)

      'If value is multi-select ComboBox, it passes an Array instead of a String
      'Check for VarType = 8204 (Array). If true, validate each element in the array. Else, the value is a string that is checked against (x, 4) array of strings
      If VarType(frmForm.Controls(varValidationRules(i, 0)).Value) = 8204 And Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getDatatypesThatAreArrayable()) Then
        Call Validator.ErrorHandler_logAndRaiseError("StringValidationHelper.allAllAllowableTextInputsAreValidStrings", varValidationRules(i, 0) & " is an array. It must be a string. This error probably occurred because you are using a Multi-select ComboBox with AllAllowableText Validator Constraint. Multi-select ComboBox can only use either IntegerInArray or StringInArray.")
      End If

      If Not Validator.StringValidationHelper_isValidAllAllowableText(varValidationRules(i, 0), varValidationRules(i, 4)(0), varValidationRules(i, 4)(1), strNotifications, frmForm) Then
        StringValidationHelper_allAllAllowableTextInputsAreValidStrings = False
      End If

    End If
  Next i

End Function



'
'StringValidationHelper_isValidAllAllowableText
'
'@param string control title
'@param integer minLength
'@param integer maxLength
'@param object frmForm
'
Private Function StringValidationHelper_isValidAllAllowableText(ByVal strTitle As String, ByVal intMinLength As Long, ByVal intMaxLength As Long, ByRef strNotifications As String, ByRef frmForm As Object) As Boolean

  Dim strDatatype As String: strDatatype = "AllAllowableText"
  Dim strFunctionName As String: strFunctionName = "Validator.StringValidationHelper_isValidAllAllowableText"

  StringValidationHelper_isValidAllAllowableText = True

  'NonNullable values should never reach here as Null
  Call Validator.Datatypes_ifNonNullableFieldIsNullRaiseError(strDatatype, frmForm.Controls(strTitle).Value, Validator.Datatypes_getNullableDataTypes, strFunctionName)
  
  'If nullable and input is Null, do not validate.
  If (Validator.ArrayHelper_isInArray(strDatatype, Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(strTitle).Value) Then
    Exit Function
  End If

  If (Validator.DataValidator_isAllAllowableText(frmForm.Controls(strTitle).Value) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.MUST_BE_ALL_ALLOWABLE_TEXT, strNotifications)
    StringValidationHelper_isValidAllAllowableText = False
  End If

  If (Validator.DataValidator_stringLengthIsInRange(frmForm.Controls(strTitle).Value, intMinLength, _
  intMaxLength) = False) Then
    Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, "Error", frmForm)
    strNotifications = Validator.StringFormatter_addLineToBody(strTitle & Validator.INVALID_STRING_LENGTH & "Min: " & _
    intMinLength & "  Max: " & intMaxLength, strNotifications)
    StringValidationHelper_isValidAllAllowableText = False
  End If
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TWO DIM VALIDATION MGR                                                  '
'                                                                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'
'TwoDimValidationMgr_get2DimensionalDataDictionaryNameAndGrpIntBasedOn2ndDimensionVariableValue
'
'Determines which #FirstVariable#Grp will be used, based on the value value of the second variable
'The function returns the integer in the #FirstVariable#Grp
'
'@param string strNameOfFirstVariableIn2DimensionalGrp
'@param string strNameOfSecondVariableIn2DimensionalGrp
'@param double dblValueOfSecondVariable
'@param variant var2DimValidationRules Array
'@return variant array
'@return varReturn(0) will be data dictionary name for the firstVariableIn2DimensionalGrp
'@return varReturn(1) will be the GrpInt
'
Public Function TwoDimValidationMgr_get2DimensionalDataDictionaryNameAndGrpIntBasedOn2ndDimensionVariableValue(ByVal strNameOfFirstVariableIn2DimensionalGrp As String, _
ByVal strNameOfSecondVariableIn2DimensionalGrp As String, ByVal dblSecondDimVariableValue As Double, _
ByVal var2DimValidationRules As Variant) As Variant
  
  Dim varReturn(1) As Variant
  'varReturn(0) will be data dictionary name for the firstVariableIn2DimensionalGrp
  'varReturn(1) will be the GrpInt
  
  'Iterate through the array until a match is found for the strFirstVariable and strSecondVariable in the array of 2Dimensional boundaries
  Dim varBoundaryValues() As Variant
  Dim i As Long
  Dim j As Long
  For i = 0 To (Validator.ArrayHelper_getSize(var2DimValidationRules) - 1)
    If ((strNameOfFirstVariableIn2DimensionalGrp = var2DimValidationRules(i, 0)) And _
    (strNameOfSecondVariableIn2DimensionalGrp = var2DimValidationRules(i, 2))) Then
      'Found the element with the requested boundary values
      
      'Set the name used by the data-dictionary
      varReturn(0) = var2DimValidationRules(i, 1)
      
      'Find the GrpInt
      varBoundaryValues = var2DimValidationRules(i, 3)

      For j = 0 To (Validator.ArrayHelper_getSize(varBoundaryValues) - 1)
  
        'Exit if this is the final boundary value
        If (j = (Validator.ArrayHelper_getSize(varBoundaryValues) - 1)) Then Exit For
    
        'Test whether the second dim variable value is less than the boundary value
        'On each iteration, check whether the value is still below the boundary.
        'x < BOUNDARY1 // RED
        'x < BOUNDARY2 // ORANGE
        'x < BOUNDARY3 // GREEN
        'x < BOUNDARY4 // ORANGE
        'x >= BOUNDARY4 // RED
        If (dblSecondDimVariableValue < varBoundaryValues(j, 0)) Then Exit For
      Next j

      'Assign the value to the return value
      varReturn(1) = varBoundaryValues(j, 1)
      Exit For
    End If
  Next i
  
  TwoDimValidationMgr_get2DimensionalDataDictionaryNameAndGrpIntBasedOn2ndDimensionVariableValue = varReturn
End Function


'
'TwoDimValidationMgr_get2DimensionalGrpIntBasedOnVariable
'
'A 2D (Two dimensional) LocationInRange.
'Add notification to strWarnings if the value is in the ORANGE range
'Add notification to strOutOfBound if the value is in the RED range
'
'@param string mFirstDimensionName
'@param double mFirstDimensionValue
'@param string mSecondDimensionName
'@param double mSecondDimensionValue
'@param string strWarnings
'@param string strOutOfBounds
'@param variant array var2DimValidationRules
'@return integer in range {1, 2, 3, 4, 5}
'
'LocationInRange codes
'1 = Red
'2 = Orange
'3 = Green
'4 = Orange
'5 = Red
'
Public Function TwoDimValidationMgr_get2DimensionalGrpIntBasedOnVariable(ByVal mFirstDimensionName As String, _
ByVal mFirstDimensionValue As Double, _
ByVal mSecondDimensionName As String, _
ByVal mSecondDimensionValue As Double, _
ByRef strWarnings As String, _
ByRef strOutOfBounds As String, _
ByVal var2DimValidationRules As Variant) As Long

  Dim strFunctionName As String: strFunctionName = "Validator.TwoDimValidationMgr_get2DimensionalGrpIntBasedOnVariable"
  
  'Get the DataDictionary group Code {int 1 to 5}
  Dim strDataDictionaryTitle As String
  Dim intGrpInt As Long
  Dim varDataDictionaryTitleAndGrpInt() As Variant
  varDataDictionaryTitleAndGrpInt = Validator.TwoDimValidationMgr_get2DimensionalDataDictionaryNameAndGrpIntBasedOn2ndDimensionVariableValue(mFirstDimensionName, _
  mSecondDimensionName, mSecondDimensionValue, var2DimValidationRules)
  
  strDataDictionaryTitle = varDataDictionaryTitleAndGrpInt(0)
  intGrpInt = varDataDictionaryTitleAndGrpInt(1)
  
  '
  'Check the mFirstDimensionDataDictionaryName and intGrpInt before passing them into the Validator_Boundaries.getGrp function
  '
  'Raise Error if the intGrpInt is Null or 0 or blank
  If IsNull(intGrpInt) Or (intGrpInt = 0) Then
    Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "Data Dictionary group number request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
  End If
  
  'Raise Error if the mFirstDimensionDataDictionaryName is Null or blank
  If IsNull(strDataDictionaryTitle) Or (strDataDictionaryTitle = "") Then
    Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
  End If
      
  'Retrieve the requested group from the Data Dictionary
  Dim arrGrp() As Double: arrGrp = Validator_Boundaries.getGrp(strDataDictionaryTitle, intGrpInt)
  
  'Get the location in range As Long
  Dim int2DLocationInRange As Long: int2DLocationInRange = Validator.BoundaryMgr_getLocationInRangeDouble(mFirstDimensionName, mFirstDimensionValue, _
  arrGrp, strWarnings, strOutOfBounds)
  
  'getLocationInRangeDouble returns a -1 whenever there was an invalid request
  If (int2DLocationInRange = -1) Then
    Call Validator.ErrorHandler_logAndRaiseError(strFunctionName, "Validator.BoundaryMgr_getLocationInRangeDouble returned an invalid value " & int2DLocationInRange & " for mFirstDimensionName: " & mFirstDimensionName & " for mFirstDimensionValue: " & " Array: " & Validator.ArrayHelper_concatAllElementsIn1DArray(arrGrp))
  End If
  
  TwoDimValidationMgr_get2DimensionalGrpIntBasedOnVariable = int2DLocationInRange

End Function







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'USER INTERFACE MGR                                                      '
'                                                                        '
'This contains the functions to change the form styles                   '
'For example, set background color for control                           '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'UserInterfaceMgr_resetBackColors()
'Set all input fields to the default background color
'This function uses the array for control names from getControlNames.
'It does not need to be edited
'
'@param array of strings varControlNamesArray
'@param Form frmForm
'@return void
'
Public Sub UserInterfaceMgr_resetBackColors(ByVal varControlNamesArray As Variant, ByRef frmForm As Object)
  Dim arrControlNames() As String: arrControlNames = varControlNamesArray
  Dim strControlName As Variant
  
  For Each strControlName In arrControlNames
    Call Validator.UserInterfaceMgr_setControlBackColor(strControlName, "Default", frmForm)
  Next strControlName
End Sub

'
'UserInterfaceMgr_setControlBackColor
'
'Set the BackColor of the ControlName Textbox
'
'@param String strControlName
'@param String strStatus
'@param Form frmForm
'
Public Sub UserInterfaceMgr_setControlBackColor(ByVal strControlName As String, ByVal strStatus As String, ByRef frmForm As Object)

  Dim strFunctionName As String: strFunctionName = "Validator.UserInterfaceMgr_setControlBackColor"

  'Guard against strControlName being Null or an Empty string
  If IsNull(strControlName) Or (strControlName = "") Then
    Exit Sub
  End If

  'Validate input
  Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, strControlName, frmForm)
  
  'Some control types (for example checkbox) do not have the BackColor method. Use BorderColor instead
  'If a different error other than Runtime 248 is raised, it will be caught in FormFuncs
  On Error GoTo UseBorderColorInsteadOfBackColor

  If strStatus = "Default" Then
    frmForm.Controls(strControlName).BackColor = Validator.TEXTBOX_BACKCOLOR_DEFAULT
    Exit Sub
  End If
  
  If strStatus = "Warning" Then
    frmForm.Controls(strControlName).BackColor = Validator.TEXTBOX_BACKCOLOR_WARNING
    Exit Sub
  End If
  
  If strStatus = "Error" Then
    frmForm.Controls(strControlName).BackColor = Validator.TEXTBOX_BACKCOLOR_ERROR
    Exit Sub
  End If
  
  Exit Sub
  
UseBorderColorInsteadOfBackColor:
  If Err.Number = 248 Then
    
    If strStatus = "Default" Then
      frmForm.Controls(strControlName).BorderColor = Validator.TEXTBOX_BACKCOLOR_DEFAULT
      Exit Sub
    End If
    
    If strStatus = "Warning" Then
      frmForm.Controls(strControlName).BorderColor = Validator.TEXTBOX_BACKCOLOR_WARNING
      Exit Sub
    End If
    
    If strStatus = "Error" Then
      frmForm.Controls(strControlName).BorderColor = Validator.TEXTBOX_BACKCOLOR_ERROR
      Exit Sub
    End If
    
  End If
  
End Sub


'
'UserInterfaceMgr_setControlBackColorFromLocationInRange
'
'Assign the background color to the field based on the LocationInRange
'This function must contain one conditional statement for each LocationInRange
'input field.
'
'@param string strTitle
'@param integer locationInRange
'@param Form frmForm
'
Public Sub UserInterfaceMgr_setControlBackColorFromLocationInRange(ByVal strTitle As String, ByVal locationInRange As Long, ByRef frmForm As Object)
  Call Validator.UserInterfaceMgr_setControlBackColor(strTitle, Validator.UserInterfaceMgr_convertLocationInRangeIntToBackColorString(locationInRange), frmForm)
End Sub

'
'UserInterfaceMgr_convertLocationInRangeIntToBackColorString
'
'Convert the locationInRange code to BackColor
'notification type
'
'@param integer intCode
'@return string strBackColor
'
Public Function UserInterfaceMgr_convertLocationInRangeIntToBackColorString(ByVal intCode As Long) As String
  Select Case intCode
    Case 1
      UserInterfaceMgr_convertLocationInRangeIntToBackColorString = "Error"
    Case 2
      UserInterfaceMgr_convertLocationInRangeIntToBackColorString = "Warning"
    Case 3
      UserInterfaceMgr_convertLocationInRangeIntToBackColorString = "Default"
    Case 4
      UserInterfaceMgr_convertLocationInRangeIntToBackColorString = "Warning"
    Case 5
      UserInterfaceMgr_convertLocationInRangeIntToBackColorString = "Error"
  End Select
End Function






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'VALIDATION RULES MGR                                                    '
'                                                                        '
'Manage the data-structures (arrays) containing the validation rules     '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'ValidationRulesMgr_getControlNames()
'
'@param variant varValidationRulesArray
'@return array of strings FormControlNames
'
Public Function ValidationRulesMgr_getControlNames(ByVal varValidationRules As Variant) As String()

  'Define the array that will store the array of control names
  Dim arrControlNames() As String: ReDim arrControlNames(0 To 0)
  
  'Only run this if the varValidationRules is not an empty array
  If Validator.ArrayHelper_getSize(varValidationRules) > 0 Then
    Dim i As Long
    'Iterate through the array and create an array of names as strings
    For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
      arrControlNames(UBound(arrControlNames)) = varValidationRules(i, 0): arrControlNames = Validator.ArrayHelper_increase1DArraySizeByOne(arrControlNames)
    Next i
  
    'The final element will be a blank element. Clip this last element of the array
    arrControlNames = Validator.ArrayHelper_clipLastElement(arrControlNames)
  End If

  ValidationRulesMgr_getControlNames = arrControlNames
End Function


'
'ValidationRulesMgr_countControlDetailsOfType
'
'Count number of elements in the array that are of the requested type
'
'@param variant array varValidationRules
'@param string strType
'@return integer
'
Public Function ValidationRulesMgr_countControlDetailsOfType(ByVal strType As String, ByVal varValidationRules As Variant) As Long
  Dim counter As Long: counter = 0
  'Count number of elements in the array that are of the requested type
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 2) = strType) Then
      counter = counter + 1
    End If
  Next i
  ValidationRulesMgr_countControlDetailsOfType = counter
End Function

'
'ValidationRulesMgr_countRequiredControlDetails
'
'Count number of elements in the array that are marked as required
'
'@param variant array varValidationRules
'@return integer
'
Public Function ValidationRulesMgr_countRequiredControlDetails(ByVal varValidationRules As Variant) As Long
  ValidationRulesMgr_countRequiredControlDetails = 0
  'Count number of elements in the array that are of the requested type
  Dim i As Long
   
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 1) = True) Then
      ValidationRulesMgr_countRequiredControlDetails = ValidationRulesMgr_countRequiredControlDetails + 1
    End If
  Next i
End Function

'
'ValidationRulesMgr_getDatatypeByControlDetailsTitle
'
'@param string strControlDetailsTitle
'@param variant varValidationRules
'@return string strDatatype
'
Public Function ValidationRulesMgr_getDatatypeByControlDetailsTitle(ByVal strControlDetailsTitle As String, ByVal varValidationRules As Variant) As String
  
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)

    If varValidationRules(i, 0) = strControlDetailsTitle Then
      ValidationRulesMgr_getDatatypeByControlDetailsTitle = varValidationRules(i, 2)
      Exit Function
    End If
    
  Next i
  Call Validator.ErrorHandler_logAndRaiseError("getDatatypeByControlDetailsTitle", "ControlName not found in ControlDetails (Validation Rules) array: " & varValidationRules(i, 0))
End Function


'
'ValidationRulesMgr_getRequiredControlDetails
'
'@param variant varValidationRulesArray
'@return array of strings RequiredControlNames
'
Public Function ValidationRulesMgr_getRequiredControlDetails(ByVal varValidationRules As Variant) As Variant()
  'Count number of elements in the array that are of the requested type
  Dim intCountRequiredElements As Long: intCountRequiredElements = Validator.ValidationRulesMgr_countRequiredControlDetails(varValidationRules)
  
  'Create the 2D array where the ControlDetails will be collected.
  Dim varRequiredControlDetails() As Variant
  
  'If there are no required controls, return an empty array
  If intCountRequiredElements > 0 Then
  
    ReDim varRequiredControlDetails(intCountRequiredElements - 1, 6) 'Each element has 7 _
    properties set by the user. (6 because arrays are zero-indexed)
    
    Dim j As Long: j = 0 'Count the index for the arrControlDetailsForRequestedType
    Dim i As Long
    
    For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
      'If the strType is required, then add the name of this field to the array
      If (varValidationRules(i, 1) = True) Then
        varRequiredControlDetails(j, 0) = varValidationRules(i, 0)
        varRequiredControlDetails(j, 1) = varValidationRules(i, 1)
        varRequiredControlDetails(j, 2) = varValidationRules(i, 2)
        varRequiredControlDetails(j, 3) = varValidationRules(i, 3)
        varRequiredControlDetails(j, 4) = varValidationRules(i, 4)
        varRequiredControlDetails(j, 5) = varValidationRules(i, 5)
        
        j = j + 1
      End If
    Next i
  
  End If

  ValidationRulesMgr_getRequiredControlDetails = varRequiredControlDetails
End Function


'
'ValidationRulesMgr_getValidationRulesWithUneditedFieldsRemoved
'
'Take the array of validation rules and return the array
'with only the edited validation rules.
'
'@param variant varValidationRules
'@param object frmForm
'@param variant varEditedValidationRules or Empty Array if no edits
'
Public Function ValidationRulesMgr_getValidationRulesWithUneditedFieldsRemoved(ByVal varValidationRules As Variant, ByRef frmForm As Object) As Variant
  
  Dim varEditedValidationRules() As Variant
  Dim intCountEditedControlsInValidationRules As Long
  intCountEditedControlsInValidationRules = Validator.ValidationRulesMgr_countEditedControlsInValidationRules(varValidationRules, frmForm)
      
  'If there are any edited fields, redim (resize) the EditedValidationArray and fill it with the validation rules
  If intCountEditedControlsInValidationRules > 0 Then
  
    'Initialize the varEditedValidationRules to the size of the count of validation rules fields that have been edited
    ReDim varEditedValidationRules(intCountEditedControlsInValidationRules - 1, 5)
  
    'Iterate through all varValidationRules. Whenever a ValidationRule is found that has an edited field, all
    'its validation rules to the varEditedValidationRules array.

    Dim i As Long
    Dim j As Long: j = 0 'j counts the index for the varEditedValidationRules array
   
    For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
      'Check that the field is a bound value.
      If Validator.FormControlHelper_isBoundField(varValidationRules(i, 0), frmForm) = True Then
        If Validator.FormControlHelper_isEditedField(varValidationRules(i, 0), frmForm) = True Then
        
          'Add the validation rules for the edited field to the EditedValidationRules array
          varEditedValidationRules(j, 0) = varValidationRules(i, 0)
          varEditedValidationRules(j, 1) = varValidationRules(i, 1)
          varEditedValidationRules(j, 2) = varValidationRules(i, 2)
          varEditedValidationRules(j, 3) = varValidationRules(i, 3)
          varEditedValidationRules(j, 4) = varValidationRules(i, 4)
          varEditedValidationRules(j, 5) = varValidationRules(i, 5)
        
          j = j + 1
        End If
      End If
    Next i
  End If
  
  'Return the EditedValidationRules, the array of validation rules whose fields have been edited
  'If there were no edits on the form, this will be an empty array
  ValidationRulesMgr_getValidationRulesWithUneditedFieldsRemoved = varEditedValidationRules
End Function

'
'ValidationRulesMgr_countEditedControls
'
'Count number of controls in the validation rules that have been edited
'
'varValidationRules must first be checked to ensure that it only
'contains bound values.
'
'@param variant array varValidationRules
'@return integer
'
Public Function ValidationRulesMgr_countEditedControlsInValidationRules(ByVal varValidationRules As Variant, ByRef frmForm As Object) As Long
  ValidationRulesMgr_countEditedControlsInValidationRules = 0
  'Count number of elements in the array that are bound values
  Dim i As Long
   
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    'Check that the field is a bound value.
    If Validator.FormControlHelper_isBoundField(varValidationRules(i, 0), frmForm) = True Then
      If Validator.FormControlHelper_isEditedField(varValidationRules(i, 0), frmForm) = True Then
        ValidationRulesMgr_countEditedControlsInValidationRules = ValidationRulesMgr_countEditedControlsInValidationRules + 1
      End If
    End If
  Next i
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'VALIDATION RULES VALIDATOR                                              '
'                                                                        '
'Validate the validation rules governing the form                        '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'ValidationRulesValidator_allValidationRulesCorrectlyDefined
'
'Check that the validation rules have been correctly defined
'
'@param varArray varValidationRulesArray
'@param varArray var1DimValidationRules
'@param varArray var2DimValidationRules
'@return Boolean
'
Public Function ValidationRulesValidator_allValidationRulesCorrectlyDefined(ByRef varValidationRules As Variant, _
  ByVal var1DimValidationRules As Variant, _
  ByVal var2DimValidationRules As Variant, ByRef frmForm As Object) As Boolean
  
  'If ValidationRules array is empty, cancel the process
  If Validator.ArrayHelper_isEmptyArray(varValidationRules) Then
    Call Validator.ErrorHandler_logAndRaiseError("Welcome to Validator", "The software has been installed. Please define the ValidationRules to start validating your input")
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If

  If Not Validator.ValidationRulesValidator_allControlDetailsCorrectlyDefined(varValidationRules, frmForm) Then
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If
  
  If Not Validator.ValidationRulesValidator_all1DimensionalVariableDetailsCorrectlyDefined(varValidationRules, var1DimValidationRules, frmForm) Then
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If
  
  If Not Validator.ValidationRulesValidator_all2DimensionalVariableDetailsCorrectlyDefined(varValidationRules, var2DimValidationRules, frmForm) Then
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If
  
  If Not Validator.ValidationRulesValidator_allControlsReferencedInValidationRulesAreRecognizedVarTypes(varValidationRules, frmForm) Then
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If
  
  If Not Validator.ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound(varValidationRules, frmForm) Then
    ValidationRulesValidator_allValidationRulesCorrectlyDefined = False
    Exit Function
  End If

  ValidationRulesValidator_allValidationRulesCorrectlyDefined = True
  
End Function


'
'ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound
'
'@param variant varValidationRules
'@param object frmForm
'@return boolean
'
Public Function ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound(ByRef varValidationRules As Variant, ByRef frmForm As Object) As Boolean
  ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound = True
  Dim i As Long

  'Iterate through the array of control details and check that the necessary rules have been defined correctly
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If Not Validator.FormControlHelper_isBoundField(varValidationRules(i, 0), frmForm) Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", "allControlsReferencedInValidationRulesAreBound: " & varValidationRules(i, 0) & " is an UNBOUND field that is included in the ValidationRules. Validator can only be used with BOUND fields")
      ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound = False
    End If
  Next i
    
End Function

'
'ValidationRulesValidator_allControlsReferencedInValidationRulesAreBound
'
'@param variant varValidationRules
'@param object frmForm
'@return boolean
'
Public Function ValidationRulesValidator_allControlsReferencedInValidationRulesAreRecognizedVarTypes(ByRef varValidationRules As Variant, ByRef frmForm As Object) As Boolean
  ValidationRulesValidator_allControlsReferencedInValidationRulesAreRecognizedVarTypes = True
  Dim i As Long

  'Iterate through the array of control details and check that the necessary rules have been defined correctly
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (Validator.Datatypes_convertIntVarTypeToStrVarType(VarType(varValidationRules(i, 0))) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", "allControlsReferencedInValidationRulesAreRecognizedVarTypes: " & varValidationRules(i, 0) & " uses a datatype in the database column that is not accepted by the Validator software")
      ValidationRulesValidator_allControlsReferencedInValidationRulesAreRecognizedVarTypes = False
    End If
  Next i
    
End Function

'
'ValidationRulesValidator_allControlDetailsCorrectlyDefined
'
'@param variant varValidationRules
'@return boolean
'
Public Function ValidationRulesValidator_allControlDetailsCorrectlyDefined(ByRef varValidationRules As Variant, ByRef frmForm As Object) As Boolean
  Dim i As Long
  
  If Not Validator.ValidationRulesValidator_allRequiredIfValidationRulesAreValid(varValidationRules, frmForm) Then
    Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", "allRequiredIfValidationRulesAreValid: Validation rules for RequiredIf not set correctly")
  End If
  
  'Update the array fields based on whether the RequiredIf conditions were met or not
  Call Validator.ValidationRulesValidator_setRequiredWhenRequiredIfIsUsed(varValidationRules, frmForm)

  'Update the array fields based on whether the RequiredIfNot conditions were met or not
  Call Validator.ValidationRulesValidator_setRequiredWhenRequiredIfNotIsUsed(varValidationRules, frmForm)

  '(x, 5) variables
  Dim j As Long
  
  'Iterate through the array of control details and check that the necessary rules have been defined correctly
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
   
    If Not ValidationRulesValidator_indexZeroValuesValid(varValidationRules(i, 0), frmForm) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
    'This must happen before indexOneValuesValid
    If Not ValidationRulesValidator_indexTwoValuesValid(varValidationRules(i, 0), varValidationRules(i, 2), Validator.Datatypes_getRecognizedDataTypes) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
    If Not ValidationRulesValidator_indexOneValuesValid(varValidationRules(i, 0), varValidationRules(i, 1), varValidationRules(i, 2)) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
    If Not ValidationRulesValidator_indexThreeValuesValid(varValidationRules(i, 0), varValidationRules(i, 3)) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
    If Not ValidationRulesValidator_indexFourValuesValid(varValidationRules(i, 0), varValidationRules(i, 2), varValidationRules(i, 4), varValidationRules(i, 5), Validator.Datatypes_getDatatypesThatUseParameters) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
    If Not ValidationRulesValidator_indexFiveValuesValid(varValidationRules(i, 0), varValidationRules(i, 2), varValidationRules(i, 5), Validator.Datatypes_getValidDatatypesForSecondDimVariable, varValidationRules) Then
      ValidationRulesValidator_allControlDetailsCorrectlyDefined = False
      Exit Function
    End If
    
  Next i
    
  ValidationRulesValidator_allControlDetailsCorrectlyDefined = True
     
End Function



Public Function ValidationRulesValidator_allRequiredIfValidationRulesAreValid(ByRef varValidationRules As Variant, ByRef frmForm As Object) As Boolean

  Dim i As Long

  'Iterate through the array of control details and check that the necessary rules have been defined correctly
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
  
    'Validate the RequiredIf values and update the array with the correct values depending on whether
    'the RequiredIf conditions were met or not met
    
    'All required fields must be a boolean or must have the "RequiredIf" option.
    '(x, 5)
    'If "RequiredIf" option used, confirm that the (x, 5) Array structure is valid
    If (varValidationRules(i, 1) = "RequiredIf") Then
    
      'The (x, 5) value must be set
      If IsNull(varValidationRules(i, 5)) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varValidationRules(i, 0) & " (x, 5) cannot be Null when RequiredIf is used")
      End If
      
      'If using "RequiredIf", the array (x, 5) must be Array("strFieldName", Array(of variants))
      If Not Validator.ArrayHelper_getSize(varValidationRules(i, 5)) = 2 Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varValidationRules(i, 0) & " (x, 5) must be Array(" & Chr(34) & "strFieldName" & Chr(34) & ", Array(1, 2)) or Array(" & Chr(34) & "strFieldName" & Chr(34) & ", Array(" & Chr(34) & "item1" & Chr(34) & ", " & Chr(34) & "item2" & Chr(34) & "))")
      End If
      
      'If using "RequiredIf", the array (x, 5)(0) must be a string
      If Not VarType(varValidationRules(i, 5)(0)) = 8 Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varValidationRules(i, 0) & " is using an (x, 5) value that is not a valid control name for RequiredIf")
      End If
      
      'If using "RequiredIf", the array (x, 5)(0) must be a valid control name
      If Not Validator.FormControlHelper_controlExists(varValidationRules(i, 5)(0), frmForm) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varValidationRules(i, 0) & " is using an (x, 5) value that is not a valid control name for RequiredIf")
      End If
      
      'If using "RequiredIf", the array (x, 5)(1) must be an Array
      If Not VarType(varValidationRules(i, 5)(1)) = 8204 Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varValidationRules(i, 0) & " is not using a valid array in (x, 5)(1)")
      End If

    End If
  Next i
  
  ValidationRulesValidator_allRequiredIfValidationRulesAreValid = True
End Function




Public Function ValidationRulesValidator_indexZeroValuesValid(ByVal varIndexZero As Variant, ByRef frmForm As Object) As Boolean

  If IsNull(varIndexZero) Or varIndexZero = "" Then
    Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", "ControlName cannot be blank. The most likely cause of this Error is that the ValidationRules array size is incorrect")
  End If
  
  '(x, 0)
  'Raise Error is control name does not exist
  If Not Validator.FormControlHelper_controlExists(varIndexZero, frmForm) Then
    Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", "ControlName does not exist: " & varIndexZero)
  End If
  
  ValidationRulesValidator_indexZeroValuesValid = True
End Function

Public Function ValidationRulesValidator_indexOneValuesValid(ByRef varIndexZero As Variant, ByRef varIndexOne As Variant, ByRef varIndexTwo As Variant) As Boolean
    '(x, 1)
    If Not (VarType(varIndexOne) = vbBoolean) Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " required field must be set as a boolean")
    End If

    If varIndexOne = False And Not Validator.ArrayHelper_isInArray(varIndexTwo, Validator.Datatypes_getNullableDataTypes) Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " is a data-type that must always be required")
    End If
  
  ValidationRulesValidator_indexOneValuesValid = True
End Function

Public Function ValidationRulesValidator_indexTwoValuesValid(ByVal varIndexZero As Variant, ByVal varIndexTwo As Variant, ByVal varRecognizedDataTypes As Variant) As Boolean
    '(x, 2)
    'All data types must be recognized terms
    If Not Validator.ArrayHelper_isInArray(varIndexTwo, varRecognizedDataTypes) Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " has an invalid data type")
    End If
    
    ValidationRulesValidator_indexTwoValuesValid = True
End Function

Public Function ValidationRulesValidator_indexThreeValuesValid(ByVal varIndexZero As Variant, ByVal varIndexThree As Variant) As Boolean
    '(x, 3)
    'LocationInRange value must ALWAYS be set to Null by the user. It is used as a flag by the system
    If Not IsNull(varIndexThree) Then
      Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 3) must be set to Null")
    End If
    
    ValidationRulesValidator_indexThreeValuesValid = True
End Function


Public Function ValidationRulesValidator_indexFourValuesValid(ByVal varIndexZero As Variant, ByVal varIndexTwo As Variant, ByVal varIndexFour As Variant, ByVal varIndexFive As Variant, ByVal varDatatypesThatUseParameters As Variant) As Boolean

    '(x, 4)
    'All data types that use (x, 4) parameters
    If Not Validator.ArrayHelper_isInArray(varIndexTwo, varDatatypesThatUseParameters) Then
      If Not IsNull(varIndexFour) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be set to Null for data-type: " & varIndexTwo)
      End If
    End If
    
    If (varIndexTwo) = "StringInArray" Then
      If Not Validator.ArrayHelper_isArrayOfStrings(varIndexFour) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array of strings for data-type: " & varIndexTwo)
      End If
    End If

    If (varIndexTwo) = "AlphaText" Or (varIndexTwo) = "AlphanumericText" Or (varIndexTwo) = "SpecialCharacterText" Or (varIndexTwo) = "AllAllowableText" Or (varIndexTwo) = "IntegerInRange" Then
      'Must be an array of two elements
      If Not Validator.ArrayHelper_getSize(varIndexFour) = 2 Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array with only min and max for data-type: " & varIndexTwo)
      End If
      
      'Must be an array of two integers
      If Not ((VarType(varIndexFour(0)) = 2) Or Not (VarType(varIndexFour(1)) = 2)) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array with only min and max As Longs for data-type: " & varIndexTwo)
      End If
      
      'Max must be greater than or equal to min
      If Not (varIndexFour(1) >= varIndexFour(0)) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) max must be greater than or equal to min for data-type: " & varIndexTwo)
      End If
    End If
    
    If (varIndexTwo) = "DecimalInRange" Then
      'Must be an array of two elements
      If Not Validator.ArrayHelper_getSize(varIndexFour) = 2 Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array with only min and max for data-type: " & varIndexTwo)
      End If
      
      'Must be an array of two doubles or integers
      If Not (VarType(varIndexFour(0)) = 2) And Not (VarType(varIndexFour(0)) = 5) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array with only min and max as doubles for data-type: " & varIndexTwo)
      End If
      
      If Not (VarType(varIndexFour(1)) = 2) And Not (VarType(varIndexFour(1)) = 5) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) must be an array with only min and max as doubles for data-type: " & varIndexTwo)
      End If
      
      'Max must be greater than or equal to min
      If Not (varIndexFour(1) >= varIndexFour(0)) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 4) max must be greater than or equal to min for data-type: " & varIndexTwo)
      End If
    End If
    
    'All data types must be recognized terms for TwoDimDecimal datatypes
    If Not Validator.ArrayHelper_isInArray(varIndexTwo, Validator.Datatypes_getDatatypesThatHaveASecondDimVariable) Then
      If Not IsNull(varIndexFive) Then
        Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 5) must be set to Null for data-type: " & varIndexTwo)
      End If
    End If
    
    ValidationRulesValidator_indexFourValuesValid = True
End Function

     
Public Function ValidationRulesValidator_indexFiveValuesValid(ByVal varIndexZero As Variant, ByVal varIndexTwo As Variant, ByVal varIndexFive As Variant, ByVal varValidDatatypesForSecondDimVariable As Variant, ByVal varValidationRules As Variant) As Boolean
    '(x, 5) control must exist and be a OneDimInteger, OneDimDecimal, TwoDimDecimal, IntegerInRange or DecimalInRange (LocationInRange can be based on IntegerInRange to make ENUM possible)
    If (varIndexTwo = "TwoDimInteger") Or (varIndexTwo = "TwoDimDecimal") Then
      'The current field is a TwoDimDecimal. This will have a corresponding fieldname as string in (x, 5) that is the second dimension variable.
      'This field name given in (x, 5) needs to be a datatype that can be used as a second dimension variable with a TwoDimDecimal.
      'Iterate through all the control details to find the control detail for the field given in (x, 5). Then, compare the datatype of that field to the
      'array of valid datatypes for second dim variables, and use that to confirm whether the fieldname is allowed to be used as a second dimension variable
      'for the TwoDimDecimal.
      Dim i As Long
      For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
        If varValidationRules(i, 0) = varIndexFive Then 'Found the fieldname that is being requested for use as a second dimension variable
          If Not Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getValidDatatypesForSecondDimVariable) Then
            Call Validator.ErrorHandler_logAndRaiseError("Validation Rules", varIndexZero & " (x, 5) must be a valid numeric input field for data-type: " & varIndexTwo)
          End If
        End If
      Next i
    End If
    
    ValidationRulesValidator_indexFiveValuesValid = True
End Function

'
'ValidationRulesValidator_all1DimensionalVariableDetailsCorrectlyDefined
'
'@param variant varValidationRules
'@param variant var1DimValidationRules
'@return boolean
'
Public Function ValidationRulesValidator_all1DimensionalVariableDetailsCorrectlyDefined(ByVal varValidationRules As Variant, ByVal var1DimValidationRules As Variant, ByRef frmForm As Object) As Boolean

  'Check that all OneDimDecimal fields in varValidationRules match with var1DimValidationRules
  
  'Create an array of all 1DimVariableDetails names from var1DimValidationRules
  Dim arr1DimVariableNames1(999) As String
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(var1DimValidationRules) - 1)
    
    '
    'Check the mFirstDimensionName and intGrpInt before passing them into the Validator_Boundaries.getGrp function
    '
    
    'Raise Error if the DimensionName is Null or blank
    If IsNull(var1DimValidationRules(i, 0)) Or (var1DimValidationRules(i, 0) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("get1DimensionalGrpIntBasedOnVariable", "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
    
    'Raise Error if the DataDictionaryDimensionName is Null or blank
    If IsNull(var1DimValidationRules(i, 1)) Or (var1DimValidationRules(i, 1) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("get1DimensionalGrpIntBasedOnVariable", "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
    
    'Raise Error if the intGrpInt is Null or 0 or blank
    If IsNull(var1DimValidationRules(i, 2)) Or (var1DimValidationRules(i, 2) = 0) Then
      Call Validator.ErrorHandler_logAndRaiseError("get1DimensionalGrpIntBasedOnVariable", "Data Dictionary group number request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
          
    'Check that each pair is represented in the DataDictionary (The DataDictionary throws an error if there is no match found)
    Call Validator_Boundaries.getGrp(var1DimValidationRules(i, 1), var1DimValidationRules(i, 2))
  
    'Create an array of all ControlDetails names
    arr1DimVariableNames1(i) = var1DimValidationRules(i, 0)
  Next i
  
  'Iterate through ControlDetails and check that each is in the 1DimVariableDetails array
  Dim j As Long
  
  For j = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    
    'Raise Error if a control name is found in ControlDetails that is marked at OneDimDecimal but does not have a match in the var1DimValidationRules
    If ((varValidationRules(j, 2) = "OneDimDecimal") Or (varValidationRules(j, 2) = "OneDimInteger")) And Not (Validator.ArrayHelper_isInArray(varValidationRules(j, 0), arr1DimVariableNames1)) Then
      Call Validator.ErrorHandler_logAndRaiseError("all1DimensionalVariableDetailsCorrectlyDefined", varValidationRules(j, 0) & " data dictionary group not correctly defined in arr1DimensionalVariableDetails")
    End If

  Next j
  
  'Check that all OneDimDecimal fields invar1DimValidationRules match with varValidationRules
  
  'Create an array of all 1DimVariableDetails names from varValidationRules
  Dim arr1DimVariableNames2(999) As String
  Dim k As Long
  Dim counter As Long: counter = 0
  For k = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)

    If (varValidationRules(k, 2) = "OneDimDecimal") Or (varValidationRules(k, 2) = "OneDimInteger") Then
      'Create an array of all ControlDetails names
      arr1DimVariableNames2(counter) = varValidationRules(k, 0)
      counter = counter + 1
    End If
    
  Next k
  
  'Iterate through var1DimValidationRules and check that each is in the 1DimVariableDetails array
  Dim l As Long
  
  For l = 0 To (Validator.ArrayHelper_getSize(var1DimValidationRules) - 1)
    
    'Raise Error if a control name is found in ControlDetails that is marked at OneDimDecimal but does not have a match in the var1DimValidationRules
    If Not Validator.ArrayHelper_isInArray(var1DimValidationRules(l, 0), arr1DimVariableNames2) Then
      Call Validator.ErrorHandler_logAndRaiseError("all1DimensionalVariableDetailsCorrectlyDefined", var1DimValidationRules(l, 0) & " data dictionary group in arr1DimensionalVariableDetails was not correctly defined as OneDimDecimal in arrValidationRules")
    End If

  Next l
  
  ValidationRulesValidator_all1DimensionalVariableDetailsCorrectlyDefined = True
  
End Function



'
'ValidationRulesValidator_all2DimensionalVariableDetailsCorrectlyDefined
'
'@param variant varValidationRules
'@param variant var2DimValidationRules
'@return boolean
'
Public Function ValidationRulesValidator_all2DimensionalVariableDetailsCorrectlyDefined(ByVal varValidationRules As Variant, ByVal var2DimValidationRules As Variant, ByRef frmForm As Object) As Boolean

  'Check that all TwoDimDecimal fields in varValidationRules match with var2DimValidationRules
  
  'Create an array of all 2DimVariableDetails names from var2DimValidationRules
  Dim arr2DimVariableNames1(999) As String
  Dim i As Long
  Dim j As Long
  For i = 0 To (Validator.ArrayHelper_getSize(var2DimValidationRules) - 1)
    '
    'Check the mFirstDimensionName and intGrpInt before passing them into the Validator_Boundaries.getGrp function
    '
    'Raise Error if the mFirstDimensionName is Null or blank
    If IsNull(var2DimValidationRules(i, 0)) Or (var2DimValidationRules(i, 0) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("get2DimensionalGrpIntBasedOnVariable", "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
    
    'Raise Error if the mFirstDimensionDataDictionaryName is Null or blank
    If IsNull(var2DimValidationRules(i, 1)) Or (var2DimValidationRules(i, 1) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("get2DimensionalGrpIntBasedOnVariable", "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
    
    'Raise Error if the mSecondDimensionDataDictionaryName is Null or blank
    If IsNull(var2DimValidationRules(i, 2)) Or (var2DimValidationRules(i, 2) = "") Then
      Call Validator.ErrorHandler_logAndRaiseError("get2DimensionalGrpIntBasedOnVariable", "Data Dictionary first dimension name request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
    End If
      
    'Iterate through array of boundary values
    For j = 0 To (Validator.ArrayHelper_getSize(var2DimValidationRules(i, 3)) - 1)
      '
      'Check the mFirstDimensionName and intGrpInt before passing them into the Validator_Boundaries.getGrp function
      '
      'Raise Error if the intGrpInt is Null or 0 or blank
      If IsNull(var2DimValidationRules(i, 3)(j, 1)) Or (var2DimValidationRules(i, 3)(j, 1) = 0) Then
        Call Validator.ErrorHandler_logAndRaiseError("get2DimensionalGrpIntBasedOnVariable", "Data Dictionary group number request invalid. This may be cause by an incorrect sizing of the ValidationRules array. Please review the validation rules")
      End If
          
      'Check that each pair is represented in the DataDictionary (The DataDictionary throws an error if there is no match found)
      Call Validator_Boundaries.getGrp(var2DimValidationRules(i, 1), var2DimValidationRules(i, 3)(j, 1))
    Next j
    
    'Each Name can only appear once in the 2Dim validation array. If the var2DimValidationRules->Name is already in the list of boundary values, throw an Error.
    If Validator.ArrayHelper_isInArray(var2DimValidationRules(i, 0), arr2DimVariableNames1) Then
      Call Validator.ErrorHandler_logAndRaiseError("all2DimensionalVariableDetailsCorrectlyDefined", var2DimValidationRules(i, 0) & " appears more than once in the 2Dim validation array. Each (x, 0) 1stDim name can only be assigned to ONE set of boundary values")
    End If
    
    'Create an array of all ValidationRules names
    arr2DimVariableNames1(i) = var2DimValidationRules(i, 0)
  Next i
  
  'Iterate through ControlDetails and check that each is in the 2DimVariableDetails array
  Dim k As Long
  
  For k = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    
    'Raise Error if a control name is found in ControlDetails that is marked at TwoDimInteger or TwoDimDecimal but does not have a match in the var2DimValidationRules
    If ((varValidationRules(k, 2) = "TwoDimInteger") Or (varValidationRules(k, 2) = "TwoDimDecimal")) And Not (Validator.ArrayHelper_isInArray(varValidationRules(k, 0), arr2DimVariableNames1)) Then
      Call Validator.ErrorHandler_logAndRaiseError("all2DimensionalVariableDetailsCorrectlyDefined", varValidationRules(k, 0) & " data dictionary group not correctly defined in arrGrp and arrBoundaryValues")
    End If

  Next k
  
  'Check that all TwoDimDecimal fields in var2DimValidationRules match with varValidationRules
  
  'Create an array of all 2DimVariableDetails names from varValidationRules
  Dim arr2DimVariableNames2(999) As String
  Dim l As Long
  Dim counter As Long: counter = 0
  For l = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)

    If ((varValidationRules(l, 2) = "TwoDimInteger") Or (varValidationRules(l, 2) = "TwoDimDecimal")) Then
      'Create an array of all ControlDetails names
      arr2DimVariableNames2(counter) = varValidationRules(l, 0)
      counter = counter + 1
    End If
    
  Next l
  
  'Iterate through var2DimValidationRules and check that each is in the 2DimVariableDetails array
  Dim m As Long
  
  For m = 0 To (Validator.ArrayHelper_getSize(var2DimValidationRules) - 1)
    
    'Raise Error if 2ndDim control name does not exist
    If Not Validator.FormControlHelper_controlExists(var2DimValidationRules(m, 2), frmForm) Then
      Call Validator.ErrorHandler_logAndRaiseError("all2DimensionalVariableDetailsCorrectlyDefined", "ControlName does not exist: " & var2DimValidationRules(m, 1))
    End If
    
    'Raise error is 2ndDim control is not an allowable 2ndDim data-type
    'Get the data-type from the ControlDetails
    If Not Validator.ArrayHelper_isInArray(Validator.ValidationRulesMgr_getDatatypeByControlDetailsTitle(var2DimValidationRules(m, 2), varValidationRules), Validator.Datatypes_getValidDatatypesForSecondDimVariable) Then
      Call Validator.ErrorHandler_logAndRaiseError("all2DimensionalVariableDetailsCorrectlyDefined", var2DimValidationRules(m, 2) & " cannot be used as a 2nd Dim variable because it is not a valid 2nd Dim Datatype")
    End If
  
    'Raise Error if a control name is found in var2DimValidationRules but does not match with varValidationRules
    If Not Validator.ArrayHelper_isInArray(var2DimValidationRules(m, 0), arr2DimVariableNames2) Then
      Call Validator.ErrorHandler_logAndRaiseError("all2DimensionalVariableDetailsCorrectlyDefined", var2DimValidationRules(m, 0) & " data dictionary group in arr2DimensionalVariableDetails was not correctly defined as TwoDimDecimal in arrValidationRules")
    End If

  Next m
  
  ValidationRulesValidator_all2DimensionalVariableDetailsCorrectlyDefined = True
  
End Function



'
'ValidationRulesValidator_setRequiredWhenRequiredIfIsUsed
'
'If the required field is marked as "RequiredIf", check whether it is required
'If the RequiredIf conditions are met, set (x, 1) to True. Else set (x, 1) to False.
'
'@param variant array varValidationRules
'@param object frmForm
'
Public Sub ValidationRulesValidator_setRequiredWhenRequiredIfIsUsed(ByRef varValidationRules As Variant, ByRef frmForm)

  'Search for "RequiredIf" in (x, 1)
  Dim i As Long
   
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 1) = "RequiredIf") Then
      
      'Check if the value in the secondary field is in the array
      If Validator.ArrayHelper_isInArray(frmForm.Controls(varValidationRules(i, 5)(0)).Value, varValidationRules(i, 5)(1)) Then
        varValidationRules(i, 1) = True 'Conditions for RequiredIf were met. Set required to True
      Else
        varValidationRules(i, 1) = False 'Conditions for RequiredIf were NOT met. Set required to False
      End If
      
      varValidationRules(i, 5) = Null 'Set (x, 5) to null because it isn't needed any longer
    End If
  Next i

End Sub


'
'ValidationRulesValidator_setRequiredWhenRequiredIfNotIsUsed
'
'If the required field is marked as "RequiredIfNot", check whether it is required
'If the RequiredIfNot conditions are met, set (x, 1) to True. Else set (x, 1) to False.
'
'@param variant array varValidationRules
'@param object frmForm
'
Public Sub ValidationRulesValidator_setRequiredWhenRequiredIfNotIsUsed(ByRef varValidationRules As Variant, ByRef frmForm)

  'Search for "RequiredIfNot" in (x, 1)
  Dim i As Long
   
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    If (varValidationRules(i, 1) = "RequiredIfNot") Then
      
      'Check if the value in the secondary field is in the array
      If Validator.ArrayHelper_isInArray(frmForm.Controls(varValidationRules(i, 5)(0)).Value, varValidationRules(i, 5)(1)) Then
        varValidationRules(i, 1) = False 'Conditions for RequiredIfNot were met. Set required to False
      Else
        varValidationRules(i, 1) = True 'Conditions for RequiredIfNot were NOT met. Set required to True
      End If
      
      varValidationRules(i, 5) = Null 'Set (x, 5) to null because it isn't needed any longer
    End If
  Next i

End Sub












''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                        '
'VALIDATOR                                                               '
'                                                                        '
'This contains functions that manipulate and validate the form           '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'TABLE OF CONTENTS
' API
' - initialize
' - save
' Getters and Setters
' - getControlValue
' - setControlValue
' Validation
' - allFieldsValid
' - allRequiredFieldsCompleted
' - sanitizeTextInputs
' - allFieldsContainValidDataTypes
' - allStringInArrayInputsAreValidStrings
' - allAlphaTextInputsAreValidStrings
' - isValidAlphaText
' - allAlphanumericTextInputsAreValidStrings
' - isValidAlphanumericText
' - allSpecialCharacterTextInputsAreValidStrings
' - isValidSpecialCharacterText
' - allAllAllowableTextInputsAreValidStrings
' - isValidAllAllowableText
' - allIntegerInArrayInputsAreValidIntegers
' - allNumericInputsAreValidNumbers
' - allDateInputsAreValidDates
' - allBooleanInputsAreValidBooleans
' - allNumericValuesUsingDataDictionaryAreWithinValidRanges
' ControlDetailsMgr
' - getControlNames
' - getRequiredControlDetails
' - countControlDetailsOfType
' - countRequiredControlDetails
' LocationInRangeMgr
' - getDataDictionaryGrpFrom1DimensionalVariable
' - setLocationForAll1DimensionalVariables
' - get1DLocationInRange
' - get2DimensionalDataDictionaryNameAndGrpIntBasedOn2ndDimensionVariableValue
' - get2DimensionalGrpIntBasedOnVariable
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'API
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'Validator_initialize
'
'Set up the form fields. This is called by Form_Current() in form's code
'the each time the user moves to a different record.
'
'@param varArray varValidationRules
'@param object frmForm
'
Public Sub Validator_initialize(ByVal varValidationRules As Variant, ByRef frmForm As Object)

  'Call the reporter if running in test mode
  If Validator_Settings.ENVIRONMENT = "testing" Then
    Call Validator.Reporter_run(frmForm)
  End If

  Call Validator.UserInterfaceMgr_resetBackColors(Validator.ValidationRulesMgr_getControlNames(varValidationRules), frmForm)
End Sub

'
'Validator_validate
'
'Call the validation code and save the updates if the input is valid
'
'@param varArray varValidationRules
'@param varArray var1DimValidationRules
'@param varArray var2DimValidationRules
'@param object frmForm
'@return boolean
'
Public Function Validator_validate(ByVal varValidationRules As Variant, _
ByVal var1DimValidationRules As Variant, _
ByVal var2DimValidationRules As Variant, _
ByRef frmForm As Object) As Boolean

  'Catch any errors that were raised in the validation functions
  On Error GoTo ErrorHandler

    'Check the validation rules to ensure that the syntax has been correctly defined
    If Not Validator.ValidationRulesValidator_allValidationRulesCorrectlyDefined(varValidationRules, _
    var1DimValidationRules, var2DimValidationRules, frmForm) Then
      Validator_validate = False
      Exit Function
    End If
    
    'Confirm that the ENVIRONMENT variable is a valid string
    If Not Validator_Settings.ENVIRONMENT = "testing" And Not Validator_Settings.ENVIRONMENT = "production" And Not Validator_Settings.ENVIRONMENT = "deactivated" Then
      Call Validator.ErrorHandler_logAndRaiseError("VALIDATOR_Validate.validate", "Validator_Settings.ENVIRONMENT has not been correctly defined. It must be " & Chr(34) & "testing" & Chr(34) & ", " & Chr(34) & "production" & Chr(34) & " " & Chr(34) & "or deactivated" & Chr(34))
      Exit Function
    End If


    'If the validator is deactivated, display warning message and return True to bypass the Validator
    If Validator_Settings.ENVIRONMENT = "deactivated" Then
      MsgBox "WARNING: Validator has been deactivated. Access is running without using the Validator software. To activate the software, refer to module VALIDATOR_GlobalSettings ENVIRONMENT variable"
      Validator_validate = True
      Exit Function
    End If

    'If running in test mode, run the reporter and display a message to say you are in test mode
    If Validator_Settings.ENVIRONMENT = "testing" Then
      Call Validator.Reporter_run(frmForm)
      MsgBox "TESTING ENVIRONMENT" & vbNewLine & vbNewLine & "Validator is running in testing mode. This mode will validate all fields, including fields that are unedited. This allows you to test that the validation rules have been correctly defined." & vbNewLine & vbNewLine & "When you are ready to release the database to production, open VALIDATOR_GlobalSettings module and change the ENVIRONMENT variable value to production"
    End If


    'In the production environment, check whether this is a new record. If it is a new record, do not remove anything from the validation rules.
    'If it is NOT a new record (updating part of an old record), remove all unedited fields from the validation rules. This is necessary for situations
    'where users are editing some fields in historical data without having the correct values for all historical data for the record. Without this, the
    'required fields would force the user to enter values without knowing what those values should be.
    '
    'Note: frmForm.NewRecord is a Boolean value, not a record object
    If Validator_Settings.ENVIRONMENT = "production" Then
      If frmForm.NewRecord = False Then
        'If this is not a new record. Remove unedited fields
        varValidationRules = Validator.ValidationRulesMgr_getValidationRulesWithUneditedFieldsRemoved(varValidationRules, frmForm)
      End If
    End If

    'Validate the form inputs and save if all inputs are valid
    Validator_validate = Validator.Validator_allFieldsValid(varValidationRules, var1DimValidationRules, _
    var2DimValidationRules, frmForm)
    
    Call Validator.Logger_logDataInputWasValid("validate", "New data was valid")

    'If there are no errors, exit the function
    Exit Function
    
ErrorHandler:
    MsgBox "There was an error. " & vbNewLine & "Source: " & Err.Source & vbNewLine & "Description: " & Err.Description
End Function






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Validator_Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Validator_allFieldsValid
'
'Call the validation functions one after another. The function has three
'stages. If stages 2 or 3 the process is interrupted with a message box.
' 1, SET UP THE USER INTERFACE
'     a, Reset the Background Colors for all fields
' 2, MAKE THE INPUTS SAFE
'     a, Check whether the required fields are complete, and mark any
'        incomplete fields as RED.
'     b, Sanitize the text fields
'     c, Check the data types. For example, if the data must be a number,
'        confirm that the input is numeric
' 3, MAKE SURE ALL NUMERIC VALUES ARE IN RANGE
'     a, Validate the numeric ranges and display an OutOfBounds error or
'        an ORANGE Warning with a Yes/No dialog box
'
'@param varArray varValidationRulesArray
'@param varArray var1DimValidationRules
'@param varArray var2DimValidationRules
'@param Form frmForm
'@return Boolean
'
Public Function Validator_allFieldsValid(ByVal varValidationRules As Variant, _
  ByVal var1DimValidationRules As Variant, _
  ByVal var2DimValidationRules As Variant, _
  ByRef frmForm As Object) As Boolean
  
  Dim strFunctionName As String: strFunctionName = "Validator.Validator_allFieldsValid"
  
  Validator_allFieldsValid = True 'Initialize return value to true

  'Return true if there are no validation rules in the array
  If Validator.ArrayHelper_getSize(varValidationRules) = 0 Then
    Validator_allFieldsValid = True
    Exit Function
  End If

  'Create blank string for MsgBox notifications
  Dim strNotifications As String: strNotifications = ""
  
  '***Step 1: SET UP THE USER INTERFACE***
  Call Validator.UserInterfaceMgr_resetBackColors(Validator.ValidationRulesMgr_getControlNames(varValidationRules), frmForm)
  
  '***Step 2: MAKE THE INPUTS SAFE***
  '2a: Check whether the required fields are complete, and mark any incomplete fields as RED.
  If Validator.Validator_allRequiredFieldsCompleted(Validator.ValidationRulesMgr_getRequiredControlDetails(varValidationRules), frmForm, strNotifications) = False Then
    Call Validator.Logger_logValidationNotice(strFunctionName & ": allRequiredFieldsCompleted", strNotifications)
    MsgBox strNotifications
    Validator_allFieldsValid = False
    Exit Function 'Cancel the process
  End If

  '2b: Sanitize the text fields
  'Sanitization is disabled by default.
  'It is not needed because Access carries out its own sanitization. However, it can be used to
  'filter out unwanted characters.
  '
  'This function filters out disallowed characters as specified in the Validator.StringFormatter_sanitize
  'function. If enabled below, it will carry out sanitization on all data in all TextBox fields
  'that are editable. Sanitization is disabled for all TextBox fields that are Locked = TRUE or Enabled = FALSE.
  '
  'To enabled sanitization, uncomment the following line of code:
  'Call Validator.StringFormatter_sanitizeTextInputs(varValidationRules, frmForm)

  '2c: Check the data types. For example, if the data must be a number, confirm that the input is numeric
  If Validator.Validator_allFieldsContainValidDataTypes(varValidationRules, strNotifications, frmForm) = False Then
    Call Validator.Logger_logValidationNotice(strFunctionName & ": allFieldsContainValidDataTypes", strNotifications)
    MsgBox strNotifications
    Validator_allFieldsValid = False
    Exit Function
  End If

  '***Step 3: MAKE SURE ALL OneDimDecimal AND TwoDimDecimal VALUES ARE IN RANGE***
  If Validator.Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges(varValidationRules, var1DimValidationRules, _
  var2DimValidationRules, frmForm) = False Then
    Validator_allFieldsValid = False
    Exit Function
  End If
    
End Function

'
'Validator_allRequiredFieldsCompleted
'
'Check that all items in the RequiredControlNamesArray
'have been completed.
'Set BackColor to error if field is null
'
'@param variant varRequiredFields
'@param object frmForm
'@param string strNotifications
'@return Boolean
'
Public Function Validator_allRequiredFieldsCompleted(ByVal varRequiredFields As Variant, _
ByRef frmForm As Object, _
ByRef strNotifications As String) As Boolean

  Dim strFunctionName As String: strFunctionName = "Validator.Validator_allRequiredFieldsCompleted"

  Validator_allRequiredFieldsCompleted = True

  'Iterate through the array of required control details
  Dim i As Long
  
  For i = 0 To (Validator.ArrayHelper_getSize(varRequiredFields) - 1)

    'Validate inputs
    Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varRequiredFields(i, 0), frmForm)

    'If Null and required, set error BackColor
    If IsNull(frmForm.Controls(varRequiredFields(i, 0)).Value) Then
      Call Validator.UserInterfaceMgr_setControlBackColor(varRequiredFields(i, 0), "Error", frmForm)
      strNotifications = Validator.StringFormatter_addLineToBody(varRequiredFields(i, 0) & Validator.FIELD_REQUIRED, strNotifications)
      Validator_allRequiredFieldsCompleted = False
    End If
  Next i
  
  'Any require field was Null then add a header to the strNotifications
  If (Validator_allRequiredFieldsCompleted = False) Then
    strNotifications = Validator.StringFormatter_addHeading("MISSING REQUIRED FIELDS", strNotifications)
  End If
  
End Function


'
'Validator_allFieldsContainValidDataTypes
'
'Check that the data entered into the fields are valid
'data types
'
'@param string strNotifications
'@param variant varValidationRulesArray
'@param object frmForm
'@return Boolean
'
Public Function Validator_allFieldsContainValidDataTypes(ByVal varValidationRules As Variant, ByRef strNotifications As String, _
ByRef frmForm As Object) As Boolean

  Dim strFunctionName As String: strFunctionName = "Validator.Validator_allFieldsContainValidDataTypes"

  Validator_allFieldsContainValidDataTypes = True 'Set default return value to TRUE
  
  'StringInArray
  If Validator.StringValidationHelper_allStringInArrayInputsAreValidStrings(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'AlphaText
  If Validator.StringValidationHelper_allAlphaTextInputsAreValidStrings(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'AlphanumericText
  If Validator.StringValidationHelper_allAlphanumericTextInputsAreValidStrings(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'SpecialCharacterText
  If Validator.StringValidationHelper_allSpecialCharacterTextInputsAreValidStrings(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'AllAllowableText
  If Validator.StringValidationHelper_allAllAllowableTextInputsAreValidStrings(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'Date Validation
  If Validator.DateValidationHelper_allDateInputsAreValidDates(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False

  'IntegerInArray
  If Validator.NumberValidationHelper_allIntegerInArrayInputsAreValidIntegers(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'Integer, Double, OneDimDecimal and TwoDimDecimal
  If Validator.NumberValidationHelper_allNumericInputsAreValidNumbers(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'Boolean Validation
  If Validator.BooleanValidationHelper_allBooleanInputsAreValidBooleans(varValidationRules, frmForm, strNotifications) = False Then Validator_allFieldsContainValidDataTypes = False
  
  'If anything was added to the strNotifications message, then add a header to the
  'string with two line-breaks to leave some space between the heading and the notification message
  If Not strNotifications = "" Then strNotifications = Validator.StringFormatter_addHeading("INCORRECT DATA TYPES ENTERED", strNotifications)
  
End Function



'
'Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges
'
'Check that the data entered is within the range set in the DataDictionary
'data types
'
'@param variant varValidationRules
'@param variant var1DimValidationRules
'@param variant var2DimValidationRules
'@param object frmForm
'@return Boolean
'
Public Function Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges(ByVal varValidationRules As Variant, _
ByVal var1DimValidationRules As Variant, _
ByVal var2DimValidationRules As Variant, _
ByRef frmForm As Object) As Boolean

  Dim strFunctionName As String: strFunctionName = "Validator.Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges"
  
  Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges = True 'Set default return value to TRUE

  Dim strWarnings As String: strWarnings = "" 'strWarnings is the ORANGE region notification with the Yes/No dialog box.
  Dim strOutOfBounds As String: strOutOfBounds = "" 'strOutOfBounds is the RED region notification that immediately cancels the process
  
  'Validate the one-dimensional variables
  Call Validator.OneDimValidationMgr_setLocationForAll1DimensionalVariables(varValidationRules, var1DimValidationRules, frmForm, strWarnings, _
  strOutOfBounds)  'Set all 1-dimensional Variables
  
  'Validate the two-dimensional variables
  Dim i As Long
  For i = 0 To (Validator.ArrayHelper_getSize(varValidationRules) - 1)
    'Check whether the ControlDetails state that it is a 2-dimensional variable
    If (varValidationRules(i, 2) = "TwoDimInteger") Or (varValidationRules(i, 2) = "TwoDimDecimal") Then

      'Validate input
      Call Validator.FormControlHelper_raiseErrorIfControlDoesNotExist(strFunctionName, varValidationRules(i, 0), frmForm)
      
      'Raise Error is control name does not exist
      If IsNull(frmForm.Controls(varValidationRules(i, 5)).Value) Then
        Call Validator.ErrorHandler_logAndRaiseError("Invalid Input", varValidationRules(i, 5) & " cannot be blank if used as a 2nd Dimension variable for " & varValidationRules(i, 0))
      End If
    
      'If nullable and input is Null, do not validate.
      If (Validator.ArrayHelper_isInArray(varValidationRules(i, 2), Validator.Datatypes_getNullableDataTypes)) And IsNull(frmForm.Controls(varValidationRules(i, 0)).Value) Then
        'Do nothing
      Else
        'Datatype is not null. Validate
        'Variable is a 2-dimensional variable
        Call Validator.UserInterfaceMgr_setControlBackColorFromLocationInRange(varValidationRules(i, 0), _
        Validator.TwoDimValidationMgr_get2DimensionalGrpIntBasedOnVariable(varValidationRules(i, 0), frmForm.Controls(varValidationRules(i, 0)).Value, _
        varValidationRules(i, 5), frmForm.Controls(varValidationRules(i, 5)).Value, strWarnings, strOutOfBounds, _
        var2DimValidationRules), frmForm)
      End If

   End If
  Next i
  
  'If there was anything added to the strOutOfBounds notifications message, add the
  'header and two line-breaks. Display the "OUT OF BOUNDS" message box and set the
  'return value of the function to False.
  If Not strOutOfBounds = "" Then
  
    'Log that the validator displayed a notification
    Call Validator.Logger_logValidationNotice(Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges & ": OUT OF BOUNDS", strOutOfBounds)

    'Log that the validator cancelled the processes
    Call Validator.Logger_logProcessCancelledByValidator("Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges: OUT OF BOUNDS", strOutOfBounds)

    strOutOfBounds = Validator.StringFormatter_addHeading("OUT OF BOUNDS", strOutOfBounds)
    MsgBox strOutOfBounds
    Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges = False
    Exit Function
  End If
  
  'If there was anything added to the strWarnings notifications message, add the
  'header and two line-breaks. Prompt the user so check the warning(s) and ask them
  'to click Yes or No. If the user knows for sure that the values are correct, they click
  'YES. If they believe that the value they entered might be wrong, they click NO to cancel the
  'save-process.
  'If the user clicks NO, the return value for the function is set to False
  Dim boolCancelledByUserBecauseOfWarning As Boolean: boolCancelledByUserBecauseOfWarning = False
  If Not strWarnings = "" Then
  
    'Log that the validator displayed a notification
    Call Validator.Logger_logValidationNotice("allNumericValuesUsingDataDictionaryAreWithinValidRanges : WARNING ", strWarnings)
    
    strWarnings = "WARNING" & vbNewLine & vbNewLine & strWarnings & vbNewLine & "Are you sure these values are correct?"
    strWarnings = strWarnings & vbNewLine & "Click YES to save. Click NO to cancel"
    boolCancelledByUserBecauseOfWarning = Validator.MsgBoxHelper_askYesNoQuestion("Warning: Value in the Orange Zone", strWarnings, _
    Validator.StringFormatter_addHeading("CANCELLED", "The new data was NOT saved to the database. Please change the incorrect values and try again"))
    
    If Not (boolCancelledByUserBecauseOfWarning) Then
      'Log that the user cancelled the processes using the Yes/No dialog box
      Call Validator.Logger_logProcessCancelledByValidator("allNumericValuesUsingDataDictionaryAreWithinValidRanges : WARNING ", "User cancelled the save process due to the warning " & strWarnings)
    End If
    
    Validator_allNumericValuesUsingDataDictionaryAreWithinValidRanges = boolCancelledByUserBecauseOfWarning
  End If

End Function







