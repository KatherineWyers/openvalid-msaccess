VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_VALIDATOR_frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
                                                                      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Welcome to the Validator. There are three areas that you
'need to edit in order to apply the validator to your Access
'database application.
'
'1, ValidationRules Array
'This is a list (an 'array') that contains each of the input fields
'in your form. This is where you tell the validator what kind of validation
'should be applied to each field. For example, if the input field is a
'textbox where the user inputs a patient's name, Validator will check
'that the text is not too long or too short.
'
'2, One-dimensional Dictionary Rules
'This is an array with details about how to validate one-dimensional
'numeric inputs. It that tells Validator which of the boundaries to
'use in the Data Dictionary. This lets you choose which input values
'give the RED and ORANGE warnings, and which input values are allowed.
'
'3, Two-dimensional Dictionary Rules
'This is an array with details about how to validate two-dimensional
'numeric inputs. It that tells Validator which of the boundaries to
'use in the Data Dictionary, and it also tells Validator which
'other input value the boundary group choice is based on. For example,
'if HeartRate is based on Age, this lets you choose which values of
'Age determine which HeartRate boundaries are selected in the Data Dictionary.
'This lets you choose which input values give the RED and ORANGE warnings,
'and which input values are allowed.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ValidationRules Array
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tutorial: How to use this function                                      '
'                                                                        '
'getValidationRules()                                                '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add a new block for each input field.                                   '
'When you add a new block, you must also increase the array index (X, 0) '
'by 1. So if the last block was (2, 0), (2, 1), (2, 2)... then your new  '
'block will be (3, 0), (3, 1), (3, 2) etc...                             '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'getValidationRules()
'
'Array containing the validation rules for the input field
'
'
'DATA TYPES
'
'STRINGS/TEXT
'StringInArray
'AlphaText
'AlphanumericText
'
'DATES
'Date
'
'NUMBERS
'IntegerInRange (min and max)
'DecimalInRange (min and max)
'OneDimDecimal
'TwoDimDecimal
'
'BOOLEAN
'Boolean
'
'
'
'(x, 0) = strTitle As String
'(x, 1) = boolRequired As Boolean
'(x, 2) = strType As String {"Plaintext", "Numeric"}
'(x, 3) = intLocationInRange As Variant (Integer if set. Default to Null for non-numeric data)
'(x, 4) = validation parameters array (AlphaText and Alphanumeric text is two integers for length. StringInArray is an array of strings. Else Null)
'(x, 5) = strSecondDimensionVariableName As String. If (x, 2) = "TwoDimDecimal" then this is the name of the 2nd dimension variable it is based on. (Default set to Null for 1Dimension numerics and for non-numerics)
'(x, 6) = boolHasBeenUpdated As Variant (Boolean if set. Default to Null)
'
'@return 2D array of form validation rules
'
Public Function getValidationRules() As Variant()
  Dim arrValidationRules(21, 5) As Variant
  
  'txtFullName
  arrValidationRules(0, 0) = "txtFullName"
  arrValidationRules(0, 1) = True
  arrValidationRules(0, 2) = "AlphaText"
  arrValidationRules(0, 3) = Null
  arrValidationRules(0, 4) = Array(3, 20) 'Min length and max length
  arrValidationRules(0, 5) = Null

  'txtSpecialCharText
  arrValidationRules(1, 0) = "txtSpecialCharText"
  arrValidationRules(1, 1) = True
  arrValidationRules(1, 2) = "SpecialCharacterText"
  arrValidationRules(1, 3) = Null
  arrValidationRules(1, 4) = Array(3, 20) 'Min length and max length
  arrValidationRules(1, 5) = Null

  'txtAllAllowableText
  arrValidationRules(2, 0) = "txtAllAllowableText"
  arrValidationRules(2, 1) = True
  arrValidationRules(2, 2) = "AllAllowableText"
  arrValidationRules(2, 3) = Null
  arrValidationRules(2, 4) = Array(3, 20) 'Min length and max length
  arrValidationRules(2, 5) = Null

  'txtAge
  arrValidationRules(3, 0) = "txtAge"
  arrValidationRules(3, 1) = True
  arrValidationRules(3, 2) = "OneDimInteger"
  arrValidationRules(3, 3) = Null
  arrValidationRules(3, 4) = Null
  arrValidationRules(3, 5) = Null

  'txtOneDimDecimal
  arrValidationRules(4, 0) = "txtOneDimDecimal"
  arrValidationRules(4, 1) = True
  arrValidationRules(4, 2) = "OneDimDecimal"
  arrValidationRules(4, 3) = Null
  arrValidationRules(4, 4) = Null
  arrValidationRules(4, 5) = Null
  
  'txtHeartRate
  arrValidationRules(5, 0) = "txtHeartRate"
  arrValidationRules(5, 1) = True
  arrValidationRules(5, 2) = "TwoDimInteger"
  arrValidationRules(5, 3) = Null
  arrValidationRules(5, 4) = Null
  arrValidationRules(5, 5) = "txtAge"
  
  'txtTwoDimDecimal
  arrValidationRules(6, 0) = "txtTwoDimDecimal"
  arrValidationRules(6, 1) = True
  arrValidationRules(6, 2) = "TwoDimDecimal"
  arrValidationRules(6, 3) = Null
  arrValidationRules(6, 4) = Null
  arrValidationRules(6, 5) = "txtAge"
  
  'txtDateOfBirth
  arrValidationRules(7, 0) = "txtDateOfBirth"
  arrValidationRules(7, 1) = True
  arrValidationRules(7, 2) = "Date"
  arrValidationRules(7, 3) = Null
  arrValidationRules(7, 4) = Null
  arrValidationRules(7, 5) = Null
  
  'cmbOptionOneOrTwo
  arrValidationRules(8, 0) = "cmbOptionOneOrTwo"
  arrValidationRules(8, 1) = True
  arrValidationRules(8, 2) = "StringInArray"
  arrValidationRules(8, 3) = Null
  arrValidationRules(8, 4) = Array("Option 1", "Option 2") 'Create the array of possible string inputs
  arrValidationRules(8, 5) = Null
  
  'cmbStringInArrayRequiredIf
  arrValidationRules(9, 0) = "cmbStringInArrayRequiredIf"
  arrValidationRules(9, 1) = "RequiredIf"
  arrValidationRules(9, 2) = "StringInArray"
  arrValidationRules(9, 3) = Null
  arrValidationRules(9, 4) = Array("This", "Is", "Required", "If", "OptionOneOrTwo", "Is", "Set", "As", "Option 1")
  arrValidationRules(9, 5) = Array("cmbOptionOneOrTwo", Array("Option 1")) 'Nested array with ("fieldName", Array("array", "of", "values", "and", 1, 2, 3, Null))

  'txtAlphanumeric
  arrValidationRules(10, 0) = "txtAlphanumeric"
  arrValidationRules(10, 1) = True
  arrValidationRules(10, 2) = "AlphanumericText"
  arrValidationRules(10, 3) = Null
  arrValidationRules(10, 4) = Array(3, 20) 'Min length and max length
  arrValidationRules(10, 5) = Null
  
  'cmbIntInArray
  arrValidationRules(11, 0) = "cmbIntInArray"
  arrValidationRules(11, 1) = True
  arrValidationRules(11, 2) = "IntegerInArray"
  arrValidationRules(11, 3) = Null
  arrValidationRules(11, 4) = Array(1, 2, 9, 12)
  arrValidationRules(11, 5) = Null

  'cmbIntInArrayRequiredIf
  arrValidationRules(12, 0) = "cmbIntInArrayRequiredIf"
  arrValidationRules(12, 1) = "RequiredIf"
  arrValidationRules(12, 2) = "IntegerInArray"
  arrValidationRules(12, 3) = Null
  arrValidationRules(12, 4) = Array(0, 1, 3, 7, 9)
  arrValidationRules(12, 5) = Array("cmbIntInArray", Array(1, 2)) 'Nested array with ("fieldName", Array("array", "of", "values", "and", 1, 2, 3, Null))

  'txtIntInRange
  arrValidationRules(13, 0) = "txtIntInRange"
  arrValidationRules(13, 1) = True
  arrValidationRules(13, 2) = "IntegerInRange"
  arrValidationRules(13, 3) = Null
  arrValidationRules(13, 4) = Array(1, 10) 'Min value and max value
  arrValidationRules(13, 5) = Null
  
  'txtDecimalInRange
  arrValidationRules(14, 0) = "txtDecimalInRange"
  arrValidationRules(14, 1) = True
  arrValidationRules(14, 2) = "DecimalInRange"
  arrValidationRules(14, 3) = Null
  arrValidationRules(14, 4) = Array(5, 12.4) 'Min value and max value
  arrValidationRules(14, 5) = Null
  
  'cbxBoolean
  arrValidationRules(15, 0) = "cbxBoolean"
  arrValidationRules(15, 1) = True
  arrValidationRules(15, 2) = "Boolean"
  arrValidationRules(15, 3) = Null
  arrValidationRules(15, 4) = Null
  arrValidationRules(15, 5) = Null
  
  'cmbStringInArrayFromDbQry
  arrValidationRules(16, 0) = "cmbStringInArrayFromDbQry"
  arrValidationRules(16, 1) = True
  arrValidationRules(16, 2) = "StringInArray"
  arrValidationRules(16, 3) = Null
  arrValidationRules(16, 4) = Validator.AccessQueryMgr_getArrayOfStringsFromAccessQuery("VALIDATOR_qryComboboxList", "listItem")
  arrValidationRules(16, 5) = Null
  
  'cmbIntegerInArrayFromDbQry
  arrValidationRules(17, 0) = "cmbIntegerInArrayFromDbQry"
  arrValidationRules(17, 1) = True
  arrValidationRules(17, 2) = "IntegerInArray"
  arrValidationRules(17, 3) = Null
  arrValidationRules(17, 4) = Validator.AccessQueryMgr_getArrayOfIntegersFromAccessQuery("VALIDATOR_qryComboboxList", "listCode")
  arrValidationRules(17, 5) = Null
  
  'lstListBoxInput
  arrValidationRules(18, 0) = "lstListBoxInput"
  arrValidationRules(18, 1) = True
  arrValidationRules(18, 2) = "IntegerInArray"
  arrValidationRules(18, 3) = Null
  arrValidationRules(18, 4) = Validator.AccessQueryMgr_getArrayOfIntegersFromAccessQuery("VALIDATOR_qryComboboxList", "listCode")
  arrValidationRules(18, 5) = Null
  
  'txtByte
  arrValidationRules(19, 0) = "txtByte"
  arrValidationRules(19, 1) = True
  arrValidationRules(19, 2) = "IntegerInRange"
  arrValidationRules(19, 3) = Null
  arrValidationRules(19, 4) = Array(0, 255)
  arrValidationRules(19, 5) = Null
  

  'txtDummyFieldForRequiredIfNot
  arrValidationRules(20, 0) = "txtDummyFieldForRequiredIfNot"
  arrValidationRules(20, 1) = False
  arrValidationRules(20, 2) = "AllAllowableText"
  arrValidationRules(20, 3) = Null
  arrValidationRules(20, 4) = Array(0, 255) 'Min length and max length
  arrValidationRules(20, 5) = Null

  'txtAlphanumericRequiredIfNot
  arrValidationRules(21, 0) = "txtAlphanumericRequiredIfNot"
  arrValidationRules(21, 1) = "RequiredIfNot"
  arrValidationRules(21, 2) = "AlphanumericText"
  arrValidationRules(21, 3) = Null
  arrValidationRules(21, 4) = Array(0, 255)
  arrValidationRules(21, 5) = Array("txtDummyFieldForRequiredIfNot", Array(Null)) 'Nested array with ("fieldName", Array("array", "of", "values", "and", 1, 2, 3, Null))
  
  getValidationRules = arrValidationRules
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'One-dimensional Dictionary Rules
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tutorial: How to use this function                                      '
'                                                                        '
'get1DimDictionaryRules()                                                '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This is where you choose which Data Dictionary boundary values to use   '
'for the input.                                                          '
'                                                                        '
'For example, ages are one-dimensional variables because the boundary    '
'values you choose are not based on what the other inputs are.           '
'                                                                        '
'We will set the txtAge input as the txtAge group 1 boundaries that are  '
'in the Validator_Boundaries.                                                  '
'                                                                        '
'  arr1DimensionalVariableDetails(0, 0) = "txtAge"                       '
'  arr1DimensionalVariableDetails(0, 1) = 1                              '
'                                                                        '
'Next, we will set the Temperature boundary values as temperature group 2'
'Notice how the left argument (input-value) has increased by one. You    '
'increment the left argument by one with every new input.                '
'The right arguments are always 0 and 1.                                 '
'                                                                        '
'  arr1DimensionalVariableDetails(1, 0) = "txtTemperature"               '
'  arr1DimensionalVariableDetails(1, 1) = 2                              '
'                                                                        '
'When you add a new block, you must also increase the array index (X, 0) '
'by 1. So if the last block was (2, 0), (2, 1), (2, 2)... then your new  '
'block will be (3, 0), (3, 1), (3, 2) etc...                             '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'
'get1DimDictionaryRules()
'
'Array of dictionary groups for 1 dimensional variables
'
'(x, 0) = strControlName As String
'(x, 1) = strDataDictionaryName As String
'(x, 2) = intGrp As Integer
'
'@return 2D array of form control details
'
Public Function get1DimDictionaryRules() As Variant()
  Dim arr1DimDictionaryRules(1, 2) As Variant

  'txtAge
  arr1DimDictionaryRules(0, 0) = "txtAge"
  arr1DimDictionaryRules(0, 1) = "Age"
  arr1DimDictionaryRules(0, 2) = 1 'Use data dictionary age Grp1
  
  'txtAge
  arr1DimDictionaryRules(1, 0) = "txtOneDimDecimal"
  arr1DimDictionaryRules(1, 1) = "OneDimDecimalDemo"
  arr1DimDictionaryRules(1, 2) = 1 'Use data dictionary age Grp1

  get1DimDictionaryRules = arr1DimDictionaryRules
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Two-dimensional Dictionary Rules
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tutorial: How to use this function                                      '
'                                                                        '
'get2DimDictionaryRules()                                                '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This function lets you choose which sets of two variables are used      '
'to determine the boundary values that are stored in the Validator_Boundaries. '
'                                                                        '
'For example, if you are dealing with HeartRate based on Age,            '
'txtHeartRate is the first variable and txtAge is the second variable.   '
'These are stored as (x, 0) and (x, 1) where x is the index of the       '
'element in the array. Next, you add the Age values where you want the   '
'boundaries to be.                                                       '
'                                                                        '
'For example, we will implement the following                            '
'                                                                        '
'txtAge < 15 ==> Use HeartRateGrp 1                                      '
'15 <= txtAge < 30 ==> Use HeartRateGrp 2                                '
'30 <= txtAge ==> Use HeartRateGrp3                                      '
'                                                                        '
'First we set the first boundary value                                   '
'arrBoundaryValues(0, 0) = 15                                            '
'arrBoundaryValues(0, 1) = 1                                             '
'                                                                        '
'Next, we set the second boundary value                                  '
'Notice how the argument (input value) on the left has increased         '
'by one. Every new boundary value will get a left-value that is          '
'one higher. The right values are ALWAYS 0 and 1.                        '
'arrBoundaryValues(1, 0) = 30                                            '
'arrBoundaryValues(1, 1) = 2                                             '
'                                                                        '
'Finally, we say what happens when the input value is higher than        '
'the highest boundary value.                                             '
'Notice how there is a Null value being entered. The final boundary      '
'always takes a Null as the first value.                                 '
'                                                                        '
'arrBoundaryValues(2, 0) = Null                                          '
'arrBoundaryValues(2, 1) = 3                                             '
'Array of dictionary groups for 1 dimensional variables                  '
'                                                                        '
'When you add a new block, you must also increase the array index (X, 0) '
'by 1. So if the last block was (2, 0), (2, 1), (2, 2)... then your new  '
'block will be (3, 0), (3, 1), (3, 2) etc...                             '
'                                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'get2DimDictionaryRules()
'
'(x, 0) = strFormNameOfFirstVariable As String
'(x, 1) = strDataDictionaryNameOfFirstVariable As String
'(x, 2) = strFormNameOfSecondVariable As String
'(x, 3) = arrGrp As Array: The Boundary Values as an Array
'
'Boundary Values as an Array
'(y, 0) = intBoundary1
'(y, 1) = intDataDictionaryGrp
'
'@return 2D array of form control details
'
Public Function get2DimDictionaryRules() As Variant()
  Dim arrBoundaryValues() As Variant

  Dim arr2DimDictionaryRules(1, 3) As Variant
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'txtHeartRate based on txtAge
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Array with (x, 0) as boundary value, (x, 1) as data dictionary group
  ReDim arrBoundaryValues(4, 1)
  
  arr2DimDictionaryRules(0, 0) = "txtHeartRate"
  arr2DimDictionaryRules(0, 1) = "HeartRate"
  arr2DimDictionaryRules(0, 2) = "txtAge"
    'Array of boundaries
    'Boundary 1
    arrBoundaryValues(0, 0) = 10
    arrBoundaryValues(0, 1) = 1
  
    'Boundary 2
    arrBoundaryValues(1, 0) = 20
    arrBoundaryValues(1, 1) = 2
  
    'Boundary 3
    arrBoundaryValues(2, 0) = 30
    arrBoundaryValues(2, 1) = 3
  
    'Boundary 4
    arrBoundaryValues(3, 0) = 40
    arrBoundaryValues(3, 1) = 4
  
    'Final Boundary Value
    arrBoundaryValues(4, 0) = Null
    arrBoundaryValues(4, 1) = 5 'The final boundary value is always ignored. Only the GrpNum matters
  
  'Assign array of boundaries to third element in first dimension
  arr2DimDictionaryRules(0, 3) = arrBoundaryValues
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'txtTwoDimDecimal based on txtAge
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Array with (x, 0) as boundary value, (x, 1) as data dictionary group
  ReDim arrBoundaryValues(6, 1)
  
  arr2DimDictionaryRules(1, 0) = "txtTwoDimDecimal"
  arr2DimDictionaryRules(1, 1) = "TwoDimDecimalDemo"
  arr2DimDictionaryRules(1, 2) = "txtAge"
    'Array of boundaries
    'Boundary 1
    arrBoundaryValues(0, 0) = 5
    arrBoundaryValues(0, 1) = 1
  
    'Boundary 2
    arrBoundaryValues(1, 0) = 30
    arrBoundaryValues(1, 1) = 2
  
    'Boundary 3
    arrBoundaryValues(2, 0) = 45
    arrBoundaryValues(2, 1) = 3
  
    'Boundary 4
    arrBoundaryValues(3, 0) = 100
    arrBoundaryValues(3, 1) = 4
  
    'Boundary 5
    arrBoundaryValues(4, 0) = 120
    arrBoundaryValues(4, 1) = 3
  
    'Boundary 6
    arrBoundaryValues(5, 0) = 125
    arrBoundaryValues(5, 1) = 2

    'Final Boundary Value
    arrBoundaryValues(6, 0) = Null
    arrBoundaryValues(6, 1) = 5 'The final boundary value is always ignored. Only the GrpNum matters
  
  'Assign array of boundaries to third element in first dimension
  arr2DimDictionaryRules(1, 3) = arrBoundaryValues
  
  'return the full array
  get2DimDictionaryRules = arr2DimDictionaryRules
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'DO NOT EDIT ANYTHING BELOW THIS HERE
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Form_Current
'
'This triggers each time the form loads or the user steps through
'next record or previous record
'
Private Sub Form_Current()
  Call Validator.Validator_initialize(getValidationRules, Me)
End Sub

'
'Form_BeforeUpdate
'
'This triggers whenever the user tries to save new
'data to the database
'
Private Sub Form_BeforeUpdate(Cancel As Integer)
  If Validator.Validator_validate(getValidationRules, get1DimDictionaryRules, get2DimDictionaryRules, Me) = False Then
    Cancel = True
  End If
End Sub


