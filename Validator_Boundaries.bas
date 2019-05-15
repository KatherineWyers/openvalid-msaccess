Attribute VB_Name = "Validator_Boundaries"
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
                                                                       

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                                               '
'Boundaries - Custom Boundary values                                                                            '
'                                                                                                               '
'Naming Conventions for boundary values:                                                                        '
'                                                                                                               '
'Each set of ranges are grouped (ie GRP1)                                                                       '
'Each set has four boundary values. All values are stored as a Decimal (AKA Floating point, AKA Decimal value)   '
'                                                                                                               '
'Each boundary MUST be at least 0.2 apart from each other for Decimal values, and 2 apart for Integers           '
'[---RED---] BOUNDARY1 [---ORANGE---] BOUNDARY2 [---GREEN---] BOUNDARY3 [---ORANGE---] BOUNDARY4 [---RED---]    '
'                                                                                                               '
'All boundaries MUST be assigned according to the following formula                                             '
'Ensure that the "or-equals-to" is on the correct side of your values                                           '
'                                                                                                               '
'x < BOUNDARY1 // RED                                                                                           '
'BOUNDARY1 <= x < BOUNDARY2 // ORANGE                                                                           '
'BOUNDARY2 <= x < BOUNDARY3 // GREEN                                                                            '
'BOUNDARY3 <= x < BOUNDARY4 // ORANGE                                                                           '
'BOUNDARY4 <= x // RED                                                                                          '
'                                                                                                               '
'                                                                                                               '
'Note about VBA Arrays Sizes                                                                                    '
'                                                                                                               '
'The size of the array starts its counting at zero. So Dim arrGrp(3) As Decimal is creating an array with FOUR   '
'elements, not three elements. This is really confusing. (It confused me!)                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'TABLE OF CONTENTS
'
'Age
'HeartRate
'Temperature
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function getGrp(ByVal strTitle As String, ByVal intGrpNum As Long) As Double()
  Dim arrGrp(3) As Double 'Create an array with FOUR elements
  Dim grpFound As Boolean: grpFound = False
  
  If strTitle = "Weight" And intGrpNum = 1 Then
    arrGrp(0) = 0
    arrGrp(1) = 0.5
    arrGrp(2) = 6
    arrGrp(3) = 15
    grpFound = True
  End If
  
  If strTitle = "Temp" And intGrpNum = 1 Then
    arrGrp(0) = 0
    arrGrp(1) = 30
    arrGrp(2) = 40
    arrGrp(3) = 60
    grpFound = True
  End If
  
  If strTitle = "HCT" And intGrpNum = 1 Then
    arrGrp(0) = 0
    arrGrp(1) = 35
    arrGrp(2) = 75
    arrGrp(3) = 100
    grpFound = True
  End If
   
  
  
  If strTitle = "Age" And intGrpNum = 1 Then
    arrGrp(0) = 0
    arrGrp(1) = 1
    arrGrp(2) = 75
    arrGrp(3) = 130
    grpFound = True
  End If

  If strTitle = "OneDimDecimalDemo" And intGrpNum = 1 Then
    arrGrp(0) = 15.6
    arrGrp(1) = 17.2
    arrGrp(2) = 28.6
    arrGrp(3) = 30
    grpFound = True
  End If

  If strTitle = "HeartRate" And intGrpNum = 1 Then
    arrGrp(0) = 10
    arrGrp(1) = 110
    arrGrp(2) = 140
    arrGrp(3) = 150
    grpFound = True
  End If

  If strTitle = "HeartRate" And intGrpNum = 2 Then
    arrGrp(0) = 20
    arrGrp(1) = 120
    arrGrp(2) = 160
    arrGrp(3) = 200
    grpFound = True
  End If

  If strTitle = "HeartRate" And intGrpNum = 3 Then
    arrGrp(0) = 30
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 130
    grpFound = True
  End If

  If strTitle = "HeartRate" And intGrpNum = 4 Then
    arrGrp(0) = 40
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 135
    grpFound = True
  End If

  If strTitle = "HeartRate" And intGrpNum = 5 Then
    arrGrp(0) = 50
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 140
    grpFound = True
  End If
  

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 1 Then
    arrGrp(0) = 10
    arrGrp(1) = 110
    arrGrp(2) = 140
    arrGrp(3) = 150
    grpFound = True
  End If

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 2 Then
    arrGrp(0) = 20
    arrGrp(1) = 120
    arrGrp(2) = 160
    arrGrp(3) = 200
    grpFound = True
  End If

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 3 Then
    arrGrp(0) = 30
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 130
    grpFound = True
  End If

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 4 Then
    arrGrp(0) = 40
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 135
    grpFound = True
  End If

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 5 Then
    arrGrp(0) = 50
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 140
    grpFound = True
  End If

  If strTitle = "TwoDimDecimalDemo" And intGrpNum = 6 Then
    arrGrp(0) = 50
    arrGrp(1) = 85
    arrGrp(2) = 90
    arrGrp(3) = 150
    grpFound = True
  End If
  
  If strTitle = "txtTemperature" And intGrpNum = 1 Then
    arrGrp(0) = 25.3
    arrGrp(1) = 36.2
    arrGrp(2) = 38.1
    arrGrp(3) = 60.6
    grpFound = True
  End If
  
  If grpFound = False Then
    Call Validator.ErrorHandler_logAndRaiseError("getGrp", "ControlName and GrpInt combination does not exist in data dictionary: " & strTitle & " and " & intGrpNum)
  End If

  getGrp = arrGrp
End Function
