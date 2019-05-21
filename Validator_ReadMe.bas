Attribute VB_Name = "VALIDATOR_README"
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
                                                                      
'
'VERSION HISTORY
'

'Version 1.06
'Bugfix: FormControlHelper function-name fixed

'Version 1.05
'Bugfix: Added dbSeeChanges to AccessQuery function to fix bug when SQL-server with autoincrement is used with Access stored query
'Restructured the Validator_Settings file

'Version 1.04
'Condensed modules into Validator module

'Version 1.03
'Added the RequiredIfNot option

'Version 1.02
'Disable the Reporter.Run function for all environments that are not "testing"

'Version 1.01
'Bugfix: Only remove unedited fields if the form is running in production mode and is not a NewRecord

'Version 1.00
'System tested. First release to production

'Version 0.60
'Changed naming convention. Real numbers are all referred to as Decimal (DecimalInRange, DecimalInArray, OneDimDecimal, TwoDimDecimal)

'Version 0.59
'Bugfix: ListBox added to allowable object-types
'Added support for Byte integer datatype
'Changed BackColor Error and Warning colors for better readibility

'Version 0.58
'Bugfix: Handle Multi-select ComboBox. Add arrayable checks for all datatypes

'Version 0.57
'Bugfix: Handle Multi-select ComboBox as array of strings or integers

'Version 0.56
'Bugfix: isAlpha, isAlphanumeric, isSpecialCharacterText, isAllAllowableText: All return True is input is zero-length string

'Version 0.55
'Added "deactivated" to VALIDATOR_GlobalSettings ENVIRONMENT variable. This allows the Validator software to be bypassed if necessary
'Bugfix: Changes all Integer parameters to Long to avoid overflow for inputs above 32,767

'Version 0.54
'Added new allowable datatypes: Single, Decimal

'Version 0.53
'Added VALIDATOR_AccessQueryMgr.getArrayOfIntegersFromAccessQuery to allow users to specify which query column to use
'Updated VALIDATOR_AccessQueryMgr.getArrayOfStringsFromAccessQuery to allow users to specify which query column to use

'Version 0.52
'Added new datatype AllAllowableText. This blacklists the Chr(10) and Chr(13) but allows all other characters. Required for Thai and Burmese fonts.
'Added txtAllAllowableText field to TestForm
'Added Validator_Settings.ENVIRONMENT variable and default installation to "testing" mode
'Added : (colon) character to SpecialCharacterText
