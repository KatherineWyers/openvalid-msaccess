Attribute VB_Name = "Validator_Settings"
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

'*****************
'ENVIRONMENT
'*****************

'Testing
'- Validate all fields, including unedited fields
'- Print form layout to the Immediate Windows (Ctrl+G)
'
'Production
'- Validate only the edited fields
'
'Deactivated
'- Bypass the Validator software
'- Display a warning message stating that the software in being bypassed

'Set environment
Public Const ENVIRONMENT As String = "testing" '"testing", "production", "deactivated"



'*****************
'LOGGING
'*****************

'Logging directory
Public Const CUSTOM_DIRECTORY_PATH As String = "" 'Leave as an empty string to default the logging to the current project directory
'Public Const CUSTOM_DIRECTORY_PATH As String = "C:\Users\katherine\Desktop\validator"'Sample custom directory path

'Optional custom subdirectory. Option appends a subdirectory to the directory path (either custom or default).
Public Const CUSTOM_SUBDIRECTORY_NAME As String = "logs" 'Leave as an empty string to default the logging to the current project directory
'Public Const CUSTOM_SUBDIRECTORY_NAME As String = "logs" 'Sample custom subdirectory name

'Filenames for logging
Public Const ERROR_LOG_FILENAME As String = "log.error.txt"
Public Const VALIDATION_NOTICE_LOG_FILENAME As String = "log.validation-notification.txt"
Public Const SUCCESSFUL_SAVE_LOG_FILENAME As String = "log.successful-save.txt"
Public Const PROCESS_CANCELLED_BY_VALIDATOR_LOG_FILENAME As String = "log.process-cancelled-by-validator.txt"


